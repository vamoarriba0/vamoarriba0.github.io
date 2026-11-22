/***************************************************************
 *  CATÁLOGO VAMO ARRIBA — Código.gs
 *  Este script es el "motor": lee el Drive, guarda la planilla y
 *  le pasa los datos a la página publicada en GitHub.
 *  Pasos (solo la primera vez):
 *   1) Servicios (+) → agregar "Drive API" (v3)
 *   2) Ejecutar configurar()          → crea carpetas y planilla
 *   3) Ejecutar activarAutomaticos()  → deja todo andando solo
 *   4) Implementar → App web (Ejecutar como: yo / Acceso: cualquiera)
 ***************************************************************/

const CONFIG = {
  IG_USUARIO: 'vamoarriba0',
  NOMBRE_CARPETA: 'Catálogo Vamo Arriba',
  LOGO_ID: '',                                // ID del PNG del sol en Drive (vacío = sol dibujado)
  CATEGORIAS: ['Mujer', 'Niño', 'Hombre', 'Free Shop'],  // orden en que aparecen los filtros
  SIN_TALLE: ['Free Shop'],                   // categorías donde no se pide talle
  CLAVE_ADMIN: 'cambiame123',                 // clave del link de "Lista de compras" (cambiala)
  DIAS_LISTA: 3,                              // días que mira la lista de compras por defecto
  MIN_PEDIDOS_FUEGO: 2,                       // desde cuántos pedidos del día aparece "🔥 Pedida X veces hoy"
  DIAS_PARA_BORRAR: 30,                       // limpiarViejos() manda a la papelera fotos más viejas
  PRECIO_MAX: 100000,                         // si lee un precio mayor, lo deja vacío para revisar a mano
  MAX_MINUTOS: 4.5                            // corte de seguridad (Google corta a los 6 min)
};

const HOJA_PRENDAS = 'Prendas';
const HOJA_INTERES = 'Interesados';
const HOJA_CONFIG = 'Config';
const COLS = ['Código', 'Categoría', 'Precio', 'Estado', 'Clicks', 'Alta', 'Archivo', 'FileId', 'Tipo', 'Comprada'];
const ESTADOS = ['Disponible', 'Vendido', 'Oculto', 'Borrado'];
const TZ = 'America/Argentina/Buenos_Aires';


/* ============================================================
 *  1. CONFIGURACIÓN INICIAL
 * ============================================================ */
function configurar() {
  const props = PropertiesService.getScriptProperties();

  // Carpeta raíz + subcarpetas por categoría
  let raiz;
  const idRaiz = props.getProperty('CARPETA_ID');
  if (idRaiz) {
    raiz = DriveApp.getFolderById(idRaiz);
  } else {
    const it = DriveApp.getFoldersByName(CONFIG.NOMBRE_CARPETA);
    raiz = it.hasNext() ? it.next() : DriveApp.createFolder(CONFIG.NOMBRE_CARPETA);
    props.setProperty('CARPETA_ID', raiz.getId());
  }
  CONFIG.CATEGORIAS.forEach(c => {
    if (!raiz.getFoldersByName(c).hasNext()) raiz.createFolder(c);
  });

  // Planilla de datos
  let ss;
  const idSS = props.getProperty('PLANILLA_ID');
  if (idSS) {
    ss = SpreadsheetApp.openById(idSS);
  } else {
    ss = SpreadsheetApp.create('Catálogo Vamo Arriba - Datos');
    DriveApp.getFileById(ss.getId()).moveTo(raiz);
    props.setProperty('PLANILLA_ID', ss.getId());
  }

  let sh = ss.getSheetByName(HOJA_PRENDAS);
  if (!sh) { sh = ss.getSheets()[0]; sh.setName(HOJA_PRENDAS); }
  sh.getRange(1, 1, 1, COLS.length).setValues([COLS])
    .setFontWeight('bold').setBackground('#1b3f8b').setFontColor('#ffffff');
  if (sh.getFrozenRows() === 0) {
    sh.setFrozenRows(1);
    sh.getRange('C:C').setNumberFormat('$#,##0');
    sh.getRange('F:F').setNumberFormat('dd/MM/yyyy HH:mm');
  }

  let si = ss.getSheetByName(HOJA_INTERES);
  if (!si) { si = ss.insertSheet(HOJA_INTERES); si.setFrozenRows(1); }
  const COLS_INT = ['Fecha', 'Código', 'Talle', 'Color', 'Usuario IG'];
  si.getRange(1, 1, 1, COLS_INT.length).setValues([COLS_INT])
    .setFontWeight('bold').setBackground('#1b3f8b').setFontColor('#ffffff');
  si.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm');

  // Config: hora de cierre de pedidos (se edita desde la app de Sheets en el celu)
  let sc = ss.getSheetByName(HOJA_CONFIG);
  if (!sc) sc = ss.insertSheet(HOJA_CONFIG);
  sc.getRange('A1').setValue('Cierre de pedidos').setFontWeight('bold');
  sc.getRange('A2').setValue('Próximo viaje').setFontWeight('bold');
  if (!(sc.getRange('B2').getValue() instanceof Date)) sc.getRange('B2').clearContent();
  sc.getRange('B2').setNumberFormat('dd/MM/yyyy');
  sc.getRange('A4').setValue('B1: hora de cierre (ej: 18:00) para hoy, o fecha y hora. ' +
    'B2: fecha del próximo viaje (ej: 28/11/2026). Vacío = no se muestra.')
    .setFontColor('#7d8798').setFontWeight('normal');
  sc.setColumnWidth(1, 160);

  Logger.log('✅ Carpeta de fotos: ' + raiz.getUrl());
  Logger.log('✅ Planilla:         ' + ss.getUrl());
}


/* ============================================================
 *  2. AUTOMÁTICOS (procesar cada 15 min, resumen y limpieza diarios)
 * ============================================================ */
function activarAutomaticos() {
  ['procesarFotos', 'enviarResumen', 'limpiarViejos', 'continuarProcesamiento']
    .forEach(borrarDisparadores_);

  ScriptApp.newTrigger('procesarFotos').timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger('enviarResumen').timeBased().atHour(22).everyDays(1).inTimezone(TZ).create();
  ScriptApp.newTrigger('limpiarViejos').timeBased().atHour(4).everyDays(1).inTimezone(TZ).create();
  Logger.log('✅ Automáticos activados');
}


/* ============================================================
 *  3. PROCESAR FOTOS NUEVAS (código + precio por OCR)
 * ============================================================ */
function procesarFotos() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;          // ya hay otra corrida en curso
  try {
    const inicio = Date.now();
    const raiz = carpeta_();
    const sh = hoja_(HOJA_PRENDAS);

    const datos = sh.getLastRow() > 1
      ? sh.getRange(2, 1, sh.getLastRow() - 1, COLS.length).getValues() : [];
    const yaCargados = new Set(datos.map(r => r[7]));

    // Último número usado por prefijo (M, N, H…)
    const contadores = {};
    datos.forEach(r => {
      const m = String(r[0]).match(/^([A-Z]+)(\d+)$/);
      if (m) contadores[m[1]] = Math.max(contadores[m[1]] || 0, +m[2]);
    });

    // Juntar fotos nuevas de todas las subcarpetas (solo las que no están en la papelera)
    const pendientes = [];
    const presentes = new Set();
    const carpetas = raiz.getFolders();
    while (carpetas.hasNext()) {
      const carpeta = carpetas.next();
      const cat = carpeta.getName();
      const files = carpeta.searchFiles('trashed = false');
      while (files.hasNext()) {
        const f = files.next();
        presentes.add(f.getId());
        const mime = f.getMimeType();
        const tipo = mime.startsWith('image/') ? 'foto' : mime.startsWith('video/') ? 'video' : null;
        if (!tipo || yaCargados.has(f.getId())) continue;
        pendientes.push({ f, cat, tipo, fecha: f.getDateCreated().getTime(), nombre: f.getName() });
      }
    }
    // Si borraron una foto del Drive (o la sacaron de la carpeta), la saca del catálogo
    datos.forEach((r, i) => {
      if (r[3] === 'Disponible' && r[7] && !presentes.has(String(r[7]))) {
        sh.getRange(i + 2, 4).setValue('Borrado');
      }
    });

    // Respeta el orden en que las subieron
    pendientes.sort((a, b) => a.fecha - b.fecha || a.nombre.localeCompare(b.nombre));

    let hechos = 0, cortado = false;
    const filaInicial = sh.getLastRow() + 1;

    for (const p of pendientes) {
      if ((Date.now() - inicio) / 60000 > CONFIG.MAX_MINUTOS) { cortado = true; break; }

      const pref = prefijo_(p.cat);
      contadores[pref] = (contadores[pref] || 0) + 1;
      const codigo = pref + String(contadores[pref]).padStart(3, '0');

      try { p.f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
      const precio = p.tipo === 'foto' ? leerPrecio_(p.f) : '';   // en videos el precio se carga a mano

      sh.appendRow([codigo, p.cat, precio, 'Disponible', 0, new Date(), p.nombre, p.f.getId(), p.tipo]);
      hechos++;
    }

    if (hechos) {
      const regla = SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS, true).build();
      sh.getRange(filaInicial, 4, hechos, 1).setDataValidation(regla);
    }

    borrarDisparadores_('continuarProcesamiento');
    if (cortado) {
      ScriptApp.newTrigger('continuarProcesamiento').timeBased().after(60 * 1000).create();
    }
    Logger.log(`Procesadas: ${hechos} | Pendientes: ${pendientes.length - hechos}`);
  } finally {
    lock.releaseLock();
  }
}

// Lo llama el disparador cuando una tanda grande no entra en una sola corrida
function continuarProcesamiento() { procesarFotos(); }


/* ============================================================
 *  4. APP WEB
 * ============================================================ */
function doGet(e) {
  const prm = (e && e.parameter) || {};

  // API para la página de GitHub
  if (prm.api) return json_(api_(prm));

  // Registro de pedidos (lo manda la página en segundo plano)
  if (prm.reg) {
    try {
      const d = JSON.parse(prm.reg);
      registrarPedido((d.i || []).map(x => ({ c: x.c, talle: x.t, color: x.col })), d.u);
      if (d.u) notificarPedido_(d);
      return json_({ ok: true });
    } catch (err) {
      Logger.log('reg: ' + err.message);
      return json_({ ok: false, error: err.message });
    }
  }

  // Guía de talles: link?guia=1
  if (prm.guia) {
    const tg = HtmlService.createTemplateFromFile('GuiaTalles');
    tg.url = ScriptApp.getService().getUrl();
    return tg.evaluate()
      .setTitle('Guía de talles · Vamo Arriba')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  const t = HtmlService.createTemplateFromFile('Index');
  const foco = String(prm.p || '').toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 10);
  const esAdmin = !!prm.admin && prm.admin === CONFIG.CLAVE_ADMIN;
  const dias = Math.min(30, Math.max(1, parseInt(prm.dias, 10) || CONFIG.DIAS_LISTA));

  t.ig = CONFIG.IG_USUARIO;                 // para los links del HTML
  t.url = ScriptApp.getService().getUrl();
  t.logo = logoDataUri_();
  t.cfg = JSON.stringify({                  // todo lo que usa el JavaScript, en un solo bloque
    ig: CONFIG.IG_USUARIO,
    url: ScriptApp.getService().getUrl(),
    pedido: String(prm.pedido || '').toUpperCase().replace(/[^A-Z0-9,]/g, '').slice(0, 300),
    admin: esAdmin ? obtenerListaCompras_(dias) : null,
    clave: esAdmin ? CONFIG.CLAVE_ADMIN : '',
    datos: esAdmin ? { items: [], categorias: [] } : obtenerCatalogo_(),
    foco: foco,                               // ?p=M045 → abre esa prenda (link de "Compartir")
    cierre: esAdmin ? null : leerCierre_()
  }).replace(/</g, '\\u003c');
  return t.evaluate()
    .setTitle('Vamo Arriba | Catálogo')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

/* ============================================================
 *  API (la usa la página publicada en GitHub)
 * ============================================================ */
function api_(prm) {
  try {
    switch (prm.api) {
      case 'catalogo':
        return Object.assign(obtenerCatalogo_(), {
          ig: CONFIG.IG_USUARIO,
          cierre: leerCierre_(),
          viaje: leerViaje_(),
          sinTalle: CONFIG.SIN_TALLE
        });
      case 'login':
        return { ok: prm.clave === CONFIG.CLAVE_ADMIN, categorias: CONFIG.CATEGORIAS };
      case 'lista': {
        if (prm.clave !== CONFIG.CLAVE_ADMIN) return { error: 'clave' };
        const dias = Math.min(30, Math.max(1, parseInt(prm.dias, 10) || CONFIG.DIAS_LISTA));
        return obtenerListaCompras_(dias);
      }
      case 'marcar':
        marcarComprada(prm.clave, prm.c, prm.v === '1');
        return { ok: true };
      default:
        return { error: 'api' };
    }
  } catch (err) {
    return { error: err.message };
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Subida de fotos desde la página de administración (una foto por pedido)
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.accion === 'subir') return json_(subirFoto_(d));
    return json_({ error: 'accion' });
  } catch (err) {
    return json_({ error: err.message });
  }
}

function subirFoto_(d) {
  if (d.clave !== CONFIG.CLAVE_ADMIN) return { error: 'clave' };
  const cat = String(d.cat || '');
  if (!CONFIG.CATEGORIAS.includes(cat)) return { error: 'categoría' };
  const precio = parseInt(String(d.precio || '').replace(/\D/g, ''), 10);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);                       // espera si justo está corriendo procesarFotos
  try {
    const raiz = carpeta_();
    const it = raiz.getFoldersByName(cat);
    const carpeta = it.hasNext() ? it.next() : raiz.createFolder(cat);
    const nombre = 'app_' + Date.now() + '.jpg';
    const blob = Utilities.newBlob(Utilities.base64Decode(d.data), 'image/jpeg', nombre);
    const f = carpeta.createFile(blob);
    try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}

    const sh = hoja_(HOJA_PRENDAS);
    const pref = prefijo_(cat);
    let max = 0;
    if (sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(r => {
        const m = String(r[0]).match(/^([A-Z]+)(\d+)$/);
        if (m && m[1] === pref) max = Math.max(max, +m[2]);
      });
    }
    const codigo = pref + String(max + 1).padStart(3, '0');
    sh.appendRow([codigo, cat, precio > 0 ? precio : '', 'Disponible', 0, new Date(), nombre, f.getId(), 'foto']);
    const regla = SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS, true).build();
    sh.getRange(sh.getLastRow(), 4).setDataValidation(regla);
    return { ok: true, codigo };
  } finally {
    lock.releaseLock();
  }
}

// Mail al instante con cada pedido nuevo (para escribirle a la clienta por DM)
function notificarPedido_(d) {
  try {
    const usuario = String(d.u || '').replace(/^@+/, '').replace(/[^A-Za-z0-9._]/g, '').slice(0, 30);
    const items = (d.i || []).slice(0, 30);
    const prendas = {};
    const sh = hoja_(HOJA_PRENDAS);
    if (sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, COLS.length).getValues()
        .forEach(r => prendas[String(r[0])] = { id: String(r[7]), p: r[2], cat: r[1] });
    }
    const esc = s => String(s || '').replace(/[<>&"]/g, '');
    const filas = items.map((x, n) => {
      const pr = prendas[x.c] || {};
      const det = [x.t && 'Talle ' + esc(x.t), esc(x.col)].filter(Boolean).join(' · ') || '—';
      return `<tr><td style="padding:6px"><img src="https://lh3.googleusercontent.com/d/${pr.id}=w200" width="70" style="border-radius:8px"></td>` +
        `<td style="padding:6px;font-family:Arial;font-size:14px"><b>${n + 1}. ${esc(x.c)}</b><br>${det}` +
        (typeof pr.p === 'number' ? `<br>$${pr.p.toLocaleString('es-AR')}` : '') + `</td></tr>`;
    }).join('');
    const html = `<div style="font-family:Arial;color:#123e75">
      <h2 style="margin:0 0 6px">🛍 Pedido nuevo de @${usuario}</h2>
      <p style="margin:0 0 12px">${items.length} prenda${items.length > 1 ? 's' : ''}</p>
      <p><a href="https://ig.me/m/${usuario}" style="background:#1d5aa6;color:#fff;padding:12px 18px;border-radius:10px;text-decoration:none;font-weight:bold">💬 Escribirle por Instagram</a></p>
      <table>${filas}</table>
      ${d.l ? `<p><a href="${esc(d.l)}">Ver las fotos del pedido</a></p>` : ''}
    </div>`;
    MailApp.sendEmail({
      to: Session.getEffectiveUser().getEmail(),
      subject: `🛍 Pedido nuevo de @${usuario} (${items.length})`,
      htmlBody: html
    });
  } catch (e) {
    Logger.log('notificarPedido: ' + e.message);
  }
}

// Fecha del próximo viaje (hoja Config, B2). Devuelve milisegundos o null.
function leerViaje_() {
  try {
    const v = hoja_(HOJA_CONFIG).getRange('B2').getValue();
    if (!(v instanceof Date)) return null;
    const hoy = Utilities.parseDate(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), TZ, 'yyyy-MM-dd');
    return v.getTime() >= hoy.getTime() ? v.getTime() : null;
  } catch (e) { return null; }
}

function obtenerCatalogo_() {
  const sh = hoja_(HOJA_PRENDAS);
  if (sh.getLastRow() < 2) return { items: [], categorias: [] };

  const filas = sh.getRange(2, 1, sh.getLastRow() - 1, COLS.length).getValues();
  const items = filas
    .filter(r => r[3] === 'Disponible' && r[7])
    .map(r => ({
      c: String(r[0]),
      cat: String(r[1]),
      p: typeof r[2] === 'number' && r[2] > 0 ? r[2] : null,
      d: r[5] instanceof Date ? Utilities.formatDate(r[5], TZ, 'yyyyMMdd') : '0',
      t: r[5] instanceof Date ? r[5].getTime() : 0,
      id: String(r[7]),
      v: r[8] === 'video'
    }));

  // Día más nuevo primero; dentro del día, en el orden en que se subieron
  items.sort((a, b) => b.d.localeCompare(a.d) || a.t - b.t);

  const presentes = [...new Set(items.map(i => i.cat))];
  const categorias = CONFIG.CATEGORIAS.filter(c => presentes.includes(c))
    .concat(presentes.filter(c => !CONFIG.CATEGORIAS.includes(c)));

  // Pedidos de hoy por prenda (para la etiqueta 🔥)
  const pedidosHoy = contarPedidosHoy_();
  items.forEach(i => { const n = pedidosHoy[i.c] || 0; if (n >= CONFIG.MIN_PEDIDOS_FUEGO) i.h = n; });

  const ultimoDia = items.length ? items[0].d : '';
  items.forEach(i => { i.n = i.d === ultimoDia; delete i.t; delete i.d; });

  return { items, categorias };
}

// Hora de cierre de pedidos (hoja Config, celda B1). Devuelve milisegundos o null.
function leerCierre_() {
  try {
    const sc = hoja_(HOJA_CONFIG);
    if (!sc) return null;
    const v = sc.getRange('B1').getValue();
    if (!(v instanceof Date)) return null;
    let d = v;
    if (v.getFullYear() < 1950) {               // solo hora → hoy a esa hora
      const hoy = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
      const hm = Utilities.formatDate(v, TZ, 'HH:mm');
      d = Utilities.parseDate(hoy + ' ' + hm, TZ, 'yyyy-MM-dd HH:mm');
    }
    const ms = d.getTime();
    return (ms > Date.now() - 12 * 3600 * 1000) ? ms : null;   // si ya pasó hace mucho, no se muestra
  } catch (e) { return null; }
}

function contarPedidosHoy_() {
  const si = hoja_(HOJA_INTERES);
  const conteo = {};
  if (!si || si.getLastRow() < 2) return conteo;
  const hoy = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd');
  const n = Math.min(si.getLastRow() - 1, 2000);              // alcanza con las últimas filas
  si.getRange(si.getLastRow() - n + 1, 1, n, 2).getValues().forEach(r => {
    if (r[0] instanceof Date && Utilities.formatDate(r[0], TZ, 'yyyyMMdd') === hoy) {
      conteo[r[1]] = (conteo[r[1]] || 0) + 1;
    }
  });
  return conteo;
}

// Lista de compras: prendas pedidas en los últimos N días, agrupadas, con talles/colores/usuarios
function obtenerListaCompras_(dias) {
  const sh = hoja_(HOJA_PRENDAS);
  const si = hoja_(HOJA_INTERES);
  if (sh.getLastRow() < 2 || si.getLastRow() < 2) return { lista: [], dias };

  const prendas = {};
  sh.getRange(2, 1, sh.getLastRow() - 1, COLS.length).getValues().forEach(r => {
    prendas[String(r[0])] = {
      id: String(r[7]), v: r[8] === 'video', cat: String(r[1]),
      p: typeof r[2] === 'number' && r[2] > 0 ? r[2] : null,
      comprada: r[9] === 'Sí', estado: String(r[3])
    };
  });

  const desde = Date.now() - dias * 86400000;
  const grupos = {};
  si.getRange(2, 1, si.getLastRow() - 1, 5).getValues().forEach(r => {
    if (!(r[0] instanceof Date) || r[0].getTime() < desde) return;
    const c = String(r[1]);
    if (!prendas[c] || !prendas[c].id) return;
    if (!grupos[c]) grupos[c] = Object.assign({ c, pedidos: [] }, prendas[c]);
    grupos[c].pedidos.push({ t: String(r[2]), col: String(r[3]), u: String(r[4]) });
  });

  const lista = Object.values(grupos)
    .sort((a, b) => (a.comprada - b.comprada) || (b.pedidos.length - a.pedidos.length));
  return { lista, dias };
}

// Tilde "Comprada" desde la lista de compras
function marcarComprada(clave, codigo, valor) {
  if (clave !== CONFIG.CLAVE_ADMIN) throw new Error('Clave incorrecta');
  const sh = hoja_(HOJA_PRENDAS);
  const codigos = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat().map(String);
  const i = codigos.indexOf(String(codigo));
  if (i >= 0) sh.getRange(i + 2, 10).setValue(valor ? 'Sí' : '');
}

// Lo llama la página cuando el cliente envía su pedido (una fila por prenda)
function registrarPedido(items, usuario) {
  if (!Array.isArray(items)) return;
  items.slice(0, 30).forEach(i => registrarInteres(i.c, i.talle, i.color, usuario));
}

function registrarInteres(codigo, talle, color, usuario) {
  codigo = limpiarTexto_(codigo, 10);
  talle = limpiarTexto_(talle, 20);
  color = limpiarTexto_(color, 30);
  usuario = limpiarTexto_(String(usuario || '').replace(/^@+/, ''), 40);
  if (usuario) usuario = '@' + usuario;
  if (!codigo) return;

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    hoja_(HOJA_INTERES).appendRow([new Date(), codigo, talle, color, usuario]);

    const sh = hoja_(HOJA_PRENDAS);
    if (sh.getLastRow() < 2) return;
    const codigos = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat();
    const i = codigos.indexOf(codigo);
    if (i >= 0) {
      const cel = sh.getRange(i + 2, 5);
      cel.setValue((+cel.getValue() || 0) + 1);
    }
  } catch (e) {
    Logger.log('registrarInteres: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}


/* ============================================================
 *  5. RESUMEN DIARIO Y LIMPIEZA
 * ============================================================ */
function enviarResumen() {
  const si = hoja_(HOJA_INTERES);
  if (si.getLastRow() < 2) return;

  const desde = Date.now() - 24 * 3600 * 1000;
  const filas = si.getRange(2, 1, si.getLastRow() - 1, 5).getValues()
    .filter(r => r[0] instanceof Date && r[0].getTime() >= desde);
  if (!filas.length) return;

  const conteo = {}, quienes = {};
  filas.forEach(r => {
    conteo[r[1]] = (conteo[r[1]] || 0) + 1;
    if (r[4]) (quienes[r[1]] = quienes[r[1]] || new Set()).add(r[4]);
  });
  const ranking = Object.entries(conteo).sort((a, b) => b[1] - a[1]);

  const cuerpo = `Hoy hubo ${filas.length} toques en "Lo quiero" sobre ${ranking.length} prendas.\n\n` +
    'Las más pedidas:\n' +
    ranking.slice(0, 20).map(([c, n]) => `  ${c}: ${n}` + (quienes[c] ? `  (${[...quienes[c]].join(', ')})` : '')).join('\n') +
    `\n\nPlanilla: ${SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('PLANILLA_ID')).getUrl()}`;

  MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'Vamo Arriba – interesados de hoy', cuerpo);
}

function limpiarViejos() {
  const sh = hoja_(HOJA_PRENDAS);
  if (sh.getLastRow() < 2) return;

  const limite = Date.now() - CONFIG.DIAS_PARA_BORRAR * 86400000;
  const filas = sh.getRange(2, 1, sh.getLastRow() - 1, COLS.length).getValues();
  let borradas = 0;

  filas.forEach((r, i) => {
    if (r[5] instanceof Date && r[5].getTime() < limite && r[3] !== 'Borrado') {
      try { DriveApp.getFileById(r[7]).setTrashed(true); } catch (e) {}
      sh.getRange(i + 2, 4).setValue('Borrado');
      borradas++;
    }
  });
  Logger.log('Fotos enviadas a la papelera: ' + borradas);
}


/* ============================================================
 *  AUXILIARES
 * ============================================================ */
function leerPrecio_(file) {
  let docId = null;
  try {
    const doc = Drive.Files.create(
      { name: 'ocr_tmp_' + file.getId(), mimeType: MimeType.GOOGLE_DOCS },
      file.getBlob(),
      { ocrLanguage: 'es', fields: 'id' }
    );
    docId = doc.id;
    const texto = DocumentApp.openById(docId).getBody().getText();
    return extraerPrecio_(texto);
  } catch (e) {
    Logger.log('OCR falló en ' + file.getName() + ': ' + e.message);
    return '';
  } finally {
    if (docId) { try { DriveApp.getFileById(docId).setTrashed(true); } catch (e) {} }
  }
}

// Devuelve el logo como imagen embebida (no hace falta compartir el archivo)
function logoDataUri_() {
  if (!CONFIG.LOGO_ID) return '';
  try {
    const blob = DriveApp.getFileById(CONFIG.LOGO_ID).getBlob();
    return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (e) {
    Logger.log('Logo: ' + e.message);
    return '';
  }
}

function extraerPrecio_(texto) {
  const t = String(texto).replace(/\s+/g, ' ');
  const conSigno = t.match(/\$\s?(\d{1,3}(?:[.,]\d{3})+|\d{3,7})/);
  const sinSigno = t.match(/\b(\d{1,3}(?:[.,]\d{3})+|\d{4,7})\b/);
  const m = conSigno || sinSigno;
  if (!m) return '';
  const n = parseInt(m[1].replace(/[.,]/g, ''), 10);
  return (n >= 500 && n <= CONFIG.PRECIO_MAX) ? n : '';
}

function prefijo_(cat) {
  const limpio = cat.normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toUpperCase().replace(/[^A-Z]/g, '');
  return limpio.charAt(0) || 'X';
}

function limpiarTexto_(s, max) {
  s = String(s || '').trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;   // evita fórmulas inyectadas en la planilla
}

function carpeta_() {
  const id = PropertiesService.getScriptProperties().getProperty('CARPETA_ID');
  if (!id) throw new Error('Primero ejecutá configurar()');
  return DriveApp.getFolderById(id);
}

function hoja_(nombre) {
  const id = PropertiesService.getScriptProperties().getProperty('PLANILLA_ID');
  if (!id) throw new Error('Primero ejecutá configurar()');
  return SpreadsheetApp.openById(id).getSheetByName(nombre);
}

function borrarDisparadores_(funcion) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === funcion)
    .forEach(t => ScriptApp.deleteTrigger(t));
}
