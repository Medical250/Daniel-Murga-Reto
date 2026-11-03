/**
 * Reto "30 días, 100 familias" — Dr. Daniel Murga
 * Backend en Google Apps Script. Recibe inscripciones y evaluaciones
 * desde reto.drmurga.com y las guarda en esta hoja de cálculo.
 *
 * Pestañas que crea setup():
 *   Hogares  → una fila por hogar (contacto único). Fuente del contador.
 *   Registro → bitácora de todo lo que llega, sin modificar.
 *   Descargas→ una fila por descarga de la guía (anónima).
 *   Resumen  → indicadores calculados.
 *
 * Columna "excluir" en Hogares: escribe "si" para sacar un hogar del
 * contador público (por ejemplo, un duplicado detectado a mano).
 */

const HOJA_HOGARES = 'Hogares';
const HOJA_LOG = 'Registro';
const HOJA_RESUMEN = 'Resumen';
const COLS = [
  'id_hogar', 'contacto', 'nombre', 'fecha_inscripcion', 'origen',
  'duplicado_declarado', 'permiso_novedades', 'puntuacion_inicial',
  'puntuacion_final', 'compartio', 'fecha_completado', 'estado',
  'intentos', 'baja', 'excluir'
];
const HOJA_DESCARGAS = 'Descargas';
const col = function (k) { return COLS.indexOf(k); };

/* ---------- Ejecutar UNA vez desde el editor ---------- */
function setup() {
  const ss = SpreadsheetApp.getActive();

  const h = ss.getSheetByName(HOJA_HOGARES) || ss.insertSheet(HOJA_HOGARES);
  h.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
  h.setFrozenRows(1);
  h.getRange('B:B').setNumberFormat('@');

  const l = ss.getSheetByName(HOJA_LOG) || ss.insertSheet(HOJA_LOG);
  l.getRange(1, 1, 1, 4).setValues([['fecha', 'tipo', 'contacto', 'datos']]).setFontWeight('bold');
  l.setFrozenRows(1);
  l.getRange('C:C').setNumberFormat('@');

  const dsc = ss.getSheetByName(HOJA_DESCARGAS) || ss.insertSheet(HOJA_DESCARGAS);
  dsc.getRange(1, 1, 1, 3).setValues([['fecha', 'origen', 'boton']]).setFontWeight('bold');
  dsc.setFrozenRows(1);

  const r = ss.getSheetByName(HOJA_RESUMEN) || ss.insertSheet(HOJA_RESUMEN);
  r.clear();
  r.getRange('A1:B11').setValues([
    ['Indicador', 'Valor'],
    ['Completaron el examen (5/5 + confirmación)', '=COUNTIFS(Hogares!L2:L,"completado",Hogares!F2:F,"<>si",Hogares!O2:O,"<>si")'],
    ['Inscritos únicos', '=COUNTA(Hogares!B2:B)'],
    ['Pendientes de completar', '=B3-B2'],
    ['Tasa de finalización', '=IF(B3=0,0,B2/B3)'],
    ['Duplicados declarados (revisar a mano)', '=COUNTIF(Hogares!F2:F,"si")'],
    ['Promedio primer intento (de 5)', '=IFERROR(AVERAGE(Hogares!H2:H),"")'],
    ['Promedio último intento (de 5)', '=IFERROR(AVERAGE(Hogares!I2:I),"")'],
    ['Bajas solicitadas', '=COUNTIF(Hogares!N2:N,"si")'],
    ['Descargas de la guía', '=COUNTA(Descargas!A2:A)'],
    ['CONTADOR PÚBLICO: familias participando (inscripción o examen)', '=COUNTIFS(Hogares!B2:B,"<>",Hogares!F2:F,"<>si",Hogares!O2:O,"<>si")']
  ]);
  r.getRange('A11:B11').setFontWeight('bold');
  r.getRange('A1:B1').setFontWeight('bold');
  r.getRange('B5').setNumberFormat('0%');
  r.getRange('B7:B8').setNumberFormat('0.0');
  r.setColumnWidth(1, 300);

  const hoja1 = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (hoja1 && ss.getSheets().length > 1) ss.deleteSheet(hoja1);
}

/* ---------- Lectura pública: solo cifras agrupadas ---------- */
// /exec                  → prueba de conexión
// /exec?accion=contador  → { completados, en_camino } para la cuadrícula
function doGet(e) {
  if (e && e.parameter && e.parameter.accion === 'contador') {
    const cache = CacheService.getScriptCache();
    const hit = cache.get('contador');
    if (hit) return out(JSON.parse(hit));
    const res = contar();
    cache.put('contador', JSON.stringify(res), 60); // se recalcula como máximo cada minuto
    return out(res);
  }
  return out({ ok: true, servicio: 'reto-100-familias' });
}

function contar() {
  const sh = SpreadsheetApp.getActive().getSheetByName(HOJA_HOGARES);
  let completados = 0, enCamino = 0;
  const n = sh.getLastRow() - 1;
  if (n > 0) {
    sh.getRange(2, 1, n, COLS.length).getValues().forEach(function (r) {
      if (r[col('excluir')] === 'si' || r[col('duplicado_declarado')] === 'si') return;
      if (r[col('estado')] === 'completado') completados++;
      else enCamino++;
    });
  }
  return { ok: true, completados: completados, en_camino: enCamino, actualizado: new Date().toISOString() };
}

/* ---------- Recepción de datos desde la landing ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    if (d.website) return out({ ok: true });              // trampa anti-bots
    const tipo = d.tipo;

    if (tipo === 'descarga') {                             // anónima: sin contacto
      SpreadsheetApp.getActive().getSheetByName(HOJA_DESCARGAS)
        .appendRow([new Date(), limpiar(d.origen, 40), limpiar(d.boton, 30)]);
      return out({ ok: true });
    }

    if (tipo !== 'inscripcion' && tipo !== 'evaluacion') return out({ ok: false, error: 'tipo' });
    const contacto = normalizar(d.contacto);
    if (!contacto) return out({ ok: false, error: 'contacto' });

    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(HOJA_HOGARES);
    const ahora = new Date();

    // bitácora sin datos sensibles adicionales
    const copia = Object.assign({}, d); delete copia.contacto; delete copia.website;
    ss.getSheetByName(HOJA_LOG).appendRow([ahora, tipo, contacto, limpiar(JSON.stringify(copia), 500)]);

    const fila = buscar(sh, contacto);
    const row = fila ? sh.getRange(fila, 1, 1, COLS.length).getValues()[0] : COLS.map(function () { return ''; });
    const set = function (k, v) { row[COLS.indexOf(k)] = v; };
    const get = function (k) { return row[COLS.indexOf(k)]; };

    if (!fila) {
      set('id_hogar', 'H' + ('0000' + sh.getLastRow()).slice(-4));
      set('contacto', contacto);
      set('intentos', 0);
      set('estado', tipo === 'inscripcion' ? 'inscrito' : 'sin inscripción');
    }

    if (tipo === 'inscripcion') {
      set('nombre', limpiar(d.nombre, 60));
      if (!get('fecha_inscripcion')) set('fecha_inscripcion', ahora);
      if (!get('origen')) set('origen', limpiar(d.origen, 40));
      set('duplicado_declarado', limpiar(d.duplicado_declarado, 5));
      set('permiso_novedades', d.permiso_novedades === true ? 'si' : 'no');
      if (get('estado') === 'sin inscripción') set('estado', 'en progreso');
    } else {
      const final = Number(d.puntuacion_final);
      if (!(final >= 0 && final <= 5)) return out({ ok: false, error: 'puntuacion' });
      // el primer intento se guarda una sola vez y nunca se sobrescribe
      if (get('puntuacion_inicial') === '') {
        const ini = Number(d.puntuacion_inicial);
        set('puntuacion_inicial', ini >= 0 && ini <= 5 ? ini : final);
      }
      set('puntuacion_final', final);
      set('intentos', Number(get('intentos') || 0) + 1);
      if (get('estado') !== 'completado') {
        set('compartio', d.compartio === true ? 'si' : 'no');
        if (final === 5 && d.compartio === true) {
          set('estado', 'completado');
          set('fecha_completado', ahora);
        } else if (get('estado') === 'inscrito') {
          set('estado', 'en progreso');
        }
      }
    }

    if (fila) sh.getRange(fila, 1, 1, COLS.length).setValues([row]);
    else sh.appendRow(row);
    CacheService.getScriptCache().remove('contador');      // la cuadrícula se actualiza al instante

    return out({ ok: true });
  } catch (err) {
    console.error(err);
    return out({ ok: false, error: 'servidor' });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- utilidades ---------- */
function normalizar(raw) {
  const v = String(raw || '').trim().toLowerCase();
  if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v)) return v.slice(0, 120);
  let dig = v.replace(/\D/g, '');
  if (dig.length === 11 && dig.indexOf('502') === 0) dig = dig.slice(3);
  return dig.length === 8 ? dig : null;
}

function limpiar(x, n) {
  // corta el largo y evita que un texto se interprete como fórmula
  return String(x == null ? '' : x).slice(0, n).replace(/^[=+\-@]/, "'");
}

function buscar(sh, contacto) {
  if (sh.getLastRow() < 2) return null;
  const hit = sh.getRange(2, 2, sh.getLastRow() - 1, 1)
    .createTextFinder(contacto).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : null;
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
