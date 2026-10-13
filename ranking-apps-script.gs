/**
 * HTML Quest — ranking y progreso compartido (Google Apps Script)
 *
 * Se pega en una planilla de Google (Extensiones > Apps Script) y se publica como
 * “Aplicación web”. El juego manda el progreso de cada alumno (POST), pide el ranking (GET ?action=list)
 * y trae el progreso completo de un alumno para seguir jugando desde otra PC (GET ?action=load&k=...).
 *
 * Para actualizar sin cambiar la URL: Implementar > Administrar implementaciones > ✏️ editar >
 * Versión: “Nueva versión” > Implementar.
 *
 * Privacidad: la planilla guarda el nombre completo (solo lo ve quien tenga acceso a la planilla),
 * pero lo que devuelve la web es únicamente “Nombre + inicial del apellido”.
 */
const HOJA = 'Progreso';
const COLS = ['id', 'k', 'nombre', 'curso', 'avatar', 'xp', 'nivel', 'misiones', 'estrellas', 'errores', 'actualizado', 'datos'];
const MAX_XP = 5000;          // tope razonable (el juego completo da ~1900 XP)
const MAX_DATOS = 40000;      // una celda admite 50.000 caracteres

function hoja_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(HOJA);
  if (!sh) {
    sh = ss.insertSheet(HOJA);
    sh.appendRow(COLS);
    sh.setFrozenRows(1);
  } else if (sh.getRange(1, COLS.length).getValue() !== COLS[COLS.length - 1]) {
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS]);   // planillas viejas: agrega la columna “datos”
  }
  return sh;
}

function texto_(v, max) { return String(v == null ? '' : v).slice(0, max); }
function num_(v, max) { const n = Math.floor(Number(v)); return isFinite(n) && n >= 0 ? Math.min(n, max) : 0; }

function filas_(sh) {
  const n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, COLS.length).getValues() : [];
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const d = JSON.parse(e.postData.contents);
    const k = texto_(d.k, 20);
    const curso = texto_(d.curso, 12).trim();
    if (!k || !curso) return salida_({ ok: false });

    let datos = '';
    if (d.p && typeof d.p === 'object') {
      datos = JSON.stringify(d.p);
      if (datos.length > MAX_DATOS) datos = '';
    }

    const sh = hoja_();
    const filas = filas_(sh);
    const i = filas.findIndex(function (f) { return String(f[1]) === k; });
    const vieja = i >= 0 ? filas[i] : null;

    // Los jugadores recuperados desde otra PC no mandan el nombre completo: se conserva el de la planilla.
    const id = texto_(d.id, 80) || (vieja ? String(vieja[0]) : '');
    const nombre = texto_(d.name, 40).trim() || (vieja ? String(vieja[2]) : '');
    if (!id || nombre.length < 2) return salida_({ ok: false });

    const fila = [
      id, k, nombre, curso, texto_(d.avatar, 8),
      num_(d.xp, MAX_XP), num_(d.lv, 99), num_(d.m, 100), num_(d.s, 300), num_(d.e, 100000), new Date(),
      datos || (vieja ? vieja[11] : '')
    ];

    if (vieja) {
      // El progreso nunca retrocede (por si el alumno juega en otra PC con menos avance).
      if (fila[5] >= (Number(vieja[5]) || 0)) sh.getRange(i + 2, 1, 1, fila.length).setValues([fila]);
    } else {
      sh.appendRow(fila);
    }
    return salida_({ ok: true });
  } catch (err) {
    return salida_({ ok: false });
  } finally {
    lock.releaseLock();
  }
}

function corto_(nombre) {
  const t = String(nombre).trim().split(/\s+/);
  return t.length < 2 ? t[0] : t[0] + ' ' + t[t.length - 1].charAt(0).toUpperCase() + '.';
}

function resumen_(f) {
  return { k: f[1], n: corto_(f[2]), c: f[3], a: f[4], xp: f[5], lv: f[6], m: f[7], s: f[8], e: f[9] };
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const filas = filas_(hoja_());

  if (p.action === 'load') {
    const f = filas.find(function (x) { return String(x[1]) === String(p.k || ''); });
    if (!f) return salida_({ ok: false });
    const r = resumen_(f);
    r.ok = true;
    try { r.d = f[11] ? JSON.parse(f[11]) : null; } catch (err) { r.d = null; }
    return salida_(r);
  }

  return salida_(filas.map(resumen_));
}

function salida_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
