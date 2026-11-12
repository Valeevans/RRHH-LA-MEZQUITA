/**
 * LA MEZQUITA — RRHH · Backend Google Sheets
 * Pegar este código en Extensiones → Apps Script de la Google Sheet "RRHH La Mezquita".
 * Implementar → Nueva implementación → Tipo: Aplicación web
 *   - Ejecutar como: Yo
 *   - Quién tiene acceso: Cualquier usuario
 * Copiar la URL que termina en /exec y pegarla en index.html (const SHEET_API_URL).
 *
 * Acciones:
 *   GET  ?action=meta                 → { years: ["2026", ...] }
 *   GET  ?action=read&sheet=NOMBRE    → { rows: [ {col: valor, ...}, ... ] }
 *   POST {action:"write", sheet, headers, rows} → reescribe la pestaña completa
 */

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || '';
    if (action === 'meta') {
      var years = {};
      ss_().getSheets().forEach(function (sh) {
        var m = sh.getName().match(/^(CAPACITACIONES|CAP_DETALLE|EVAL_DESEMPENO)_(\d{4})$/);
        if (m) years[m[2]] = true;
      });
      return json_({ years: Object.keys(years).sort() });
    }
    if (action === 'read') {
      var sh = ss_().getSheetByName(e.parameter.sheet);
      if (!sh || sh.getLastRow() < 1) return json_({ rows: [] });
      var values = sh.getDataRange().getDisplayValues();
      var headers = values.shift();
      var rows = values
        .filter(function (r) { return r.some(function (c) { return c !== ''; }); })
        .map(function (r) {
          var o = {};
          headers.forEach(function (h, i) { if (h) o[h] = r[i]; });
          return o;
        });
      return json_({ rows: rows });
    }
    return json_({ error: 'Acción no reconocida: ' + action });
  } catch (err) {
    return json_({ error: String(err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var body = JSON.parse(e.postData.contents);
    if (body.action !== 'write') return json_({ error: 'Acción no reconocida' });
    if (!/^[A-Z_0-9]+$/.test(body.sheet)) return json_({ error: 'Nombre de pestaña inválido' });
    var headers = body.headers || [];
    var rows = body.rows || [];
    var ss = ss_();
    var sh = ss.getSheetByName(body.sheet) || ss.insertSheet(body.sheet);
    sh.clearContents();
    var data = [headers].concat(rows.map(function (r) {
      return headers.map(function (h) { return r[h] == null ? '' : String(r[h]); });
    }));
    var range = sh.getRange(1, 1, data.length, headers.length);
    range.setNumberFormat('@');            // texto plano: las fechas no se convierten
    range.setValues(data);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#E2EFDA').setFontColor('#2F5233');
    sh.setFrozenRows(1);
    return json_({ ok: true, rows: rows.length });
  } catch (err) {
    return json_({ error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}
