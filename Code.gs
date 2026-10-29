/**
 * Google Apps Script backend – Metal printing cost estimator
 * Storage: the Google Sheet this script is bound to (Extensions > Apps Script).
 * Tabs created automatically: Quotes, Powders, Machines
 *
 * Deploy: Deploy > New deployment > Web app
 *   Execute as: Me      Who has access: Anyone
 * Copy the /exec URL into API_URL at the top of script.js.
 */
const QUOTE_HEADERS = ['ID', 'Saved', 'Customer', 'Part', 'Machine', 'Material', 'Laser', 'Parts/Build', 'Layer (um)', 'Print hrs',
  'Material kg', 'Sets', 'Units', 'Total cost', 'Round-off', 'Per part', 'State JSON'];
const POWDER_HEADERS = ['Powder', 'Price / kg', 'Density g/cc', 'Wastage', 'HT time (hrs)', 'HT power (hrs)'];
const POWDER_KEYS = ['n', 'p', 'dens', 'waste', 'ht', 'fp'];
const MACHINE_HEADERS = ['Machine', 'Lasers', 'Price', 'Interest %', 'Years', 'Utilisation %', 'Fixed rate per hr'];
const MACHINE_KEYS = ['n', 'laser', 'price', 'rate', 'yrs', 'util', 'fixed'];

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(headers); sh.setFrozenRows(1); sh.getRange(1, 1, 1, headers.length).setFontWeight('bold'); }
  return sh;
}
function setup() { sheet_('Quotes', QUOTE_HEADERS); sheet_('Powders', POWDER_HEADERS); sheet_('Machines', MACHINE_HEADERS); }   // run once to authorise
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
const blank_ = v => v === null || v === undefined || (typeof v === 'number' && !isFinite(v)) ? '' : v;

function doGet(e) {
  try {
    const a = (e.parameter || {}).action;
    if (a === 'list') {
      const rows = sheet_('Quotes', QUOTE_HEADERS).getDataRange().getValues().slice(1);
      return json_({ ok: true, quotes: rows.filter(r => r[0]).map(r => ({ id: r[0], saved: r[1], total: r[13], round: r[14], perPart: r[15], state: JSON.parse(r[16]) })).reverse() });
    }
    if (a === 'masters') {   // empty lists -> the site uses its built-in defaults taken from the Excel
      const mats = sheet_('Powders', POWDER_HEADERS).getDataRange().getValues().slice(1).filter(r => r[0])
        .map(r => Object.fromEntries(POWDER_KEYS.map((k, i) => [k, i === 0 ? String(r[i]) : +r[i]])));
      const mach = sheet_('Machines', MACHINE_HEADERS).getDataRange().getValues().slice(1).filter(r => r[0])
        .map(r => Object.fromEntries(MACHINE_KEYS.map((k, i) => [k, i < 2 ? String(r[i]) : (r[i] === '' ? null : +r[i])])));
      return json_({ ok: true, materials: mats, machines: mach });
    }
    return json_({ ok: true, message: 'Cost estimator API is running' });
  } catch (err) { return json_({ ok: false, error: String(err) }); }
}

function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.action === 'save') {
      const sh = sheet_('Quotes', QUOTE_HEADERS), i = d.state.i;
      const row = [d.id, d.saved, i.customer, i.part, i.machine, i.material, i.laser, i.parts, i.layer, (i.days || 0) * 24 + (i.hrs || 0),
        i.kg, i.sets, i.units, d.total, d.round, d.perPart, JSON.stringify(d.state)].map(blank_);
      const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(r => r[0]), at = ids.indexOf(d.id);
      if (at >= 0) sh.getRange(at + 2, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
      return json_({ ok: true, id: d.id });
    }
    if (d.action === 'delete') {
      const sh = sheet_('Quotes', QUOTE_HEADERS), ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(r => r[0]), at = ids.indexOf(d.id);
      if (at >= 0) sh.deleteRow(at + 2);
      return json_({ ok: true });
    }
    if (d.action === 'saveMasters') {
      const mats = d.materials || [], mach = d.machines || [];
      const p = sheet_('Powders', POWDER_HEADERS); p.clearContents(); p.getRange(1, 1, 1, POWDER_HEADERS.length).setValues([POWDER_HEADERS]).setFontWeight('bold');
      if (mats.length) p.getRange(2, 1, mats.length, POWDER_KEYS.length).setValues(mats.map(x => POWDER_KEYS.map(k => blank_(x[k]))));
      const m = sheet_('Machines', MACHINE_HEADERS); m.clearContents(); m.getRange(1, 1, 1, MACHINE_HEADERS.length).setValues([MACHINE_HEADERS]).setFontWeight('bold');
      if (mach.length) m.getRange(2, 1, mach.length, MACHINE_KEYS.length).setValues(mach.map(x => MACHINE_KEYS.map(k => blank_(x[k]))));
      SpreadsheetApp.flush();
      return json_({ ok: true, powders: mats.length, machines: mach.length });
    }
    return json_({ ok: false, error: 'Unknown action' });
  } catch (err) { return json_({ ok: false, error: String(err) }); }
  finally { lock.releaseLock(); }
}
