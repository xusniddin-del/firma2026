/**
 * Firma admin paneli — Google Sheets bazasi (Apps Script veb-ilova)
 *
 * O'rnatish:
 *  1) Google Sheets'da yangi jadval oching (masalan "Firma 2026 baza").
 *  2) Kengaytmalar → Apps Script. Ochilgan oynadagi kodni o'chirib, shu faylni to'liq qo'ying va saqlang.
 *  3) (Ixtiyoriy, tavsiya etiladi) pastdagi PAROL qatoriga o'zingizning parolingizni yozing.
 *  4) Joylashtirish (Deploy) → Yangi joylashtirish → turi: Veb-ilova.
 *       Kim nomidan ishlaydi: Men
 *       Kim kira oladi: Hamma (Anyone)
 *     "Joylashtirish" ni bosing, ruxsat bering va chiqqan "Veb-ilova URL" ni nusxalang.
 *  5) Saytda: Sozlamalar → URL va parolni kiriting → "Saqlash va ulanish".
 *
 * Kod o'zgartirilsa: Joylashtirish → Joylashtirishlarni boshqarish → tahrirlash → Yangi versiya.
 */

const PAROL = '';   // masalan: 'firma2026-maxfiy'. Saytdagi "Parol" maydoniga ham shuni yozing.

const SHEETS = {
  items:    "Yo'l varaqalari",
  drivers:  'Haydovchilar',
  staff:    'Ichki xodimlar',
  firms:    'Firmalar',
  payments: "To'lovlar",
  fond:     'Fond',
  mehnat:   'Mehnat shartnomalari',
  readN:    "O'qilgan bildirishnomalar",
  tpl:      'Shablonlar'
};
const MAX_CELL = 49000;   // Google Sheets katakka 50 000 belgigacha sig'adi
const CHUNK = 40000;

function doGet(e)  { return handle_((e && e.parameter) || {}); }
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch (err) {}
  return handle_(Object.assign({}, (e && e.parameter) || {}, body));
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function handle_(req) {
  const parol = PropertiesService.getScriptProperties().getProperty('PAROL') || PAROL;
  if (parol && req.token !== parol) return out_({ ok: false, error: 'parol' });
  const lock = LockService.getScriptLock();
  try {
    if (req.action === 'ping') {
      return out_({ ok: true, name: SpreadsheetApp.getActive().getName() });
    }
    if (req.action === 'all') {
      const data = {};
      Object.keys(SHEETS).forEach(function (k) {
        data[k] = k === 'tpl' ? readChunks_(SHEETS[k]) : readList_(SHEETS[k]);
      });
      return out_({ ok: true, data: data });
    }
    if (req.action === 'save') {
      if (!SHEETS[req.key]) return out_({ ok: false, error: 'key' });
      lock.waitLock(20000);
      if (req.key === 'tpl') writeChunks_(SHEETS.tpl, req.value || {});
      else writeList_(SHEETS[req.key], req.value || []);
      return out_({ ok: true, at: new Date().toISOString() });
    }
    return out_({ ok: false, error: 'action' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function sheet_(name) {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/* Ro'yxat: 1-qator — maydon nomlari, keyingi qatorlar — yozuvlar (hammasi matn ko'rinishida). */
function writeList_(name, list) {
  const sh = sheet_(name);
  const rows = (list || []).map(function (x) { return (x !== null && typeof x === 'object') ? x : { key: x }; });
  const old = sh.getLastColumn() > 0 ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].filter(String) : [];
  const keys = old.slice();
  if (keys.indexOf('id') < 0 && rows.some(function (r) { return 'id' in r; })) keys.unshift('id');
  rows.forEach(function (r) { Object.keys(r).forEach(function (k) { if (keys.indexOf(k) < 0) keys.push(k); }); });
  sh.clearContents();
  if (!keys.length) return;
  const values = [keys].concat(rows.map(function (r) {
    return keys.map(function (k) {
      let v = r[k];
      if (v === undefined || v === null) return '';
      if (typeof v === 'object') v = JSON.stringify(v);
      v = String(v);
      return v.length > MAX_CELL ? '' : v;
    });
  }));
  const range = sh.getRange(1, 1, values.length, keys.length);
  range.setNumberFormat('@');
  range.setValues(values);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, keys.length).setFontWeight('bold').setBackground('#E8EDFF');
}

function readList_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues();
  const keys = v[0];
  const onlyKey = keys.length === 1 && keys[0] === 'key';
  return v.slice(1).filter(function (r) { return r.join('') !== ''; }).map(function (r) {
    if (onlyKey) return r[0];
    const o = {};
    keys.forEach(function (k, i) {
      if (!k || r[i] === '') return;
      let val = r[i];
      if ((val[0] === '{' && val[val.length - 1] === '}') || (val[0] === '[' && val[val.length - 1] === ']')) {
        try { val = JSON.parse(val); } catch (e) {}
      }
      o[k] = val;
    });
    return o;
  });
}

/* Shablonlar: katta matn bo'laklarga bo'linib saqlanadi. */
function writeChunks_(name, obj) {
  const sh = sheet_(name);
  const s = JSON.stringify(obj || {});
  const rows = [['qism', "ma'lumot"]];
  for (let i = 0; i < s.length; i += CHUNK) rows.push([String(rows.length), s.slice(i, i + CHUNK)]);
  sh.clearContents();
  const range = sh.getRange(1, 1, rows.length, 2);
  range.setNumberFormat('@');
  range.setValues(rows);
}

function readChunks_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return {};
  const v = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues().map(function (r) { return r[0]; }).join('');
  try { return JSON.parse(v); } catch (e) { return {}; }
}
