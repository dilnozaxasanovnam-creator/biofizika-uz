/**
 * Biofizika.uz — o‘quvchilar ro‘yxati, reyting va kunlik hisobot.
 * Google Sheets → Kengaytmalar → Apps Script ga shu kodni to‘liq joylang.
 * Ma’lumotlar faqat sizning Google jadvalingizda saqlanadi.
 */

const OQUVCHILAR = 'O‘quvchilar';
const FAOLLIK = 'Faollik';
const HISOBOT_SOATI = 20; // har kuni soat 20:00 da hisobot
const TZ = 'Asia/Tashkent';
const USTUNLAR = ['ID', 'Ism familiya', 'Guruh', 'Ro‘yxatdan o‘tgan', 'Oxirgi faollik',
  'Testlar', 'Testlar jami', 'Masalalar', 'Masalalar jami', 'O‘yin', 'O‘yin jami',
  'Tasdiqlangan topshiriqlar', 'Xato urinishlar', 'Ball'];

/** Ball: topshiriq 10, masala 5, test 2, o‘yin juftligi 1. */
function ball_(t, m, o, top) { return top * 10 + m * 5 + t * 2 + o; }

/** 1-marta shu funksiyani ishga tushiring: varaqlar, ustoz kaliti va kunlik hisobotni sozlaydi. */
function sozlash() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TZ);
  varaq_(OQUVCHILAR, USTUNLAR);
  varaq_(FAOLLIK, ['Vaqt', 'ID', 'Ism familiya', 'Guruh', 'Turi', 'Topshiriq ID', 'Nomi', 'Natija']);
  const props = PropertiesService.getScriptProperties();
  let kalit = props.getProperty('USTOZ_KALITI');
  if (!kalit) {
    kalit = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
    props.setProperty('USTOZ_KALITI', kalit);
  }
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'kunlikHisobot') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('kunlikHisobot').timeBased().everyDays(1).atHour(HISOBOT_SOATI).create();
  MailApp.sendEmail({
    to: Session.getEffectiveUser().getEmail(),
    subject: 'Biofizika.uz — ustoz sahifasi kaliti',
    htmlBody: '<p style="font-family:sans-serif">Ustoz sahifasiga kirish kaliti: <b style="font-size:18px">' + kalit +
      '</b></p><p style="font-family:sans-serif">Bu kalitni o‘quvchilarga bermang.</p>'
  });
  Logger.log('Ustoz kaliti: ' + kalit);
}

/** Kalitni almashtirish kerak bo‘lsa shu funksiyani ishga tushiring. */
function kalitniYangilash() {
  PropertiesService.getScriptProperties().deleteProperty('USTOZ_KALITI');
  sozlash();
}

function varaq_(nom, sarlavhalar) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nom);
  if (!sh) sh = ss.insertSheet(nom);
  if (sh.getLastRow() === 0 && sarlavhalar.length) {
    sh.appendRow(sarlavhalar);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, sarlavhalar.length).setFontWeight('bold');
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Sayt shu manzilga natijalarni yuboradi; ustoz sahifasi esa hisobotni so‘raydi. */
function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  if (d.action === 'report') return hisobot_(d);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const toza = v => String(v == null ? '' : v).slice(0, 200);
    const son = v => Math.max(0, Math.min(10000, Number(v) || 0));
    if (!d.sid) return json_({ ok: false });
    const vaqt = new Date();
    const natija = d.ok === true ? 'to‘g‘ri' : d.ok === false ? 'xato' : '';
    varaq_(FAOLLIK, []).appendRow([vaqt, toza(d.sid), toza(d.name), toza(d.group), toza(d.type),
      toza(d.task), toza(d.title), natija]);

    const sh = varaq_(OQUVCHILAR, USTUNLAR);
    const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => r[0]) : [];
    let qator = ids.indexOf(toza(d.sid)) + 2;
    if (qator < 2) {
      sh.appendRow([toza(d.sid), toza(d.name), toza(d.group), vaqt, vaqt, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
      qator = sh.getLastRow();
    }
    const old = sh.getRange(qator, 1, 1, USTUNLAR.length).getValues()[0];
    const s = d.snap || {};
    const t = son(s.test), m = son(s.masala), o = son(s.oyin), top = son(s.topshiriq);
    const xato = Number(old[12] || 0) + (d.ok === false ? 1 : 0);
    sh.getRange(qator, 2, 1, USTUNLAR.length - 1).setValues([[toza(d.name), toza(d.group), old[3] || vaqt, vaqt,
      t, son(s.testN), m, son(s.masalaN), o, son(s.oyinN), top, xato, ball_(t, m, o, top)]]);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

/** Ustoz sahifasi uchun: o‘quvchilar, reyting va so‘nggi faollik. */
function hisobot_(d) {
  const kalit = PropertiesService.getScriptProperties().getProperty('USTOZ_KALITI');
  if (!kalit || String(d.key || '') !== kalit) return json_({ ok: false, error: 'kalit' });
  const oquv = varaq_(OQUVCHILAR, USTUNLAR);
  const log = varaq_(FAOLLIK, []);
  const rows = oquv.getLastRow() > 1 ? oquv.getRange(2, 1, oquv.getLastRow() - 1, USTUNLAR.length).getValues() : [];
  const n = log.getLastRow() - 1;
  const boshi = Math.max(2, log.getLastRow() - 2999);
  const ev = n > 0 ? log.getRange(boshi, 1, log.getLastRow() - boshi + 1, 8).getValues() : [];
  const bugun = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const kun = {};
  ev.forEach(r => {
    if (Utilities.formatDate(new Date(r[0]), TZ, 'yyyy-MM-dd') !== bugun) return;
    const k = kun[r[1]] || (kun[r[1]] = { ok: 0, xato: 0, topshiriq: 0 });
    if (r[7] === 'to‘g‘ri') k.ok++; else if (r[7] === 'xato') k.xato++;
    if (r[4] === 'topshiriq' && r[7] === 'to‘g‘ri') k.topshiriq++;
  });
  const students = rows.map(r => ({
    id: r[0], name: r[1], group: r[2], joined: r[3] ? new Date(r[3]).toISOString() : '',
    last: r[4] ? new Date(r[4]).toISOString() : '', test: Number(r[5]) || 0, testN: Number(r[6]) || 0,
    masala: Number(r[7]) || 0, masalaN: Number(r[8]) || 0, oyin: Number(r[9]) || 0, oyinN: Number(r[10]) || 0,
    topshiriq: Number(r[11]) || 0, xato: Number(r[12]) || 0, ball: Number(r[13]) || 0,
    today: kun[r[0]] || { ok: 0, xato: 0, topshiriq: 0 }
  }));
  const events = ev.slice(-400).reverse().map(r => ({
    ts: new Date(r[0]).toISOString(), id: r[1], name: r[2], group: r[3], type: r[4], task: r[5], title: r[6], result: r[7]
  }));
  return json_({ ok: true, generated: new Date().toISOString(), students: students, events: events });
}

function doGet() {
  return ContentService.createTextOutput('Biofizika.uz hisobot xizmati ishlayapti.');
}

function esc_(v) {
  return String(v == null ? '' : v).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

/** Oxirgi 24 soatdagi faollik bo‘yicha sizga elektron xat yuboradi. */
function kunlikHisobot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = varaq_(FAOLLIK, []);
  const oquv = varaq_(OQUVCHILAR, USTUNLAR);
  const chegara = new Date(Date.now() - 24 * 3600 * 1000);
  const qatorlar = log.getLastRow() > 1 ? log.getRange(2, 1, log.getLastRow() - 1, 8).getValues() : [];

  const kim = {};
  qatorlar.filter(r => r[0] >= chegara).forEach(r => {
    const id = r[1];
    const k = kim[id] || (kim[id] = { ism: r[2], guruh: r[3], yangi: false, topshiriq: [], xatoT: 0, test: 0, testXato: 0, masala: 0 });
    if (r[4] === 'royxat') k.yangi = true;
    if (r[4] === 'topshiriq') { if (r[7] === 'to‘g‘ri') k.topshiriq.push(r[6]); else k.xatoT++; }
    if (r[4] === 'test') { if (r[7] === 'to‘g‘ri') k.test++; else k.testXato++; }
    if (r[4] === 'masala' && r[7] === 'to‘g‘ri') k.masala++;
  });

  const reyting = oquv.getLastRow() > 1 ? oquv.getRange(2, 1, oquv.getLastRow() - 1, USTUNLAR.length).getValues()
    .sort((a, b) => (Number(b[13]) || 0) - (Number(a[13]) || 0)).slice(0, 10) : [];
  const jami = Math.max(oquv.getLastRow() - 1, 0);
  const faol = Object.keys(kim).length;
  const sana = Utilities.formatDate(new Date(), TZ, 'dd.MM.yyyy');
  const td = 'style="border:1px solid #DCE8E1;padding:6px 8px;vertical-align:top"';
  let html = '<h2 style="font-family:sans-serif">Biofizika.uz — kunlik hisobot, ' + sana + '</h2>' +
    '<p style="font-family:sans-serif">Ro‘yxatdagi o‘quvchilar: <b>' + jami + '</b>. Oxirgi 24 soatda faol bo‘lganlar: <b>' + faol + '</b>.</p>';
  if (faol) {
    html += '<h3 style="font-family:sans-serif">Bugun kim nima qildi</h3><table style="border-collapse:collapse;font-family:sans-serif;font-size:14px"><tr>' +
      ['O‘quvchi', 'Guruh', 'Tasdiqlangan topshiriqlar', 'Xato urinish', 'Testlar (to‘g‘ri / xato)', 'Masalalar'].map(h => '<th ' + td + '>' + h + '</th>').join('') + '</tr>';
    Object.keys(kim).map(id => kim[id]).sort((a, b) => b.topshiriq.length - a.topshiriq.length).forEach(k => {
      html += '<tr><td ' + td + '>' + esc_(k.ism) + (k.yangi ? ' <i>(yangi)</i>' : '') + '</td><td ' + td + '>' + esc_(k.guruh) + '</td><td ' + td + '>' +
        k.topshiriq.length + (k.topshiriq.length ? '<br><small>' + k.topshiriq.map(esc_).join('; ') + '</small>' : '') + '</td><td ' + td + '>' + k.xatoT +
        '</td><td ' + td + '>' + k.test + ' / ' + k.testXato + '</td><td ' + td + '>' + k.masala + '</td></tr>';
    });
    html += '</table>';
  }
  if (reyting.length) {
    html += '<h3 style="font-family:sans-serif">Umumiy reyting (TOP-10)</h3><table style="border-collapse:collapse;font-family:sans-serif;font-size:14px"><tr>' +
      ['O‘rin', 'O‘quvchi', 'Guruh', 'Ball'].map(h => '<th ' + td + '>' + h + '</th>').join('') + '</tr>';
    reyting.forEach((r, i) => {
      html += '<tr><td ' + td + '>' + (i + 1) + '</td><td ' + td + '>' + esc_(r[1]) + '</td><td ' + td + '>' + esc_(r[2]) + '</td><td ' + td + '>' + (Number(r[13]) || 0) + '</td></tr>';
    });
    html += '</table>';
  }
  html += '<p style="font-family:sans-serif;color:#4F6559">To‘liq ma’lumot: <a href="' + ss.getUrl() + '">Google jadval</a> · ' +
    '<a href="https://dilnozaxasanovnam-creator.github.io/biofizika-uz/ustoz.html">Ustoz sahifasi</a></p>';
  MailApp.sendEmail({ to: Session.getEffectiveUser().getEmail(), subject: 'Biofizika.uz — kunlik hisobot ' + sana, htmlBody: html });
}
