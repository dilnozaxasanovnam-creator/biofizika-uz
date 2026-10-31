/**
 * Biofizika.uz — o‘quvchilar ro‘yxati, reyting va kunlik hisobot.
 * Mustaqil Apps Script loyihasi; jadvalga JADVAL_ID orqali ulanadi.
 * Ma’lumotlar faqat sizning Google jadvalingizda saqlanadi.
 */

const OQUVCHILAR = 'O‘quvchilar';
const FAOLLIK = 'Faollik';
const HISOBOT_SOATI = 20; // har kuni soat 20:00 da hisobot
const TZ = 'Asia/Tashkent';
const GOOGLE_CLIENT_ID = '668867285337-i7v5dlhn9tnpjmvi4bchjhgl6e3q1930.apps.googleusercontent.com'; // Google Cloud → OAuth Client ID (Web)
const JADVAL_ID = '1ErxHBYG6P8TVESLtYuTyzsklvWvCZ5KSmQ_q4Uc9hjo'; // «Biofizika.uz o‘quvchilar» jadvali
function jadval_() { return SpreadsheetApp.openById(JADVAL_ID); }
const USTUNLAR = ['ID', 'Ism familiya', 'Guruh', 'Ro‘yxatdan o‘tgan', 'Oxirgi faollik',
  'Testlar', 'Testlar jami', 'Masalalar', 'Masalalar jami', 'O‘yin', 'O‘yin jami',
  'Tasdiqlangan topshiriqlar', 'Xato urinishlar', 'Ball'];

/** Ball: topshiriq 10, masala 5, test 2, o‘yin juftligi 1. */
function ball_(t, m, o, top) { return top * 10 + m * 5 + t * 2 + o; }

/** 1-marta shu funksiyani ishga tushiring: varaqlar, ustoz kaliti va kunlik hisobotni sozlaydi. */
function sozlash() {
  const ss = jadval_();
  ss.setSpreadsheetTimeZone(TZ);
  varaq_(OQUVCHILAR, USTUNLAR);
  varaq_(FAOLLIK, ['Vaqt', 'ID', 'Ism familiya', 'Guruh', 'Turi', 'Topshiriq ID', 'Nomi', 'Natija']);
  varaq_(HISOBLAR, H_USTUN);
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
  const ss = jadval_();
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

/** Sayt shu manzilga murojaat qiladi: hisob, natijalar va ustoz hisoboti. */
function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  if (d.action === 'report') return hisobot_(d);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (d.action === 'register') return royxat_(d);
    if (d.action === 'login') return kirish_(d);
    if (d.action === 'sync') return sinx_(d);
    if (d.action === 'reset') return tiklash_(d);
    if (d.action === 'google') return google_(d);
    if (d.action === 'ustoz_google') return ustozGoogle_(d);
    if (d.action === 'ustoz_login') return ustozLogin_(d);
    if (d.action === 'ustoz_parol') return ustozParol_(d);
    if (d.action === 'ustoz_royxat') return ustozRoyxat_(d);
    if (!d.sid) return json_({ ok: false });
    yozish_(d);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

const toza_ = v => String(v == null ? '' : v).slice(0, 200);
const son_ = v => Math.max(0, Math.min(10000, Number(v) || 0));

/** Faollik jurnaliga yozish va o‘quvchi qatorini yangilash. */
function yozish_(d) {
  const vaqt = new Date();
  const natija = d.ok === true ? 'to‘g‘ri' : d.ok === false ? 'xato' : '';
  varaq_(FAOLLIK, []).appendRow([vaqt, toza_(d.sid), toza_(d.name), toza_(d.group), toza_(d.type),
    toza_(d.task), toza_(d.title), natija]);
  const sh = varaq_(OQUVCHILAR, USTUNLAR);
  const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => r[0]) : [];
  let qator = ids.indexOf(toza_(d.sid)) + 2;
  if (qator < 2) {
    sh.appendRow([toza_(d.sid), toza_(d.name), toza_(d.group), vaqt, vaqt, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
    qator = sh.getLastRow();
  }
  const old = sh.getRange(qator, 1, 1, USTUNLAR.length).getValues()[0];
  const s = d.snap || {};
  const t = son_(s.test), m = son_(s.masala), o = son_(s.oyin), top = son_(s.topshiriq);
  const xato = Number(old[12] || 0) + (d.ok === false ? 1 : 0);
  sh.getRange(qator, 2, 1, USTUNLAR.length - 1).setValues([[toza_(d.name) || old[1], toza_(d.group) || old[2], old[3] || vaqt, vaqt,
    t, son_(s.testN), m, son_(s.masalaN), o, son_(s.oyinN), top, xato, ball_(t, m, o, top)]]);
}

/* ---------- O‘quvchi hisoblari ---------- */
const HISOBLAR = 'Hisoblar';
const H_USTUN = ['Login', 'Salt', 'Parol xeshi', 'ID', 'Ism familiya', 'Guruh', 'Yaratilgan', 'Token', 'Natijalar (JSON)', 'Yangilangan'];

function xesh_(salt, parol) {
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + parol, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(b);
}
function loginToza_(v) { return String(v || '').trim().toLowerCase().slice(0, 80); }
function hisobTop_(login) {
  const sh = varaq_(HISOBLAR, H_USTUN);
  const n = sh.getLastRow() - 1;
  if (n < 1) return { sh: sh, qator: 0 };
  const logins = sh.getRange(2, 1, n, 1).getValues().map(r => String(r[0]));
  const i = logins.indexOf(login);
  return { sh: sh, qator: i < 0 ? 0 : i + 2 };
}
function hisobId_(sid) {
  const sh = varaq_(HISOBLAR, H_USTUN);
  const n = sh.getLastRow() - 1;
  if (n < 1) return { sh: sh, qator: 0 };
  const ids = sh.getRange(2, 4, n, 1).getValues().map(r => String(r[0]));
  const i = ids.indexOf(String(sid || ''));
  return { sh: sh, qator: i < 0 ? 0 : i + 2 };
}
function yangiToken_() { return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8); }
function urinish_(login, xato) {
  const c = CacheService.getScriptCache(), k = 'urinish_' + Utilities.base64EncodeWebSafe(login).slice(0, 200);
  const n = Number(c.get(k) || 0);
  if (xato) c.put(k, String(n + 1), 600);
  return n;
}

function royxat_(d) {
  const login = loginToza_(d.login), parol = String(d.password || '');
  const ism = toza_(d.name), guruh = toza_(d.group);
  if (!/^[a-z0-9._@+-]{3,80}$/.test(login)) return json_({ ok: false, error: 'login' });
  if (parol.length < 6 || parol.length > 100) return json_({ ok: false, error: 'parol' });
  if (ism.length < 3 || guruh.length < 2) return json_({ ok: false, error: 'ism' });
  const h = hisobTop_(login);
  if (h.qator) return json_({ ok: false, error: 'band' });
  const sid = String(d.sid || '').match(/^s[a-z0-9]{6,30}$/) && !hisobId_(d.sid).qator ? String(d.sid) :
    's' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  const salt = Utilities.getUuid(), token = yangiToken_(), vaqt = new Date();
  const holat = JSON.stringify(d.state || {}).slice(0, 45000);
  h.sh.appendRow([login, salt, xesh_(salt, parol), sid, ism, guruh, vaqt, token, holat, vaqt]);
  yozish_({ sid: sid, name: ism, group: guruh, type: 'royxat', snap: d.snap || {} });
  return json_({ ok: true, sid: sid, token: token, login: login, name: ism, group: guruh });
}

function kirish_(d) {
  const login = loginToza_(d.login), parol = String(d.password || '');
  if (urinish_(login, false) >= 8) return json_({ ok: false, error: 'kop' });
  const h = hisobTop_(login);
  if (!h.qator) { urinish_(login, true); return json_({ ok: false, error: 'topilmadi' }); }
  const r = h.sh.getRange(h.qator, 1, 1, H_USTUN.length).getValues()[0];
  if (r[1] === 'google' && !r[2]) return json_({ ok: false, error: 'google' });
  if (xesh_(r[1], parol) !== r[2]) { urinish_(login, true); return json_({ ok: false, error: 'parol' }); }
  const token = yangiToken_();
  h.sh.getRange(h.qator, 8).setValue(token);
  let holat = {};
  try { holat = JSON.parse(r[8] || '{}'); } catch (err) { holat = {}; }
  yozish_({ sid: r[3], name: r[4], group: r[5], type: 'kirish', snap: d.snap || {} });
  return json_({ ok: true, sid: r[3], token: token, login: r[0], name: r[4], group: r[5], state: holat });
}

function sinx_(d) {
  const h = hisobId_(d.sid);
  if (!h.qator) return json_({ ok: false, error: 'topilmadi' });
  const r = h.sh.getRange(h.qator, 1, 1, H_USTUN.length).getValues()[0];
  if (!d.token || String(d.token) !== String(r[7])) return json_({ ok: false, error: 'token' });
  const qiymat = [[JSON.stringify(d.state || {}).slice(0, 45000), new Date()]];
  h.sh.getRange(h.qator, 9, 1, 2).setValues(qiymat);
  if (d.name || d.group) h.sh.getRange(h.qator, 5, 1, 2).setValues([[toza_(d.name) || r[4], toza_(d.group) || r[5]]]);
  return json_({ ok: true });
}

/** Google bilan kirish: ID tokenni Google’da tekshiradi. */
function google_(d) {
  if (!GOOGLE_CLIENT_ID) return json_({ ok: false, error: 'sozlanmagan' });
  let info;
  try {
    const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(String(d.credential || '')), { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return json_({ ok: false, error: 'token' });
    info = JSON.parse(res.getContentText());
  } catch (err) { return json_({ ok: false, error: 'token' }); }
  if (info.aud !== GOOGLE_CLIENT_ID || String(info.email_verified) !== 'true' || !info.email) return json_({ ok: false, error: 'token' });
  const login = loginToza_(info.email);
  const h = hisobTop_(login);
  if (h.qator) {
    const r = h.sh.getRange(h.qator, 1, 1, H_USTUN.length).getValues()[0];
    const token = yangiToken_();
    h.sh.getRange(h.qator, 8).setValue(token);
    let holat = {};
    try { holat = JSON.parse(r[8] || '{}'); } catch (err) { holat = {}; }
    yozish_({ sid: r[3], name: r[4], group: r[5], type: 'kirish', snap: d.snap || {} });
    return json_({ ok: true, sid: r[3], token: token, login: r[0], name: r[4], group: r[5], state: holat });
  }
  const ism = toza_(d.name) || toza_([info.given_name, info.family_name].filter(Boolean).join(' ')) || toza_(info.name);
  const guruh = toza_(d.group);
  if (!guruh) return json_({ ok: false, error: 'yangi', email: login, name: ism });
  const sid = String(d.sid || '').match(/^s[a-z0-9]{6,30}$/) && !hisobId_(d.sid).qator ? String(d.sid) :
    's' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  const token = yangiToken_(), vaqt = new Date();
  h.sh.appendRow([login, 'google', '', sid, ism, guruh, vaqt, token, JSON.stringify(d.state || {}).slice(0, 45000), vaqt]);
  yozish_({ sid: sid, name: ism, group: guruh, type: 'royxat', snap: d.snap || {} });
  return json_({ ok: true, sid: sid, token: token, login: login, name: ism, group: guruh, state: d.state || {} });
}

/* ---------- Ustoz hisoblari: faqat ruxsat berilgan Gmail manzillari ---------- */
function egasi_() { return String(Session.getEffectiveUser().getEmail() || '').toLowerCase(); }
function prop_(k, d) { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(k) || '') ; } catch (e) { return d; } }
function propSet_(k, v) { PropertiesService.getScriptProperties().setProperty(k, JSON.stringify(v)); }
function ustozlar_() {
  const list = prop_('USTOZLAR', []) || [];
  const e = egasi_();
  return [e].concat(list.filter(x => x && x !== e));
}
function ustozTokenBer_(email) {
  const t = yangiToken_(), now = Date.now();
  const list = (prop_('USTOZ_TOKENLAR', []) || []).filter(x => x.exp > now).slice(-29);
  list.push({ t: t, e: email, exp: now + 30 * 864e5 });
  propSet_('USTOZ_TOKENLAR', list);
  return t;
}
function ustozKim_(d) {
  const kalit = PropertiesService.getScriptProperties().getProperty('USTOZ_KALITI');
  if (kalit && d.key && String(d.key) === kalit) return egasi_();
  if (!d.token) return '';
  const now = Date.now(), x = (prop_('USTOZ_TOKENLAR', []) || []).find(y => y.t === String(d.token) && y.exp > now);
  return x && ustozlar_().indexOf(x.e) >= 0 ? x.e : '';
}
function ustozOk_(d) { return !!ustozKim_(d); }
function ustozJavob_(email) {
  const parollar = prop_('USTOZ_PAROLLAR', {}) || {};
  const owner = email === egasi_();
  return { ok: true, token: ustozTokenBer_(email), email: email, hasPassword: !!parollar[email], owner: owner, teachers: owner ? ustozlar_() : undefined };
}
function ustozGoogle_(d) {
  if (!GOOGLE_CLIENT_ID) return json_({ ok: false, error: 'sozlanmagan' });
  let info;
  try {
    const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(String(d.credential || '')), { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return json_({ ok: false, error: 'token' });
    info = JSON.parse(res.getContentText());
  } catch (err) { return json_({ ok: false, error: 'token' }); }
  if (info.aud !== GOOGLE_CLIENT_ID || String(info.email_verified) !== 'true') return json_({ ok: false, error: 'token' });
  const email = String(info.email || '').toLowerCase();
  if (ustozlar_().indexOf(email) < 0) return json_({ ok: false, error: 'ustoz_emas', email: email });
  return json_(ustozJavob_(email));
}
function ustozLogin_(d) {
  const login = loginToza_(d.login), parol = String(d.password || '');
  if (urinish_('ustoz:' + login, false) >= 6) return json_({ ok: false, error: 'kop' });
  const parollar = prop_('USTOZ_PAROLLAR', {}) || {};
  if (ustozlar_().indexOf(login) < 0 || !parollar[login]) { urinish_('ustoz:' + login, true); return json_({ ok: false, error: 'xato' }); }
  const sp = String(parollar[login]).split('|');
  if (xesh_(sp[0], parol) !== sp[1]) { urinish_('ustoz:' + login, true); return json_({ ok: false, error: 'xato' }); }
  return json_(ustozJavob_(login));
}
function ustozParol_(d) {
  const email = ustozKim_(d);
  if (!email) return json_({ ok: false, error: 'kalit' });
  const parol = String(d.password || '');
  if (parol.length < 8 || parol.length > 100) return json_({ ok: false, error: 'parol' });
  const parollar = prop_('USTOZ_PAROLLAR', {}) || {};
  const salt = Utilities.getUuid();
  parollar[email] = salt + '|' + xesh_(salt, parol);
  propSet_('USTOZ_PAROLLAR', parollar);
  return json_({ ok: true, email: email });
}
/** Faqat sayt egasi: ustoz qo‘shish yoki olib tashlash. */
function ustozRoyxat_(d) {
  if (ustozKim_(d) !== egasi_()) return json_({ ok: false, error: 'egasi_emas' });
  const email = loginToza_(d.email);
  let list = (prop_('USTOZLAR', []) || []).filter(x => x !== egasi_());
  if (d.op === 'add') {
    if (!/^[^@\s]+@[^@\s]+\.[a-z]{2,}$/.test(email)) return json_({ ok: false, error: 'email' });
    if (list.indexOf(email) < 0) list.push(email);
  } else if (d.op === 'remove') {
    list = list.filter(x => x !== email);
    const parollar = prop_('USTOZ_PAROLLAR', {}) || {}; delete parollar[email]; propSet_('USTOZ_PAROLLAR', parollar);
    propSet_('USTOZ_TOKENLAR', (prop_('USTOZ_TOKENLAR', []) || []).filter(x => x.e !== email));
  }
  propSet_('USTOZLAR', list.slice(0, 50));
  return json_({ ok: true, teachers: ustozlar_() });
}

/** Ustoz o‘quvchi parolini tiklaydi: vaqtinchalik 6 xonali parol qaytaradi. */
function tiklash_(d) {
  if (!ustozOk_(d)) return json_({ ok: false, error: 'kalit' });
  const h = hisobId_(d.sid);
  if (!h.qator) return json_({ ok: false, error: 'topilmadi' });
  const yangi = String(Math.floor(100000 + Math.random() * 900000));
  const salt = Utilities.getUuid();
  h.sh.getRange(h.qator, 2, 1, 2).setValues([[salt, xesh_(salt, yangi)]]);
  h.sh.getRange(h.qator, 8).setValue('');
  return json_({ ok: true, password: yangi, login: h.sh.getRange(h.qator, 1).getValue() });
}

/** Ustoz sahifasi uchun: o‘quvchilar, reyting va so‘nggi faollik. */
function hisobot_(d) {
  if (!ustozOk_(d)) return json_({ ok: false, error: 'kalit' });
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
  const hs = varaq_(HISOBLAR, H_USTUN);
  const loginlar = {};
  if (hs.getLastRow() > 1) hs.getRange(2, 1, hs.getLastRow() - 1, 4).getValues().forEach(x => { loginlar[x[3]] = x[0]; });
  const students = rows.map(r => ({
    id: r[0], login: loginlar[r[0]] || '', name: r[1], group: r[2], joined: r[3] ? new Date(r[3]).toISOString() : '',
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
  const ss = jadval_();
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
