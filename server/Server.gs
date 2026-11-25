/**
 * שרת הרשמה וכניסה – מערכת המזכירות
 * ---------------------------------------------
 * רץ ב-Google Apps Script שמחובר לגיליון Google Sheets, בחשבון של בעל המערכת.
 * מה הוא עושה:
 *   • הרשמה רק עם קוד קופון שבעל המערכת נתן
 *   • כניסה עם קוד חד-פעמי שנשלח למייל (בלי סיסמאות)
 *   • תוקף גישה: ימי ניסיון / חודשי / ללא הגבלה (מוכן להרחבה לתשלום)
 *   • (לא חובה) קריאת טפסים סרוקים בעזרת Claude, אם הוגדר מפתח API
 *
 * גיליונות: משתמשים · קופונים · יומן
 * התקנה: ראו admin-setup.html באתר.
 */

const SH_USERS = 'משתמשים', SH_COUPONS = 'קופונים', SH_LOG = 'יומן';
const U_COLS = ['מייל', 'שם', 'מוסד', 'סטטוס', 'תוכנית', 'בתוקף עד', 'קופון', 'נרשם', 'כניסה אחרונה', 'הערות תשלום'];
const K_COLS = ['קוד', 'תיאור', 'תוכנית', 'ימי גישה', 'מקסימום שימושים', 'נוצלו', 'בתוקף עד', 'פעיל'];
const U = {}, K = {};
U_COLS.forEach((c, i) => U[c] = i); K_COLS.forEach((c, i) => K[c] = i);
const SESSION_DAYS = 30, CODE_MINUTES = 15, MAX_TRIES = 5;
const APP_NAME = 'מערכת המזכירות';

// ============================ התקנה ============================
function setup() {
  const ss = SpreadsheetApp.getActive();
  const mk = (name, cols, color) => {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      const first = ss.getSheets()[0];
      sh = (first.getLastRow() === 0 && /^(Sheet|גיליון)\d*$/.test(first.getName())) ? first.setName(name) : ss.insertSheet(name);
    }
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground(color).setFontColor('#fff');
      sh.setFrozenRows(1); sh.setRightToLeft(true);
    }
    return sh;
  };
  const users = mk(SH_USERS, U_COLS, '#1d5fa8');
  users.getRange(2, U['סטטוס'] + 1, 999).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['פעיל', 'חסום'], true).build());
  users.getRange(2, U['בתוקף עד'] + 1, 999).setNumberFormat('dd/MM/yyyy');
  users.getRange(2, U['נרשם'] + 1, 999, 2).setNumberFormat('dd/MM/yyyy HH:mm');
  users.setColumnWidth(1, 220);
  const coupons = mk(SH_COUPONS, K_COLS, '#15803d');
  coupons.getRange(2, K['בתוקף עד'] + 1, 999).setNumberFormat('dd/MM/yyyy');
  coupons.getRange(2, K['קוד'] + 1, 999).setNumberFormat('@');
  coupons.getRange(2, K['פעיל'] + 1, 999).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['כן', 'לא'], true).build());
  coupons.getRange(2, K['תוכנית'] + 1, 999).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['ניסיון', 'חודשי', 'שנתי', 'מלא'], false).build());
  if (coupons.getLastRow() < 2) {
    coupons.appendRow(['NISAYON14', 'דוגמה: 14 ימי ניסיון', 'ניסיון', 14, 20, 0, '', 'כן']);
    coupons.appendRow([newCouponCode_(), 'דוגמה: גישה מלאה לאדם אחד', 'מלא', 0, 1, 0, '', 'כן']);
  }
  mk(SH_LOG, ['זמן', 'פעולה', 'מייל', 'פרטים'], '#6b7280');
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('SECRET')) p.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  return 'ההתקנה הושלמה. עכשיו: פריסה ← פריסה חדשה ← אפליקציית אינטרנט (גישה: כל אחד).';
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🔑 מערכת הרשמה')
    .addItem('התקנה ראשונית', 'menuSetup')
    .addItem('יצירת קוד קופון חדש', 'menuCoupon')
    .addItem('שליחת מייל בדיקה אליי', 'menuTestMail')
    .addItem('הגדרת מפתח לקריאת טפסים (לא חובה)', 'menuApiKey')
    .addToUi();
}
function menuSetup() { SpreadsheetApp.getUi().alert(setup()); }
function menuTestMail() {
  const me = Session.getEffectiveUser().getEmail();
  sendCodeMail_(me, '123456', 'בדיקה');
  SpreadsheetApp.getUi().alert('נשלח מייל בדיקה אל ' + me + '. אם הגיע – החיבור למייל תקין.');
}
function menuCoupon() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('קופון חדש', 'למי הקופון? (תיאור קצר, למשל: "בית ספר אור – מזכיר")', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const d = ui.prompt('ימי גישה', 'כמה ימים הגישה תקפה? (0 = ללא הגבלה, 14 = שבועיים ניסיון, 30 = חודש)', ui.ButtonSet.OK_CANCEL);
  if (d.getSelectedButton() !== ui.Button.OK) return;
  const days = parseInt(d.getResponseText(), 10) || 0;
  const code = newCouponCode_();
  SpreadsheetApp.getActive().getSheetByName(SH_COUPONS).appendRow([code, r.getResponseText(), days ? (days <= 30 ? 'ניסיון' : 'חודשי') : 'מלא', days, 1, 0, '', 'כן']);
  ui.alert('הקופון נוצר: ' + code + '\n\nשלחו אותו לאדם יחד עם הקישור לאתר. הקופון טוב לאדם אחד.');
}
function menuApiKey() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('מפתח API של Anthropic', 'כדי שהמערכת תקרא טפסים סרוקים, הדביקו כאן מפתח מ-console.anthropic.com (בתשלום לפי שימוש).\nכדי לבטל – השאירו ריק.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const k = r.getResponseText().trim(), p = PropertiesService.getScriptProperties();
  if (k) p.setProperty('ANTHROPIC_KEY', k); else p.deleteProperty('ANTHROPIC_KEY');
  ui.alert(k ? 'המפתח נשמר. קריאת טפסים פעילה.' : 'קריאת טפסים כובתה.');
}
function newCouponCode_() {
  const abc = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; let s = '';
  for (let i = 0; i < 8; i++) s += abc[Math.floor(Math.random() * abc.length)];
  return s.slice(0, 4) + '-' + s.slice(4);
}

// ============================ API ============================
// האתר שולח POST עם JSON (Content-Type: text/plain כדי לא להיתקע ב-CORS)
function doPost(e) {
  let req = {};
  try { req = JSON.parse(e.postData.contents || '{}'); } catch (err) { return out_({ ok: false, error: 'בקשה לא תקינה' }); }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const fn = { register: apiRegister_, login: apiLogin_, verify: apiVerify_, me: apiMe_, extract: apiExtract_ }[req.action];
    if (!fn) return out_({ ok: false, error: 'פעולה לא מוכרת' });
    return out_(fn(req));
  } catch (err) {
    log_('שגיאה', req.email || '', String(err));
    return out_({ ok: false, error: 'שגיאה בשרת. נסו שוב בעוד רגע.' });
  } finally { try { lock.releaseLock(); } catch (e2) {} }
}
function doGet() { return out_({ ok: true, app: APP_NAME, ai: !!PropertiesService.getScriptProperties().getProperty('ANTHROPIC_KEY') }); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

const normEmail_ = e => String(e || '').trim().toLowerCase();
const validEmail_ = e => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e);

/** הרשמה: שם, מוסד, מייל, קופון → שולח קוד כניסה */
function apiRegister_(r) {
  const email = normEmail_(r.email), coupon = String(r.coupon || '').trim().toUpperCase();
  if (!validEmail_(email)) return { ok: false, error: 'כתובת המייל לא תקינה' };
  if (!String(r.name || '').trim()) return { ok: false, error: 'חסר שם' };
  if (!coupon) return { ok: false, error: 'חסר קוד קופון' };
  if (tooMany_('reg:' + email, 6)) return { ok: false, error: 'יותר מדי ניסיונות. נסו שוב בעוד שעה.' };
  const existing = findUser_(email);
  if (existing && existing.status === 'פעיל' && !expired_(existing)) {
    sendLoginCode_(email); return { ok: true, sent: true, note: 'המייל כבר רשום – שלחנו קוד כניסה' };
  }
  const k = findCoupon_(coupon);
  if (!k) { log_('קופון שגוי', email, coupon); return { ok: false, error: 'קוד הקופון לא נמצא' }; }
  if (k.active !== 'כן') return { ok: false, error: 'הקופון כבר לא פעיל' };
  if (k.validUntil && k.validUntil < new Date()) return { ok: false, error: 'תוקף הקופון פג' };
  if (k.max > 0 && k.used >= k.max) return { ok: false, error: 'הקופון כבר נוצל' };
  const until = k.days > 0 ? new Date(Date.now() + k.days * 86400000) : '';
  const sh = SpreadsheetApp.getActive().getSheetByName(SH_USERS);
  if (existing) {
    // משתמש שפג תוקפו מחדש עם קופון חדש
    const row = existing.row;
    sh.getRange(row, U['סטטוס'] + 1).setValue('פעיל');
    sh.getRange(row, U['תוכנית'] + 1).setValue(k.plan);
    sh.getRange(row, U['בתוקף עד'] + 1).setValue(until);
    sh.getRange(row, U['קופון'] + 1).setValue(coupon);
  } else {
    sh.appendRow([email, String(r.name).trim(), String(r.school || '').trim(), 'פעיל', k.plan, until, coupon, new Date(), '', '']);
  }
  SpreadsheetApp.getActive().getSheetByName(SH_COUPONS).getRange(k.row, K['נוצלו'] + 1).setValue(k.used + 1);
  log_('הרשמה', email, coupon + ' · ' + k.plan);
  notifyOwner_('נרשם משתמש חדש: ' + r.name + ' (' + email + ')', 'מוסד: ' + (r.school || '-') + '\nקופון: ' + coupon + '\nתוכנית: ' + k.plan + (until ? '\nבתוקף עד: ' + fmtD_(until) : ''));
  sendLoginCode_(email);
  return { ok: true, sent: true };
}

/** כניסה של משתמש רשום → שולח קוד */
function apiLogin_(r) {
  const email = normEmail_(r.email);
  if (!validEmail_(email)) return { ok: false, error: 'כתובת המייל לא תקינה' };
  if (tooMany_('login:' + email, 6)) return { ok: false, error: 'יותר מדי ניסיונות. נסו שוב בעוד שעה.' };
  const u = findUser_(email);
  if (!u) return { ok: false, error: 'המייל הזה לא רשום. להרשמה צריך קוד קופון.', needCoupon: true };
  if (u.status !== 'פעיל') return { ok: false, error: 'הגישה חסומה. פנו למנהל המערכת.' };
  if (expired_(u)) return { ok: false, error: 'תקופת הגישה הסתיימה. להארכה פנו למנהל המערכת, או הזינו קוד קופון חדש.', expired: true, needCoupon: true };
  sendLoginCode_(email);
  return { ok: true, sent: true };
}

/** אימות הקוד מהמייל → טוקן לכניסה */
function apiVerify_(r) {
  const email = normEmail_(r.email), code = String(r.code || '').replace(/\D/g, '');
  const cache = CacheService.getScriptCache(), key = 'otp:' + email, raw = cache.get(key);
  if (!raw) return { ok: false, error: 'הקוד פג תוקף. בקשו קוד חדש.' };
  const o = JSON.parse(raw);
  if (o.tries >= MAX_TRIES) { cache.remove(key); return { ok: false, error: 'יותר מדי ניסיונות. בקשו קוד חדש.' }; }
  if (o.code !== code) { o.tries++; cache.put(key, JSON.stringify(o), CODE_MINUTES * 60); return { ok: false, error: 'הקוד לא נכון' }; }
  cache.remove(key);
  const u = findUser_(email);
  if (!u || u.status !== 'פעיל' || expired_(u)) return { ok: false, error: 'אין גישה פעילה' };
  SpreadsheetApp.getActive().getSheetByName(SH_USERS).getRange(u.row, U['כניסה אחרונה'] + 1).setValue(new Date());
  log_('כניסה', email, '');
  return { ok: true, token: makeToken_(email), user: publicUser_(u) };
}

/** בדיקת טוקן בכל פתיחה של האתר */
function apiMe_(r) {
  const email = checkToken_(r.token);
  if (!email) return { ok: false, error: 'צריך להתחבר מחדש', relogin: true };
  const u = findUser_(email);
  if (!u || u.status !== 'פעיל') return { ok: false, error: 'הגישה חסומה', relogin: true };
  if (expired_(u)) return { ok: false, error: 'תקופת הגישה הסתיימה', relogin: true, expired: true };
  return { ok: true, user: publicUser_(u), ai: !!PropertiesService.getScriptProperties().getProperty('ANTHROPIC_KEY') };
}

/** קריאת טופס סרוק (לא חובה – רק אם הוגדר מפתח) */
function apiExtract_(r) {
  const email = checkToken_(r.token);
  if (!email) return { ok: false, error: 'צריך להתחבר מחדש', relogin: true };
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_KEY');
  if (!key) return { ok: false, error: 'קריאת טפסים לא הופעלה במערכת' };
  if (tooMany_('ai:' + email, 200)) return { ok: false, error: 'הגעתם למכסת הקריאה לשעה' };
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: PropertiesService.getScriptProperties().getProperty('MODEL') || 'claude-haiku-5-5', max_tokens: 1500,
      messages: [{ role: 'user', content: [
        { type: 'image', source: { type: 'base64', media_type: r.mediaType || 'image/jpeg', data: r.image } },
        { type: 'text', text: String(r.prompt || '').slice(0, 4000) }] }]
    })
  });
  if (res.getResponseCode() !== 200) { log_('קריאת טופס נכשלה', email, res.getContentText().slice(0, 300)); return { ok: false, error: 'הקריאה נכשלה (' + res.getResponseCode() + ')' }; }
  const j = JSON.parse(res.getContentText());
  log_('קריאת טופס', email, '');
  return { ok: true, text: (j.content || []).filter(c => c.type === 'text').map(c => c.text).join('') };
}

// ============================ עזרים ============================
function findUser_(email) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH_USERS), n = sh.getLastRow() - 1;
  if (n < 1) return null;
  const v = sh.getRange(2, 1, n, U_COLS.length).getValues();
  for (let i = 0; i < v.length; i++) if (normEmail_(v[i][U['מייל']]) === email) {
    return { row: i + 2, email, name: v[i][U['שם']], school: v[i][U['מוסד']], status: v[i][U['סטטוס']] || 'פעיל', plan: v[i][U['תוכנית']],
      until: v[i][U['בתוקף עד']] instanceof Date ? v[i][U['בתוקף עד']] : null };
  }
  return null;
}
function findCoupon_(code) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH_COUPONS), n = sh.getLastRow() - 1;
  if (n < 1) return null;
  const v = sh.getRange(2, 1, n, K_COLS.length).getValues();
  for (let i = 0; i < v.length; i++) if (String(v[i][K['קוד']]).trim().toUpperCase() === code) {
    return { row: i + 2, plan: v[i][K['תוכנית']] || 'מלא', days: +v[i][K['ימי גישה']] || 0, max: +v[i][K['מקסימום שימושים']] || 0,
      used: +v[i][K['נוצלו']] || 0, validUntil: v[i][K['בתוקף עד']] instanceof Date ? v[i][K['בתוקף עד']] : null, active: String(v[i][K['פעיל']] || 'כן').trim() };
  }
  return null;
}
function expired_(u) {
  if (!u.until) return false;
  const end = new Date(u.until); end.setHours(23, 59, 59);
  return end < new Date();
}
function publicUser_(u) {
  const daysLeft = u.until ? Math.max(0, Math.ceil((new Date(u.until).setHours(23, 59, 59) - Date.now()) / 86400000)) : null;
  return { email: u.email, name: u.name, school: u.school, plan: u.plan, until: u.until ? fmtD_(u.until) : '', daysLeft };
}
function sendLoginCode_(email) {
  const code = String(Math.floor(100000 + Math.random() * 900000));
  CacheService.getScriptCache().put('otp:' + email, JSON.stringify({ code, tries: 0 }), CODE_MINUTES * 60);
  sendCodeMail_(email, code);
}
function sendCodeMail_(email, code, note) {
  const html = `<div dir="rtl" style="font-family:Arial,sans-serif;max-width:480px;margin:auto;color:#1f2a37">
    <div style="background:#1d5fa8;color:#fff;padding:14px 18px;border-radius:12px 12px 0 0;font-size:19px;font-weight:bold">🔑 ${APP_NAME}</div>
    <div style="border:1px solid #e3e7ed;border-top:0;border-radius:0 0 12px 12px;padding:20px;text-align:center">
    <p style="font-size:16px;margin:0 0 10px">קוד הכניסה שלך${note ? ' (' + note + ')' : ''}:</p>
    <div style="font-size:36px;font-weight:bold;letter-spacing:8px;direction:ltr;background:#e7f0fb;color:#1d5fa8;border-radius:10px;padding:12px">${code}</div>
    <p style="color:#6b7280;font-size:14px;margin-top:14px">הקוד תקף ל-${CODE_MINUTES} דקות. אם לא ביקשת להיכנס – אפשר להתעלם מהמייל.</p></div></div>`;
  MailApp.sendEmail({ to: email, subject: 'קוד הכניסה שלך: ' + code + ' – ' + APP_NAME, htmlBody: html, body: 'קוד הכניסה שלך: ' + code + ' (תקף ל-' + CODE_MINUTES + ' דקות)', name: APP_NAME });
}
function notifyOwner_(subject, body) {
  try { MailApp.sendEmail(Session.getEffectiveUser().getEmail(), '🔔 ' + subject, body, { name: APP_NAME }); } catch (e) {}
}
function makeToken_(email) {
  const exp = Date.now() + SESSION_DAYS * 86400000, payload = email + '|' + exp;
  return Utilities.base64EncodeWebSafe(payload) + '.' + sign_(payload);
}
function checkToken_(t) {
  if (!t || String(t).indexOf('.') < 0) return null;
  const parts = String(t).split('.');
  let payload;
  try { payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString(); } catch (e) { return null; }
  if (sign_(payload) !== parts[1]) return null;
  const [email, exp] = payload.split('|');
  return +exp > Date.now() ? email : null;
}
function sign_(s) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(s, PropertiesService.getScriptProperties().getProperty('SECRET'))).replace(/=+$/, '');
}
function tooMany_(key, max) {
  const c = CacheService.getScriptCache(), k = 'rate:' + key, n = +(c.get(k) || 0);
  if (n >= max) return true;
  c.put(k, String(n + 1), 3600); return false;
}
function fmtD_(d) { return Utilities.formatDate(new Date(d), Session.getScriptTimeZone() || 'Asia/Jerusalem', 'dd/MM/yyyy'); }
function log_(a, email, info) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH_LOG);
  if (sh) sh.appendRow([new Date(), a, email, info]);
}
