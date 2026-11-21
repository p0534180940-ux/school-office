/**
 * מערכת משימות ותזכורות למזכיר
 * -----------------------------------------
 * רץ בתוך Google Apps Script שמחובר לגיליון Google Sheets.
 * התקנה: ראו tasks-setup.html באתר (או README.md).
 *
 * גיליונות:
 *   משימות  – כל המשימות (שורה למשימה)
 *   הגדרות  – שעות עבודה, זמני תזכורות וכו' (משנים שם, בלי קוד)
 *   חגים    – תאריכים שבהם לא נשלחות תזכורות
 *   יומן    – רישום של כל מייל שנשלח
 */

// ======================= קבועים =======================
const SH_TASKS = 'משימות', SH_SET = 'הגדרות', SH_HOL = 'חגים', SH_LOG = 'יומן';
const COLS = ['מזהה', 'כותרת', 'פרטים', 'תאריך', 'שעה', 'עדיפות', 'סוג', 'סטטוס', 'חזרה', 'מקור',
              'תזכורת אחרונה', 'מספר תזכורות', 'בוצע בתאריך', 'נדחה עד', 'נוצר'];
const C = {}; COLS.forEach((n, i) => C[n] = i);
const PRIORITIES = ['דחוף', 'גבוה', 'רגיל', 'נמוך'];
const STATUSES = ['פתוח', 'בוצע', 'נדחה', 'בוטל'];
const REPEATS = ['', 'יומי', 'שבועי', 'חודשי'];
const DEFAULT_SETTINGS = [
  ['מייל לתזכורות', '', 'ריק = המייל של בעל החשבון'],
  ['ימי עבודה', 'א,ב,ג,ד,ה', 'אותיות הימים, מופרדות בפסיק (ו = שישי)'],
  ['שעת התחלה', '07:30', ''],
  ['שעת סיום', '16:00', ''],
  ['שעת סיום ביום שישי', '12:00', 'רלוונטי רק אם ו נמצא בימי העבודה'],
  ['מייל בוקר', 'כן', 'כן / לא'],
  ['שעת מייל בוקר', '07:45', ''],
  ['דקות לפני סוף היום לסיכום', '15', 'הסיכום היומי נשלח X דקות לפני שעת הסיום'],
  ['העתק סיכום אל', '', 'מייל נוסף לסיכום היומי (למשל המנהל). ריק = אין'],
  ['תזכורת מראש (דקות)', '30', 'כמה דקות לפני משימה עם שעה'],
  ['חזרה – דחוף (דקות)', '30', '0 = בלי תזכורת חוזרת'],
  ['חזרה – גבוה (דקות)', '60', ''],
  ['חזרה – רגיל (דקות)', '120', 'רק אחרי שהמועד עבר'],
  ['חזרה – נמוך (דקות)', '0', ''],
  ['מקסימום מיילי תזכורת בשעה', '4', ''],
  ['סדר סוגים', 'הורים,תשלומים,טלפון,מייל,מסמכים,כללי', 'משפיע על הסדר כשהעדיפות והשעה זהות'],
  ['תווית Gmail לקליטה', 'משימה', 'מייל עם התווית הזו הופך למשימה'],
  ['הסלמה למנהל – מייל', '', 'ריק = כבוי'],
  ['הסלמה אחרי (דקות)', '120', 'משימה דחופה שלא בוצעה'],
];
const DAY_LETTERS = ['א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ש'];

// ======================= התקנה =======================
/** מריצים פעם אחת: יוצר גיליונות, הגדרות וטריגר של 10 דקות */
function setup() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SH_TASKS);
  if (!sh) {
    sh = ss.getSheets()[0].getName().match(/^(Sheet|גיליון)\d*$/) && ss.getSheets()[0].getLastRow() === 0
      ? ss.getSheets()[0].setName(SH_TASKS) : ss.insertSheet(SH_TASKS);
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold').setBackground('#1d5fa8').setFontColor('#fff');
    sh.setFrozenRows(1); sh.setRightToLeft(true);
    sh.setColumnWidth(C['כותרת'] + 1, 260); sh.setColumnWidth(C['פרטים'] + 1, 220);
    const rule = (list) => SpreadsheetApp.newDataValidation().requireValueInList(list, true).build();
    sh.getRange(2, C['עדיפות'] + 1, 999).setDataValidation(rule(PRIORITIES));
    sh.getRange(2, C['סטטוס'] + 1, 999).setDataValidation(rule(STATUSES));
    sh.getRange(2, C['חזרה'] + 1, 999).setDataValidation(rule(REPEATS.filter(Boolean)));
    sh.getRange(2, C['תאריך'] + 1, 999).setNumberFormat('dd/MM/yyyy');
    sh.getRange(2, C['שעה'] + 1, 999).setNumberFormat('@');
    sh.getRange(2, C['תזכורת אחרונה'] + 1, 999, 1).setNumberFormat('dd/MM HH:mm');
    sh.getRange(2, C['בוצע בתאריך'] + 1, 999, 2).setNumberFormat('dd/MM HH:mm');
    sh.getRange(2, C['נוצר'] + 1, 999).setNumberFormat('dd/MM HH:mm');
    // צבעי עדיפות וסטטוס
    const r = sh.getRange('A2:O1000');
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$H2="בוצע"').setFontColor('#9aa3ad').setRanges([r]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$F2="דחוף"').setBackground('#fde2e2').setRanges([r]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$F2="גבוה"').setBackground('#fdf0d9').setRanges([r]).build(),
    ]);
    // משימות לדוגמה
    const today = new Date();
    addTaskRow_(sh, { title: 'דוגמה: להתקשר להורים לגבי טופס רישום', date: today, time: '10:00', priority: 'גבוה', type: 'הורים' });
    addTaskRow_(sh, { title: 'דוגמה: להזמין ציוד משרדי', date: today, priority: 'רגיל', type: 'כללי' });
  }
  let st = ss.getSheetByName(SH_SET);
  if (!st) {
    st = ss.insertSheet(SH_SET); st.setRightToLeft(true);
    st.getRange(1, 1, 1, 3).setValues([['הגדרה', 'ערך', 'הסבר']]).setFontWeight('bold').setBackground('#1d5fa8').setFontColor('#fff');
    st.getRange(2, 1, DEFAULT_SETTINGS.length, 3).setValues(DEFAULT_SETTINGS);
    st.getRange(2, 2, DEFAULT_SETTINGS.length).setNumberFormat('@');
    st.setColumnWidth(1, 230); st.setColumnWidth(2, 260); st.setColumnWidth(3, 380);
  }
  let hl = ss.getSheetByName(SH_HOL);
  if (!hl) {
    hl = ss.insertSheet(SH_HOL); hl.setRightToLeft(true);
    hl.getRange(1, 1, 1, 2).setValues([['תאריך', 'שם']]).setFontWeight('bold');
    hl.getRange(2, 1, 999).setNumberFormat('dd/MM/yyyy');
  }
  let lg = ss.getSheetByName(SH_LOG);
  if (!lg) {
    lg = ss.insertSheet(SH_LOG); lg.setRightToLeft(true);
    lg.getRange(1, 1, 1, 4).setValues([['זמן', 'סוג', 'נושא', 'משימות']]).setFontWeight('bold');
  }
  installTrigger();
  // מפתח סודי לקישורים במייל
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('SECRET')) p.setProperty('SECRET', Utilities.getUuid());
  return 'ההתקנה הושלמה. התזכורות פעילות כל 10 דקות.';
}

function installTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'tick').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('tick').timeBased().everyMinutes(10).create();
}
function stopReminders() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'tick').forEach(t => ScriptApp.deleteTrigger(t));
  return 'התזכורות הופסקו.';
}
function remindersActive() { return ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'tick'); }

/** תפריט בגיליון */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('📋 משימות')
    .addItem('התקנה / הפעלת תזכורות', 'menuSetup')
    .addItem('הרצת בדיקה עכשיו', 'menuTick')
    .addItem('שליחת סיכום יומי עכשיו (בדיקה)', 'menuSummary')
    .addItem('הפסקת התזכורות', 'menuStop')
    .addToUi();
}
function menuSetup() { SpreadsheetApp.getUi().alert(setup()); }
function menuTick() { const r = tick(true); SpreadsheetApp.getUi().alert(r || 'הבדיקה הסתיימה.'); }
function menuSummary() { sendSummary_(true); SpreadsheetApp.getUi().alert('הסיכום נשלח למייל.'); }
function menuStop() { SpreadsheetApp.getUi().alert(stopReminders()); }

// ======================= הגדרות =======================
let SETTINGS_CACHE = null;
function settings_() {
  if (SETTINGS_CACHE) return SETTINGS_CACHE;
  const s = {};
  DEFAULT_SETTINGS.forEach(r => s[r[0]] = r[1]);
  const sh = SpreadsheetApp.getActive().getSheetByName(SH_SET);
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues().forEach(r => { if (r[0]) s[r[0].trim()] = String(r[1]).trim(); });
  const num = (k, d) => { const v = parseFloat(s[k]); return isNaN(v) ? d : v; };
  SETTINGS_CACHE = {
    email: s['מייל לתזכורות'] || Session.getEffectiveUser().getEmail(),
    days: String(s['ימי עבודה']).split(/[,\s]+/).map(x => DAY_LETTERS.indexOf(x.replace(/['׳]/g, ''))).filter(x => x >= 0),
    start: hm_(s['שעת התחלה'], 450), end: hm_(s['שעת סיום'], 960), friEnd: hm_(s['שעת סיום ביום שישי'], 720),
    morning: /^כן|yes|true|1/i.test(s['מייל בוקר']), morningAt: hm_(s['שעת מייל בוקר'], 465),
    summaryBefore: num('דקות לפני סוף היום לסיכום', 15), summaryCc: s['העתק סיכום אל'],
    lead: num('תזכורת מראש (דקות)', 30),
    repeat: { 'דחוף': num('חזרה – דחוף (דקות)', 30), 'גבוה': num('חזרה – גבוה (דקות)', 60), 'רגיל': num('חזרה – רגיל (דקות)', 120), 'נמוך': num('חזרה – נמוך (דקות)', 0) },
    maxPerHour: num('מקסימום מיילי תזכורת בשעה', 4),
    typeOrder: String(s['סדר סוגים']).split(/\s*,\s*/),
    label: s['תווית Gmail לקליטה'] || 'משימה',
    escalateTo: s['הסלמה למנהל – מייל'], escalateAfter: num('הסלמה אחרי (דקות)', 120),
  };
  return SETTINGS_CACHE;
}
/** "07:30" → 450 דקות מחצות */
function hm_(v, def) {
  if (v instanceof Date) return v.getHours() * 60 + v.getMinutes();
  const m = String(v || '').match(/(\d{1,2})[:.](\d{2})/);
  return m ? (+m[1]) * 60 + (+m[2]) : def;
}
const pad2_ = n => (n < 10 ? '0' : '') + n;
const fmtHM_ = m => pad2_(Math.floor(m / 60)) + ':' + pad2_(m % 60);
const tz_ = () => Session.getScriptTimeZone() || 'Asia/Jerusalem';
const dayKey_ = d => Utilities.formatDate(d, tz_(), 'yyyy-MM-dd');
const minutesOf_ = d => +Utilities.formatDate(d, tz_(), 'H') * 60 + +Utilities.formatDate(d, tz_(), 'm');
const dow_ = d => +Utilities.formatDate(d, tz_(), 'u') % 7; // 0 = ראשון

function isHoliday_(d) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH_HOL);
  if (!sh || sh.getLastRow() < 2) return false;
  const k = dayKey_(d);
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().some(r => r[0] instanceof Date && dayKey_(r[0]) === k);
}
function workWindow_(d) {
  const s = settings_(), dw = dow_(d);
  if (s.days.indexOf(dw) < 0 || isHoliday_(d)) return null;
  return { start: s.start, end: dw === 5 ? Math.min(s.end, s.friEnd) : s.end };
}

// ======================= קריאה וכתיבה של משימות =======================
function tasksSheet_() { return SpreadsheetApp.getActive().getSheetByName(SH_TASKS); }
function readTasks_() {
  const sh = tasksSheet_(), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const vals = sh.getRange(2, 1, n, COLS.length).getValues();
  const disp = sh.getRange(2, C['שעה'] + 1, n, 1).getDisplayValues();
  const fixes = [];
  const list = vals.map((r, i) => {
    const t = {
      row: i + 2, id: String(r[C['מזהה']] || ''), title: String(r[C['כותרת']] || '').trim(), details: String(r[C['פרטים']] || ''),
      date: r[C['תאריך']] instanceof Date ? r[C['תאריך']] : null, timeTxt: disp[i][0],
      time: r[C['שעה']] === '' ? null : hm_(r[C['שעה']] instanceof Date ? r[C['שעה']] : disp[i][0], null),
      priority: PRIORITIES.indexOf(r[C['עדיפות']]) >= 0 ? r[C['עדיפות']] : 'רגיל',
      type: String(r[C['סוג']] || 'כללי'), status: STATUSES.indexOf(r[C['סטטוס']]) >= 0 ? r[C['סטטוס']] : 'פתוח',
      repeat: String(r[C['חזרה']] || ''), source: String(r[C['מקור']] || ''),
      lastReminder: r[C['תזכורת אחרונה']] instanceof Date ? r[C['תזכורת אחרונה']] : null,
      count: +r[C['מספר תזכורות']] || 0, doneAt: r[C['בוצע בתאריך']] instanceof Date ? r[C['בוצע בתאריך']] : null,
      snoozeUntil: r[C['נדחה עד']] instanceof Date ? r[C['נדחה עד']] : null,
    };
    // השלמת שורות שנכתבו ידנית בגיליון
    if (t.title && !t.id) { t.id = newId_(); fixes.push([t.row, C['מזהה'], t.id]); }
    if (t.title && !r[C['סטטוס']]) fixes.push([t.row, C['סטטוס'], 'פתוח']);
    if (t.title && !r[C['עדיפות']]) fixes.push([t.row, C['עדיפות'], 'רגיל']);
    if (t.title && !r[C['נוצר']]) fixes.push([t.row, C['נוצר'], new Date()]);
    if (t.title && !t.source) fixes.push([t.row, C['מקור'], 'גיליון']);
    if (t.status === 'בוצע' && !t.doneAt && t.title) { t.doneAt = new Date(); fixes.push([t.row, C['בוצע בתאריך'], t.doneAt]); t.needRepeat = true; }
    return t;
  }).filter(t => t.title);
  fixes.forEach(f => sh.getRange(f[0], f[1] + 1).setValue(f[2]));
  list.filter(t => t.needRepeat).forEach(t => createNextRepeat_(t));
  return list;
}
function newId_() { return 'M' + Utilities.getUuid().slice(0, 8); }
function addTaskRow_(sh, o) {
  const id = newId_(), row = new Array(COLS.length).fill('');
  row[C['מזהה']] = id; row[C['כותרת']] = o.title; row[C['פרטים']] = o.details || '';
  row[C['תאריך']] = o.date || new Date(); row[C['שעה']] = o.time || '';
  row[C['עדיפות']] = o.priority || 'רגיל'; row[C['סוג']] = o.type || 'כללי'; row[C['סטטוס']] = 'פתוח';
  row[C['חזרה']] = o.repeat || ''; row[C['מקור']] = o.source || 'ידני'; row[C['מספר תזכורות']] = 0; row[C['נוצר']] = new Date();
  sh.appendRow(row);
  // התאריך נשמר כתאריך בלבד, השעה כטקסט
  const r = sh.getLastRow();
  sh.getRange(r, C['תאריך'] + 1).setValue(o.date ? new Date(o.date) : new Date()).setNumberFormat('dd/MM/yyyy');
  sh.getRange(r, C['שעה'] + 1).setNumberFormat('@').setValue(o.time || '');
  return id;
}
function setCells_(row, map) { const sh = tasksSheet_(); Object.keys(map).forEach(k => sh.getRange(row, C[k] + 1).setValue(map[k])); }
function findTask_(id) { return readTasks_().find(t => t.id === id); }

function dueAt_(t) { // מועד המשימה כ-Date (אם יש שעה)
  if (t.time == null) return null;
  const d = new Date(t.date || new Date()); const k = dayKey_(d).split('-');
  return new Date(+k[0], +k[1] - 1, +k[2], Math.floor(t.time / 60), t.time % 60);
}
function isToday_(t, now) { return !t.date || dayKey_(t.date) === dayKey_(now); }
function isPastDay_(t, now) { return t.date && dayKey_(t.date) < dayKey_(now); }
function isOpen_(t) { return t.status === 'פתוח' || t.status === 'נדחה'; }

/** ציון למיון: קודם באיחור, אחר כך מתקרבת, עדיפות, שעה, סוג */
function sortTasks_(list, now) {
  const s = settings_();
  const key = t => {
    const due = dueAt_(t), late = (due && due <= now) || isPastDay_(t, now) ? 0 : 1;
    const soon = due && due > now && (due - now) <= 30 * 60000 ? 0 : 1;
    const ti = s.typeOrder.indexOf(t.type);
    return [late, soon, PRIORITIES.indexOf(t.priority), t.time == null ? 9999 : t.time, ti < 0 ? 99 : ti];
  };
  return list.slice().sort((a, b) => { const x = key(a), y = key(b); for (let i = 0; i < x.length; i++) if (x[i] !== y[i]) return x[i] - y[i]; return 0; });
}

// ======================= הלב: בדיקה כל 10 דקות =======================
function tick(manual) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return 'בדיקה אחרת כבר רצה';
  try {
    SETTINGS_CACHE = null;
    const now = new Date(), s = settings_(), win = workWindow_(now), m = minutesOf_(now);
    try { importFromGmail_(); } catch (e) { console.warn('Gmail import: ' + e); }
    const tasks = readTasks_();
    if (!win) return 'היום אינו יום עבודה – לא נשלחו תזכורות.';
    const props = PropertiesService.getScriptProperties(), today = dayKey_(now);
    const out = [];
    // מייל בוקר
    if (s.morning && m >= s.morningAt && m < win.end && props.getProperty('MORNING') !== today) {
      sendMorning_(tasks, now); props.setProperty('MORNING', today); out.push('נשלח מייל בוקר');
      // מייל הבוקר כבר מזכיר את הכול – הדחופות והגבוהות נספרות כאילו קיבלו תזכורת עכשיו
      tasks.filter(t => isOpen_(t) && (isToday_(t, now) || isPastDay_(t, now)) && (t.priority === 'דחוף' || t.priority === 'גבוה') && t.time == null)
        .forEach(t => setCells_(t.row, { 'תזכורת אחרונה': now }));
      return out.join(' · ');
    }
    // תזכורות (אחרי מייל הבוקר, אם הוא פעיל)
    const morningPending = s.morning && props.getProperty('MORNING') !== today && m < s.morningAt;
    if (m >= win.start && m < win.end - s.summaryBefore && !(morningPending && !tasks.some(t => isOpen_(t) && dueAt_(t) && isToday_(t, now) && (dueAt_(t) - now) / 60000 <= s.lead))) {
      const r = sendReminders_(tasks, now); if (r) out.push(r);
    }
    // סיכום יומי
    if (m >= win.end - s.summaryBefore && props.getProperty('SUMMARY') !== today) {
      sendSummary_(); props.setProperty('SUMMARY', today); out.push('נשלח סיכום יומי');
    }
    return out.join(' · ') || 'אין כרגע משימות שצריכות תזכורת.';
  } finally { lock.releaseLock(); }
}

function needsReminder_(t, now) {
  const s = settings_();
  if (!isOpen_(t)) return false;
  if (t.snoozeUntil && t.snoozeUntil > now) return false;
  if (t.date && dayKey_(t.date) > dayKey_(now)) return false; // משימה עתידית
  const interval = s.repeat[t.priority] || 0, since = t.lastReminder ? (now - t.lastReminder) / 60000 : Infinity;
  const due = dueAt_(t);
  if (due && isToday_(t, now)) {
    if (now < due) return !t.lastReminder && (due - now) / 60000 <= s.lead;   // תזכורת מראש – פעם אחת
    if (!t.lastReminder || t.lastReminder < due) return true;                  // הגיע המועד
    return interval > 0 && since >= interval;                                  // באיחור – חוזר לפי עדיפות
  }
  // בלי שעה (או מיום קודם)
  if (isPastDay_(t, now)) return interval > 0 ? since >= interval : !t.lastReminder || dayKey_(t.lastReminder) !== dayKey_(now);
  if (t.priority === 'דחוף' || t.priority === 'גבוה') return interval > 0 && since >= interval;
  return false; // רגיל / נמוך בלי שעה – רק בסיכומים
}

function sendReminders_(tasks, now) {
  const s = settings_();
  let due = tasks.filter(t => needsReminder_(t, now));
  if (!due.length) return '';
  // איחוד: משימות שהיו מגיעות לתזכורת ב-20 הדקות הקרובות נכנסות כבר למייל הזה (פחות מיילים)
  const soon = new Date(+now + 20 * 60000);
  due = sortTasks_(due.concat(tasks.filter(t => due.indexOf(t) < 0 && needsReminder_(t, soon) && !(dueAt_(t) && dueAt_(t) > now && (dueAt_(t) - now) / 60000 > settings_().lead))), now);
  // תקרת מיילים בשעה
  const props = PropertiesService.getScriptProperties();
  const sent = JSON.parse(props.getProperty('SENT') || '[]').filter(x => now - x < 3600000);
  if (sent.length >= s.maxPerHour) return 'הגעתי לתקרת המיילים לשעה – התזכורת תישלח בסבב הבא';
  const subject = '⏰ ' + (due.length === 1 ? 'תזכורת: ' + due[0].title : due.length + ' משימות לטיפול');
  const html = mailFrame_('תזכורת למשימות', `<p style="margin:0 0 12px">${due.length === 1 ? 'משימה אחת מחכה לך:' : 'אלה המשימות שמחכות לך, לפי סדר חשיבות:'}</p>`
    + due.map(t => taskRow_(t, now, true)).join('')
    + `<p style="color:#6b7280;font-size:13px;margin-top:14px">משימות אחרות להיום: ${tasks.filter(t => isOpen_(t) && isToday_(t, now) && due.indexOf(t) < 0).length}. <a href="${appUrl_()}">למסך המשימות</a></p>`);
  GmailApp.sendEmail(s.email, subject, plain_(due), { htmlBody: html, name: 'מערכת המשימות' });
  sent.push(+now); props.setProperty('SENT', JSON.stringify(sent));
  due.forEach(t => setCells_(t.row, { 'תזכורת אחרונה': now, 'מספר תזכורות': t.count + 1 }));
  log_('תזכורת', subject, due.map(t => t.title).join(' | '));
  // הסלמה
  if (s.escalateTo) {
    const esc = due.filter(t => t.priority === 'דחוף' && dueAt_(t) && (now - dueAt_(t)) / 60000 >= s.escalateAfter && !PropertiesService.getScriptProperties().getProperty('ESC_' + t.id));
    if (esc.length) {
      GmailApp.sendEmail(s.escalateTo, '⚠ משימות דחופות שעדיין פתוחות', plain_(esc), { htmlBody: mailFrame_('משימות דחופות פתוחות', esc.map(t => taskRow_(t, now, false)).join('')), name: 'מערכת המשימות' });
      esc.forEach(t => PropertiesService.getScriptProperties().setProperty('ESC_' + t.id, '1'));
      log_('הסלמה', 'משימות דחופות', esc.map(t => t.title).join(' | '));
    }
  }
  return 'נשלחה תזכורת על ' + due.length + ' משימות';
}

function sendMorning_(tasks, now) {
  const s = settings_();
  const list = sortTasks_(tasks.filter(t => isOpen_(t) && (isToday_(t, now) || isPastDay_(t, now))), now);
  const subject = '☀️ בוקר טוב – ' + list.length + ' משימות להיום';
  const html = mailFrame_('המשימות של היום', list.length
    ? list.map(t => taskRow_(t, now, true)).join('')
    : '<p>אין משימות פתוחות להיום. יום טוב! 🌤️</p>');
  GmailApp.sendEmail(s.email, subject, plain_(list), { htmlBody: html, name: 'מערכת המשימות' });
  log_('בוקר', subject, list.length);
}

function sendSummary_(manual) {
  SETTINGS_CACHE = null;
  const s = settings_(), now = new Date(), tasks = readTasks_();
  const today = dayKey_(now);
  const done = tasks.filter(t => t.status === 'בוצע' && t.doneAt && dayKey_(t.doneAt) === today).sort((a, b) => a.doneAt - b.doneAt);
  const open = sortTasks_(tasks.filter(t => isOpen_(t) && (isToday_(t, now) || isPastDay_(t, now))), now);
  const tomorrow = nextWorkday_(now);
  const tmr = sortTasks_(tasks.filter(t => isOpen_(t) && t.date && dayKey_(t.date) === dayKey_(tomorrow)), tomorrow);
  const total = done.length + open.length;
  const tm = d => Utilities.formatDate(d, tz_(), 'HH:mm');
  const sec = (title, color, body) => `<h3 style="color:${color};font-size:17px;margin:20px 0 8px">${title}</h3>${body}`;
  const html = mailFrame_('סיכום יום העבודה',
    `<div style="background:#e7f0fb;border-radius:10px;padding:14px;font-size:18px;text-align:center">היום בוצעו <b>${done.length}</b> מתוך <b>${total}</b> משימות${open.length ? '' : ' – הכול בוצע! 🎉'}</div>`
    + sec('✔ מה בוצע היום', '#15803d', done.length ? done.map(t => `<div style="padding:5px 0;border-bottom:1px dashed #e3e7ed;color:#4b5563">✔ ${esc_(t.title)} <span style="color:#9aa3ad">(${tm(t.doneAt)})</span></div>`).join('') : '<p style="color:#6b7280">לא סומנו משימות כבוצעות היום.</p>')
    + sec('⏳ מה נשאר פתוח', '#b91c1c', open.length ? open.map(t => taskRow_(t, now, true)).join('')
      + (appUrl_() ? `<p style="margin-top:12px"><a href="${actionUrl_('alltomorrow', 'all')}" style="background:#1d5fa8;color:#fff;padding:8px 14px;border-radius:8px;text-decoration:none;font-weight:bold">↪ העבר את כל הפתוחות ל${dayName_(tomorrow)}</a></p>` : '') : '<p>אין משימות פתוחות. 👏</p>')
    + sec('📅 מה מחכה ב' + dayName_(tomorrow), '#1d5fa8', tmr.length ? tmr.map(t => `<div style="padding:4px 0">• ${esc_(t.title)}${t.time != null ? ' <span style="color:#6b7280">(' + fmtHM_(t.time) + ')</span>' : ''}</div>`).join('') : '<p style="color:#6b7280">עדיין אין משימות.</p>'));
  const subject = `📊 סיכום יומי – בוצעו ${done.length} מתוך ${total}`;
  const opt = { htmlBody: html, name: 'מערכת המשימות' };
  if (s.summaryCc) opt.cc = s.summaryCc;
  GmailApp.sendEmail(s.email, subject, `בוצעו ${done.length} מתוך ${total}. פתוחות: ${open.map(t => t.title).join(', ')}`, opt);
  log_('סיכום', subject, '');
}
function nextWorkday_(d) {
  const x = new Date(d);
  for (let i = 0; i < 10; i++) { x.setDate(x.getDate() + 1); if (workWindow_(x)) return x; }
  x.setDate(new Date(d).getDate() + 1); return x;
}
function dayName_(d) {
  if (dayKey_(d) === dayKey_(new Date(Date.now() + 86400000))) return 'מחר';
  return 'יום ' + ['ראשון', 'שני', 'שלישי', 'רביעי', 'חמישי', 'שישי', 'שבת'][dow_(d)];
}

// ======================= קליטה מ-Gmail =======================
function importFromGmail_() {
  const s = settings_(), lbl = GmailApp.getUserLabelByName(s.label);
  if (!lbl) { GmailApp.createLabel(s.label); GmailApp.createLabel(s.label + '/נקלט'); return; }
  const doneLbl = GmailApp.getUserLabelByName(s.label + '/נקלט') || GmailApp.createLabel(s.label + '/נקלט');
  const threads = lbl.getThreads(0, 20);
  if (!threads.length) return;
  const sh = tasksSheet_();
  threads.forEach(th => {
    const msg = th.getMessages()[th.getMessageCount() - 1], subj = th.getFirstMessageSubject() || '(ללא נושא)';
    const from = msg.getFrom().replace(/<.*>/, '').replace(/"/g, '').trim();
    addTaskRow_(sh, {
      title: subj.replace(/^(re|fw|fwd)\s*:\s*/ig, ''), details: 'מאת: ' + from + '\n' + msg.getPlainBody().slice(0, 300) + '\nhttps://mail.google.com/mail/u/0/#all/' + th.getId(),
      priority: /דחוף|urgent/i.test(subj) ? 'דחוף' : 'רגיל', type: 'מייל', source: 'מייל',
    });
    th.removeLabel(lbl); th.addLabel(doneLbl);
    log_('קליטה', subj, 'נוצרה משימה ממייל');
  });
}

// ======================= פעולות: בוצע / דחייה / מחר =======================
function markDone_(id) {
  const t = findTask_(id); if (!t) return null;
  if (t.status !== 'בוצע') { setCells_(t.row, { 'סטטוס': 'בוצע', 'בוצע בתאריך': new Date(), 'נדחה עד': '' }); createNextRepeat_(t); }
  return t;
}
function createNextRepeat_(t) {
  if (!t.repeat) return;
  const base = new Date(t.date || new Date()), next = new Date(base);
  if (t.repeat === 'יומי') { do { next.setDate(next.getDate() + 1); } while (!workWindow_(next) && next - base < 15 * 86400000); }
  else if (t.repeat === 'שבועי') next.setDate(next.getDate() + 7);
  else if (t.repeat === 'חודשי') next.setMonth(next.getMonth() + 1);
  else return;
  addTaskRow_(tasksSheet_(), { title: t.title, details: t.details, date: next, time: t.time != null ? fmtHM_(t.time) : '', priority: t.priority, type: t.type, repeat: t.repeat, source: 'חוזרת' });
}
function snooze_(id, minutes) {
  const t = findTask_(id); if (!t) return null;
  const until = new Date(Date.now() + minutes * 60000);
  setCells_(t.row, { 'סטטוס': 'נדחה', 'נדחה עד': until });
  return t;
}
function toTomorrow_(id) {
  const t = findTask_(id); if (!t) return null;
  const d = nextWorkday_(new Date());
  tasksSheet_().getRange(t.row, C['תאריך'] + 1).setValue(d);
  setCells_(t.row, { 'סטטוס': 'פתוח', 'נדחה עד': '', 'תזכורת אחרונה': '' });
  return t;
}

// ======================= Web App: קישורים ומסך ניהול =======================
function appUrl_() { try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; } }
function sig_(id) {
  const secret = PropertiesService.getScriptProperties().getProperty('SECRET') || '';
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(id, secret)).slice(0, 16);
}
function actionUrl_(action, id) { return appUrl_() + '?a=' + action + '&id=' + encodeURIComponent(id) + '&s=' + sig_(action + ':' + id); }

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.a) {
    if (p.s !== sig_(p.a + ':' + p.id)) return page_('הקישור לא תקין', 'נסו לפתוח את מסך המשימות.', '#b91c1c');
    let t, msg;
    if (p.a === 'done') { t = markDone_(p.id); msg = t ? '✔ סומן כבוצע: ' + t.title : null; }
    else if (p.a === 'snooze') { t = snooze_(p.id, 60); msg = t ? '⏸ תזכורת חוזרת בעוד שעה: ' + t.title : null; }
    else if (p.a === 'tomorrow') { t = toTomorrow_(p.id); msg = t ? '↪ הועבר ל' + dayName_(nextWorkday_(new Date())) + ': ' + t.title : null; }
    else if (p.a === 'alltomorrow') {
      const now = new Date(), open = readTasks_().filter(x => isOpen_(x) && (isToday_(x, now) || isPastDay_(x, now)));
      open.forEach(x => toTomorrow_(x.id)); msg = '↪ ' + open.length + ' משימות הועברו ל' + dayName_(nextWorkday_(now)); t = true;
    }
    return t ? page_(msg, 'אפשר לסגור את החלון.', '#15803d') : page_('המשימה לא נמצאה', 'ייתכן שנמחקה מהגיליון.', '#b91c1c');
  }
  return HtmlService.createTemplateFromFile('App').evaluate()
    .setTitle('המשימות שלי').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function page_(title, sub, color) {
  return HtmlService.createHtmlOutput(`<div dir="rtl" style="font-family:Arial,sans-serif;text-align:center;padding:60px 20px">
    <div style="font-size:26px;font-weight:bold;color:${color}">${esc_(title)}</div><p style="color:#6b7280;font-size:18px">${esc_(sub)}</p>
    <p><a href="${appUrl_()}" style="color:#1d5fa8;font-size:18px">למסך המשימות</a></p></div>`)
    .setTitle('המשימות שלי').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- פונקציות שמסך הניהול (App.html) קורא להן ---
function apiList() {
  SETTINGS_CACHE = null;
  const now = new Date();
  const toObj = t => ({ id: t.id, title: t.title, details: t.details, date: t.date ? dayKey_(t.date) : '', time: t.time != null ? fmtHM_(t.time) : '',
    priority: t.priority, type: t.type, status: t.status, repeat: t.repeat, source: t.source,
    late: isOpen_(t) && ((dueAt_(t) && dueAt_(t) <= now && isToday_(t, now)) || isPastDay_(t, now)),
    doneAt: t.doneAt ? Utilities.formatDate(t.doneAt, tz_(), 'HH:mm') : '', doneToday: !!(t.doneAt && dayKey_(t.doneAt) === dayKey_(now)) });
  const all = readTasks_();
  return {
    today: sortTasks_(all.filter(t => isOpen_(t) && (isToday_(t, now) || isPastDay_(t, now))), now).map(toObj),
    future: all.filter(t => isOpen_(t) && t.date && dayKey_(t.date) > dayKey_(now)).sort((a, b) => a.date - b.date || (a.time || 0) - (b.time || 0)).slice(0, 50).map(toObj),
    done: all.filter(t => t.status === 'בוצע' && t.doneAt && dayKey_(t.doneAt) === dayKey_(now)).map(toObj),
    active: remindersActive(), email: settings_().email, sheetUrl: SpreadsheetApp.getActive().getUrl(),
    types: settings_().typeOrder, todayKey: dayKey_(now),
  };
}
function apiAdd(o) {
  if (!o || !String(o.title || '').trim()) throw new Error('חסרה כותרת');
  const d = o.date ? new Date(o.date + 'T12:00:00') : new Date();
  addTaskRow_(tasksSheet_(), { title: String(o.title).trim(), details: o.details || '', date: d, time: o.time || '', priority: o.priority, type: o.type, repeat: o.repeat, source: 'מסך ניהול' });
  return apiList();
}
function apiAction(action, id) {
  if (action === 'done') markDone_(id);
  else if (action === 'snooze') snooze_(id, 60);
  else if (action === 'tomorrow') toTomorrow_(id);
  else if (action === 'reopen') { const t = findTask_(id); if (t) setCells_(t.row, { 'סטטוס': 'פתוח', 'בוצע בתאריך': '' }); }
  else if (action === 'cancel') { const t = findTask_(id); if (t) setCells_(t.row, { 'סטטוס': 'בוטל' }); }
  return apiList();
}
function apiRunNow() { return tick(true); }
function apiSummaryNow() { sendSummary_(true); return 'הסיכום נשלח ל-' + settings_().email; }
function apiSetup() { return setup(); }
function apiStop() { return stopReminders(); }

// ======================= עיצוב המיילים =======================
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function mailFrame_(title, body) {
  return `<div dir="rtl" style="font-family:Arial,Helvetica,sans-serif;max-width:640px;margin:0 auto;color:#1f2a37;font-size:15px;line-height:1.6">
  <div style="background:#1d5fa8;color:#fff;padding:14px 18px;border-radius:12px 12px 0 0;font-size:19px;font-weight:bold">📋 ${esc_(title)}</div>
  <div style="border:1px solid #e3e7ed;border-top:0;border-radius:0 0 12px 12px;padding:16px 18px">${body}</div>
  <div style="text-align:center;color:#9aa3ad;font-size:12px;margin-top:8px">נשלח אוטומטית ממערכת המשימות · <a href="${appUrl_()}" style="color:#9aa3ad">מסך המשימות</a></div></div>`;
}
const PCOL = { 'דחוף': ['#fde2e2', '#a61b1b'], 'גבוה': ['#fdeccc', '#8a4f06'], 'רגיל': ['#e3edf9', '#1d5fa8'], 'נמוך': ['#eef1f5', '#4b5563'] };
function taskRow_(t, now, buttons) {
  const c = PCOL[t.priority] || PCOL['רגיל'], due = dueAt_(t);
  const late = (due && due <= now && isToday_(t, now)) || isPastDay_(t, now);
  const when = [t.time != null ? fmtHM_(t.time) : '', isPastDay_(t, now) ? 'מ-' + Utilities.formatDate(t.date, tz_(), 'dd/MM') : '', late ? '<b style="color:#b91c1c">באיחור</b>' : ''].filter(Boolean).join(' · ');
  const btn = (a, label, main) => `<a href="${actionUrl_(a, t.id)}" style="display:inline-block;margin:4px 0 0 6px;padding:5px 12px;border-radius:7px;text-decoration:none;font-size:13px;font-weight:bold;${main ? 'background:#15803d;color:#fff' : 'background:#eef1f5;color:#1f2a37'}">${label}</a>`;
  return `<div style="padding:10px 0;border-bottom:1px solid #eef1f5">
    <span style="background:${c[0]};color:${c[1]};border-radius:20px;padding:1px 9px;font-size:12px;font-weight:bold">${t.priority}</span>
    <b style="font-size:16px"> ${esc_(t.title)}</b> <span style="color:#6b7280;font-size:13px">${t.type ? '· ' + esc_(t.type) : ''}${when ? ' · ' + when : ''}</span>
    ${t.details ? `<div style="color:#6b7280;font-size:13px;margin-top:2px">${esc_(t.details.split('\n')[0]).slice(0, 140)}</div>` : ''}
    ${buttons && appUrl_() ? `<div>${btn('done', '✔ בוצע', true)}${btn('snooze', '⏸ עוד שעה')}${btn('tomorrow', '↪ ' + dayName_(nextWorkday_(now)))}</div>` : ''}</div>`;
}
function plain_(list) { return list.map(t => `[${t.priority}] ${t.title}${t.time != null ? ' (' + fmtHM_(t.time) + ')' : ''}`).join('\n'); }
function log_(kind, subject, info) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH_LOG);
  if (sh) sh.appendRow([new Date(), kind, subject, String(info)]);
}
