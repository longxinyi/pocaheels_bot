/**
 * Poca Heels Mini App backend. Paste into Extensions > Apps Script in the team's Google Sheet.
 *
 * No secrets live in this file. The bot token goes in Project Settings > Script Properties as BOT_TOKEN.
 * Every request must carry Telegram's signed initData, which is verified before anything is read or written.
 */
const TZ = 'Asia/Singapore';
const TABLES = {
  People: ['id', 'name', 'username', 'tg_name', 'role', 'pays', 'joined', 'left'],
  Classes: ['date', 'special', 'cancelled', 'start', 'end', 'location', 'kind', 'title'],
  Status: ['date', 'person_id', 'status', 'note', 'marked_at'],
  Hangouts: ['id', 'date', 'title', 'start', 'end', 'location', 'by', 'going', 'deleted'],
  Payments: ['month', 'person_id', 'state', 'amount', 'at', 'by'],
  Config: ['key', 'value'],
  Log: ['at', 'actor', 'action', 'detail'],
};
const DEFAULTS = {
  START_MONTH: '2026-12', // First month paid through this app. Earlier months count as already paid.
  RENT: '5',
  TRAINING: '25',
  LOW_COUNT: '5',
  DEFAULT_START: '20:00',
  DEFAULT_END: '23:00',
  DEFAULT_LOCATION: 'The Velvet Room',
  PAYNOW: '',
  GROUP_CHAT_ID: '',
  GROUP_TITLE: '',
  TOPIC_PAYMENTS: '',
  TOPIC_HANGOUTS: '',
  APP_LINK: '',
  GENERATED_THROUGH: '',
};
const MONTHS_AHEAD = 6;
const MINI_APP_LINK = 'https://t.me/pocaheels_bot?startapp';
const KINDS = { paid: 'Training', free: 'Extra training', performance: 'Performance', hangout: 'Crew hangout' };
function isPaid(c) { return !c.kind || c.kind === 'paid'; }
function isTraining(c) { return isPaid(c) || c.kind === 'free'; }
function itemName(c) { return c.title || KINDS[c.kind || 'paid']; }

// ---------- setup, run once by hand ----------
function setup() {
  resetCaches();
  clearFastCache();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  Object.keys(TABLES).forEach((name) => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.getRange(1, 1, sh.getMaxRows(), TABLES[name].length).setNumberFormat('@');
    sh.getRange(1, 1, 1, TABLES[name].length).setValues([TABLES[name]]);
  });
  const config = readTable('Config');
  Object.keys(DEFAULTS).forEach((k) => {
    if (!config.some((r) => r.key === k)) config.push({ key: k, value: DEFAULTS[k] });
  });
  writeTable('Config', config);
  ensureMondays();
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === 'tick') ScriptApp.deleteTrigger(t);
  });
  // Once a day: the payment messages go out in the morning.
  ScriptApp.newTrigger('tick').timeBased().atHour(10).everyDays(1).inTimezone(TZ).create();
}

// ---------- storage ----------
// Each request opens the Sheet once and reads each tab at most once, which is most of the save time.
let BOOK = null;
let TABLE_CACHE = {};
function resetCaches() { BOOK = null; TABLE_CACHE = {}; }
// Sheet reads take seconds, so tabs are also kept in Google's fast cache between requests.
// onEdit clears it when someone changes the Sheet by hand, and it expires after 6 hours anyway.
const CACHE_SECONDS = 21600;
function cacheKey(name) { return 'table_' + name; }
function fastGet(name) {
  try {
    const raw = CacheService.getScriptCache().get(cacheKey(name));
    return raw ? JSON.parse(raw) : null;
  } catch (err) {
    return null;
  }
}
function fastPut(name, rows) {
  try {
    CacheService.getScriptCache().put(cacheKey(name), JSON.stringify(rows), CACHE_SECONDS);
  } catch (err) {
    CacheService.getScriptCache().remove(cacheKey(name)); // Too big to cache: read from the Sheet instead.
  }
}
function clearFastCache() { CacheService.getScriptCache().removeAll(Object.keys(TABLES).map(cacheKey)); }
function onEdit() { clearFastCache(); }
function readTable(name) {
  if (!TABLE_CACHE[name]) {
    const cached = fastGet(name);
    if (cached) {
      TABLE_CACHE[name] = cached;
    } else {
      const values = bookOrActive().getSheetByName(name).getDataRange().getDisplayValues();
      const head = TABLES[name];
      TABLE_CACHE[name] = values.slice(1).filter((r) => r.join('') !== '').map((r) => {
        const o = {};
        head.forEach((h, i) => { o[h] = r[i] === undefined ? '' : String(r[i]); });
        return o;
      });
      fastPut(name, TABLE_CACHE[name]);
    }
  }
  return TABLE_CACHE[name].map((r) => Object.assign({}, r));
}
function bookOrActive() {
  if (BOOK) return BOOK;
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  BOOK = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  return BOOK;
}
function writeTable(name, rows) {
  const sh = bookOrActive().getSheetByName(name);
  const head = TABLES[name];
  const need = rows.length + 1;
  if (need > sh.getMaxRows()) {
    const start = sh.getMaxRows();
    sh.insertRowsAfter(start, need - start + 200);
    sh.getRange(start + 1, 1, sh.getMaxRows() - start, head.length).setNumberFormat('@');
  }
  const last = sh.getLastRow();
  if (last > rows.length + 1) sh.getRange(rows.length + 2, 1, last - rows.length - 1, head.length).clearContent();
  const cells = rows.map((r) => head.map((h) => (r[h] === undefined ? '' : String(r[h]))));
  if (rows.length) sh.getRange(2, 1, rows.length, head.length).setValues(cells);
  TABLE_CACHE[name] = cells.map((row) => {
    const o = {};
    head.forEach((h, i) => { o[h] = row[i]; });
    return o;
  });
  fastPut(name, TABLE_CACHE[name]);
}
function logAction(actor, action, detail) {
  bookOrActive().getSheetByName('Log').appendRow([nowStamp(), actor, action, JSON.stringify(detail || {})]);
}
function config() {
  const c = Object.assign({}, DEFAULTS);
  readTable('Config').forEach((r) => { c[r.key] = r.value; });
  return c;
}
function setConfig(key, value) {
  const rows = readTable('Config');
  const row = rows.find((r) => r.key === key);
  if (row) row.value = value; else rows.push({ key, value });
  writeTable('Config', rows);
}

// ---------- dates ----------
function today() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function nowStamp() { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss"); }
function nowHour() { return Number(Utilities.formatDate(new Date(), TZ, 'H')); }
function pad(n) { return String(n).padStart(2, '0'); }
function isoOf(d) { return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate()); }
function dateOf(s) { const p = s.split('-').map(Number); return new Date(Date.UTC(p[0], p[1] - 1, p[2] || 1)); }
function addMonths(m, n) { const d = dateOf(m + '-01'); d.setUTCMonth(d.getUTCMonth() + n); return isoOf(d).slice(0, 7); }
function lastDay(m) { const d = dateOf(addMonths(m, 1) + '-01'); d.setUTCDate(0); return d.getUTCDate(); }
function dueDate(m) { const prev = addMonths(m, -1); return prev + '-' + pad(Math.min(30, lastDay(prev))); }
function niceDate(s) {
  const d = dateOf(s);
  return ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][d.getUTCDay()] + ' ' + d.getUTCDate() + ' ' + ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][d.getUTCMonth()];
}
function niceTime(t) {
  const p = String(t).split(':').map(Number);
  if (p.length < 2 || isNaN(p[0])) return '';
  const h = p[0] % 12 || 12;
  return h + (p[1] ? '.' + pad(p[1]) : '') + (p[0] < 12 ? 'am' : 'pm');
}
function whenWhere(c) {
  const time = c.start ? niceTime(c.start) + (c.end ? ' to ' + niceTime(c.end) : '') : '';
  return [time, c.location].filter(Boolean).join(', ');
}
function monthName(m) {
  return ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'][Number(m.slice(5)) - 1];
}
function paymentMonth(cfg) {
  const t = today();
  const m = Number(t.slice(8)) >= 15 ? addMonths(t.slice(0, 7), 1) : t.slice(0, 7);
  return m < cfg.START_MONTH ? cfg.START_MONTH : m;
}
function ensureMondays() {
  const t = today();
  const cfg = config();
  const through = addMonths(t.slice(0, 7), MONTHS_AHEAD);
  const from = cfg.GENERATED_THROUGH ? addMonths(cfg.GENERATED_THROUGH, 1) : t.slice(0, 7);
  if (from > through) return;
  const classes = readTable('Classes');
  const known = {};
  classes.forEach((c) => { known[c.date] = true; });
  for (let m = from; m <= through; m = addMonths(m, 1)) {
    const n = lastDay(m);
    for (let i = 1; i <= n; i++) {
      const s = m + '-' + pad(i);
      if (s > t && dateOf(s).getUTCDay() === 1 && !known[s]) classes.push({ date: s, special: '', cancelled: '', start: cfg.DEFAULT_START, end: cfg.DEFAULT_END, location: cfg.DEFAULT_LOCATION, kind: 'paid', title: '' });
    }
  }
  classes.sort((a, b) => (a.date < b.date ? -1 : 1));
  writeTable('Classes', classes);
  setConfig('GENERATED_THROUGH', through);
}

// ---------- Telegram sign-in ----------
function hex(bytes) { return bytes.map((b) => ('0' + (b & 255).toString(16)).slice(-2)).join(''); }
function verifyInitData(initData) {
  const token = PropertiesService.getScriptProperties().getProperty('BOT_TOKEN');
  if (!token || !initData) throw new Error('Not signed in');
  const pairs = {};
  String(initData).split('&').forEach((kv) => {
    const i = kv.indexOf('=');
    pairs[decodeURIComponent(kv.slice(0, i))] = decodeURIComponent(kv.slice(i + 1).replace(/\+/g, ' '));
  });
  const hash = pairs.hash;
  if (!hash) throw new Error('Not signed in');
  delete pairs.hash;
  const check = Object.keys(pairs).sort().map((k) => k + '=' + pairs[k]).join('\n');
  const key = Utilities.computeHmacSha256Signature(Utilities.newBlob(token).getBytes(), Utilities.newBlob('WebAppData').getBytes());
  const sig = hex(Utilities.computeHmacSha256Signature(Utilities.newBlob(check).getBytes(), key));
  if (sig !== hash) throw new Error('Not signed in');
  if (Date.now() / 1000 - Number(pairs.auth_date) > 86400) throw new Error('Session expired, please reopen the app');
  return JSON.parse(pairs.user);
}

// ---------- money ----------
/**
 * The rules, all in one place:
 * - Each class costs RENT + TRAINING. An absence costs RENT only, so an absence marked after paying
 *   automatically leaves TRAINING as credit for the next payment.
 * - A cancelled class costs nothing, and so does anything that isn't paid training: extra training,
 *   performances and hangouts never affect what anyone owes.
 * - Months before START_MONTH were paid outside this app: they cost nothing, and an absence there
 *   gives TRAINING back as credit.
 */
function buildLedger(data, cfg) {
  const rent = Number(cfg.RENT), training = Number(cfg.TRAINING), full = rent + training;
  const statusBy = {};
  data.status.forEach((s) => { statusBy[s.person_id + '|' + s.date] = s; });
  const payBy = {};
  data.payments.forEach((p) => { payBy[p.person_id + '|' + p.month] = p; });
  function charge(person, c) {
    if (c.date < person.joined || (person.left && c.date >= person.left)) return 0;
    const month = c.date.slice(0, 7);
    const s = statusBy[person.id + '|' + c.date];
    const absent = s && s.status === 'absent';
    if (c.cancelled || !isPaid(c)) return 0;
    if (month < cfg.START_MONTH) return absent ? -training : 0;
    return absent ? rent : full;
  }
  function summary(person, month) {
    let before = 0, thisMonth = 0;
    const lines = [];
    let here = 0, away = 0;
    const offsets = [];
    const prev = addMonths(month, -1);
    data.classes.forEach((c) => {
      const m = c.date.slice(0, 7);
      if (m > month) return;
      const amt = charge(person, c);
      if (m < month) {
        before -= amt;
        const s = statusBy[person.id + '|' + c.date];
        if (m === prev && s && s.status === 'absent' && !c.cancelled && isPaid(c) && c.date >= person.joined) {
          const p = payBy[person.id + '|' + prev];
          if (prev < cfg.START_MONTH || (p && s.marked_at > p.at)) offsets.push(c.date);
        }
        return;
      }
      thisMonth += amt;
      if (c.cancelled || amt === 0) return;
      if (amt === rent) away += 1;
      else here += 1;
    });
    data.payments.forEach((p) => {
      if (p.person_id === person.id && p.month < month) before += Number(p.amount);
    });
    if (here) lines.push({ l: here + (here === 1 ? ' class' : ' classes') + ' at $' + full, v: '$' + here * full });
    if (away) lines.push({ l: away + (away === 1 ? ' absence' : ' absences') + ', rent only at $' + rent, v: '$' + away * rent });
    let left = before;
    offsets.forEach((d) => {
      lines.push({ l: 'Offset: absent ' + niceDate(d).slice(4), v: '-$' + training });
      left -= training;
    });
    if (left > 0) lines.push({ l: 'Credit from before', v: '-$' + left });
    if (left < 0) lines.push({ l: 'Owing from before', v: '$' + -left });
    const pay = payBy[person.id + '|' + month];
    const paid = pay ? Number(pay.amount) : 0;
    const toPay = Math.max(0, thisMonth - before);
    return {
      id: person.id, name: person.name, lines, toPay,
      state: pay ? pay.state : (toPay === 0 ? 'nothing' : 'none'),
      paid, after: before + paid - thisMonth,
    };
  }
  return { summary };
}
function loadData() {
  return {
    people: readTable('People'),
    classes: readTable('Classes').sort((a, b) => (a.date < b.date ? -1 : 1)),
    status: readTable('Status'),
    payments: readTable('Payments'),
  };
}
function billable(people) { return people.filter((p) => (p.role === 'member' || p.role === 'admin') && p.pays === 'yes'); }

// ---------- what each person is allowed to see ----------
function stateFor(me, month) {
  const cfg = config();
  const data = loadData();
  const pm = month || paymentMonth(cfg);
  const led = buildLedger(data, cfg);
  const t = today();
  const out = {
    me: { name: me.name, role: me.role, pays: me.pays === 'yes' },
    today: t, month: pm, due: dueDate(pm), paymentMonth: paymentMonth(cfg), startMonth: cfg.START_MONTH,
    rent: Number(cfg.RENT), training: Number(cfg.TRAINING), paynow: cfg.PAYNOW,
  };
  const mine = {};
  data.status.forEach((s) => { if (s.person_id === me.id) mine[s.date] = s; });
  const everyone = billable(data.people);
  function counts(c) {
    const active = everyone.filter((p) => c.date >= p.joined && !(p.left && c.date >= p.left));
    const away = active.filter((p) => data.status.some((s) => s.date === c.date && s.person_id === p.id && s.status === 'absent'));
    return { coming: active.length - away.length, total: active.length };
  }
  out.myClasses = data.classes.filter((c) => !c.cancelled && c.date >= t).map((c) => ({
    date: c.date, start: c.start, end: c.end, location: c.location, kind: c.kind || 'paid', title: c.title,
    status: mine[c.date] ? mine[c.date].status : 'coming', note: mine[c.date] ? mine[c.date].note : '',
    paid: c.date.slice(0, 7) < cfg.START_MONTH || data.payments.some((p) => p.person_id === me.id && p.month === c.date.slice(0, 7)),
  }));
  if (me.pays === 'yes') out.myPayment = led.summary(me, pm);
  const nameOf = {};
  data.people.forEach((p) => { nameOf[p.id] = p.name || p.tg_name; });
  out.hangouts = readTable('Hangouts').filter((h) => !h.deleted).map((h) => {
    const going = h.going ? h.going.split(',') : [];
    return {
      id: h.id, date: h.date, title: h.title, start: h.start, end: h.end, location: h.location,
      byName: nameOf[h.by] || '', going: going.map((id) => nameOf[id]).filter(Boolean),
      imGoing: going.indexOf(me.id) >= 0, canEdit: me.role === 'admin' || h.by === me.id,
    };
  }).sort((a, b) => (a.date < b.date ? -1 : 1));
  if (me.role !== 'admin') {
    out.calendar = data.classes.map((c) => Object.assign({
      date: c.date, cancelled: !!c.cancelled, start: c.start, end: c.end, location: c.location, kind: c.kind || 'paid', title: c.title,
    }, counts(c)));
  }
  if (me.role === 'admin') {
    const people = billable(data.people);
    out.admin = {
      lowCount: Number(cfg.LOW_COUNT),
      defaults: { start: cfg.DEFAULT_START, end: cfg.DEFAULT_END, location: cfg.DEFAULT_LOCATION },
      classes: data.classes.map((c) => {
        const rows = data.status.filter((s) => s.date === c.date);
        const active = people.filter((p) => c.date >= p.joined && !(p.left && c.date >= p.left));
        const absent = active.filter((p) => rows.some((s) => s.person_id === p.id && s.status === 'absent'));
        const late = active.filter((p) => rows.some((s) => s.person_id === p.id && s.status === 'late'));
        return {
          date: c.date, cancelled: !!c.cancelled, start: c.start, end: c.end, location: c.location, kind: c.kind || 'paid', title: c.title,
          coming: active.length - absent.length, total: active.length,
          absent: absent.map((p) => p.name),
          late: late.map((p) => {
            const n = rows.find((s) => s.person_id === p.id).note;
            return n ? p.name + ' (' + n + ')' : p.name;
          }),
          everyone: active.map((p) => {
            const s = rows.find((x) => x.person_id === p.id);
            return { id: p.id, name: p.name, status: s ? s.status : 'coming', note: s ? s.note : '' };
          }),
        };
      }),
      payments: people.map((p) => led.summary(p, pm)),
      people: data.people.filter((p) => p.role === 'member' || p.role === 'admin').map((p) => ({ id: p.id, name: p.name, role: p.role, pays: p.pays === 'yes' })),
      pending: data.people.filter((p) => p.role === 'pending').map((p) => ({ id: p.id, tgName: p.tg_name, username: p.username })),
      group: { id: cfg.GROUP_CHAT_ID, title: cfg.GROUP_TITLE, topics: { payments: cfg.TOPIC_PAYMENTS ? (cfg.TOPIC_PAYMENTS_NAME || 'Saved topic') : '', hangouts: cfg.TOPIC_HANGOUTS ? (cfg.TOPIC_HANGOUTS_NAME || 'Saved topic') : '' } },
      sheetUrl: 'https://docs.google.com/spreadsheets/d/' + bookOrActive().getId() + '/export?format=xlsx',
    };
  }
  return out;
}

// ---------- requests ----------
function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const tgUser = verifyInitData(req.initData);
    return reply(handle(tgUser, req.action, req.payload || {}));
  } catch (err) {
    return reply({ error: err.message });
  }
}
function doGet() { return reply({ ok: true }); }
function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function handle(tgUser, action, p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    resetCaches();
    ensureMondays();
    const id = String(tgUser.id);
    let people = readTable('People');
    let me = people.find((x) => x.id === id);
    if (!me) {
      me = { id, name: '', username: tgUser.username || '', tg_name: [tgUser.first_name, tgUser.last_name].filter(Boolean).join(' '), role: 'pending', pays: '', joined: '', left: '' };
      people.push(me);
      writeTable('People', people);
      logAction(id, 'requested_access', { name: me.tg_name });
    }
    if (me.role === 'pending') return { pending: true, name: me.tg_name };
    if (me.role === 'removed') return { removed: true };
    const fn = me.role === 'admin' ? Object.assign({}, MEMBER_ACTIONS, ADMIN_ACTIONS)[action] : MEMBER_ACTIONS[action];
    if (action !== 'state' && !fn) throw new Error('You do not have access to this');
    if (fn) {
      fn(me, p);
      if (action !== 'state') logAction(id, action, p);
    }
    me = readTable('People').find((x) => x.id === id);
    const out = stateFor(me, p.month);
    if (action === 'findGroups') out.groups = fetchGroups();
    return out;
  } finally {
    lock.releaseLock();
  }
}
function upsertStatus(personId, date, status, note) {
  const rows = readTable('Status');
  const row = rows.find((r) => r.person_id === personId && r.date === date);
  const stamp = nowStamp();
  if (row) {
    if (row.status !== status) row.marked_at = stamp;
    row.status = status; row.note = note || '';
  } else {
    rows.push({ date, person_id: personId, status, note: note || '', marked_at: stamp });
  }
  writeTable('Status', rows);
}
function validStatus(s) { if (['coming', 'late', 'absent'].indexOf(s) < 0) throw new Error('Unknown status'); }
function savePayment(month, personId, state, amount, by) {
  const rows = readTable('Payments').filter((r) => !(r.month === month && r.person_id === personId));
  if (state) rows.push({ month, person_id: personId, state, amount, at: nowStamp(), by });
  writeTable('Payments', rows);
}
function currentToPay(person, month) {
  return buildLedger(loadData(), config()).summary(person, month).toPay;
}
const MEMBER_ACTIONS = {
  state() {},
  setStatus(me, p) {
    validStatus(p.status);
    if (p.date < today()) throw new Error('This class has already passed. Please ask your teacher to change it.');
    const c = readTable('Classes').find((x) => x.date === p.date);
    if (!c || c.cancelled) throw new Error('This class is not on the schedule');
    upsertStatus(me.id, p.date, p.status, p.note);
  },
  claim(me, p) {
    if (readTable('Payments').some((r) => r.month === p.month && r.person_id === me.id)) return;
    savePayment(p.month, me.id, 'claimed', currentToPay(me, p.month), me.id);
  },
  saveHangout(me, p) {
    const title = String(p.title || '').trim().slice(0, 60);
    if (!title) throw new Error('Please enter a name');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(p.date)) throw new Error('Invalid date');
    ['start', 'end'].forEach((k) => {
      if (p[k] && !/^([01]\d|2[0-3]):[0-5]\d$/.test(p[k])) throw new Error('Please enter a valid time');
    });
    const location = String(p.location || '').trim().slice(0, 100);
    const rows = hangoutRows();
    if (p.id) {
      const h = editableHangout(rows, me, p.id);
      const before = h.title + hangoutWhen(h);
      h.title = title; h.date = p.date; h.start = p.start || ''; h.end = p.end || ''; h.location = location;
      writeTable('Hangouts', rows);
      if (p.notify && before !== h.title + hangoutWhen(h)) notifyGroup(esc('‼️ Hangout updated: ' + title + ', ' + hangoutWhen(h) + '.'), 'TOPIC_HANGOUTS');
      return;
    }
    if (p.date < today() && me.role !== 'admin') throw new Error('Please pick a date that has not passed');
    const h = { id: Date.now().toString(36) + Math.floor(Math.random() * 1e6).toString(36), date: p.date, title, start: p.start || '', end: p.end || '', location, by: me.id, going: me.id, deleted: '' };
    rows.push(h);
    writeTable('Hangouts', rows);
    if (p.notify) notifyGroup(esc('😆 New hangout from ' + (me.name || 'a teammate') + ': ' + title + ', ' + hangoutWhen(h) + '.'), 'TOPIC_HANGOUTS');
  },
  deleteHangout(me, p) {
    const rows = hangoutRows();
    const h = editableHangout(rows, me, p.id);
    h.deleted = 'yes';
    writeTable('Hangouts', rows);
    if (h.date >= today()) notifyGroup(esc('😭 Hangout cancelled: ' + h.title + ', ' + hangoutWhen(h) + '.'), 'TOPIC_HANGOUTS');
  },
  toggleGoing(me, p) {
    const rows = hangoutRows();
    const h = rows.find((x) => x.id === String(p.id) && !x.deleted);
    if (!h) throw new Error('This hangout no longer exists');
    const going = h.going ? h.going.split(',') : [];
    const i = going.indexOf(me.id);
    if (i >= 0) going.splice(i, 1); else going.push(me.id);
    h.going = going.join(',');
    writeTable('Hangouts', rows);
  },
  unclaim(me, p) {
    const row = readTable('Payments').find((r) => r.month === p.month && r.person_id === me.id);
    if (row && row.state === 'claimed') savePayment(p.month, me.id, null);
  },
};
function hangoutRows() { return readTable('Hangouts'); }
function editableHangout(rows, me, id) {
  const h = rows.find((x) => x.id === String(id) && !x.deleted);
  if (!h) throw new Error('This hangout no longer exists');
  if (me.role !== 'admin' && h.by !== me.id) throw new Error('Only the person who added this can change it');
  return h;
}
function hangoutWhen(h) {
  const w = whenWhere(h);
  return niceDate(h.date) + (w ? ', ' + w : '');
}
function personById(id) {
  const person = readTable('People').find((x) => x.id === String(id));
  if (!person) throw new Error('Person not found');
  return person;
}
const ADMIN_ACTIONS = {
  setStatusFor(me, p) {
    validStatus(p.status);
    upsertStatus(String(p.id), p.date, p.status, p.note);
  },
  saveClass(me, p) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(p.date)) throw new Error('Invalid date');
    ['start', 'end'].forEach((k) => {
      if (p[k] !== undefined && !/^([01]\d|2[0-3]):[0-5]\d$/.test(p[k])) throw new Error('Please enter a valid time');
    });
    if (p.location !== undefined && !String(p.location).trim()) throw new Error('Please enter a location');
    if (p.kind !== undefined && !KINDS[p.kind]) throw new Error('Unknown type');
    const cfg = config();
    const rows = readTable('Classes');
    let c = rows.find((x) => x.date === p.date);
    const isNew = !c;
    if (isNew) { c = { date: p.date, special: '', cancelled: '', start: cfg.DEFAULT_START, end: cfg.DEFAULT_END, location: cfg.DEFAULT_LOCATION, kind: 'free', title: '' }; rows.push(c); }
    const before = whenWhere(c);
    if (p.cancelled !== undefined) c.cancelled = p.cancelled ? 'yes' : '';
    if (p.start !== undefined) c.start = p.start;
    if (p.end !== undefined) c.end = p.end;
    if (p.location !== undefined) c.location = String(p.location).trim().slice(0, 100);
    if (p.kind !== undefined) c.kind = p.kind;
    if (p.title !== undefined) c.title = String(p.title).trim().slice(0, 60);
    rows.sort((a, b) => (a.date < b.date ? -1 : 1));
    writeTable('Classes', rows);
    (p.statuses || []).forEach((u) => {
      validStatus(u.status);
      const old = readTable('Status').find((r) => r.person_id === String(u.id) && r.date === p.date);
      upsertStatus(String(u.id), p.date, u.status, old ? old.note : '');
    });
    if (p.cancelled && p.notify) notifyGroup(esc('‼️ ' + niceDate(p.date) + ': class is cancelled.'));
    else if (p.notify && isNew) notifyGroup(esc((isTraining(c) ? '🫪 New class' : '💃 New event') + ' on ' + niceDate(p.date) + ': ' + itemName(c) + ', ' + whenWhere(c) + '.'));
    else if (p.notify && before !== whenWhere(c)) notifyGroup(esc('‼️ ' + niceDate(p.date) + ' ' + itemName(c) + ' is now ' + whenWhere(c) + '.'));
  },
  confirm(me, p) {
    const row = readTable('Payments').find((r) => r.month === p.month && r.person_id === String(p.id));
    if (row) savePayment(p.month, row.person_id, 'paid', row.amount, me.id);
  },
  markPaid(me, p) {
    const id = String(p.id);
    const amount = p.amount === undefined || p.amount === null ? currentToPay(personById(id), p.month) : Number(p.amount);
    if (isNaN(amount) || amount < 0) throw new Error('Please enter a valid amount');
    savePayment(p.month, id, 'paid', amount, me.id);
  },
  resetPay(me, p) { (p.ids || [p.id]).forEach((id) => savePayment(p.month, String(id), null)); },
  markAll(me, p) {
    const existing = readTable('Payments').filter((r) => r.month === p.month);
    billable(readTable('People')).forEach((person) => {
      const row = existing.find((r) => r.person_id === person.id);
      if (row) {
        if (row.state === 'claimed') savePayment(p.month, person.id, 'paid', row.amount, me.id);
        return;
      }
      const toPay = currentToPay(person, p.month);
      if (toPay > 0) savePayment(p.month, person.id, 'paid', toPay, me.id);
    });
  },
  remind(me, p) {
    const people = readTable('People').filter((x) => (p.ids || []).indexOf(x.id) >= 0);
    if (!people.length) throw new Error('Tick at least one person first');
    sendReminder(people);
  },
  approve(me, p) {
    const rows = readTable('People');
    const row = rows.find((x) => x.id === String(p.id));
    const name = String(p.name || '').trim();
    if (!row || !name) throw new Error('Please enter a name');
    if (rows.some((x) => x !== row && x.name.toLowerCase() === name.toLowerCase() && x.role !== 'removed')) throw new Error('That name is already used');
    row.name = name;
    row.role = p.role === 'admin' ? 'admin' : 'member';
    row.pays = p.pays ? 'yes' : 'no';
    row.joined = today();
    row.left = '';
    writeTable('People', rows);
  },
  decline(me, p) {
    writeTable('People', readTable('People').filter((x) => !(x.id === String(p.id) && x.role === 'pending')));
  },
  removePerson(me, p) {
    if (String(p.id) === me.id) throw new Error('You cannot remove yourself');
    const rows = readTable('People');
    const row = rows.find((x) => x.id === String(p.id));
    if (row) { row.role = 'removed'; row.left = today(); writeTable('People', rows); }
  },
  restorePerson(me, p) {
    const rows = readTable('People');
    const row = rows.find((x) => x.id === String(p.id));
    if (row) { row.role = p.role === 'admin' ? 'admin' : 'member'; row.left = ''; writeTable('People', rows); }
  },
  setPaynow(me, p) { setConfig('PAYNOW', String(p.value || '').slice(0, 100)); },
  findGroups() {},
  setTopic(me, p) {
    const key = p.which === 'hangouts' ? 'TOPIC_HANGOUTS' : 'TOPIC_PAYMENTS';
    const link = String(p.link || '').trim();
    if (!link) { setConfig(key, ''); setConfig(key + '_NAME', ''); return; }
    const cfg = config();
    if (!cfg.GROUP_CHAT_ID) throw new Error('Connect your group chat first');
    const thread = threadFromLink(link, cfg.GROUP_CHAT_ID);
    // Bots can't look up a topic's name, but a message sent to a topic comes back with it.
    // So this sends one silent message to read the name, then deletes it straight away.
    const probe = telegram('sendMessage', { chat_id: cfg.GROUP_CHAT_ID, message_thread_id: thread, text: 'setting up...', disable_notification: true });
    const created = probe.reply_to_message && probe.reply_to_message.forum_topic_created;
    try {
      telegram('deleteMessage', { chat_id: cfg.GROUP_CHAT_ID, message_id: probe.message_id });
    } catch (err) {
      logAction('system', 'probe_delete_failed', { error: err.message });
    }
    setConfig(key, String(thread));
    setConfig(key + '_NAME', created ? created.name : '');
  },
  setGroup(me, p) {
    setConfig('GROUP_CHAT_ID', String(p.id));
    setConfig('GROUP_TITLE', String(p.title || ''));
    sendToGroup('hello~ this is poca heels crew bot! i will be everyone\'s secretary, in charge of payment and scheduling. open the <a href="' + MINI_APP_LINK + '">miniapp</a> to see what i can do 🕺🏻🪩');
  },
};

// ---------- Telegram messages ----------
function esc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function telegram(method, body) {
  const token = PropertiesService.getScriptProperties().getProperty('BOT_TOKEN');
  const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/' + method, {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true,
  });
  const json = JSON.parse(res.getContentText());
  if (!json.ok) {
    const err = new Error('Telegram: ' + json.description);
    err.migrateTo = json.parameters && json.parameters.migrate_to_chat_id;
    throw err;
  }
  return json.result;
}
/** Reads a topic's number from a link like t.me/c/1234567890/78 or t.me/c/1234567890/78/90. */
function threadFromLink(link, groupId) {
  const m = /t\.me\/c\/(\d+)\/(\d+)(?:\/\d+)?(?:\?thread=(\d+))?/.exec(link);
  if (!m) throw new Error('That does not look like a link to a topic in your group');
  if (String(groupId).replace(/^-100/, '') !== m[1]) throw new Error('That link is from a different group');
  return Number(m[3] || m[2]);
}
function sendToGroup(html, topicKey) {
  const cfg = config();
  if (!cfg.GROUP_CHAT_ID) throw new Error('Connect your group chat first, on the People tab');
  const body = { chat_id: cfg.GROUP_CHAT_ID, text: html, parse_mode: 'HTML', disable_web_page_preview: true };
  if (cfg.APP_LINK) body.reply_markup = { inline_keyboard: [[{ text: 'Open app', url: cfg.APP_LINK }]] };
  if (topicKey && cfg[topicKey]) body.message_thread_id = Number(cfg[topicKey]);
  try {
    telegram('sendMessage', body);
  } catch (err) {
    if (/thread not found/i.test(err.message) && body.message_thread_id) {
      // The topic was deleted or closed: post in General so the message isn't lost.
      delete body.message_thread_id;
      telegram('sendMessage', body);
      return;
    }
    if (!err.migrateTo) throw err;
    // Telegram gives a group a new ID when it upgrades (for example, when topics are turned on).
    setConfig('GROUP_CHAT_ID', String(err.migrateTo));
    body.chat_id = err.migrateTo;
    telegram('sendMessage', body);
  }
}
/** For updates that shouldn't undo a saved change: no group yet, or a Telegram hiccup, is only logged. */
function notifyGroup(html, topicKey) {
  try {
    if (config().GROUP_CHAT_ID) sendToGroup(html, topicKey);
  } catch (err) {
    logAction('system', 'group_message_failed', { error: err.message });
  }
}
function sendReminder(people) {
  const tags = people.map((p) => '<a href="tg://user?id=' + encodeURIComponent(p.id) + '">' + esc(p.name) + '</a>').join(' ');
  sendToGroup('gentle reminder for payment 🙏🏻 ' + tags + ' mark in the <a href="' + MINI_APP_LINK + '">miniapp</a> once transferred!', 'TOPIC_PAYMENTS');
}
function fetchGroups() {
  const updates = telegram('getUpdates', { limit: 100 });
  const seen = {};
  const upgraded = {};
  updates.forEach((u) => {
    const m = u.message || u.my_chat_member || u.edited_message;
    const chat = m && m.chat;
    if (!chat || (chat.type !== 'group' && chat.type !== 'supergroup')) return;
    seen[chat.id] = chat.title;
    if (m.migrate_to_chat_id) { upgraded[chat.id] = true; seen[m.migrate_to_chat_id] = chat.title; }
    if (m.migrate_from_chat_id) upgraded[m.migrate_from_chat_id] = true;
  });
  return Object.keys(seen).filter((id) => !upgraded[id]).map((id) => ({ id, title: seen[id] }));
}

// ---------- automatic messages, run twice a day ----------
function tick() {
  resetCaches();
  const cfg = config();
  if (!cfg.GROUP_CHAT_ID) return;
  const props = PropertiesService.getScriptProperties();
  const once = (key, fn) => { if (!props.getProperty(key)) { fn(); props.setProperty(key, '1'); } };
  const t = today();
  const hour = nowHour();
  if (hour >= 10) {
    const month = paymentMonth(cfg);
    const due = dueDate(month);
    if (t === due) {
      once('sent_due_' + month, () => sendToGroup('hi!! ' + monthName(month) + ' payment is due today. open the <a href="' + MINI_APP_LINK + '">miniapp</a> to see your amount due 💰💰', 'TOPIC_PAYMENTS'));
    }
    const dayAfter = dateOf(due);
    dayAfter.setUTCDate(dayAfter.getUTCDate() + 1);
    if (t === isoOf(dayAfter)) {
      const data = loadData();
      const led = buildLedger(data, cfg);
      const unpaid = billable(data.people).filter((p) => led.summary(p, month).state === 'none');
      if (unpaid.length) once('sent_overdue_' + month, () => sendReminder(unpaid));
    }
  }
}
