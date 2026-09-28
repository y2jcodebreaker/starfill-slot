/**
 * STARFILL Lucky Spin — backend (Google Apps Script web app)
 *
 * The server is the single source of truth. It decides every spin, so the
 * limits below cannot be bypassed by reloading, incognito tabs or a second
 * device. See SETUP.md for deployment steps.
 *
 * Rules enforced here:
 *   - One record per mobile number (normalised 10-digit Indian mobile).
 *   - MAX_SPINS_PER_DAY spins per mobile per conference day (IST).
 *   - A winner cannot spin again for WIN_LOCK_HOURS.
 *   - Exactly WINNERS_PER_DAY winners per day across all devices
 *     (never more; see prize pacing in CONFIG).
 *
 * Sheets written (created automatically, append-only):
 *   Registrations — first check-in of each mobile per day
 *   Spins         — one row per spin
 */

var CONFIG = {
  TIMEZONE: 'Asia/Kolkata',
  MAX_SPINS_PER_DAY: 3,
  WINNERS_PER_DAY: 2,        // exactly this many winners each day (never more)
  // Prize pacing. The day's winning spins are placed at random among the first
  // PRIZE_WINDOW_SPINS spins: each spin wins with chance
  //   prizes left / spins left in the window
  // which is the same as drawing 2 winning tickets out of 120 without
  // replacement. The first spin has a 2/120 = 1.7% chance. Both prizes are
  // guaranteed to be awarded by spin 120 (~42 visitors at ~2.9 spins each).
  PRIZE_WINDOW_SPINS: 120,
  // Safety net for a slow day: from this time (IST), any prize still left is
  // awarded on the next spin(s), so the day still ends with exactly 2 winners.
  // Set this to about an hour before the booth closes. '' disables it.
  LAST_CALL_TIME: '16:00',
  WIN_LOCK_HOURS: 24,
  RESET_SPINS_DAILY: true,  // false = 3 spins for the whole event
};

// Which prize a winner gets keeps the original 5 : 3 : 2 distribution.
var SYMBOLS = [
  { id: 'plus',    label: 'STARFILL PLUS',         weight: 5, prefix: 'SFP' },
  { id: 'deep',    label: 'STARFILL DEEP PLUS',    weight: 3, prefix: 'SFD' },
  { id: 'implant', label: 'STARFILL IMPLANT PLUS', weight: 2, prefix: 'SFI' },
];

var REG_HEADERS  = ['Timestamp (IST)', 'Date', 'Name', 'Mobile'];
var SPIN_HEADERS = ['Timestamp (IST)', 'Date', 'Name', 'Mobile', 'Spin #',
                    'Reel 1', 'Reel 2', 'Reel 3', 'Result', 'Prize', 'Claim Code'];

// ─────────────────────────────────────────────
// ENTRY POINTS
// ─────────────────────────────────────────────
function doPost(e) {
  var req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad_request' });
  }
  try {
    switch (req.action) {
      case 'register':     return json_(withLock_(function () { return register_(req); }));
      case 'spin':         return json_(withLock_(function () { return spin_(req); }));
      case 'admin_days':   return json_(adminDays_(req));
      case 'admin_export': return json_(adminExport_(req));
      case 'admin_verify': return json_(adminVerify_(req));
      default:             return json_({ ok: false, error: 'unknown_action' });
    }
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'server_error' });
  }
}

// Open the /exec URL in a browser to confirm the deployment is live
function doGet() {
  return ContentService.createTextOutput('STARFILL backend is running.');
}

/**
 * Run this ONCE from the Apps Script editor after testing and before the
 * conference opens (select resetAllData → Run). It clears every player's
 * spin/win state and the daily prize counters, and renames the Registrations
 * and Spins tabs to "... (archived <time>)" so no test data counts on the day.
 * The ADMIN_KEY property is kept.
 */
function resetAllData() {
  var p = props_();
  Object.keys(p.getProperties()).forEach(function (k) {
    if (/^(u|d|code):/.test(k)) p.deleteProperty(k);
  });
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var suffix = ' (archived ' + stamp_(Date.now()) + ')';
  ['Registrations', 'Spins'].forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (sh) sh.setName(name + suffix);
  });
}

// ─────────────────────────────────────────────
// PLAYER ACTIONS
// ─────────────────────────────────────────────
function register_(req) {
  var phone = normalizePhone_(req.phone);
  var name  = cleanName_(req.name);
  if (!phone) return { ok: false, error: 'invalid_phone' };
  if (!name)  return { ok: false, error: 'invalid_name' };

  var now = Date.now();
  var today = dayKey_(now);
  var u = getUser_(phone) || {};

  if (u.lockedUntil && now < u.lockedUntil) {
    return { ok: true, status: 'won_locked', name: u.name, lockedUntil: u.lockedUntil };
  }

  var firstVisitToday = u.day !== today;
  if (firstVisitToday) {
    if (CONFIG.RESET_SPINS_DAILY || !u.day) u.used = 0;
    u.day = today;
    u.last = null;
  }
  if (firstVisitToday || !u.name) u.name = name;  // first name entered each day sticks
  setUser_(phone, u);

  if (firstVisitToday) {
    sheet_('Registrations', REG_HEADERS).appendRow([text_(stamp_(now)), text_(today), safeCell_(name), text_(phone)]);
  }

  var spinsLeft = Math.max(0, CONFIG.MAX_SPINS_PER_DAY - (u.used || 0));
  return {
    ok: true,
    status: spinsLeft > 0 ? 'ok' : 'no_spins',
    name: u.name,
    spinsLeft: spinsLeft,
    maxSpins: CONFIG.MAX_SPINS_PER_DAY,
  };
}

function spin_(req) {
  var phone = normalizePhone_(req.phone);
  if (!phone) return { ok: false, error: 'invalid_phone' };
  var spinId = String(req.spinId || '').slice(0, 64);

  var now = Date.now();
  var today = dayKey_(now);
  var u = getUser_(phone);
  if (!u || u.day !== today) return { ok: false, error: 'not_registered' };

  // Idempotent retry: a client that lost the response gets the same result
  if (spinId && u.last && u.last.spinId === spinId) return u.last.result;

  if (u.lockedUntil && now < u.lockedUntil) return { ok: false, error: 'won_locked' };
  if ((u.used || 0) >= CONFIG.MAX_SPINS_PER_DAY) return { ok: false, error: 'no_spins' };

  // Test wins (?admin&forceWin=...) need the real admin key and never count
  // toward the daily prize cap or the player's spin allowance.
  var forced = null;
  if (req.forceWin && isAdmin_(req.adminKey)) forced = symbolById_(req.forceWin);

  var stats = getDay_(today);
  var won, reels, prize = null;
  if (forced) {
    won = true;
    prize = forced;
    reels = [forced.id, forced.id, forced.id];
  } else {
    won = Math.random() < winChance_(stats, now);
    if (won) {
      prize = weightedSymbol_();
      reels = [prize.id, prize.id, prize.id];
    } else {
      reels = losingReels_();
    }
  }

  var code = '';
  if (won) {
    code = newClaimCode_(prize.prefix);
    props_().setProperty('code:' + code, JSON.stringify({
      name: u.name, phone: phone, prize: prize.label, issued: stamp_(now), test: !!forced,
    }));
  }
  if (!forced) {
    u.used = (u.used || 0) + 1;
    stats.spins += 1;
    if (won) {
      stats.wins += 1;
      u.lockedUntil = now + CONFIG.WIN_LOCK_HOURS * 3600 * 1000;
    }
    setDay_(today, stats);
  }

  sheet_('Spins', SPIN_HEADERS).appendRow([
    text_(stamp_(now)), text_(today), safeCell_(u.name), text_(phone),
    forced ? 'TEST' : u.used,
    labelOf_(reels[0]), labelOf_(reels[1]), labelOf_(reels[2]),
    forced ? 'TEST WIN' : (won ? 'WON' : 'miss'),
    won ? prize.label : '',
    code,
  ]);

  var spinsLeft = won && !forced ? 0 : Math.max(0, CONFIG.MAX_SPINS_PER_DAY - u.used);
  var result = {
    ok: true, reels: reels, won: won,
    prize: won ? prize.label : '',
    code: code,
    spinsLeft: spinsLeft,
    lockedUntil: won && !forced ? u.lockedUntil : null,
  };
  u.last = { spinId: spinId, result: result };
  setUser_(phone, u);
  return result;
}

function winChance_(stats, now) {
  var prizesLeft = CONFIG.WINNERS_PER_DAY - stats.wins;
  if (prizesLeft <= 0) return 0;
  if (CONFIG.LAST_CALL_TIME && timeOfDay_(now) >= CONFIG.LAST_CALL_TIME) return 1;
  var spinsLeftInWindow = CONFIG.PRIZE_WINDOW_SPINS - stats.spins;
  if (spinsLeftInWindow <= prizesLeft) return 1;
  return prizesLeft / spinsLeftInWindow;
}

function weightedSymbol_() {
  var total = SYMBOLS.reduce(function (s, x) { return s + x.weight; }, 0);
  var r = Math.random() * total;
  for (var i = 0; i < SYMBOLS.length; i++) {
    r -= SYMBOLS[i].weight;
    if (r < 0) return SYMBOLS[i];
  }
  return SYMBOLS[0];
}

// Three weighted symbols that are guaranteed not to all match
function losingReels_() {
  while (true) {
    var r = [weightedSymbol_().id, weightedSymbol_().id, weightedSymbol_().id];
    if (!(r[0] === r[1] && r[1] === r[2])) return r;
  }
}

// ─────────────────────────────────────────────
// ADMIN ACTIONS
// ─────────────────────────────────────────────
function adminDays_(req) {
  if (!isAdmin_(req.adminKey)) return { ok: false, error: 'unauthorized' };
  var days = {};
  rows_('Registrations').forEach(function (r) {
    var d = String(r[1]);
    days[d] = days[d] || { date: d, players: 0, spins: 0, winners: 0 };
    days[d].players += 1;
  });
  rows_('Spins').forEach(function (r) {
    var d = String(r[1]);
    if (r[8] === 'TEST WIN') return;
    days[d] = days[d] || { date: d, players: 0, spins: 0, winners: 0 };
    days[d].spins += 1;
    if (r[8] === 'WON') days[d].winners += 1;
  });
  var list = Object.keys(days).sort().reverse().map(function (k) { return days[k]; });
  return { ok: true, days: list, today: dayKey_(Date.now()), winnersPerDay: CONFIG.WINNERS_PER_DAY };
}

// Staff check a code shown on a winner's phone
function adminVerify_(req) {
  if (!isAdmin_(req.adminKey)) return { ok: false, error: 'unauthorized' };
  var code = normalizeCode_(req.code);
  if (!isWellFormedCode_(code)) return { ok: true, valid: false, code: code, reason: 'not_a_real_code' };
  var raw = props_().getProperty('code:' + code);
  if (!raw) return { ok: true, valid: false, code: code, reason: 'not_issued' };
  var info = JSON.parse(raw);
  return { ok: true, valid: !info.test, test: !!info.test, code: code,
           name: info.name, phone: info.phone, prize: info.prize, issued: info.issued };
}

function adminExport_(req) {
  if (!isAdmin_(req.adminKey)) return { ok: false, error: 'unauthorized' };
  var date = String(req.date || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return { ok: false, error: 'bad_date' };

  var byPhone = {};
  var order = [];
  rows_('Registrations').forEach(function (r) {
    if (String(r[1]) !== date) return;
    var phone = String(r[3]);
    if (!byPhone[phone]) {
      byPhone[phone] = { name: String(r[2]), phone: phone, firstSeen: String(r[0]),
                         spins: 0, result: 'No win', prize: '', code: '', lastSpin: '' };
      order.push(phone);
    }
  });

  var spins = [];
  rows_('Spins').forEach(function (r) {
    if (String(r[1]) !== date) return;
    var s = { time: String(r[0]), name: String(r[2]), phone: String(r[3]), spinNo: String(r[4]),
              reel1: String(r[5]), reel2: String(r[6]), reel3: String(r[7]),
              result: String(r[8]), prize: String(r[9]), code: String(r[10]) };
    spins.push(s);
    var p = byPhone[s.phone];
    if (!p) {
      p = byPhone[s.phone] = { name: s.name, phone: s.phone, firstSeen: s.time,
                               spins: 0, result: 'No win', prize: '', code: '', lastSpin: '' };
      order.push(s.phone);
    }
    if (s.result === 'TEST WIN') return;
    p.name = s.name;
    p.spins += 1;
    p.lastSpin = s.time;
    if (s.result === 'WON') { p.result = 'WON'; p.prize = s.prize; p.code = s.code; }
  });

  return { ok: true, date: date, participants: order.map(function (k) { return byPhone[k]; }), spins: spins };
}

// ─────────────────────────────────────────────
// STORAGE
// ─────────────────────────────────────────────
function props_() { return PropertiesService.getScriptProperties(); }

function getUser_(phone) {
  var raw = props_().getProperty('u:' + phone);
  return raw ? JSON.parse(raw) : null;
}
function setUser_(phone, u) { props_().setProperty('u:' + phone, JSON.stringify(u)); }

function getDay_(day) {
  var raw = props_().getProperty('d:' + day);
  return raw ? JSON.parse(raw) : { spins: 0, wins: 0 };
}
function setDay_(day, stats) { props_().setProperty('d:' + day, JSON.stringify(stats)); }

function sheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  return sh;
}

// Data rows (header skipped) of a sheet, as display strings
function rows_(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getDisplayValues();
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: 'busy' };
  try { return fn(); } finally { lock.releaseLock(); }
}

// ─────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────
function isAdmin_(key) {
  var real = props_().getProperty('ADMIN_KEY');
  return !!real && typeof key === 'string' && key === real;
}

// Accepts "98765 43210", "+91 98765-43210", "09876543210" → "9876543210"
function normalizePhone_(raw) {
  var d = String(raw || '').replace(/\D/g, '');
  if (d.length === 12 && d.indexOf('91') === 0) d = d.slice(2);
  else if (d.length === 11 && d.charAt(0) === '0') d = d.slice(1);
  return /^[6-9]\d{9}$/.test(d) ? d : null;
}

function cleanName_(raw) {
  return String(raw || '').replace(/[\u0000-\u001f\u007f]/g, '').replace(/\s+/g, ' ').trim().slice(0, 60);
}

// Stop names like "=HYPERLINK(...)" being run as spreadsheet formulas
function safeCell_(s) { return /^[=+\-@]/.test(s) ? "'" + s : s; }

// Leading quote keeps dates and mobiles as plain text (no 9.98E+9, no date reformatting)
function text_(s) { return "'" + s; }

function symbolById_(id) {
  for (var i = 0; i < SYMBOLS.length; i++) if (SYMBOLS[i].id === id) return SYMBOLS[i];
  return null;
}
function labelOf_(id) { var s = symbolById_(id); return s ? s.label : id; }

// ─────────────────────────────────────────────
// CLAIM CODES — e.g. SFI-7KQ4XM9
//   prefix (prize) + 6 random characters + 1 check character.
//   Characters exclude look-alikes (0/O, 1/I/L, U) so codes are easy to read out.
//   Randomness comes from Utilities.getUuid() (a secure random UUID v4), not
//   Math.random(). The check character makes typos and made-up codes fail
//   validation, and every issued code is stored so each one is unique.
// ─────────────────────────────────────────────
var CODE_ALPHABET = '23456789ABCDEFGHJKMNPQRSTVWXYZ'; // 30 characters
var CODE_BODY_LEN = 6;                                // 30^6 ≈ 729 million codes per prize

function newClaimCode_(prefix) {
  for (var attempt = 0; attempt < 20; attempt++) {
    var body = secureRandomChars_(CODE_BODY_LEN);
    var code = prefix + '-' + body + codeCheckChar_(prefix, body);
    if (!props_().getProperty('code:' + code)) return code;
  }
  throw new Error('could not generate a unique claim code');
}

// Uniform random characters from CODE_ALPHABET using UUID bytes
// (bytes >= 240 are skipped so every character is equally likely)
function secureRandomChars_(n) {
  var out = '';
  while (out.length < n) {
    var hex = Utilities.getUuid().replace(/-/g, '');
    for (var i = 0; i + 1 < hex.length && out.length < n; i += 2) {
      if (i === 12 || i === 16) continue;  // UUID version/variant bytes aren't random
      var b = parseInt(hex.substr(i, 2), 16);
      if (b < 240) out += CODE_ALPHABET.charAt(b % CODE_ALPHABET.length);
    }
  }
  return out;
}

// Check character: Luhn mod 30 over the prize letter + the random body.
// Catches every single wrong character (including a changed prize prefix,
// e.g. SFP → SFI) and almost every swap of two neighbouring characters.
var CODE_PRIZE_CHAR = { SFP: 'P', SFD: 'D', SFI: 'J' };

function codeCheckChar_(prefix, body) {
  var s = CODE_PRIZE_CHAR[prefix] + body;
  var n = CODE_ALPHABET.length, factor = 2, sum = 0;
  for (var i = s.length - 1; i >= 0; i--) {
    var addend = factor * CODE_ALPHABET.indexOf(s.charAt(i));
    addend = Math.floor(addend / n) + (addend % n);
    sum += addend;
    factor = factor === 2 ? 1 : 2;
  }
  return CODE_ALPHABET.charAt((n - (sum % n)) % n);
}

function normalizeCode_(raw) {
  return String(raw || '').toUpperCase().replace(/\s/g, '');
}

function isWellFormedCode_(code) {
  var m = /^(SF[PDI])-([2-9A-HJKMNP-TV-Z]{6})([2-9A-HJKMNP-TV-Z])$/.exec(code);
  return !!m && codeCheckChar_(m[1], m[2]) === m[3];
}

function dayKey_(ms) { return Utilities.formatDate(new Date(ms), CONFIG.TIMEZONE, 'yyyy-MM-dd'); }
function timeOfDay_(ms) { return Utilities.formatDate(new Date(ms), CONFIG.TIMEZONE, 'HH:mm'); }
function stamp_(ms)  { return Utilities.formatDate(new Date(ms), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
