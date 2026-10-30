/**
 * Shared helpers for every script in this kit.
 *
 * Apps Script runs all files of a project in ONE global scope, so a function name used twice
 * silently overrides the first. Helpers therefore end in "_" (Apps Script also hides those from
 * the editor's Run menu); entry points do not. test/structure.test.js fails on duplicate names.
 *
 * Configuration lives in Script Properties (Project Settings > Script properties), never in code.
 */

/** Trimmed script property, or `fallback` when it is missing or blank. */
function getProp_(name, fallback) {
  var value = PropertiesService.getScriptProperties().getProperty(name);
  if (value === null || value === undefined) return fallback;
  value = String(value).trim();
  return value === '' ? fallback : value;
}

/** Secrets are returned exactly as stored (no trimming); null when missing or empty. */
function getSecret_(name) {
  var value = PropertiesService.getScriptProperties().getProperty(name);
  return value === null || value === undefined || value === '' ? null : String(value);
}

/** Whole-number property. A malformed value is a configuration error, not a silent default. */
function getIntProp_(name, fallback) {
  var raw = getProp_(name, null);
  if (raw === null) return fallback;
  if (!/^\d+$/.test(raw)) throw new Error('Script property ' + name + ' must be a whole number, got "' + raw + '"');
  return parseInt(raw, 10);
}

/** Comma/semicolon/space separated list property -> array of non-empty strings. */
function getListProp_(name) {
  var raw = getProp_(name, '');
  return raw === '' ? [] : raw.split(/[,;\s]+/).filter(function (s) { return s !== ''; });
}

/** Realm-safe Date check (also true for Dates created by Sheets). */
function isDate_(value) {
  return Object.prototype.toString.call(value) === '[object Date]';
}

/** SPREADSHEET_ID property if set, otherwise the spreadsheet this script is bound to. */
function getSpreadsheet_() {
  var id = getProp_('SPREADSHEET_ID', null);
  var ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No spreadsheet: set the SPREADSHEET_ID script property or bind the script to a spreadsheet.');
  return ss;
}

function getSheetByName_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Sheet "' + name + '" not found.');
  return sheet;
}

/** Collapses every kind of whitespace (including non-breaking and zero-width spaces) and trims. */
function normalizeText_(value) {
  if (value === null || value === undefined) return '';
  return String(value).replace(/[\u200b\u200c\u200d\ufeff]/g, '').replace(/\s+/g, ' ').trim();
}

function normalizeKey_(value) {
  return normalizeText_(value).toLowerCase();
}

/** Plain text for a cell value. Dates become yyyy-MM-dd, with the time only when it is not midnight. */
function toText_(value, tz) {
  if (value === null || value === undefined) return '';
  if (isDate_(value)) {
    if (isNaN(value.getTime())) return '';
    var day = Utilities.formatDate(value, tz, 'yyyy-MM-dd');
    var time = Utilities.formatDate(value, tz, 'HH:mm:ss');
    return time === '00:00:00' ? day : day + ' ' + time;
  }
  return String(value);
}

function isBlankCell_(value) {
  return value === '' || value === null || value === undefined || (typeof value === 'string' && value.trim() === '');
}

function isBlankRow_(row) {
  for (var i = 0; i < row.length; i++) if (!isBlankCell_(row[i])) return false;
  return true;
}

/** Reads the whole used range. Row 1 is the header; `rows[i]` is sheet row i + 2. */
function readTable_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return { headers: [], rows: [] };
  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  return {
    headers: values[0].map(function (h) { return normalizeText_(h); }),
    rows: values.slice(1)
  };
}

/** Header row only (cheap even on a sheet with many rows). */
function readHeaders_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (sheet.getLastRow() < 1 || lastCol < 1) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return normalizeText_(h); });
}

/** 0-based column of a header (case-insensitive, whitespace-insensitive), or -1. */
function headerIndex_(headers, name) {
  var wanted = normalizeKey_(name);
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] !== '' && headers[i].toLowerCase() === wanted) return i;
  }
  return -1;
}

/**
 * Makes sure each named header exists, appending missing ones to the right of the last used
 * column of row 1 (existing cells are never moved). `table` is updated in place so its rows stay
 * aligned. Returns {name: 0-based column}.
 */
function ensureColumns_(sheet, table, names) {
  var out = {};
  names.forEach(function (name) {
    var idx = headerIndex_(table.headers, name);
    if (idx < 0) {
      idx = table.headers.length;
      sheet.getRange(1, idx + 1, 1, 1).setNumberFormat('@').setValues([[name]]);
      table.headers.push(name);
      table.rows.forEach(function (row) { row.push(''); });
    }
    out[name] = idx;
  });
  return out;
}

/** Writes text only: the '@' format stops Sheets turning "=SUM(1)", "007" or "2026-01-05" into something else. */
function writeTextBlock_(sheet, row, col, matrix) {
  var range = sheet.getRange(row, col, matrix.length, matrix[0].length);
  range.setNumberFormat('@');
  range.setValues(matrix);
}

/**
 * Runs fn while holding the script lock. Returns {locked:false} when the lock could not be
 * obtained within timeoutMs, otherwise {locked:true, value}. The lock is always released.
 */
function withScriptLock_(timeoutMs, fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(timeoutMs)) return { locked: false };
  try {
    return { locked: true, value: fn() };
  } finally {
    lock.releaseLock();
  }
}

/** Apps Script returns HMAC bytes as signed numbers (-128..127); mask before formatting. */
function bytesToHex_(bytes) {
  var out = '';
  for (var i = 0; i < bytes.length; i++) {
    var b = bytes[i] & 0xff;
    out += (b < 16 ? '0' : '') + b.toString(16);
  }
  return out;
}

/** HMAC-SHA256 as lowercase hex. Both arguments are strings (UTF-8). */
function hmacSha256Hex_(message, key) {
  return bytesToHex_(Utilities.computeHmacSha256Signature(message, key));
}

/**
 * Compares without returning at the first differing character. Best effort only: JavaScript gives
 * no timing guarantee, which is acceptable here because the compared value is a one-time HMAC.
 */
function constantTimeEquals_(a, b) {
  a = String(a);
  b = String(b);
  var diff = a.length ^ b.length;
  var n = Math.max(a.length, b.length);
  for (var i = 0; i < n; i++) diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return diff === 0;
}

/**
 * Strict ISO-8601 with a time zone (what Date#toISOString emits), like webhook-kit's parseTimestamp.
 * Stricter than Date.parse on purpose: V8 rolls "2026-02-30" over to March 2. Returns epoch ms or null.
 */
function parseIsoTimestamp_(text) {
  if (typeof text !== 'string') return null;
  var m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(?:\.\d{1,6})?(?:Z|[+-]\d{2}:\d{2})$/.exec(text);
  if (!m) return null;
  var y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
  if (mo < 1 || mo > 12 || d < 1 || Number(m[4]) > 23 || Number(m[5]) > 59 || Number(m[6]) > 59) return null;
  if (d > new Date(Date.UTC(y, mo, 0)).getUTCDate()) return null;
  var ms = Date.parse(text);
  return isNaN(ms) ? null : ms;
}

function isEmail_(text) {
  return /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]{2,}$/.test(text) && text.indexOf('..') === -1 && text.charAt(0) !== '.';
}

function escapeHtml_(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/** One line, capped, safe to store in a cell or a log. */
function oneLine_(value, max) {
  var s = String(value).replace(/\s+/g, ' ').trim();
  return s.length > max ? s.slice(0, max - 3) + '...' : s;
}
