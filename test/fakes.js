'use strict';
/**
 * Hand-written fakes of the Apps Script services the kit uses. They model only what the scripts
 * call, plus a few documented Google behaviours that bite in practice (and are asserted in tests):
 *  - writing a string into a cell that is not formatted as plain text ('@') lets Sheets reinterpret it:
 *    "=..." becomes a formula, "007" becomes the number 7, "2026-01-05" becomes a date, "TRUE" a boolean
 *  - getRange() with zero rows/columns throws, setValues() with the wrong dimensions throws
 *  - a cell holds at most 50,000 characters
 *  - Utilities.computeHmacSha256Signature returns SIGNED bytes (-128..127)
 *  - CacheService caps the TTL at 6 hours
 *  - UrlFetchApp throws on 4xx/5xx unless muteHttpExceptions is true
 *  - time-based triggers accept everyMinutes(1|5|10|15|30) only
 * These behaviours are modelled from Google's documentation and common experience; they have NOT been
 * checked against Google itself (see selfTestTextCells / selfTestHmac for the in-Google check).
 */
const crypto = require('node:crypto');

const MAX_CELL_CHARS = 50000;

function isDate(v) {
  return Object.prototype.toString.call(v) === '[object Date]';
}

// ---------------------------------------------------------------------------------------------
// Spreadsheet

class FakeRange {
  constructor(sheet, row, col, numRows, numCols) {
    if (!(row >= 1) || !(col >= 1)) throw new Error('Exception: The coordinates of the range are outside the dimensions of the sheet.');
    if (!(numRows >= 1)) throw new Error('Exception: The number of rows in the range must be at least 1.');
    if (!(numCols >= 1)) throw new Error('Exception: The number of columns in the range must be at least 1.');
    this.sheet = sheet;
    this.row = row;
    this.col = col;
    this.numRows = numRows;
    this.numCols = numCols;
  }

  getRow() { return this.row; }
  getColumn() { return this.col; }
  getNumRows() { return this.numRows; }
  getNumColumns() { return this.numCols; }
  getSheet() { return this.sheet; }

  _each(fn) {
    for (let r = 0; r < this.numRows; r++) {
      for (let c = 0; c < this.numCols; c++) fn(this.sheet._cell(this.row + r, this.col + c), r, c);
    }
  }

  getValues() {
    this.sheet._count('getValues');
    const out = [];
    for (let r = 0; r < this.numRows; r++) {
      const line = [];
      for (let c = 0; c < this.numCols; c++) line.push(this.sheet._peek(this.row + r, this.col + c).v);
      out.push(line);
    }
    return out;
  }

  getValue() { return this.getValues()[0][0]; }

  setValues(values) {
    this.sheet._count('setValues');
    if (!Array.isArray(values) || values.length !== this.numRows) {
      throw new Error(`Exception: The number of rows in the data does not match the number of rows in the range. The data has ${Array.isArray(values) ? values.length : 'n/a'} but the range has ${this.numRows}.`);
    }
    values.forEach((line) => {
      if (!Array.isArray(line) || line.length !== this.numCols) {
        throw new Error(`Exception: The number of columns in the data does not match the number of columns in the range. The data has ${Array.isArray(line) ? line.length : 'n/a'} but the range has ${this.numCols}.`);
      }
    });
    this._each((cell, r, c) => this.sheet._write(cell, values[r][c]));
    return this;
  }

  setValue(value) {
    this.sheet._count('setValue');
    this._each((cell) => this.sheet._write(cell, value));
    return this;
  }

  setNumberFormat(format) {
    this.sheet._count('setNumberFormat');
    this._each((cell) => { cell.fmt = format; });
    return this;
  }

  getNumberFormats() {
    const out = [];
    for (let r = 0; r < this.numRows; r++) {
      const line = [];
      for (let c = 0; c < this.numCols; c++) line.push(this.sheet._peek(this.row + r, this.col + c).fmt);
      out.push(line);
    }
    return out;
  }

  clearContent() {
    this._each((cell) => { cell.v = ''; cell.formula = false; });
    return this;
  }
}

class FakeSheet {
  constructor(ss, name) {
    this._ss = ss;
    this._name = name;
    this._rows = [];
    this.calls = {};
  }

  _count(name) { this.calls[name] = (this.calls[name] || 0) + 1; }

  _cell(r, c) {
    while (this._rows.length < r) this._rows.push([]);
    const row = this._rows[r - 1];
    while (row.length < c) row.push({ v: '', fmt: '', formula: false });
    return row[c - 1];
  }

  _peek(r, c) {
    const cell = this._rows[r - 1] && this._rows[r - 1][c - 1];
    return cell || { v: '', fmt: '', formula: false };
  }

  _write(cell, raw) {
    let v = raw === null || raw === undefined ? '' : raw;
    if (typeof v === 'string' && v.length > MAX_CELL_CHARS) {
      throw new Error('Exception: Your input contains more than the maximum of 50000 characters in a single cell.');
    }
    cell.formula = false;
    if (cell.fmt === '@') {
      if (isDate(v)) throw new Error('fake: a Date was written into a text-formatted cell; convert it to text first');
      cell.v = typeof v === 'string' ? v : String(v);
      return;
    }
    if (typeof v === 'string') {
      if (v.startsWith('=')) cell.formula = true;
      else if (/^\s*[-+]?(\d+(\.\d+)?|\.\d+)\s*$/.test(v)) v = Number(v);
      else if (/^\d{4}-\d{2}-\d{2}$/.test(v)) v = new Date(v + 'T00:00:00Z');
      else if (v === 'TRUE' || v === 'FALSE') v = v === 'TRUE';
    }
    cell.v = v;
  }

  getName() { return this._name; }
  getParent() { return this._ss; }

  getLastRow() {
    for (let r = this._rows.length; r >= 1; r--) {
      if (this._rows[r - 1].some((cell) => cell.v !== '')) return r;
    }
    return 0;
  }

  getLastColumn() {
    let max = 0;
    this._rows.forEach((row) => {
      for (let c = row.length; c >= 1; c--) {
        if (row[c - 1].v !== '') { max = Math.max(max, c); break; }
      }
    });
    return max;
  }

  getRange(row, col, numRows = 1, numCols = 1) { return new FakeRange(this, row, col, numRows, numCols); }

  getDataRange() {
    return new FakeRange(this, 1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1));
  }

  appendRow(values) {
    if (!Array.isArray(values) || values.length === 0) throw new Error('Exception: The parameters (number[]) do not match the method signature for SpreadsheetApp.Sheet.appendRow.');
    this._count('appendRow');
    const r = this.getLastRow() + 1;
    values.forEach((v, i) => this._write(this._cell(r, i + 1), v));
  }

  deleteRows(start, count) {
    if (!(start >= 1) || !(count >= 1)) throw new Error('Exception: The parameters do not match the method signature for SpreadsheetApp.Sheet.deleteRows.');
    this._count('deleteRows');
    this._rows.splice(start - 1, count);
  }

  deleteRow(n) { this.deleteRows(n, 1); }

  clear() { this._rows = []; }

  // --- test helpers (not part of the Apps Script API) ---------------------------------------
  /** Puts exact values (typed) into the sheet, bypassing the write-time reinterpretation. */
  seed(rows, formats) {
    this._rows = [];
    rows.forEach((line, r) => line.forEach((v, c) => {
      const cell = this._cell(r + 1, c + 1);
      cell.v = v === null || v === undefined ? '' : v;
      if (formats && formats[r] && formats[r][c]) cell.fmt = formats[r][c];
    }));
    return this;
  }

  /** All used cells as a 2D array of values. */
  values() {
    const rows = this.getLastRow();
    const cols = this.getLastColumn();
    const out = [];
    for (let r = 1; r <= rows; r++) {
      const line = [];
      for (let c = 1; c <= cols; c++) line.push(this._peek(r, c).v);
      out.push(line);
    }
    return out;
  }

  formatAt(r, c) { return this._peek(r, c).fmt; }
  isFormula(r, c) { return this._peek(r, c).formula; }
}

class FakeSpreadsheet {
  constructor(id, tz = 'UTC') {
    this._id = id;
    this._tz = tz;
    this._sheets = [];
    this._active = null;
  }

  getId() { return this._id; }
  getSpreadsheetTimeZone() { return this._tz; }
  getSheets() { return this._sheets.slice(); }
  getSheetByName(name) { return this._sheets.find((s) => s._name === name) || null; }
  getActiveSheet() { return this._active || this._sheets[0] || null; }

  insertSheet(name) {
    if (this.getSheetByName(name)) throw new Error(`Exception: A sheet with the name "${name}" already exists. Please enter another name.`);
    const sheet = new FakeSheet(this, name);
    this._sheets.push(sheet);
    if (!this._active) this._active = sheet;
    return sheet;
  }

  deleteSheet(sheet) {
    this._sheets = this._sheets.filter((s) => s !== sheet);
    if (this._active === sheet) this._active = this._sheets[0] || null;
  }

  // test helper
  addSheet(name, rows, formats) {
    const sheet = this.insertSheet(name);
    if (rows) sheet.seed(rows, formats);
    this._active = sheet;
    return sheet;
  }
}

class FakeUi {
  constructor() {
    this.alerts = [];
    this.answers = [];
    this.menus = [];
    this.Button = { OK: 'OK', YES: 'YES', NO: 'NO', CLOSE: 'CLOSE' };
    this.ButtonSet = { OK: 'OK', YES_NO: 'YES_NO' };
  }

  alert(...args) {
    this.alerts.push(args);
    return this.answers.length ? this.answers.shift() : this.Button.OK;
  }

  createMenu(name) {
    const menu = { name, items: [] };
    const builder = {
      addItem: (label, fn) => { menu.items.push([label, fn]); return builder; },
      addSeparator: () => builder,
      addToUi: () => { this.menus.push(menu); },
    };
    return builder;
  }
}

// ---------------------------------------------------------------------------------------------
// Other services

function makePropertiesService() {
  const store = new Map();
  const props = {
    getProperty: (k) => (store.has(k) ? store.get(k) : null),
    setProperty: (k, v) => { store.set(k, String(v)); return props; },
    setProperties: (o) => { Object.keys(o).forEach((k) => store.set(k, String(o[k]))); return props; },
    deleteProperty: (k) => { store.delete(k); return props; },
    getProperties: () => Object.fromEntries(store),
  };
  return { getScriptProperties: () => props, _props: props };
}

function makeCacheService(clock) {
  const store = new Map();
  const cache = {
    reads: 0,
    puts: 0,
    get(key) {
      cache.reads++;
      const hit = store.get(key);
      if (!hit) return null;
      if (hit.expires <= clock.now) { store.delete(key); return null; }
      return hit.value;
    },
    put(key, value, ttlSeconds = 600) {
      if (String(key).length > 250) throw new Error('Exception: Argument too large: key');
      if (String(value).length > 100 * 1024) throw new Error('Exception: Argument too large: value');
      cache.puts++;
      store.set(key, { value: String(value), expires: clock.now + Math.min(ttlSeconds, 21600) * 1000 });
    },
    remove(key) { store.delete(key); },
    clear() { store.clear(); },
  };
  return { getScriptCache: () => cache, _cache: cache };
}

function makeLockService(clock) {
  const state = { held: false, heldByOther: false, acquired: 0, releases: 0 };
  const lock = {
    tryLock(ms) {
      if (state.heldByOther || state.held) { clock.now += ms; return false; }
      state.held = true;
      state.acquired++;
      return true;
    },
    waitLock(ms) {
      if (!lock.tryLock(ms)) throw new Error(`Exception: Could not obtain lock after ${ms} milliseconds.`);
    },
    releaseLock() { state.held = false; state.releases++; },
    hasLock() { return state.held; },
  };
  return { getScriptLock: () => lock, state };
}

function makeUtilities(clock) {
  const util = {
    sleeps: [],
    computeHmacSha256Signature(value, key) {
      if (typeof value !== 'string' || typeof key !== 'string') throw new Error('fake: only the (string, string) overload is modelled');
      const digest = crypto.createHmac('sha256', key).update(value, 'utf8').digest();
      return Array.from(digest, (b) => (b > 127 ? b - 256 : b));
    },
    sleep(ms) {
      util.sleeps.push(ms);
      clock.now += ms;
    },
    getUuid: () => crypto.randomUUID(),
    formatDate(date, tz, pattern) {
      const parts = new Intl.DateTimeFormat('en-GB', {
        timeZone: tz, hourCycle: 'h23', year: 'numeric', month: '2-digit', day: '2-digit',
        hour: '2-digit', minute: '2-digit', second: '2-digit',
      }).formatToParts(new Date(date.getTime()));
      const p = {};
      parts.forEach((x) => { p[x.type] = x.value; });
      const tokens = { yyyy: p.year, MM: p.month, dd: p.day, HH: p.hour, mm: p.minute, ss: p.second };
      return pattern.replace(/'([^']*)'|yyyy|MM|dd|HH|mm|ss|[A-Za-z]/g, (m, lit) => {
        if (lit !== undefined) return lit;
        if (tokens[m] === undefined) throw new Error(`fake: date pattern token "${m}" is not modelled`);
        return tokens[m];
      });
    },
  };
  return util;
}

function makeUrlFetchApp(clock) {
  const app = {
    queue: [],
    requests: [],
    /** Next response: {code, body, headers, latencyMs} | Error (thrown) | function(request) -> either. */
    enqueue(...items) { app.queue.push(...items); },
    fetch(url, params = {}) {
      const request = { url, params, at: clock.now };
      app.requests.push(request);
      if (app.queue.length === 0) throw new Error('fake UrlFetchApp: no scripted response left');
      let item = app.queue.shift();
      if (typeof item === 'function') item = item(request);
      if (item instanceof Error) throw item;
      if (item.latencyMs) clock.now += item.latencyMs;
      const response = {
        getResponseCode: () => item.code,
        getContentText: () => (item.body === undefined ? '' : String(item.body)),
        getHeaders: () => item.headers || {},
        getAllHeaders: () => item.headers || {},
      };
      if (!params.muteHttpExceptions && item.code >= 400) {
        throw new Error(`Exception: Request failed for ${url} returned code ${item.code}. Truncated server response: ${response.getContentText()}`);
      }
      return response;
    },
  };
  return app;
}

function makeMailApp() {
  const mail = {
    sent: [],
    quota: 100,
    failNext: null,
    sendEmail(a, subject, body) {
      if (mail.failNext) { const err = mail.failNext; mail.failNext = null; throw err; }
      const msg = typeof a === 'object' && a !== null ? { ...a } : { to: a, subject, body };
      if (!msg.to) throw new Error('Exception: Invalid argument: to');
      mail.sent.push(msg);
    },
    getRemainingDailyQuota() {
      const used = mail.sent.reduce((n, m) => n + [m.to, m.cc, m.bcc].filter(Boolean).join(',').split(',').filter(Boolean).length, 0);
      return mail.quota - used;
    },
  };
  return mail;
}

function makeScriptApp() {
  const app = {
    triggers: [],
    EventType: { CLOCK: 'CLOCK', ON_FORM_SUBMIT: 'ON_FORM_SUBMIT' },
    getProjectTriggers: () => app.triggers.slice(),
    deleteTrigger: (t) => { app.triggers = app.triggers.filter((x) => x !== t); },
    newTrigger(handler) {
      const spec = { handler, type: null, sourceId: null, everyMinutes: null, everyDays: null, atHour: null };
      const finish = () => {
        const trigger = {
          getHandlerFunction: () => spec.handler,
          getEventType: () => spec.type,
          getTriggerSourceId: () => spec.sourceId,
          spec,
        };
        app.triggers.push(trigger);
        return trigger;
      };
      return {
        timeBased() {
          spec.type = 'CLOCK';
          const clockBuilder = {
            everyMinutes(n) {
              if (![1, 5, 10, 15, 30].includes(n)) throw new Error('Exception: Invalid number of minutes for time-based trigger');
              spec.everyMinutes = n;
              return clockBuilder;
            },
            everyDays(n) { spec.everyDays = n; return clockBuilder; },
            atHour(h) {
              if (spec.everyDays === null) throw new Error('fake: atHour needs everyDays first');
              if (!(h >= 0 && h <= 23)) throw new Error('Exception: Invalid hour');
              spec.atHour = h;
              return clockBuilder;
            },
            create: finish,
          };
          return clockBuilder;
        },
        forForm(idOrForm) {
          spec.sourceId = typeof idOrForm === 'string' ? idOrForm : idOrForm.getId();
          return {
            onFormSubmit() {
              spec.type = 'ON_FORM_SUBMIT';
              return { create: finish };
            },
          };
        },
      };
    },
  };
  return app;
}

function makeContentService() {
  return {
    MimeType: { JSON: 'JSON', TEXT: 'TEXT' },
    createTextOutput(text) {
      const out = {
        _text: text,
        _mime: 'TEXT',
        setMimeType(m) { out._mime = m; return out; },
        getContent: () => out._text,
        getMimeType: () => out._mime,
      };
      return out;
    },
  };
}

// --- Forms ----------------------------------------------------------------------------------

class FakeForm {
  constructor(id, title) {
    this._id = id;
    this._title = title;
    this._responses = [];
  }

  getId() { return this._id; }
  getTitle() { return this._title; }
  getResponses() { return this._responses.slice(); }

  /** answers: [[questionTitle, value], ...] */
  makeResponse({ id, timestamp = '2026-11-01T10:00:00.000Z', email = null, answers = [] }) {
    return {
      getId: () => id,
      getTimestamp: () => new Date(timestamp),
      getRespondentEmail: () => email,
      getItemResponses: () => answers.map(([title, value]) => ({
        getItem: () => ({ getTitle: () => title }),
        getResponse: () => value,
      })),
    };
  }

  addResponse(spec) {
    const response = this.makeResponse(spec);
    this._responses.push(response);
    return response;
  }
}

module.exports = {
  FakeSpreadsheet, FakeSheet, FakeUi, FakeForm,
  makePropertiesService, makeCacheService, makeLockService, makeUtilities,
  makeUrlFetchApp, makeMailApp, makeScriptApp, makeContentService,
  isDate,
};
