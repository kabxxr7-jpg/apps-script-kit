'use strict';
/**
 * Loads the files in src/ into a bare vm context that has only JavaScript built-ins plus the fakes,
 * the way Apps Script runs them: every file in one global scope, no require/process/Buffer/fetch/
 * setTimeout/URL. Anything the scripts use that Apps Script would not provide fails here.
 */
const vm = require('node:vm');
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const fakes = require('./fakes');

// KIT_SRC_DIR lets tools/mutation-check.js run the suite against a mutated copy without touching src/.
const SRC = process.env.KIT_SRC_DIR || path.join(__dirname, '..', 'src');
const LOAD_ORDER = ['common.js', 'dedupe.js', 'clean.js', 'receiver.js', 'sync.js', 'digest.js', 'forms.js', 'selftest.js', 'menu.js'];

function sourceFiles() {
  return LOAD_ORDER.map((f) => ({ name: f, code: fs.readFileSync(path.join(SRC, f), 'utf8') }));
}

/** Plain outer-realm copy, so assert.deepStrictEqual works on results that came out of the vm. */
function plain(value) {
  return value === undefined ? undefined : JSON.parse(JSON.stringify(value));
}

function createEnv(options = {}) {
  const clock = { now: Date.parse(options.now || '2026-11-01T12:00:00.000Z') };
  const logs = { log: [], warn: [], error: [] };
  const ctx = vm.createContext({});

  ctx.__clock = clock;
  vm.runInContext(
    `globalThis.Date = class extends Date {
       constructor(...args) { if (args.length === 0) { super(__clock.now); } else { super(...args); } }
       static now() { return __clock.now; }
     };`,
    ctx
  );

  const ss = new fakes.FakeSpreadsheet(options.spreadsheetId || 'ss-main', options.timeZone || 'UTC');
  const ui = new fakes.FakeUi();
  const properties = fakes.makePropertiesService();
  const cache = fakes.makeCacheService(clock);
  const lock = fakes.makeLockService(clock);
  const utilities = fakes.makeUtilities(clock);
  const urlFetch = fakes.makeUrlFetchApp(clock);
  const mail = fakes.makeMailApp();
  const scriptApp = fakes.makeScriptApp();
  const forms = new Map();
  const spreadsheets = new Map([[ss.getId(), ss]]);

  ctx.SpreadsheetApp = {
    getActiveSpreadsheet: () => (options.unbound ? null : ss),
    openById: (id) => {
      const found = spreadsheets.get(id);
      if (!found) throw new Error(`Exception: Unexpected error while getting the method or property openById on object SpreadsheetApp.`);
      return found;
    },
    getUi: () => ui,
  };
  ctx.PropertiesService = properties;
  ctx.CacheService = cache;
  ctx.LockService = lock;
  ctx.Utilities = utilities;
  ctx.UrlFetchApp = urlFetch;
  ctx.MailApp = mail;
  ctx.ScriptApp = scriptApp;
  ctx.ContentService = fakes.makeContentService();
  ctx.FormApp = {
    openById: (id) => {
      const form = forms.get(id);
      if (!form) throw new Error(`Exception: No item with the given ID could be found, or you do not have permission to access it.`);
      return form;
    },
  };
  ctx.console = {
    log: (...a) => logs.log.push(a.join(' ')),
    warn: (...a) => logs.warn.push(a.join(' ')),
    error: (...a) => logs.error.push(a.join(' ')),
  };

  sourceFiles().forEach(({ name, code }) => {
    new vm.Script(code, { filename: `src/${name}` }).runInContext(ctx);
  });

  const env = {
    ctx, clock, logs, ss, ui, cache: cache._cache, lock: lock.state, utilities, urlFetch, mail, scriptApp,
    props: properties._props,
    forms,
    /** Calls a global function from src/ and returns a plain (outer-realm) copy of its result. */
    call(name, ...args) {
      if (typeof ctx[name] !== 'function') throw new Error(`no such function in src: ${name}`);
      return plain(ctx[name](...args));
    },
    /** Same, but returns the raw value (for objects with methods, e.g. ContentService output). */
    callRaw(name, ...args) {
      if (typeof ctx[name] !== 'function') throw new Error(`no such function in src: ${name}`);
      return ctx[name](...args);
    },
    setProps(obj) {
      Object.keys(obj).forEach((k) => (obj[k] === null ? env.props.deleteProperty(k) : env.props.setProperty(k, obj[k])));
      return env;
    },
    addSpreadsheet(id, tz) {
      const other = new fakes.FakeSpreadsheet(id, tz);
      spreadsheets.set(id, other);
      return other;
    },
    addForm(id, title) {
      const form = new fakes.FakeForm(id, title);
      forms.set(id, form);
      return form;
    },
  };
  return env;
}

// --- webhook signing, written independently of src/ (Node crypto) ----------------------------

function signPayload(secret, timestamp, body) {
  return crypto.createHmac('sha256', secret).update(timestamp).update('.').update(body).digest('hex');
}

/** Builds the doPost event for a signed delivery. */
function signedEvent({ secret = 'whsec_test', body, now, ageMs = 0, id, type, sig, ts, contentType = 'application/json' }) {
  const timestamp = ts !== undefined ? ts : new Date(now - ageMs).toISOString();
  const parameter = { ts: timestamp, sig: sig !== undefined ? sig : signPayload(secret, timestamp, body) };
  if (id !== undefined) parameter.id = id;
  if (type !== undefined) parameter.type = type;
  return { parameter, postData: { contents: body, type: contentType, length: body.length } };
}

module.exports = { createEnv, plain, signPayload, signedEvent, sourceFiles, SRC, LOAD_ORDER };
