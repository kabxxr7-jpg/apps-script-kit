'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { createEnv } = require('./harness');

const HEADER = ['ID', 'Customer', 'Total', 'Ordered'];
const ok = (extra = {}) => ({ code: 200, body: '{"ok":true}', ...extra });

function setup(rows, props = {}) {
  const env = createEnv();
  env.setProps({ API_URL: 'https://api.example.com/orders', API_TOKEN: 'tok_abc', ...props });
  const sheet = env.ss.addSheet('Orders', [HEADER, ...rows]);
  const run = () => env.call('syncRows');
  const status = () => sheet.values().slice(1).map((r) => r[4]);
  return { env, sheet, run, status };
}

const ROWS3 = [
  ['A-1', 'Ann', 10, new Date('2026-10-30T00:00:00Z')],
  ['A-2', 'Bo', 20.5, ''],
  ['A-3', 'Cy', 30, ''],
];

test('happy path: every row is POSTed once and marked synced with a timestamp', () => {
  const { env, sheet, run } = setup(ROWS3);
  env.urlFetch.enqueue(ok(), ok(), ok());
  const r = run();
  assert.deepEqual(r, { status: 'done', processed: 3, synced: 3, failed: 0, rejected: 0, skipped: 0, remaining: 0 });
  assert.equal(env.urlFetch.requests.length, 3);
  assert.deepEqual(sheet.values()[0], [...HEADER, 'Sync Status', 'Sync Error', 'Synced At']);
  assert.deepEqual(sheet.values().slice(1).map((r) => r.slice(4)), Array(3).fill(['synced', '', '2026-11-01T12:00:00.000Z']));
});

test('request shape: JSON body keyed by header, dates ISO, blanks null, bearer token, idempotency key = ID', () => {
  const { env, run } = setup(ROWS3);
  env.urlFetch.enqueue(ok(), ok(), ok());
  run();
  const [first, second] = env.urlFetch.requests;
  assert.equal(first.url, 'https://api.example.com/orders');
  assert.equal(first.params.method, 'post');
  assert.equal(first.params.contentType, 'application/json');
  assert.equal(first.params.muteHttpExceptions, true);
  assert.deepEqual(JSON.parse(first.params.payload), { ID: 'A-1', Customer: 'Ann', Total: 10, Ordered: '2026-10-30' });
  assert.deepEqual(JSON.parse(second.params.payload), { ID: 'A-2', Customer: 'Bo', Total: 20.5, Ordered: null });
  assert.equal(first.params.headers.Authorization, 'Bearer tok_abc');
  assert.equal(first.params.headers['Idempotency-Key'], 'A-1');
  assert.equal(second.params.headers['Idempotency-Key'], 'A-2');
});

test('the token is optional, and API_URL must be https', () => {
  const noToken = setup(ROWS3.slice(0, 1), { API_TOKEN: null });
  noToken.env.urlFetch.enqueue(ok());
  noToken.run();
  assert.equal('Authorization' in noToken.env.urlFetch.requests[0].params.headers, false);

  const insecure = setup(ROWS3, { API_URL: 'http://api.example.com/orders' });
  assert.throws(() => insecure.env.ctx.syncRows(), /must start with https/);
  assert.equal(insecure.env.urlFetch.requests.length, 0);

  const missing = setup(ROWS3, { API_URL: null });
  assert.throws(() => missing.env.ctx.syncRows(), /Set the API_URL/);
});

test('transient failure then success: 500, 503, then 200 is retried with doubling backoff and ends synced', () => {
  const { env, run, sheet } = setup(ROWS3.slice(0, 1));
  env.urlFetch.enqueue({ code: 500, body: 'oops' }, { code: 503, body: 'busy' }, ok());
  const r = run();
  assert.equal(r.synced, 1);
  assert.equal(env.urlFetch.requests.length, 3);
  assert.deepEqual(env.utilities.sleeps, [1000, 2000]);
  assert.equal(sheet.values()[1][4], 'synced');
  assert.equal(sheet.values()[1][5], '', 'old error text is cleared');
});

test('network exceptions are retried like 5xx', () => {
  const { env, run } = setup(ROWS3.slice(0, 1));
  env.urlFetch.enqueue(new Error('DNS error: api.example.com'), ok());
  assert.equal(run().synced, 1);
  assert.deepEqual(env.utilities.sleeps, [1000]);
});

test('429 honours Retry-After (seconds, case-insensitive header), capped at 60s', () => {
  const { env, run } = setup(ROWS3.slice(0, 1));
  env.urlFetch.enqueue({ code: 429, headers: { 'retry-after': '7' } }, { code: 429, headers: { 'Retry-After': '9999' } }, ok());
  run();
  assert.deepEqual(env.utilities.sleeps, [7000, 60000]);
});

test('backoff is capped at 30 seconds and SYNC_MAX_ATTEMPTS / SYNC_BACKOFF_MS are honoured', () => {
  const { env, run, sheet } = setup(ROWS3.slice(0, 1), { SYNC_MAX_ATTEMPTS: '8', SYNC_BACKOFF_MS: '10000' });
  for (let i = 0; i < 8; i++) env.urlFetch.enqueue({ code: 500, body: 'x' });
  const r = run();
  assert.equal(r.failed, 1);
  assert.equal(env.urlFetch.requests.length, 8);
  assert.deepEqual(env.utilities.sleeps, [10000, 20000, 30000, 30000, 30000, 30000, 30000]);
  assert.match(sheet.values()[1][5], /\[8 attempts\]$/);
});

test('retries exhausted: row is "failed" with the last error; no sleep after the last attempt', () => {
  const { env, run, sheet } = setup(ROWS3.slice(0, 1));
  env.urlFetch.enqueue({ code: 502, body: 'bad gateway' }, { code: 502, body: 'bad gateway' }, { code: 502, body: 'bad gateway' }, { code: 502, body: 'upstream\nreset' });
  const r = run();
  assert.equal(r.failed, 1);
  assert.equal(sheet.values()[1][4], 'failed');
  assert.equal(sheet.values()[1][5], 'HTTP 502: upstream reset [4 attempts]');
  assert.equal(sheet.values()[1][6], '', 'no Synced At');
  assert.equal(env.utilities.sleeps.length, 3);
});

test('permanent client error (422) is NOT retried, the row is "rejected", and the run carries on', () => {
  const { env, run, sheet } = setup(ROWS3);
  env.urlFetch.enqueue({ code: 422, body: '{"error":"total must be positive"}' }, ok(), ok());
  const r = run();
  assert.deepEqual([r.synced, r.rejected, r.failed], [2, 1, 0]);
  assert.equal(env.urlFetch.requests.length, 3);
  assert.equal(sheet.values()[1][4], 'rejected');
  assert.equal(sheet.values()[1][5], 'HTTP 422: {"error":"total must be positive"}');
  assert.deepEqual(env.utilities.sleeps, []);
});

test('401/403/404 and unexpected redirects are permanent too; 408 is transient', () => {
  const { env, run } = setup(ROWS3);
  env.urlFetch.enqueue({ code: 401 }, { code: 404 }, { code: 302, headers: { Location: 'https://x' } });
  assert.equal(run().rejected, 3);
  const second = setup(ROWS3.slice(0, 1));
  second.env.urlFetch.enqueue({ code: 408 }, ok());
  assert.equal(second.run().synced, 1);
});

test('resumable: a second run retries "failed" rows, skips "synced" and "rejected", and never re-sends them', () => {
  const { env, run, sheet } = setup(ROWS3);
  env.urlFetch.enqueue(ok(), { code: 400, body: 'bad' }, { code: 500 }, { code: 500 }, { code: 500 }, { code: 500 });
  const first = run();
  assert.deepEqual([first.synced, first.rejected, first.failed], [1, 1, 1]);
  const requestsAfterFirst = env.urlFetch.requests.length;

  env.urlFetch.enqueue(ok());
  const second = run();
  assert.deepEqual(second, { status: 'done', processed: 1, synced: 1, failed: 0, rejected: 0, skipped: 2, remaining: 0 });
  assert.equal(env.urlFetch.requests.length, requestsAfterFirst + 1);
  assert.equal(JSON.parse(env.urlFetch.requests.at(-1).params.payload).ID, 'A-3');
  assert.deepEqual(sheet.values().slice(1).map((r) => r[4]), ['synced', 'rejected', 'synced']);

  // clearing a rejected row's status makes it eligible again
  sheet.getRange(3, 5, 1, 1).setValues([['']]);
  env.urlFetch.enqueue(ok());
  assert.equal(run().synced, 1);
});

test('a finished sheet makes zero API calls on the next run', () => {
  const { env, run } = setup(ROWS3);
  env.urlFetch.enqueue(ok(), ok(), ok());
  run();
  const n = env.urlFetch.requests.length;
  const again = run();
  assert.equal(again.processed, 0);
  assert.equal(again.skipped, 3);
  assert.equal(env.urlFetch.requests.length, n);
});

test('time budget: the run stops cleanly before the limit, reports what is left, and the next run finishes it', () => {
  const rows = Array.from({ length: 10 }, (_, i) => [`R-${i}`, `c${i}`, i, '']);
  const { env, run, sheet } = setup(rows, { SYNC_MAX_RUNTIME_MS: '5000' });
  for (let i = 0; i < 10; i++) env.urlFetch.enqueue(ok({ latencyMs: 1000 }));
  const first = run();
  assert.equal(first.status, 'stopped_early');
  assert.equal(first.synced, 5);
  assert.equal(first.remaining, 5);
  assert.deepEqual(sheet.values().slice(1).map((r) => r[4]), [...Array(5).fill('synced'), ...Array(5).fill('')]);

  const second = run();
  assert.equal(second.status, 'done');
  assert.equal(second.synced, 5);
  const ids = env.urlFetch.requests.map((r) => JSON.parse(r.params.payload).ID);
  assert.deepEqual(ids, Array.from({ length: 10 }, (_, i) => `R-${i}`), 'every record sent exactly once');
});

test('a backoff that would run past the deadline stops retrying; the row stays "failed" for the next run', () => {
  const { env, run, sheet } = setup(ROWS3.slice(0, 1), { SYNC_MAX_RUNTIME_MS: '1500', SYNC_BACKOFF_MS: '1000' });
  env.urlFetch.enqueue({ code: 500 }, { code: 500 }, { code: 500 });
  const r = run();
  assert.equal(r.failed, 1);
  assert.match(sheet.values()[1][5], /\(stopped: out of time\)/);
  assert.equal(env.urlFetch.requests.length, 2, 'attempt 2 happened at t=1s; the 2s wait would pass the deadline');
});

test('error text and status cells are plain text, so server output cannot become a formula', () => {
  const { env, run, sheet } = setup(ROWS3.slice(0, 1));
  env.urlFetch.enqueue({ code: 400, body: '=HYPERLINK("http://evil","x")' });
  run();
  assert.equal(sheet.values()[1][5], 'HTTP 400: =HYPERLINK("http://evil","x")');
  assert.equal(sheet.isFormula(2, 6), false);
  assert.equal(sheet.formatAt(2, 6), '@');
});

test('long error bodies are cut to one short line', () => {
  const { env, run, sheet } = setup(ROWS3.slice(0, 1));
  env.urlFetch.enqueue({ code: 400, body: 'x'.repeat(5000) });
  run();
  assert.ok(sheet.values()[1][5].length < 200);
});

test('blank rows are skipped; rows without an ID column fall back to a sheet:row idempotency key', () => {
  const env = createEnv();
  env.setProps({ API_URL: 'https://api.example.com/x', SYNC_SHEET: 'Orders' });
  const sheet = env.ss.addSheet('Orders', [['Customer', 'Total'], ['Ann', 1], ['', ''], ['Bo', 2]]);
  env.urlFetch.enqueue(ok(), ok());
  const r = env.call('syncRows');
  assert.equal(r.synced, 2);
  assert.equal(r.skipped, 1);
  assert.deepEqual(env.urlFetch.requests.map((q) => q.params.headers['Idempotency-Key']), ['Orders:row2', 'Orders:row4']);
  assert.equal(sheet.values()[2][2], '', 'the blank row got no status');
});

test('SYNC_FIELDS limits the payload; a field that is not a column is a configuration error before anything is sent', () => {
  const { env, run } = setup(ROWS3.slice(0, 1), { SYNC_FIELDS: 'ID, Total' });
  env.urlFetch.enqueue(ok());
  run();
  assert.deepEqual(JSON.parse(env.urlFetch.requests[0].params.payload), { ID: 'A-1', Total: 10 });

  const bad = setup(ROWS3, { SYNC_FIELDS: 'ID,Nope' });
  assert.throws(() => bad.env.ctx.syncRows(), /no such column/);
  assert.equal(bad.env.urlFetch.requests.length, 0);
});

test('existing Sync columns in another position are reused, and the three cells are written separately', () => {
  const env = createEnv();
  env.setProps({ API_URL: 'https://api.example.com/x' });
  const sheet = env.ss.addSheet('Orders', [['ID', 'Sync Status', 'Customer', 'Synced At', 'Sync Error'], ['1', '', 'Ann', '', '']]);
  env.urlFetch.enqueue(ok());
  env.call('syncRows');
  assert.deepEqual(sheet.values()[1], ['1', 'synced', 'Ann', '2026-11-01T12:00:00.000Z', '']);
  assert.deepEqual(JSON.parse(env.urlFetch.requests[0].params.payload), { ID: '1', Customer: 'Ann' });
});

test('another run holds the lock: status busy, nothing sent', () => {
  const { env, run } = setup(ROWS3);
  env.lock.heldByOther = true;
  const r = run();
  assert.equal(r.status, 'busy');
  assert.equal(env.urlFetch.requests.length, 0);
});

test('lock is released even when the sheet is missing', () => {
  const env = createEnv();
  env.setProps({ API_URL: 'https://api.example.com/x', SYNC_SHEET: 'Nope' });
  assert.throws(() => env.ctx.syncRows(), /Sheet "Nope" not found/);
  assert.equal(env.lock.held, false);
});

test('empty or header-only sheet: nothing_to_do and no API calls', () => {
  const env = createEnv();
  env.setProps({ API_URL: 'https://api.example.com/x' });
  env.ss.addSheet('Orders', [['ID', 'Customer']]);
  assert.equal(env.call('syncRows').status, 'nothing_to_do');
  assert.equal(env.urlFetch.requests.length, 0);
});

test('installSyncTrigger creates one 15-minute trigger and replaces rather than duplicates', () => {
  const env = createEnv();
  env.call('installSyncTrigger');
  env.call('installSyncTrigger');
  const triggers = env.scriptApp.getProjectTriggers();
  assert.equal(triggers.length, 1);
  assert.equal(triggers[0].getHandlerFunction(), 'syncRows');
  assert.equal(triggers[0].spec.everyMinutes, 15);
});

test('the token never appears in logs or in sheet cells', () => {
  const { env, run, sheet } = setup(ROWS3);
  env.urlFetch.enqueue({ code: 500, body: 'x' }, { code: 500, body: 'x' }, { code: 500, body: 'x' }, { code: 500, body: 'x' }, ok(), ok());
  run();
  const everything = JSON.stringify([sheet.values(), env.logs]);
  assert.doesNotMatch(everything, /tok_abc/);
});
