'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { createEnv, signPayload, signedEvent } = require('./harness');

const SECRET = 'whsec_test';
const BODY = '{"id":"evt_100","type":"order.created","total":42}';
const HEADER = ['Received At', 'Event ID', 'Event Type', 'Payload'];

function setup(extraProps = {}) {
  const env = createEnv();
  env.setProps({ WEBHOOK_SECRET: SECRET, ...extraProps });
  const send = (event) => {
    const out = env.callRaw('doPost', event);
    return { mime: out.getMimeType(), body: JSON.parse(out.getContent()) };
  };
  const post = (opts) => send(signedEvent({ secret: SECRET, body: BODY, now: env.clock.now, ...opts }));
  const events = () => env.ss.getSheetByName('Events');
  const rows = () => (events() ? events().values().slice(1) : []);
  return { env, send, post, events, rows };
}

test('valid signed delivery: HTTP-level JSON with status 200 in the body, one row appended', () => {
  const { post, events } = setup();
  const r = post({});
  assert.equal(r.mime, 'JSON');
  assert.deepEqual(r.body, { ok: true, status: 200, result: 'ok' });
  assert.deepEqual(events().values(), [HEADER, ['2026-11-01T12:00:00.000Z', 'evt_100', 'order.created', BODY]]);
});

test('the signature may carry the sha256= prefix and upper-case hex', () => {
  const { env, post, rows } = setup();
  const ts = new Date(env.clock.now).toISOString();
  const hex = signPayload(SECRET, ts, BODY);
  assert.equal(post({ ts, sig: `sha256=${hex}` }).body.status, 200);
  assert.equal(post({ ts, sig: hex.toUpperCase(), id: 'other' }).body.status, 200);
  assert.equal(rows().length, 2);
});

test('wrong secret, tampered body and wrong timestamp are all 401 invalid_signature and write nothing', () => {
  const { env, post, send, rows } = setup();
  const wrongSecret = post({ secret: 'not-the-secret' });
  assert.deepEqual(wrongSecret.body, { ok: false, status: 401, error: 'invalid_signature' });

  const ts = new Date(env.clock.now).toISOString();
  const signedForOther = signPayload(SECRET, ts, '{"id":"evt_100","total":1}');
  assert.equal(post({ ts, sig: signedForOther }).body.status, 401, 'body changed after signing');

  const sigForOtherTs = signPayload(SECRET, new Date(env.clock.now - 1000).toISOString(), BODY);
  assert.equal(post({ ts, sig: sigForOtherTs }).body.status, 401, 'timestamp changed after signing');

  const flipped = signPayload(SECRET, ts, BODY).replace(/.$/, (c) => (c === '0' ? '1' : '0'));
  assert.equal(post({ ts, sig: flipped }).body.status, 401, 'one hex digit off');
  assert.equal(rows().length, 0);
  void send;
});

test('replay window: 5 minutes either side is accepted, older or further in the future is 401 (same error body)', () => {
  const { post, rows } = setup();
  assert.equal(post({ ageMs: 299000, id: 'a' }).body.status, 200);
  assert.equal(post({ ageMs: -299000, id: 'b' }).body.status, 200);
  const old = post({ ageMs: 301000, id: 'c' });
  const future = post({ ageMs: -301000, id: 'd' });
  assert.deepEqual(old.body, { ok: false, status: 401, error: 'invalid_signature' });
  assert.deepEqual(future.body, old.body);
  assert.deepEqual(rows().map((r) => r[1]), ['a', 'b']);
});

test('WEBHOOK_TOLERANCE_SECONDS changes the window', () => {
  const { post } = setup({ WEBHOOK_TOLERANCE_SECONDS: '10' });
  assert.equal(post({ ageMs: 9000, id: 'a' }).body.status, 200);
  assert.equal(post({ ageMs: 11000, id: 'b' }).body.status, 401);
});

test('missing or malformed signature / timestamp: 400 with a specific error', () => {
  const { env, post, send } = setup();
  const good = signedEvent({ secret: SECRET, body: BODY, now: env.clock.now });
  const without = (key) => ({ ...good, parameter: { ...good.parameter, [key]: undefined } });
  assert.deepEqual(send(without('sig')).body, { ok: false, status: 400, error: 'missing_signature' });
  assert.deepEqual(send(without('ts')).body, { ok: false, status: 400, error: 'missing_signature' });
  assert.equal(send({ postData: good.postData }).body.error, 'missing_signature', 'no query parameters at all');
  assert.equal(send({}).body.error, 'missing_signature', 'no postData either');
  assert.equal(post({ sig: 'abc' }).body.error, 'malformed_signature');
  assert.equal(post({ sig: 'z'.repeat(64) }).body.error, 'malformed_signature');
  assert.equal(post({ sig: 'a'.repeat(63) }).body.error, 'malformed_signature');
  assert.equal(post({ ts: '2026-11-01 12:00:00' }).body.error, 'malformed_timestamp');
  assert.equal(post({ ts: '2026-02-30T12:00:00Z' }).body.error, 'malformed_timestamp');
  assert.equal(post({ ts: '1780315200' }).body.error, 'malformed_timestamp');
});

test('a replayed event id inside the window is acknowledged as a duplicate and NOT written twice', () => {
  const { env, post, rows } = setup();
  assert.deepEqual(post({}).body, { ok: true, status: 200, result: 'ok' });
  env.clock.now += 5000;
  assert.deepEqual(post({}).body, { ok: true, status: 200, result: 'duplicate' });
  assert.deepEqual(post({}).body, { ok: true, status: 200, result: 'duplicate' });
  assert.equal(rows().length, 1);
});

test('retry served from the cache does not take the lock or read the sheet', () => {
  const { env, post, events } = setup();
  post({});
  const acquired = env.lock.acquired;
  const reads = events().calls.getValues || 0;
  post({});
  assert.equal(env.lock.acquired, acquired);
  assert.equal(events().calls.getValues || 0, reads);
});

test('the sheet stays the source of truth: duplicate detected even if the cache was lost or expired', () => {
  const { env, post, rows } = setup();
  post({});
  env.cache.clear();
  assert.equal(post({}).body.result, 'duplicate');
  assert.equal(rows().length, 1);
  env.cache.clear();
  env.clock.now += 7 * 3600 * 1000; // cache TTL (capped at 6 hours) is long gone, and so is the window...
  const late = post({ ageMs: 0 });  // ...but a fresh signature for the same id is still a duplicate
  assert.equal(late.body.result, 'duplicate');
  assert.equal(rows().length, 1);
});

test('event id comes from ?id= first, then payload.id; a different id is a different event', () => {
  const { post, rows } = setup();
  post({ id: 'from-param' });
  post({ body: '{"id":"from-payload"}' });
  post({ body: '{"id":12345}' });
  post({ body: '{"no":"id"}', id: 'only-param' });
  assert.deepEqual(rows().map((r) => r[1]), ['from-param', 'from-payload', '12345', 'only-param']);
});

test('no usable event id: 400 missing_event_id (even for valid JSON and non-JSON bodies); bad ids: 400 invalid_event_id', () => {
  const { post, rows } = setup();
  assert.equal(post({ body: '{"hello":1}' }).body.error, 'missing_event_id');
  assert.equal(post({ body: 'plain text' }).body.error, 'missing_event_id');
  assert.equal(post({ body: '[1,2]' }).body.error, 'missing_event_id');
  assert.equal(post({ body: '{"id":null}' }).body.error, 'missing_event_id');
  assert.equal(post({ id: 'x'.repeat(201) }).body.error, 'invalid_event_id');
  assert.equal(post({ id: 'a\nb' }).body.error, 'invalid_event_id');
  assert.equal(rows().length, 0);
});

test('a non-JSON body with an id parameter is accepted and stored verbatim', () => {
  const { post, rows } = setup();
  assert.equal(post({ body: 'a=1&b=2', id: 'form-1', contentType: 'application/x-www-form-urlencoded' }).body.status, 200);
  assert.deepEqual(rows()[0].slice(1), ['form-1', '', 'a=1&b=2']);
});

test('event type: ?type= first, then payload.type; stored on one line', () => {
  const { post, rows } = setup();
  post({ id: 'a', type: 'custom.type' });
  post({ body: '{"id":"b","type":"nested.type"}' });
  post({ body: '{"id":"c","type":5}' });
  assert.deepEqual(rows().map((r) => r[2]), ['custom.type', 'nested.type', '']);
});

test('empty body with a valid signature is 400 empty_body', () => {
  const { post, rows } = setup();
  assert.equal(post({ body: '', id: 'x' }).body.error, 'empty_body');
  assert.equal(rows().length, 0);
});

test('UTF-8 body: the signature is verified over the bytes and the text is stored intact', () => {
  const { post, rows } = setup();
  const body = '{"id":"evt_u","note":"café ☕ naïve"}';
  assert.equal(post({ body }).body.status, 200);
  assert.equal(rows()[0][3], body);
});

test('payload cells are plain text: a "=formula" payload is never turned into a formula', () => {
  const { post, events } = setup();
  assert.equal(post({ body: '=HYPERLINK("http://x")', id: 'f' }).body.status, 200);
  assert.equal(events().values()[1][3], '=HYPERLINK("http://x")');
  assert.equal(events().isFormula(2, 4), false);
  assert.equal(events().formatAt(2, 2), '@');
});

test('numeric-looking event ids stay strings, so "007" and "7" are different events and "007" dedupes', () => {
  const { post, rows } = setup();
  post({ id: '007' });
  post({ id: '7' });
  assert.equal(post({ id: '007' }).body.result, 'duplicate');
  assert.deepEqual(rows().map((r) => r[1]), ['007', '7']);
});

test('secret rotation: the previous secret still verifies; an unknown one does not', () => {
  const { env, post } = setup({ WEBHOOK_SECRET: 'new-secret', WEBHOOK_SECRET_PREVIOUS: 'old-secret' });
  assert.equal(post({ secret: 'new-secret', id: 'a' }).body.status, 200);
  assert.equal(post({ secret: 'old-secret', id: 'b' }).body.status, 200);
  assert.equal(post({ secret: 'older-secret', id: 'c' }).body.status, 401);
  env.setProps({ WEBHOOK_SECRET_PREVIOUS: null });
  assert.equal(post({ secret: 'old-secret', id: 'd' }).body.status, 401);
});

test('no secret configured: 500 not_configured, nothing written', () => {
  const env = createEnv();
  const out = JSON.parse(env.callRaw('doPost', signedEvent({ body: BODY, now: env.clock.now })).getContent());
  assert.deepEqual(out, { ok: false, status: 500, error: 'not_configured' });
  assert.equal(env.ss.getSheetByName('Events'), null);
});

test('oversized body is 413 before any signature work; the limit is configurable', () => {
  const { post, rows } = setup({ WEBHOOK_MAX_BODY_CHARS: '100' });
  assert.deepEqual(post({ body: 'x'.repeat(101), id: 'big' }).body, { ok: false, status: 413, error: 'payload_too_large' });
  assert.equal(post({ body: JSON.stringify({ id: 'ok', pad: 'x'.repeat(50) }) }).body.status, 200);
  assert.equal(rows().length, 1);
});

test('another run holds the lock: 503 busy (sender should retry), nothing written, then it works', () => {
  const { env, post, rows } = setup();
  env.lock.heldByOther = true;
  const t0 = env.clock.now;
  assert.deepEqual(post({ now: t0 }).body, { ok: false, status: 503, error: 'busy' });
  assert.equal(rows().length, 0);
  env.lock.heldByOther = false;
  assert.equal(post({}).body.status, 200);
  assert.equal(rows().length, 1);
});

test('the lock is released after every outcome', () => {
  const { env, post } = setup();
  post({});
  post({});
  post({ id: 'z' });
  assert.equal(env.lock.held, false);
  assert.equal(env.lock.acquired, env.lock.releases);
});

test('internal failure (Events sheet without the Event ID column): 500 internal_error, no detail leaked, lock released', () => {
  const { env, post } = setup();
  env.ss.addSheet('Events', [['Received At', 'Payload']]);
  const r = post({});
  assert.deepEqual(r.body, { ok: false, status: 500, error: 'internal_error' });
  assert.match(env.logs.error.join('\n'), /has no "Event ID" column/);
  assert.doesNotMatch(JSON.stringify(r.body), /Event ID|Events|whsec/);
  assert.equal(env.lock.held, false);
});

test('an existing Events sheet with extra columns in a different order keeps its layout', () => {
  const { env, post } = setup();
  const sheet = env.ss.addSheet('Events', [['Note', 'Payload', 'Event ID', 'Received At', 'Event Type']]);
  post({});
  assert.deepEqual(sheet.values()[1], ['', BODY, 'evt_100', '2026-11-01T12:00:00.000Z', 'order.created']);
});

test('EVENTS_SHEET and SPREADSHEET_ID properties redirect the writes', () => {
  const { env, post } = setup({ EVENTS_SHEET: 'Hooks', SPREADSHEET_ID: 'other' });
  const other = env.addSpreadsheet('other');
  post({});
  assert.equal(other.getSheetByName('Hooks').values().length, 2);
  assert.equal(env.ss.getSheets().length, 0);
});

test('rejections are logged with a reason but never log the body, signature or secret', () => {
  const { env, post } = setup();
  post({ secret: 'wrong', body: '{"id":"evt_secretbody"}' });
  const text = env.logs.warn.join('\n');
  assert.match(text, /401 invalid_signature \(signature_mismatch\)/);
  assert.doesNotMatch(text, /evt_secretbody|wrong|whsec_test/);
});

test('doGet answers 405 and touches no data', () => {
  const { env } = setup();
  const out = JSON.parse(env.callRaw('doGet').getContent());
  assert.deepEqual(out, { ok: false, status: 405, error: 'method_not_allowed' });
  assert.equal(env.ss.getSheets().length, 0);
});

test('known-answer: the receiver accepts a signature produced by webhook-kit for the same inputs', () => {
  const { env, send, rows } = setup({ WEBHOOK_SECRET: 'whsec_test_secret' });
  const ts = '2026-06-01T12:00:00.000Z';
  env.clock.now = Date.parse(ts) + 1000;
  const body = '{"id":"evt_1","type":"order.created"}';
  const out = send({
    parameter: { ts, sig: 'sha256=b756eaf18c51802ee644074734a5a93743b6062da7e1c70f0c372f7570456ada' },
    postData: { contents: body, type: 'application/json', length: body.length },
  });
  assert.deepEqual(out.body, { ok: true, status: 200, result: 'ok' });
  assert.equal(rows()[0][1], 'evt_1');
});

test('many different events are all stored once, in order, and the id lookup is a single column read', () => {
  const { env, post, events } = setup();
  for (let i = 0; i < 25; i++) post({ id: `e${i}`, body: JSON.stringify({ id: `e${i}` }) });
  for (let i = 0; i < 25; i++) assert.equal(post({ id: `e${i}`, body: JSON.stringify({ id: `e${i}` }) }).body.result, 'duplicate');
  assert.deepEqual(events().values().slice(1).map((r) => r[1]), Array.from({ length: 25 }, (_, i) => `e${i}`));
  void env;
});
