'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { createEnv, signPayload } = require('./harness');

// Vectors produced once by webhook-kit's signPayload (src/sign.ts), embedded so this test has no
// runtime dependency on the other repo.
const WEBHOOK_KIT_VECTORS = [
  { secret: 'whsec_test_secret', ts: '2026-06-01T12:00:00.000Z', body: '{"id":"evt_1","type":"order.created"}', expected: 'b756eaf18c51802ee644074734a5a93743b6062da7e1c70f0c372f7570456ada' },
  { secret: 'whsec_test_secret', ts: '2026-06-01T12:00:00.000Z', body: '{"id":"evt_2","note":"café ☕ naïve"}', expected: 'a45f29e16fd2af94218f232d76087e5362d3ac2087a9973afed76f6128c6f3b8' },
  { secret: 'k', ts: '2026-06-01T12:00:00.000Z', body: '', expected: '6d26b252feca52059e438a6b6f1800168bda5d1ab8fabcc460f3f4348e19322c' },
];

test('HMAC hex matches webhook-kit byte for byte (ASCII, UTF-8 and empty bodies)', () => {
  const env = createEnv();
  WEBHOOK_KIT_VECTORS.forEach((v) => {
    assert.equal(env.call('hmacSha256Hex_', `${v.ts}.${v.body}`, v.secret), v.expected);
    assert.equal(signPayload(v.secret, v.ts, v.body), v.expected, 'harness signer agrees too');
  });
});

test('selfTestHmac passes against the same vectors and logs PASS', () => {
  const env = createEnv();
  assert.equal(env.call('selfTestHmac'), true);
  assert.match(env.logs.log.at(-1), /^PASS selfTestHmac/);
});

test('bytesToHex_ handles Apps Script signed bytes (negative numbers) and pads single digits', () => {
  const env = createEnv();
  assert.equal(env.call('bytesToHex_', [-1, 0, 1, 15, 16, 127, -128]), 'ff00010f107f80');
});

test('constantTimeEquals_: equal, different, different length, empty', () => {
  const env = createEnv();
  assert.equal(env.call('constantTimeEquals_', 'abc', 'abc'), true);
  assert.equal(env.call('constantTimeEquals_', 'abc', 'abd'), false);
  assert.equal(env.call('constantTimeEquals_', 'abc', 'abcd'), false);
  assert.equal(env.call('constantTimeEquals_', 'abc', 'ab'), false);
  assert.equal(env.call('constantTimeEquals_', '', ''), true);
  assert.equal(env.call('constantTimeEquals_', 'a\u0000', 'a'), false, 'NUL padding must not equal a shorter string');
});

test('parseIsoTimestamp_: strict ISO with zone; rejects rolled-over dates and loose formats', () => {
  const env = createEnv();
  const p = (s) => env.call('parseIsoTimestamp_', s);
  assert.equal(p('2026-06-01T12:00:00.000Z'), Date.parse('2026-06-01T12:00:00.000Z'));
  assert.equal(p('2026-06-01T12:00:00Z'), Date.parse('2026-06-01T12:00:00Z'));
  assert.equal(p('2026-06-01T12:00:00+05:30'), Date.parse('2026-06-01T12:00:00+05:30'));
  assert.equal(p('2026-02-30T00:00:00Z'), null);
  assert.equal(p('2026-13-01T00:00:00Z'), null);
  assert.equal(p('2026-06-01T24:00:00Z'), null);
  assert.equal(p('2026-06-01'), null);
  assert.equal(p('2026-06-01T12:00:00'), null, 'no time zone');
  assert.equal(p('1'), null);
  assert.equal(p(1780315200000), null);
  assert.equal(p(''), null);
  assert.equal(p('2028-02-29T00:00:00Z') !== null, true, 'leap day is real');
  assert.equal(p('2027-02-29T00:00:00Z'), null);
});

test('normalizeText_ collapses all whitespace kinds and removes zero-width characters', () => {
  const env = createEnv();
  assert.equal(env.call('normalizeText_', '  a  b \t c\n d​ '), 'a b c d');
  assert.equal(env.call('normalizeText_', null), '');
  assert.equal(env.call('normalizeText_', 5), '5');
});

test('toText_: dates become yyyy-MM-dd, with the time only when it is not midnight', () => {
  const env = createEnv();
  const date1 = new env.ctx.Date('2026-03-04T00:00:00Z');
  const date2 = new env.ctx.Date('2026-03-04T15:30:05Z');
  assert.equal(env.ctx.toText_(date1, 'UTC'), '2026-03-04');
  assert.equal(env.ctx.toText_(date2, 'UTC'), '2026-03-04 15:30:05');
  assert.equal(env.ctx.toText_(date2, 'Asia/Kolkata'), '2026-03-04 21:00:05');
  assert.equal(env.ctx.toText_(new env.ctx.Date('nope'), 'UTC'), '');
  assert.equal(env.ctx.toText_(null, 'UTC'), '');
  assert.equal(env.ctx.toText_(12, 'UTC'), '12');
});

test('getProp_/getIntProp_/getListProp_/getSecret_: blanks fall back, bad numbers throw, secrets are not trimmed', () => {
  const env = createEnv();
  env.setProps({ A: '  x  ', B: '   ', N: '42', BAD: '4x2', L: 'a@x.com, b@x.com;c@x.com  d@x.com', S: ' s3cret ' });
  assert.equal(env.call('getProp_', 'A', 'f'), 'x');
  assert.equal(env.call('getProp_', 'B', 'f'), 'f');
  assert.equal(env.call('getProp_', 'MISSING', 'f'), 'f');
  assert.equal(env.call('getIntProp_', 'N', 1), 42);
  assert.equal(env.call('getIntProp_', 'MISSING', 7), 7);
  assert.throws(() => env.ctx.getIntProp_('BAD', 1), /BAD must be a whole number/);
  assert.deepEqual(env.call('getListProp_', 'L'), ['a@x.com', 'b@x.com', 'c@x.com', 'd@x.com']);
  assert.deepEqual(env.call('getListProp_', 'MISSING'), []);
  assert.equal(env.call('getSecret_', 'S'), ' s3cret ');
  assert.equal(env.call('getSecret_', 'MISSING'), null);
});

test('getSpreadsheet_: SPREADSHEET_ID wins, else the bound sheet, else a clear error', () => {
  const env = createEnv();
  const other = env.addSpreadsheet('other-id');
  assert.equal(env.ctx.getSpreadsheet_(), env.ss);
  env.setProps({ SPREADSHEET_ID: 'other-id' });
  assert.equal(env.ctx.getSpreadsheet_(), other);
  env.setProps({ SPREADSHEET_ID: 'does-not-exist' });
  assert.throws(() => env.ctx.getSpreadsheet_(), /openById/);
  const unbound = createEnv({ unbound: true });
  assert.throws(() => unbound.ctx.getSpreadsheet_(), /No spreadsheet/);
});

test('withScriptLock_ releases the lock when fn throws, and reports when it cannot lock', () => {
  const env = createEnv();
  env.ctx.__boom = () => { throw new Error('boom'); };
  assert.throws(() => env.ctx.withScriptLock_(100, env.ctx.__boom), /boom/);
  assert.equal(env.lock.held, false);
  assert.equal(env.lock.acquired, 1);
  assert.equal(env.lock.releases, 1);

  env.lock.heldByOther = true;
  const t0 = env.clock.now;
  assert.deepEqual(env.call('withScriptLock_', 250, () => 'never'), { locked: false });
  assert.equal(env.clock.now - t0, 250);
  env.lock.heldByOther = false;
  assert.deepEqual(env.call('withScriptLock_', 250, () => 'v'), { locked: true, value: 'v' });
});

test('ensureColumns_ appends only the missing headers, keeps existing cells, and keeps rows aligned', () => {
  const env = createEnv();
  const sheet = env.ss.addSheet('T', [['ID', 'Sync Status'], [1, 'x'], [2, '']]);
  const table = env.ctx.readTable_(sheet);
  const cols = env.call('ensureColumns_', sheet, table, ['sync status', 'Sync Error']);
  assert.deepEqual(cols, { 'sync status': 1, 'Sync Error': 2 });
  assert.deepEqual(sheet.values(), [['ID', 'Sync Status', 'Sync Error'], [1, 'x', ''], [2, '', '']]);
  assert.deepEqual(table.rows.map((r) => r.length), [3, 3]);
});

test('escapeHtml_, isEmail_, oneLine_', () => {
  const env = createEnv();
  assert.equal(env.call('escapeHtml_', `<a href="x">&'</a>`), '&lt;a href=&quot;x&quot;&gt;&amp;&#39;&lt;/a&gt;');
  assert.equal(env.call('isEmail_', 'a@b.co'), true);
  ['a@b', 'a b@c.com', '@b.com', 'a@@b.com', 'a@b..com', '.a@b.com', 'a@b.c'].forEach((bad) => assert.equal(env.call('isEmail_', bad), false, bad));
  assert.equal(env.call('oneLine_', 'a\n\n b   c', 100), 'a b c');
  assert.equal(env.call('oneLine_', 'x'.repeat(50), 10), 'xxxxxxx...');
});

test('writeTextBlock_ keeps tricky strings as text; the fake proves the unformatted write would not', () => {
  const env = createEnv();
  const sheet = env.ss.addSheet('T');
  env.ctx.writeTextBlock_(sheet, 1, 1, [['007', '=1+1', '2026-01-05', 'TRUE']]);
  assert.deepEqual(sheet.values(), [['007', '=1+1', '2026-01-05', 'TRUE']]);
  assert.equal(sheet.isFormula(1, 2), false);
  sheet.getRange(2, 1, 1, 4).setValues([['007', '=1+1', '2026-01-05', 'TRUE']]);
  const raw = sheet.values()[1];
  assert.equal(raw[0], 7);
  assert.equal(sheet.isFormula(2, 2), true);
  assert.equal(Object.prototype.toString.call(raw[2]), '[object Date]');
  assert.equal(raw[3], true);
});
