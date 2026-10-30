'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { createEnv } = require('./harness');

test('selfTestTextCells passes against the fake, leaves no temporary sheet behind, and reports the unformatted behaviour', () => {
  const env = createEnv();
  env.ss.addSheet('Keep', [['x']]);
  assert.equal(env.call('selfTestTextCells'), true);
  assert.deepEqual(env.ss.getSheets().map((s) => s.getName()), ['Keep']);
  assert.match(env.logs.log.at(-1), /^PASS selfTestTextCells textFormatBack=\["007","=1\+1","2026-01-05","\+15551234567","TRUE"\] \(info\) unformatted "007" reads back as number 7/);
});

test('selfTestTextCells deletes its sheet even when a check throws', () => {
  const env = createEnv();
  env.ss.addSheet('Keep', [['x']]);
  const original = env.ss.insertSheet.bind(env.ss);
  env.ss.insertSheet = (name) => {
    const sheet = original(name);
    sheet.getRange = () => { throw new Error('boom'); };
    return sheet;
  };
  assert.throws(() => env.ctx.selfTestTextCells(), /boom/);
  assert.deepEqual(env.ss.getSheets().map((s) => s.getName()), ['Keep']);
});

test('selfTestHmac reports FAIL if the runtime computed something different', () => {
  const env = createEnv();
  env.ctx.Utilities = { computeHmacSha256Signature: () => [1, 2, 3] };
  assert.equal(env.call('selfTestHmac'), false);
  assert.match(env.logs.log.at(-1), /^FAIL selfTestHmac/);
});
