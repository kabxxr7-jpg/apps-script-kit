'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const { execFileSync, spawnSync } = require('node:child_process');
const { createEnv } = require('./harness');
const { build, parseArgs, shellQuote } = require('../tools/sign');

const SIGN = path.join(__dirname, '..', 'tools', 'sign.js');

test('tools/sign.js produces a request the receiver accepts (end to end through doPost)', () => {
  const env = createEnv();
  env.setProps({ WEBHOOK_SECRET: 'cli-secret' });
  const body = '{"id":"evt_cli","type":"x"}';
  const r = build({ body, id: 'evt_cli' }, 'cli-secret', env.clock.now);
  const out = env.callRaw('doPost', { parameter: { ts: r.ts, sig: r.sig, id: 'evt_cli' }, postData: { contents: body } });
  assert.deepEqual(JSON.parse(out.getContent()), { ok: true, status: 200, result: 'ok' });
});

test('tools/sign.js matches the webhook-kit known answer', () => {
  const r = build({ body: '{"id":"evt_1","type":"order.created"}', ts: '2026-06-01T12:00:00.000Z' }, 'whsec_test_secret', 0);
  assert.equal(r.sig, 'b756eaf18c51802ee644074734a5a93743b6062da7e1c70f0c372f7570456ada');
});

test('CLI: reads the secret from the environment only and prints a usable query string and curl line', () => {
  const out = execFileSync(process.execPath, [SIGN, '--body', '{"id":"e1"}', '--id', 'e1', '--ts', '2026-06-01T12:00:00.000Z', '--url', 'https://example.test/exec'], {
    env: { ...process.env, WEBHOOK_SECRET: 's' },
    encoding: 'utf8',
  });
  assert.match(out, /^query: ts=2026-06-01T12%3A00%3A00\.000Z&sig=[0-9a-f]{64}&id=e1/);
  assert.match(out, /printf '%s' '\{"id":"e1"\}' \| curl -L -H .* 'https:\/\/example\.test\/exec\?ts=/);
});

test('CLI: missing secret or body exits non-zero with a message', () => {
  const noSecret = spawnSync(process.execPath, [SIGN, '--body', 'x'], { env: { PATH: process.env.PATH }, encoding: 'utf8' });
  assert.equal(noSecret.status, 1);
  assert.match(noSecret.stderr, /WEBHOOK_SECRET/);
  const noBody = spawnSync(process.execPath, [SIGN], { env: { ...process.env, WEBHOOK_SECRET: 's' }, encoding: 'utf8' });
  assert.equal(noBody.status, 1);
  assert.match(noBody.stderr, /--body/);
});

test('parseArgs rejects stray arguments and flags without values', () => {
  assert.throws(() => parseArgs(['oops']), /unexpected argument/);
  assert.throws(() => parseArgs(['--body']), /needs a value/);
});

test('the printed shell command delivers the exact signed bytes (quotes, newlines, UTF-8, no trailing newline)', () => {
  const bodies = ['{"id":"e1"}', `{"note":"it's \"quoted\"\\n"}`, '{"a":\n  1}', '{"note":"café ☕"}', ''];
  bodies.forEach((body) => {
    const out = execFileSync('sh', ['-c', `printf '%s' ${shellQuote(body)}`], { encoding: 'utf8' });
    assert.equal(out, body);
  });
});
