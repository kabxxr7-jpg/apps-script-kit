'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { createEnv, sourceFiles, SRC } = require('./harness');

test('every file in src/ is loaded by the harness (nothing is silently untested)', () => {
  const onDisk = fs.readdirSync(SRC).filter((f) => f.endsWith('.js')).sort();
  assert.deepEqual(sourceFiles().map((f) => f.name).sort(), onDisk);
});

test('all files load into one bare global scope with no top-level side effects', () => {
  const env = createEnv();
  assert.equal(env.mail.sent.length, 0);
  assert.equal(env.urlFetch.requests.length, 0);
  assert.equal(env.ss.getSheets().length, 0);
  assert.equal(env.scriptApp.triggers.length, 0);
});

test('no function name is declared twice (Apps Script would silently keep the last one)', () => {
  const seen = new Map();
  const dupes = [];
  sourceFiles().forEach(({ name, code }) => {
    for (const m of code.matchAll(/^function\s+([A-Za-z0-9_$]+)\s*\(/gm)) {
      if (seen.has(m[1])) dupes.push(`${m[1]} (${seen.get(m[1])} and ${name})`);
      seen.set(m[1], name);
    }
  });
  assert.deepEqual(dupes, []);
});

test('only entry points are public; helpers end in an underscore', () => {
  const entryPoints = new Set([
    'dedupeSheet', 'removeDuplicateRows', 'cleanActiveSheet', 'doGet', 'doPost', 'syncRows', 'installSyncTrigger',
    'sendDigest', 'initDigestBaseline', 'installDigestTrigger', 'appendFormResponse', 'backfillFormResponses',
    'installFormTriggers', 'selfTestHmac', 'selfTestTextCells', 'onOpen',
  ]);
  const publicNames = [];
  sourceFiles().forEach(({ code }) => {
    for (const m of code.matchAll(/^function\s+([A-Za-z0-9_$]+)\s*\(/gm)) if (!m[1].endsWith('_')) publicNames.push(m[1]);
  });
  assert.deepEqual(publicNames.filter((n) => !entryPoints.has(n)), [], 'unexpected public function');
  assert.deepEqual([...entryPoints].filter((n) => !publicNames.includes(n)), [], 'documented entry point missing');
});

test('conservative syntax and platform guard: nothing that only Node or newer engines provide', () => {
  const banned = [
    [/\?\./, 'optional chaining'],
    [/\?\?/, 'nullish coalescing'],
    [/\.replaceAll\(/, 'replaceAll'],
    [/\.at\(/, 'Array.prototype.at'],
    [/\basync\b|\bawait\b/, 'async/await'],
    [/\brequire\(|\bimport\s|module\.exports|\bprocess\./, 'Node modules/globals'],
    [/\bsetTimeout\b|\bsetInterval\b|\bBuffer\b|\bTextEncoder\b|\bstructuredClone\b/, 'Node/web-only API'],
    [/(^|[^.\w])fetch\(/, 'global fetch (use UrlFetchApp)'],
    [/\bLogger\.log\b/, 'Logger.log (this kit logs with console)'],
  ];
  sourceFiles().forEach(({ name, code }) => {
    banned.forEach(([re, label]) => assert.ok(!re.test(code), `${name} uses ${label}`));
  });
});

test('src/ is pure ASCII (non-ASCII characters are written as \\u escapes so no editor or encoding can change them)', () => {
  sourceFiles().forEach(({ name, code }) => {
    const bad = [...code].filter((c) => c.charCodeAt(0) > 127);
    assert.deepEqual(bad, [], `${name} contains non-ASCII characters`);
  });
});

test('no secrets, tokens or personal addresses are committed in src/', () => {
  sourceFiles().forEach(({ name, code }) => {
    assert.ok(!/AKIA[0-9A-Z]{12,}|sk_live|-----BEGIN|ya29\./.test(code), `${name} looks like it holds a credential`);
    assert.ok(!/@gmail\.com|@googlemail\.com/.test(code), `${name} contains a personal address`);
  });
});

test('appsscript.json: V8 runtime, explicit scopes, anonymous web app that runs as the deployer', () => {
  const manifest = JSON.parse(fs.readFileSync(path.join(SRC, 'appsscript.json'), 'utf8'));
  assert.equal(manifest.runtimeVersion, 'V8');
  assert.ok(Array.isArray(manifest.oauthScopes) && manifest.oauthScopes.length >= 4);
  assert.deepEqual(manifest.webapp, { executeAs: 'USER_DEPLOYING', access: 'ANYONE_ANONYMOUS' });
});

test('onOpen adds the menu that points at real functions', () => {
  const env = createEnv();
  env.call('onOpen');
  const menu = env.ui.menus[0];
  assert.equal(menu.name, 'Sheet tools');
  menu.items.forEach(([, fn]) => assert.equal(typeof env.ctx[fn], 'function', fn));
});
