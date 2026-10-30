'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { createEnv } = require('./harness');

const HEADER = ['Name', 'Email', 'Plan'];

function setup(rows) {
  const env = createEnv();
  const sheet = env.ss.addSheet('Leads', [HEADER, ...rows]);
  return { env, sheet };
}

test('dry run reports duplicates and changes nothing', () => {
  const { env, sheet } = setup([['A', 'a@x.com', 1], ['B', 'b@x.com', 2], ['A2', 'A@x.com ', 3]]);
  const before = sheet.values();
  const report = env.call('dedupeSheet', 'Leads', 'Email', 'first');
  assert.equal(report.dryRun, true);
  assert.deepEqual(report.removed, [{ row: 4, key: 'A@x.com', duplicateOfRow: 2 }]);
  assert.deepEqual(sheet.values(), before);
  assert.equal(env.ss.getSheetByName('_Removed Rows'), null);
  assert.equal(sheet.calls.deleteRows, undefined);
});

test('keeps the first occurrence by default and removes later ones; key is trimmed and case-insensitive', () => {
  const { env, sheet } = setup([['A', 'a@x.com', 1], ['B', 'b@x.com', 2], ['A2', ' A@X.com', 3], ['A3', 'a@x.com', 4]]);
  const report = env.call('dedupeSheet', 'Leads', 'email', 'first', false);
  assert.deepEqual(report.removed.map((r) => r.row), [4, 5]);
  assert.deepEqual(sheet.values(), [HEADER, ['A', 'a@x.com', 1], ['B', 'b@x.com', 2]]);
});

test('keep "last" keeps the most recent row of each group', () => {
  const { env, sheet } = setup([['old', 'a@x.com', 1], ['B', 'b@x.com', 2], ['new', 'a@x.com', 3]]);
  const report = env.call('dedupeSheet', 'Leads', 'Email', 'last', false);
  assert.deepEqual(report.removed, [{ row: 2, key: 'a@x.com', duplicateOfRow: 4 }]);
  assert.deepEqual(sheet.values(), [HEADER, ['B', 'b@x.com', 2], ['new', 'a@x.com', 3]]);
});

test('blank keys are never duplicates of each other and are counted', () => {
  const { env, sheet } = setup([['A', '', 1], ['B', '  ', 2], ['C', null, 3], ['D', 'd@x.com', 4]]);
  const report = env.call('dedupeSheet', 'Leads', 'Email', 'first', false);
  assert.equal(report.removed.length, 0);
  assert.equal(report.blankKeyRows, 3);
  assert.equal(sheet.values().length, 5);
});

test('empty sheet and header-only sheet return an empty report without throwing', () => {
  const env = createEnv();
  env.ss.addSheet('Empty');
  env.ss.addSheet('HeaderOnly', [HEADER]);
  const a = env.call('dedupeSheet', 'Empty', 'Email', 'first', false);
  const b = env.call('dedupeSheet', 'HeaderOnly', 'Email', 'first', false);
  [a, b].forEach((r) => {
    assert.equal(r.scanned, 0);
    assert.deepEqual(r.removed, []);
  });
  assert.equal(env.ss.getSheetByName('_Removed Rows'), null);
});

test('unknown key column gives an error that lists the headers; nothing is deleted', () => {
  const { env, sheet } = setup([['A', 'a@x.com', 1], ['A', 'a@x.com', 1]]);
  assert.throws(() => env.ctx.dedupeSheet('Leads', 'E-mail', 'first', false), /Key column "E-mail" not found\. Headers found: Name, Email, Plan/);
  assert.equal(sheet.values().length, 3);
});

test('unknown sheet gives a clear error', () => {
  const env = createEnv();
  assert.throws(() => env.ctx.dedupeSheet('Nope', 'Email', 'first'), /Sheet "Nope" not found/);
});

test('removed rows are archived as text before deletion, with a header row and provenance', () => {
  const { env, sheet } = setup([['A', 'a@x.com', 1], ['A2', 'a@x.com', 2]]);
  env.clock.now = Date.parse('2026-11-02T08:00:00.000Z');
  env.call('dedupeSheet', 'Leads', 'Email', 'first', false);
  const archive = env.ss.getSheetByName('_Removed Rows');
  assert.deepEqual(archive.values(), [
    ['Removed At', 'Source Sheet', 'Source Row', 'Key', 'Name', 'Email', 'Plan'],
    ['2026-11-02T08:00:00.000Z', 'Leads', '3', 'a@x.com', 'A2', 'a@x.com', '2'],
  ]);
  assert.equal(archive.formatAt(2, 5), '@');
  assert.deepEqual(sheet.values(), [HEADER, ['A', 'a@x.com', 1]]);
});

test('a second batch appends below the first and brings its own header row', () => {
  const { env, sheet } = setup([['A', 'a@x.com', 1], ['A2', 'a@x.com', 2]]);
  env.call('dedupeSheet', 'Leads', 'Email', 'first', false);
  sheet.appendRow(['A3', 'a@x.com', 3]);
  env.call('dedupeSheet', 'Leads', 'Email', 'first', false);
  const archive = env.ss.getSheetByName('_Removed Rows').values();
  assert.equal(archive.length, 4);
  assert.equal(archive[2][0], 'Removed At');
  assert.equal(archive[3][4], 'A3');
});

test('archive keeps Dates and formula-looking text safe', () => {
  const env = createEnv();
  const sheet = env.ss.addSheet('Leads', [
    ['Email', 'Joined', 'Note'],
    ['a@x.com', new Date('2026-01-05T00:00:00Z'), 'ok'],
    ['a@x.com', new Date('2026-02-06T10:11:12Z'), '=HYPERLINK("http://evil")'],
  ]);
  env.call('dedupeSheet', 'Leads', 'Email', 'first', false);
  const archive = env.ss.getSheetByName('_Removed Rows');
  assert.deepEqual(archive.values()[1].slice(4), ['a@x.com', '2026-02-06 10:11:12', '=HYPERLINK("http://evil")']);
  assert.equal(archive.isFormula(2, 7), false);
  assert.equal(sheet.values().length, 2);
});

test('consecutive duplicate rows are deleted with ONE deleteRows call per run, bottom-up', () => {
  const rows = [['A', 'a@x.com', 1]];
  for (let i = 0; i < 40; i++) rows.push([`dup${i}`, 'a@x.com', i]);
  rows.push(['B', 'b@x.com', 1], ['B2', 'b@x.com', 2], ['C', 'c@x.com', 1]);
  const { env, sheet } = setup(rows);
  env.call('dedupeSheet', 'Leads', 'Email', 'first', false);
  assert.equal(sheet.calls.deleteRows, 2, 'one run of 40 rows and one single row');
  assert.deepEqual(sheet.values().map((r) => r[0]), ['Name', 'A', 'B', 'C']);
});

test('numbers and dates work as keys (7 and "7" match; same day twice matches)', () => {
  const env = createEnv();
  const sheet = env.ss.addSheet('S', [['Key', 'V'], [7, 'a'], ['7', 'b'], [new Date('2026-01-05T00:00:00Z'), 'c'], [new Date('2026-01-05T00:00:00Z'), 'd']]);
  env.call('dedupeSheet', 'S', 'Key', 'first', false);
  assert.deepEqual(sheet.values().map((r) => r[1]), ['V', 'a', 'c']);
});

test('running it again finds nothing (idempotent)', () => {
  const { env } = setup([['A', 'a@x.com', 1], ['A2', 'a@x.com', 2]]);
  assert.equal(env.call('dedupeSheet', 'Leads', 'Email', 'first', false).removed.length, 1);
  assert.equal(env.call('dedupeSheet', 'Leads', 'Email', 'first', false).removed.length, 0);
});

test('menu action: previews, asks, and removes only after YES', () => {
  const { env, sheet } = setup([['A', 'a@x.com', 1], ['A2', 'a@x.com', 2]]);
  env.ui.answers = ['NO'];
  const declined = env.call('removeDuplicateRows');
  assert.equal(declined.dryRun, true);
  assert.equal(sheet.values().length, 3);
  assert.match(env.ui.alerts[0][0], /Remove 1 duplicate row\(s\)\?/);
  assert.match(env.ui.alerts[0][1], /Row 3 \("a@x.com"\) duplicates row 2/);

  env.ui.answers = ['YES'];
  const done = env.call('removeDuplicateRows');
  assert.equal(done.dryRun, false);
  assert.equal(sheet.values().length, 2);
  assert.match(env.ui.alerts.at(-1)[0], /Removed 1 row\(s\)/);
});

test('menu action: says so when there is nothing to remove, and honours DEDUPE_KEY_HEADER / DEDUPE_KEEP', () => {
  const { env, sheet } = setup([['A', 'a@x.com', 1], ['B', 'b@x.com', 1]]);
  env.call('removeDuplicateRows');
  assert.match(env.ui.alerts[0][0], /No duplicates found/);
  env.setProps({ DEDUPE_KEY_HEADER: 'Plan', DEDUPE_KEEP: 'last' });
  env.ui.answers = ['YES'];
  env.call('removeDuplicateRows');
  assert.deepEqual(sheet.values(), [HEADER, ['B', 'b@x.com', 1]]);
});
