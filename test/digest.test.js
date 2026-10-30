'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { createEnv } = require('./harness');

const HEADER = ['Name', 'Email', 'Source', 'Digested At'];
const ROWS = [
  ['Ann', 'ann@x.com', 'web', ''],
  ['Bo', 'bo@x.com', 'ads', ''],
  ['Cy', 'cy@x.com', 'web', 'baseline'],
];

function setup(rows = ROWS, props = {}) {
  const env = createEnv();
  env.setProps({ DIGEST_RECIPIENTS: 'boss@client.com, ops@client.com', ...props });
  const sheet = env.ss.addSheet('Leads', [HEADER, ...rows]);
  return { env, sheet, run: () => env.call('sendDigest') };
}

test('DRY RUN IS THE DEFAULT: with no DRY_RUN property nothing is sent and no row is marked', () => {
  const { env, sheet, run } = setup();
  const before = sheet.values();
  const r = run();
  assert.equal(r.status, 'dry_run');
  assert.equal(r.rows, 2);
  assert.deepEqual(r.wouldSendTo, ['boss@client.com', 'ops@client.com']);
  assert.equal(r.testSentTo, null);
  assert.equal(env.mail.sent.length, 0);
  assert.deepEqual(sheet.values(), before);
  assert.match(env.logs.log.join('\n'), /DRY RUN: would send "\[DRY RUN\] New rows \(2\)" to boss@client.com, ops@client.com/);
});

test('only the exact value "false" turns sending on; anything else stays a dry run', () => {
  ['', 'true', 'TRUE', 'no', '0', 'off', 'falsey', ' ', 'f'].forEach((value) => {
    const { env, run } = setup(ROWS, { DRY_RUN: value });
    assert.equal(run().status, 'dry_run', `DRY_RUN=${JSON.stringify(value)}`);
    assert.equal(env.mail.sent.length, 0);
  });
  ['false', 'FALSE', ' False '].forEach((value) => {
    const { env, run } = setup(ROWS, { DRY_RUN: value });
    assert.equal(run().status, 'sent', `DRY_RUN=${JSON.stringify(value)}`);
    assert.equal(env.mail.sent.length, 1);
  });
});

test('live run: ONE email to the list, new rows marked with a timestamp, old rows untouched', () => {
  const { env, sheet, run } = setup(ROWS, { DRY_RUN: 'false' });
  const r = run();
  assert.deepEqual(r, { status: 'sent', rows: 2, sentTo: ['boss@client.com', 'ops@client.com'], subject: 'New rows (2)' });
  assert.equal(env.mail.sent.length, 1);
  const mail = env.mail.sent[0];
  assert.equal(mail.to, 'boss@client.com,ops@client.com');
  assert.equal(mail.subject, 'New rows (2)');
  assert.match(mail.body, /^2 new row\(s\)\./);
  assert.match(mail.body, /Name: Ann \| Email: ann@x.com \| Source: web/);
  assert.doesNotMatch(mail.body, /Cy|Digested At/);
  assert.deepEqual(sheet.values().slice(1).map((r) => r[3]), ['2026-11-01T12:00:00.000Z', '2026-11-01T12:00:00.000Z', 'baseline']);
});

test('a second live run has nothing new and sends nothing (no duplicate digests)', () => {
  const { env, run } = setup(ROWS, { DRY_RUN: 'false' });
  run();
  const again = run();
  assert.equal(again.status, 'nothing_new');
  assert.equal(env.mail.sent.length, 1);
});

test('rows added after a digest are picked up by the next one, and only those', () => {
  const { env, sheet, run } = setup(ROWS, { DRY_RUN: 'false' });
  run();
  sheet.appendRow(['Di', 'di@x.com', 'web']);
  env.clock.now += 86400000;
  const r = run();
  assert.equal(r.rows, 1);
  assert.match(env.mail.sent[1].body, /Di/);
  assert.doesNotMatch(env.mail.sent[1].body, /Ann|Bo/);
});

test('empty list-sheet, or no new rows: nothing_new in both modes', () => {
  const env = createEnv();
  env.setProps({ DIGEST_RECIPIENTS: 'a@x.com', DRY_RUN: 'false' });
  env.ss.addSheet('Leads', [HEADER]);
  assert.equal(env.call('sendDigest').status, 'nothing_new');
  const all = setup([['Ann', 'a@x.com', 'web', 'baseline']], { DRY_RUN: 'false' });
  assert.equal(all.run().status, 'nothing_new');
  assert.equal(all.env.mail.sent.length, 0);
});

test('a live run refuses to start until the "Digested At" column exists, so history is not mailed by accident', () => {
  const env = createEnv();
  env.setProps({ DIGEST_RECIPIENTS: 'a@x.com', DRY_RUN: 'false' });
  const sheet = env.ss.addSheet('Leads', [['Name', 'Email'], ['Ann', 'ann@x.com'], ['Bo', 'bo@x.com']]);
  const r = env.call('sendDigest');
  assert.equal(r.status, 'needs_baseline');
  assert.equal(env.mail.sent.length, 0);
  assert.deepEqual(sheet.values()[0], ['Name', 'Email'], 'the sheet was not modified');
});

test('initDigestBaseline adds the column and stamps existing rows; the first live run then sends only newer rows', () => {
  const env = createEnv();
  env.setProps({ DIGEST_RECIPIENTS: 'a@x.com', DRY_RUN: 'false' });
  const sheet = env.ss.addSheet('Leads', [['Name', 'Email'], ['Ann', 'ann@x.com'], ['', ''], ['Bo', 'bo@x.com']]);
  assert.deepEqual(env.call('initDigestBaseline'), { columnCreated: true, stamped: 2 });
  assert.deepEqual(sheet.values().map((r) => r[2]), ['Digested At', 'baseline', '', 'baseline']);
  assert.equal(env.call('sendDigest').status, 'nothing_new');
  sheet.appendRow(['Cy', 'cy@x.com']);
  assert.equal(env.call('sendDigest').rows, 1);
  assert.deepEqual(env.call('initDigestBaseline'), { columnCreated: false, stamped: 0 });
});

test('dry run does not need the column and previews every row as new', () => {
  const env = createEnv();
  env.setProps({ DIGEST_RECIPIENTS: 'a@x.com' });
  const sheet = env.ss.addSheet('Leads', [['Name', 'Email'], ['Ann', 'ann@x.com'], ['Bo', 'bo@x.com']]);
  const r = env.call('sendDigest');
  assert.equal(r.status, 'dry_run');
  assert.equal(r.rows, 2);
  assert.deepEqual(sheet.values()[0], ['Name', 'Email']);
});

test('dry run with DIGEST_TEST_RECIPIENT sends the preview to that one address only, prefixed, and marks nothing', () => {
  const { env, sheet, run } = setup(ROWS, { DIGEST_TEST_RECIPIENT: 'me@agency.com' });
  const before = sheet.values();
  const r = run();
  assert.equal(r.testSentTo, 'me@agency.com');
  assert.equal(env.mail.sent.length, 1);
  assert.equal(env.mail.sent[0].to, 'me@agency.com');
  assert.match(env.mail.sent[0].subject, /^\[DRY RUN\] New rows \(2\)$/);
  assert.deepEqual(sheet.values(), before);
  assert.doesNotMatch(JSON.stringify(env.mail.sent), /boss@client.com|ops@client.com/);
});

test('a test recipient that is really a list, or malformed, is refused', () => {
  const list = setup(ROWS, { DIGEST_TEST_RECIPIENT: 'a@x.com,b@x.com' });
  assert.equal(list.run().status, 'invalid_recipient');
  assert.equal(list.env.mail.sent.length, 0);
});

test('DIGEST_TEST_RECIPIENT is ignored in live mode (the real list is used)', () => {
  const { env, run } = setup(ROWS, { DRY_RUN: 'false', DIGEST_TEST_RECIPIENT: 'me@agency.com' });
  run();
  assert.equal(env.mail.sent[0].to, 'boss@client.com,ops@client.com');
});

test('recipient guards: none, too many, malformed - nothing is sent and rows stay unmarked', () => {
  const none = setup(ROWS, { DRY_RUN: 'false', DIGEST_RECIPIENTS: null });
  assert.equal(none.run().status, 'no_recipients');

  const many = setup(ROWS, { DRY_RUN: 'false', DIGEST_RECIPIENTS: Array.from({ length: 11 }, (_, i) => `u${i}@x.com`).join(',') });
  assert.deepEqual(many.run(), { status: 'too_many_recipients', count: 11, max: 10 });

  const raised = setup(ROWS, { DRY_RUN: 'false', DIGEST_MAX_RECIPIENTS: '12', DIGEST_RECIPIENTS: Array.from({ length: 11 }, (_, i) => `u${i}@x.com`).join(',') });
  assert.equal(raised.run().status, 'sent');

  const bad = setup(ROWS, { DRY_RUN: 'false', DIGEST_RECIPIENTS: 'ok@x.com, not-an-email' });
  assert.deepEqual(bad.run(), { status: 'invalid_recipient', invalid: ['not-an-email'] });

  [none, many, bad].forEach(({ env, sheet }) => {
    assert.equal(env.mail.sent.length, 0);
    assert.deepEqual(sheet.values().slice(1).map((r) => r[3]), ['', '', 'baseline']);
  });
});

test('mail quota too low: nothing is sent, rows are not marked', () => {
  const { env, sheet, run } = setup(ROWS, { DRY_RUN: 'false' });
  env.mail.quota = 1;
  assert.deepEqual(run(), { status: 'quota_too_low', remaining: 1, needed: 2 });
  assert.equal(env.mail.sent.length, 0);
  assert.equal(sheet.values()[1][3], '');
});

test('MailApp failure: the run throws and the rows stay unmarked so the next run retries them', () => {
  const { env, sheet, run } = setup(ROWS, { DRY_RUN: 'false' });
  env.mail.failNext = new Error('Exception: Service invoked too many times for one day: email.');
  assert.throws(() => env.ctx.sendDigest(), /too many times/);
  assert.deepEqual(sheet.values().slice(1).map((r) => r[3]), ['', '', 'baseline']);
  assert.equal(env.lock.held, false);
  assert.equal(run().status, 'sent');
});

test('cell values are HTML-escaped in the HTML body', () => {
  const { env, run } = setup([['<script>alert(1)</script>', 'a&b@x.com', '"q"', '']], { DRY_RUN: 'false' });
  run();
  const html = env.mail.sent[0].htmlBody;
  assert.ok(html.includes('&lt;script&gt;alert(1)&lt;/script&gt;'));
  assert.ok(html.includes('a&amp;b@x.com'));
  assert.ok(!html.includes('<script>'));
});

test('large digests show DIGEST_MAX_ROWS rows and say how many more there are; all rows are still marked', () => {
  const rows = Array.from({ length: 12 }, (_, i) => [`n${i}`, `n${i}@x.com`, 'web', '']);
  const { env, sheet, run } = setup(rows, { DRY_RUN: 'false', DIGEST_MAX_ROWS: '5' });
  const r = run();
  assert.equal(r.rows, 12);
  const mail = env.mail.sent[0];
  assert.match(mail.body, /n4/);
  assert.doesNotMatch(mail.body, /n5/);
  assert.match(mail.body, /\.\.\.and 7 more row\(s\) in the sheet\./);
  assert.ok(mail.htmlBody.includes('and 7 more row(s)'));
  assert.equal(sheet.values().slice(1).filter((r) => r[3] === '2026-11-01T12:00:00.000Z').length, 12);
});

test('dates in cells are shown as dates, and DIGEST_SUBJECT is used', () => {
  const env = createEnv();
  env.setProps({ DIGEST_RECIPIENTS: 'a@x.com', DRY_RUN: 'false', DIGEST_SUBJECT: 'Daily leads' });
  env.ss.addSheet('Leads', [['Name', 'Joined', 'Digested At'], ['Ann', new Date('2026-10-31T00:00:00Z'), '']]);
  env.call('sendDigest');
  assert.equal(env.mail.sent[0].subject, 'Daily leads (1)');
  assert.match(env.mail.sent[0].body, /Joined: 2026-10-31/);
});

test('another run holds the lock: status busy, nothing sent', () => {
  const { env, run } = setup(ROWS, { DRY_RUN: 'false' });
  env.lock.heldByOther = true;
  assert.equal(run().status, 'busy');
  assert.equal(env.mail.sent.length, 0);
});

test('installDigestTrigger: one daily trigger at DIGEST_HOUR (default 8), replaced not duplicated; bad hour rejected', () => {
  const env = createEnv();
  env.call('installDigestTrigger');
  env.setProps({ DIGEST_HOUR: '17' });
  env.call('installDigestTrigger');
  const triggers = env.scriptApp.getProjectTriggers();
  assert.equal(triggers.length, 1);
  assert.equal(triggers[0].getHandlerFunction(), 'sendDigest');
  assert.equal(triggers[0].spec.atHour, 17);
  assert.equal(triggers[0].spec.everyDays, 1);
  env.setProps({ DIGEST_HOUR: '24' });
  assert.throws(() => env.ctx.installDigestTrigger(), /0-23/);
});

test('installing the digest trigger does not turn sending on', () => {
  const { env } = setup();
  env.call('installDigestTrigger');
  assert.equal(env.call('sendDigest').status, 'dry_run');
  assert.equal(env.mail.sent.length, 0);
});
