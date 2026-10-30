'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { createEnv } = require('./harness');

const CENTRAL = ['Timestamp', 'Form Origin', 'Response ID', 'Full name', 'Email', 'Plan', 'Respondent Email', 'Other Answers'];

function setup(centralRows = []) {
  const env = createEnv();
  const central = env.ss.addSheet('Submissions', [CENTRAL, ...centralRows]);
  const contact = env.addForm('form-contact', 'Contact us');
  const trial = env.addForm('form-trial', '  Free   trial\nsign-up ');
  const submit = (form, spec) => {
    const response = form.makeResponse(spec);
    return env.call('appendFormResponse', { source: form, response });
  };
  return { env, central, contact, trial, submit };
}

const ANN = {
  id: 'r1', timestamp: '2026-11-01T09:00:00.000Z', email: 'ann@x.com',
  answers: [['Full name', 'Ann Lee'], ['Email', 'ann@x.com'], ['Plan', 'Pro']],
};

test('appends one row with Form Origin = the form title, and fills the matching columns', () => {
  const { central, contact, submit } = setup();
  assert.deepEqual(submit(contact, ANN), { status: 'appended', unmapped: [] });
  assert.deepEqual(central.values()[1], [
    '2026-11-01T09:00:00.000Z', 'Contact us', 'form-contact/r1', 'Ann Lee', 'ann@x.com', 'Pro', 'ann@x.com', '',
  ]);
});

test('two forms feed one sheet; each row carries its own origin (titles are tidied to one line)', () => {
  const { central, contact, trial, submit } = setup();
  submit(contact, ANN);
  submit(trial, { id: 'r1', answers: [['Full name', 'Bo']] });
  submit(contact, { ...ANN, id: 'r2' });
  assert.deepEqual(central.values().slice(1).map((r) => [r[1], r[2]]), [
    ['Contact us', 'form-contact/r1'],
    ['Free trial sign-up', 'form-trial/r1'],
    ['Contact us', 'form-contact/r2'],
  ]);
});

test('EXISTING ROWS ARE NEVER TOUCHED: they are byte-identical after any number of submissions', () => {
  const existing = [
    ['2026-10-01T00:00:00.000Z', 'Old form', 'old/1', 'Pre Existing', 'pre@x.com', 'Free', '', ''],
    ['2026-10-02T00:00:00.000Z', 'Old form', 'old/2', '=SUM(1,1)', '007', 'Pro', '', 'note'],
  ];
  const { env, central, contact, trial, submit } = setup(existing);
  const before = JSON.stringify(central.values().slice(0, 3));
  const formatsBefore = [2, 3].map((r) => [1, 2, 3, 4, 5].map((c) => central.formatAt(r, c)));
  submit(contact, ANN);
  submit(trial, { id: 'x', answers: [['Plan', 'Free']] });
  submit(contact, { ...ANN, id: 'r9', answers: [['Full name', 'Zed']] });
  assert.equal(JSON.stringify(central.values().slice(0, 3)), before);
  assert.deepEqual([2, 3].map((r) => [1, 2, 3, 4, 5].map((c) => central.formatAt(r, c))), formatsBefore);
  assert.equal(central.values().length, 6);
  assert.equal(central.calls.deleteRows, undefined);
  assert.equal(central.calls.appendRow, undefined);
  void env;
});

test('writes go only to the new row: every setValues call lands below the previous last row', () => {
  const { central, contact, submit } = setup([['t', 'o', 'old/1', 'A', 'a@x.com', 'Free', '', '']]);
  const writes = [];
  const original = central.getRange.bind(central);
  central.getRange = (row, col, nr, nc) => {
    const range = original(row, col, nr, nc);
    const set = range.setValues.bind(range);
    range.setValues = (v) => { writes.push(row); return set(v); };
    return range;
  };
  submit(contact, ANN);
  submit(contact, { ...ANN, id: 'r2' });
  assert.deepEqual(writes, [3, 4]);
});

test('a retried trigger (same response id) is skipped, not written twice', () => {
  const { central, contact, submit } = setup();
  assert.equal(submit(contact, ANN).status, 'appended');
  assert.equal(submit(contact, ANN).status, 'duplicate');
  assert.equal(central.values().length, 2);
});

test('the same response id on two different forms are two different submissions', () => {
  const { central, contact, trial, submit } = setup();
  submit(contact, { id: 'same', answers: [['Plan', 'A']] });
  submit(trial, { id: 'same', answers: [['Plan', 'B']] });
  assert.equal(central.values().length, 3);
});

test('answers with no matching column go to "Other Answers" (and are listed), nothing is dropped silently', () => {
  const { central, contact, submit } = setup();
  const r = submit(contact, { id: 'r3', answers: [['Full name', 'Ann'], ['How did you hear about us?', 'Friend'], ['Budget', '$500']] });
  assert.deepEqual(r.unmapped, ['How did you hear about us?', 'Budget']);
  assert.equal(central.values()[1][7], 'How did you hear about us?: Friend\nBudget: $500');
});

test('without an "Other Answers" column unmapped answers are logged, and the headers are NOT changed', () => {
  const env = createEnv();
  const headers = ['Timestamp', 'Form Origin', 'Response ID', 'Plan'];
  const central = env.ss.addSheet('Submissions', [headers]);
  const form = env.addForm('f', 'F');
  env.call('appendFormResponse', { source: form, response: form.makeResponse({ id: '1', answers: [['Plan', 'Pro'], ['Surprise', 'x']] }) });
  assert.deepEqual(central.values()[0], headers);
  assert.deepEqual(central.values()[1].slice(3), ['Pro']);
  assert.match(env.logs.warn.join('\n'), /no column for Surprise/);
});

test('header matching ignores case and spacing; a question named like a system column cannot overwrite it', () => {
  const { central, contact, submit } = setup();
  submit(contact, { id: 'r4', answers: [['  FULL   NAME ', 'Ann'], ['Form Origin', 'Spoofed'], ['Response ID', 'Spoofed'], ['Timestamp', 'yesterday']] });
  const row = central.values()[1];
  assert.equal(row[3], 'Ann');
  assert.equal(row[1], 'Contact us');
  assert.equal(row[2], 'form-contact/r4');
  assert.equal(row[0], '2026-11-01T10:00:00.000Z');
  assert.match(row[7], /Form Origin: Spoofed\nResponse ID: Spoofed\nTimestamp: yesterday/);
});

test('checkbox and grid answers are flattened to readable text; missing answers are blank', () => {
  const { central, contact, submit } = setup();
  submit(contact, { id: 'r5', answers: [['Plan', ['Pro', 'Team']], ['Email', null], ['Full name', [['a', 'b'], ['c']]]] });
  const row = central.values()[1];
  assert.equal(row[5], 'Pro; Team');
  assert.equal(row[4], '');
  assert.equal(row[3], 'a, b; c');
});

test('answers that look like formulas, numbers or dates are stored as typed text', () => {
  const { central, contact, submit } = setup();
  submit(contact, { id: 'r6', answers: [['Full name', '=HYPERLINK("http://evil","click")'], ['Email', '007'], ['Plan', '2026-01-05']] });
  const row = central.values()[1];
  assert.deepEqual(row.slice(3, 6), ['=HYPERLINK("http://evil","click")', '007', '2026-01-05']);
  assert.equal(central.isFormula(2, 4), false);
  [1, 2, 3, 4, 5, 6].forEach((c) => assert.equal(central.formatAt(2, c), '@'));
});

test('no respondent email collected: the column stays blank; empty title gets a placeholder origin', () => {
  const { env, central, submit } = setup();
  const untitled = env.addForm('form-x', '');
  submit(untitled, { id: 'r7', email: null, answers: [['Plan', 'Free']] });
  assert.equal(central.values()[1][1], '(untitled form)');
  assert.equal(central.values()[1][6], '');
});

test('new central sheet: created with the standard headers on first use', () => {
  const env = createEnv();
  const form = env.addForm('f', 'Survey');
  env.call('appendFormResponse', { source: form, response: form.makeResponse({ id: '1', answers: [['Q', 'A']] }) });
  const sheet = env.ss.getSheetByName('Submissions');
  assert.deepEqual(sheet.values()[0], ['Timestamp', 'Form Origin', 'Response ID', 'Respondent Email', 'Other Answers']);
  assert.equal(sheet.values()[1][4], 'Q: A');
});

test('central sheet missing a required column: clear error and NOTHING is written', () => {
  const env = createEnv();
  const sheet = env.ss.addSheet('Submissions', [['Timestamp', 'Plan'], ['t', 'x']]);
  const form = env.addForm('f', 'F');
  assert.throws(
    () => env.ctx.appendFormResponse({ source: form, response: form.makeResponse({ id: '1', answers: [] }) }),
    /needs a "Form Origin" column in row 1/
  );
  assert.deepEqual(sheet.values(), [['Timestamp', 'Plan'], ['t', 'x']]);
  assert.equal(env.lock.held, false);
});

test('lock not available within the wait: the handler throws (visible failure) and writes nothing', () => {
  const { env, central, contact } = setup();
  env.lock.heldByOther = true;
  assert.throws(() => env.ctx.appendFormResponse({ source: contact, response: contact.makeResponse(ANN) }), /Could not get the script lock/);
  assert.equal(central.values().length, 1);
});

test('FORMS_SHEET and SPREADSHEET_ID route the rows to another spreadsheet', () => {
  const env = createEnv();
  env.setProps({ SPREADSHEET_ID: 'central', FORMS_SHEET: 'Intake' });
  const central = env.addSpreadsheet('central');
  const form = env.addForm('f', 'F');
  env.call('appendFormResponse', { source: form, response: form.makeResponse({ id: '1', answers: [['Q', 'A']] }) });
  assert.equal(central.getSheetByName('Intake').values().length, 2);
  assert.equal(env.ss.getSheets().length, 0);
});

test('backfill adds only the responses that are missing', () => {
  const { env, central, contact, submit } = setup();
  contact.addResponse({ ...ANN, id: 'b1' });
  contact.addResponse({ ...ANN, id: 'b2' });
  contact.addResponse({ ...ANN, id: 'b3' });
  submit(contact, { ...ANN, id: 'b2' });
  assert.deepEqual(env.call('backfillFormResponses', 'form-contact'), { formId: 'form-contact', added: 2 });
  assert.deepEqual(central.values().slice(1).map((r) => r[2]).sort(), ['form-contact/b1', 'form-contact/b2', 'form-contact/b3']);
  assert.deepEqual(env.call('backfillFormResponses', 'form-contact'), { formId: 'form-contact', added: 0 });
  assert.throws(() => env.ctx.backfillFormResponses('nope'), /No item with the given ID/);
});

test('installFormTriggers creates one on-submit trigger per form id and is safe to run twice', () => {
  const env = createEnv();
  env.setProps({ FORM_IDS: 'form-a, form-b' });
  assert.deepEqual(env.call('installFormTriggers'), { created: ['form-a', 'form-b'], alreadyInstalled: [] });
  env.setProps({ FORM_IDS: 'form-a, form-b, form-c' });
  assert.deepEqual(env.call('installFormTriggers'), { created: ['form-c'], alreadyInstalled: ['form-a', 'form-b'] });
  const triggers = env.scriptApp.getProjectTriggers();
  assert.equal(triggers.length, 3);
  triggers.forEach((t) => {
    assert.equal(t.getHandlerFunction(), 'appendFormResponse');
    assert.equal(t.getEventType(), 'ON_FORM_SUBMIT');
  });
  assert.deepEqual(triggers.map((t) => t.getTriggerSourceId()), ['form-a', 'form-b', 'form-c']);
});

test('installFormTriggers without FORM_IDS is a clear error', () => {
  assert.throws(() => createEnv().ctx.installFormTriggers(), /Set the FORM_IDS/);
});
