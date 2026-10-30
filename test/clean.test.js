'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { createEnv } = require('./harness');

function clean(rows, options = {}, tz) {
  const env = createEnv(tz ? { timeZone: tz } : {});
  const source = env.ss.addSheet('Raw', rows);
  const result = env.call('cleanSheet_', env.ss, 'Raw', options);
  return { env, source, result, out: env.ss.getSheetByName('Cleaned') };
}

const call = (name, ...args) => createEnv().call(name, ...args);

// --- cell cleaners ---------------------------------------------------------------------------

test('names: ALL CAPS / lower case become Title Case; mixed case is left alone; digits are flagged', () => {
  assert.deepEqual(call('cleanName_', '  JOHN   SMITH '), { value: 'John Smith' });
  assert.deepEqual(call('cleanName_', "o'brien-jones"), { value: "O'Brien-Jones" });
  assert.deepEqual(call('cleanName_', 'McDonald'), { value: 'McDonald' });
  assert.deepEqual(call('cleanName_', 'Anne de Vries'), { value: 'Anne de Vries' });
  assert.deepEqual(call('cleanName_', 'JOSÉ ÁLVAREZ'), { value: 'José Álvarez' });
  assert.deepEqual(call('cleanName_', 'R2D2'), { value: 'R2D2', issue: 'name contains digits' });
  assert.deepEqual(call('cleanName_', ''), { value: '' });
});

test('emails: lower-cased and trimmed; mailto/angle brackets stripped; bad ones flagged, never altered beyond case', () => {
  assert.deepEqual(call('cleanEmail_', '  Ann@Example.COM '), { value: 'ann@example.com' });
  assert.deepEqual(call('cleanEmail_', 'mailto:ann@example.com'), { value: 'ann@example.com' });
  assert.deepEqual(call('cleanEmail_', '<ann@example.com>'), { value: 'ann@example.com' });
  assert.equal(call('cleanEmail_', 'ann@example').issue, 'invalid email');
  assert.equal(call('cleanEmail_', 'ann@@example.com').issue, 'invalid email');
  assert.match(call('cleanEmail_', 'a@x.com; b@x.com').issue, /more than one email/);
  assert.match(call('cleanEmail_', 'a@x.com b@x.com').issue, /more than one email/);
  assert.deepEqual(call('cleanEmail_', ''), { value: '' });
});

test('phones: US/Canada, international, 00 prefix, extension, national numbers with a country code', () => {
  const p = (v, cc) => call('cleanPhone_', v, cc);
  assert.deepEqual(p('(555) 123-4567', '1'), { value: '+15551234567' });
  assert.deepEqual(p('1-555-123-4567', '1'), { value: '+15551234567' });
  assert.deepEqual(p(5551234567, '1'), { value: '+15551234567' }, 'a number cell');
  assert.deepEqual(p('+44 20 7946 0958', ''), { value: '+442079460958' });
  assert.deepEqual(p('0044 20 7946 0958', ''), { value: '+442079460958' });
  assert.deepEqual(p('555.123.4567 ext. 89', '1'), { value: '+15551234567 x89' });
  assert.deepEqual(p('098765 43210', '91'), { value: '+919876543210' });
  assert.deepEqual(p('9876543210', '91'), { value: '+919876543210' });
});

test('phones: what it refuses to guess is flagged and returned unchanged', () => {
  const p = (v, cc) => call('cleanPhone_', v, cc);
  assert.match(p('555-1234', '1').issue, /10-digit/);
  assert.match(p('5551234567', '').issue, /no country code/);
  assert.equal(p('call me', '1').issue, 'invalid phone');
  assert.equal(p('+1 555 12', '').issue, 'phone length is not 8-15 digits');
  assert.equal(p('+1+555', '').issue, 'invalid phone');
  assert.match(p('919876543210', '91').issue, /length unclear/, '12 digits: could already contain the country code');
  assert.deepEqual(p('', '1'), { value: '' });
  assert.equal(p('555-1234', '1').value, '555-1234');
});

test('dates: ISO, slashes, dots, month names, with times, and Date objects', () => {
  const d = (v, order, tz) => call('cleanDate_', v, order || '', tz || 'UTC');
  assert.deepEqual(d('2026-01-05'), { value: '2026-01-05' });
  assert.deepEqual(d('2026/1/5'), { value: '2026-01-05' });
  assert.deepEqual(d('2026-01-05T23:59:00Z'), { value: '2026-01-05' });
  assert.deepEqual(d('5 Jan 2026'), { value: '2026-01-05' });
  assert.deepEqual(d('5th January, 2026'), { value: '2026-01-05' });
  assert.deepEqual(d('Jan 5, 2026'), { value: '2026-01-05' });
  assert.deepEqual(d('Sept 9 2026'), { value: '2026-09-09' });
  assert.deepEqual(d('25/12/2026'), { value: '2026-12-25' }, 'first number > 12 can only be the day');
  assert.deepEqual(d('12/25/2026'), { value: '2026-12-25' }, 'second number > 12 can only be the day');
  assert.deepEqual(d('03/03/2026'), { value: '2026-03-03' }, 'same number twice is unambiguous');
  assert.deepEqual(d('', ''), { value: '' });
});

test('dates: ambiguous ones are flagged unless DATE_ORDER says how to read them', () => {
  const d = (v, order) => call('cleanDate_', v, order || '', 'UTC');
  assert.match(d('01/02/2026').issue, /ambiguous/);
  assert.equal(d('01/02/2026').value, '01/02/2026', 'left untouched');
  assert.deepEqual(d('01/02/2026', 'DMY'), { value: '2026-02-01' });
  assert.deepEqual(d('01/02/2026', 'MDY'), { value: '2026-01-02' });
});

test('dates: impossible and unrecognised dates are flagged, never rolled over', () => {
  const d = (v) => call('cleanDate_', v, 'DMY', 'UTC');
  assert.equal(d('2026-02-30').issue, 'not a real calendar date');
  assert.equal(d('31/04/2026').issue, 'not a real calendar date');
  assert.equal(d('2027-02-29').issue, 'not a real calendar date');
  assert.equal(d('1899-12-31').issue, 'not a real calendar date');
  assert.match(d('5/1/26').issue, /4-digit year/);
  assert.match(d('yesterday').issue, /unrecognised/);
  assert.match(d('31 Foo 2026').issue, /unrecognised/);
  assert.equal(d(45678).issue, 'date is a bare number');
});

test('dates: Date objects are formatted in the spreadsheet time zone (the day the user sees)', () => {
  const env = createEnv();
  const midnightKolkata = new env.ctx.Date('2026-01-04T18:30:00Z');
  assert.deepEqual(plainResult(env.ctx.cleanDate_(midnightKolkata, '', 'Asia/Kolkata')), { value: '2026-01-05' });
  assert.deepEqual(plainResult(env.ctx.cleanDate_(midnightKolkata, '', 'UTC')), { value: '2026-01-04' });
  assert.deepEqual(plainResult(env.ctx.cleanDate_(new env.ctx.Date('nope'), '', 'UTC')), { value: '', issue: 'invalid date' });
});

function plainResult(r) { return JSON.parse(JSON.stringify(r)); }

// --- whole-sheet behaviour -------------------------------------------------------------------

test('cleans a messy sheet into "Cleaned" and leaves the source untouched', () => {
  const raw = [
    ['Full Name', 'Email', 'Mobile', 'Signup Date', 'Notes'],
    ['  JANE   DOE ', ' Jane@Example.COM', '(555) 123-4567', '5 Jan 2026', '  hello   there '],
    ['bob smith', 'bob@bad', '555-12', '2026-02-30', ''],
    ['', '', '', '', ''],
    ['Cy Young', 'cy@example.com', '+44 20 7946 0958', '25/12/2026', 'ok'],
  ];
  const { source, result, out } = clean(raw, { countryCode: '1' });
  assert.deepEqual(source.values(), raw, 'source unchanged');
  assert.equal(result.rows, 3);
  assert.equal(result.blankRowsSkipped, 1);
  assert.equal(result.flagged, 1);
  assert.deepEqual(out.values(), [
    ['Full Name', 'Email', 'Mobile', 'Signup Date', 'Notes', 'Clean Issues'],
    ['Jane Doe', 'jane@example.com', '+15551234567', '2026-01-05', 'hello there', ''],
    ['Bob Smith', 'bob@bad', '555-12', '2026-02-30', '', 'Email: invalid email; Mobile: not a 10-digit US/Canada number; Signup Date: not a real calendar date'],
    ['Cy Young', 'cy@example.com', '+442079460958', '2026-12-25', 'ok', ''],
  ]);
  assert.deepEqual(result.issues.map((i) => i.row), [3]);
});

test('IDs with leading zeros and plus signs survive, numeric columns stay numbers', () => {
  const { out } = clean([
    ['Ref', 'Amount', 'Qty'],
    ['007', 12.5, 3],
    ['0042', 100, ''],
    ['+123', 7, 1],
  ], {});
  assert.deepEqual(out.values(), [['Ref', 'Amount', 'Qty', 'Clean Issues'], ['007', 12.5, 3, ''], ['0042', 100, '', ''], ['+123', 7, 1, '']]);
  assert.equal(out.formatAt(2, 1), '@');
  assert.notEqual(out.formatAt(2, 2), '@', 'numbers are written as numbers');
});

test('a mixed column (numbers and text) is written as text, so nothing is reinterpreted', () => {
  const { out } = clean([['Code'], [5], ['A-5'], ['0123']], {});
  assert.deepEqual(out.values().slice(1).map((r) => r[0]), ['5', 'A-5', '0123']);
});

test('Date cells in an unlabeled column are treated as dates (ISO text), using the sheet time zone', () => {
  const { out } = clean([['Joined'], [new Date('2026-01-04T18:30:00Z')]], {}, 'Asia/Kolkata');
  assert.deepEqual(out.values()[1], ['2026-01-05', '']);
});

test('text that looks like a formula is kept as plain text', () => {
  const { out } = clean([['Comment'], ['=HYPERLINK("http://x")']], {});
  assert.equal(out.values()[1][0], '=HYPERLINK("http://x")');
  assert.equal(out.isFormula(2, 1), false);
});

test('required columns: blank values are listed as issues', () => {
  const { result, out } = clean([['Name', 'Email'], ['Ann', ''], ['Bo', 'bo@x.com']], { required: ['Email'] });
  assert.equal(result.flagged, 1);
  assert.equal(out.values()[1][2], 'Email: missing');
});

test('required column that does not exist is a clear error', () => {
  const env = createEnv();
  env.ss.addSheet('Raw', [['Name'], ['Ann']]);
  assert.throws(() => env.ctx.cleanSheet_(env.ss, 'Raw', { required: ['Phone'] }), /Required column "Phone" not found/);
});

test('explicit rules override the header guess; "none" copies untouched; bad rule names are rejected', () => {
  const { out } = clean([['Contact', 'Ref'], ['  x  ', '  ABC  ']], { rules: { Contact: 'none', ref: 'text' } });
  assert.deepEqual(out.values()[1].slice(0, 2), ['  x  ', 'ABC']);
  const env = createEnv();
  env.ss.addSheet('Raw', [['A'], ['x']]);
  assert.throws(() => env.ctx.cleanSheet_(env.ss, 'Raw', { rules: { A: 'telepathy' } }), /Unknown clean rule "telepathy" for "A"/);
});

test('header guesses: email, phone, date, name headers, but not "Company Name" or "Username"', () => {
  const env = createEnv();
  const g = (h) => env.call('inferCleanRule_', h, []);
  assert.equal(g('E-mail'), 'email');
  assert.equal(g('Work Email'), 'email');
  assert.equal(g('Mobile'), 'phone');
  assert.equal(g('Phone number'), 'phone');
  assert.equal(g('Date of Birth'), 'date');
  assert.equal(g('DOB'), 'date');
  assert.equal(g('First Name'), 'name');
  assert.equal(g('Name'), 'name');
  assert.equal(g('Company Name'), 'text');
  assert.equal(g('Username'), 'text');
  assert.equal(g('Updated'), 'text', 'no date guess from "Updated" unless the cells are Dates');
});

test('re-running replaces the previous "Cleaned" sheet instead of appending to it', () => {
  const env = createEnv();
  env.ss.addSheet('Raw', [['Name'], ['ann']]);
  env.call('cleanSheet_', env.ss, 'Raw', {});
  env.call('cleanSheet_', env.ss, 'Raw', {});
  assert.deepEqual(env.ss.getSheetByName('Cleaned').values(), [['Name', 'Clean Issues'], ['Ann', '']]);
});

test('empty sheet, header-only sheet and all-blank rows produce no output sheet and no error', () => {
  const env = createEnv();
  env.ss.addSheet('Empty');
  env.ss.addSheet('HeaderOnly', [['Name']]);
  env.ss.addSheet('Blanks', [['Name'], [''], [' ']]);
  ['Empty', 'HeaderOnly', 'Blanks'].forEach((n) => assert.equal(env.call('cleanSheet_', env.ss, n, {}).rows, 0));
  assert.equal(env.ss.getSheetByName('Cleaned'), null);
});

test('the output sheet may not be the source sheet', () => {
  const env = createEnv();
  env.ss.addSheet('Raw', [['Name'], ['x']]);
  assert.throws(() => env.ctx.cleanSheet_(env.ss, 'Raw', { outputName: 'Raw' }), /different name/);
});

test('menu action reads Script Properties and reports counts', () => {
  const env = createEnv();
  env.ss.addSheet('Raw', [['Name', 'Mobile', 'Joined'], ['ann lee', '5551234567', '01/02/2026'], ['bo', '5551234567', '13/02/2026']]);
  env.setProps({ DEFAULT_COUNTRY_CODE: '1', DATE_ORDER: 'dmy', CLEAN_REQUIRED: 'Name', CLEAN_RULES: '{"Joined":"date"}' });
  const result = env.call('cleanActiveSheet');
  assert.equal(result.rows, 2);
  assert.equal(result.flagged, 0);
  assert.deepEqual(env.ss.getSheetByName('Cleaned').values()[1], ['Ann Lee', '+15551234567', '2026-02-01', '']);
  assert.match(env.ui.alerts[0][0], /Cleaned 2 row\(s\) into "Cleaned"\. 0 row\(s\) need a look/);
});

test('menu action: invalid CLEAN_RULES JSON is reported as a configuration error', () => {
  const env = createEnv();
  env.ss.addSheet('Raw', [['Name'], ['x']]);
  env.setProps({ CLEAN_RULES: '{nope' });
  assert.throws(() => env.ctx.cleanActiveSheet(), /CLEAN_RULES is not valid JSON/);
});
