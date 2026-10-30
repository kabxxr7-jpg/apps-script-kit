/**
 * Clean a messy sheet into a new "Cleaned" sheet. The source sheet is never modified.
 *
 * What it does, per column (the rule comes from the header text, or from CLEAN_RULES):
 *   name  - trims; fixes ALL CAPS / all lower case to Title Case (mixed case is left alone)
 *   email - trims, lower-cases, flags anything that is not one plain address
 *   phone - to +<country><number> (E.164 style); flags numbers it cannot place
 *   date  - to yyyy-mm-dd; flags ambiguous (01/02/2026) and impossible (2026-02-30) dates
 *   text  - trims and collapses spaces (the default); numbers are left as numbers
 *   none  - copied untouched
 * Every problem it cannot fix is listed in a "Clean Issues" column; it never guesses.
 *
 * Script properties (all optional): CLEAN_RULES (JSON, e.g. {"Mobile":"phone","Ref":"none"}),
 * CLEAN_REQUIRED (comma list of headers that must not be blank), DEFAULT_COUNTRY_CODE (e.g. 1 or 91),
 * DATE_ORDER (DMY or MDY, used only for dates that are ambiguous without it).
 */

function cleanRuleNames_() {
  return ['name', 'email', 'phone', 'date', 'text', 'none'];
}

function inferCleanRule_(header, values) {
  var h = normalizeKey_(header);
  if (/e-?mail/.test(h)) return 'email';
  if (/(^|\b)(phone|mobile|cell|tel|telephone|fax)(\b|$)/.test(h)) return 'phone';
  if (/(^|\b)(date|dob|birthday|birth date)(\b|$)/.test(h)) return 'date';
  if (/^(contact |customer |client )?(name|first name|last name|full name|middle name|given name|family name|surname)$/.test(h)) return 'name';
  for (var i = 0; i < values.length; i++) if (isDate_(values[i])) return 'date';
  return 'text';
}

function resolveCleanRules_(table, overrides) {
  var byHeader = {};
  Object.keys(overrides || {}).forEach(function (k) {
    var rule = String(overrides[k]).toLowerCase();
    if (cleanRuleNames_().indexOf(rule) < 0) {
      throw new Error('Unknown clean rule "' + overrides[k] + '" for "' + k + '". Use: ' + cleanRuleNames_().join(', '));
    }
    byHeader[normalizeKey_(k)] = rule;
  });
  return table.headers.map(function (h, c) {
    var key = normalizeKey_(h);
    if (Object.prototype.hasOwnProperty.call(byHeader, key)) return byHeader[key];
    return inferCleanRule_(h, table.rows.map(function (row) { return row[c]; }));
  });
}

function cleanName_(value) {
  var text = normalizeText_(value);
  if (text === '') return { value: '' };
  var issue = /\d/.test(text) ? 'name contains digits' : null;
  if (!issue && /[A-Za-z\u00c0-\u024f]/.test(text) && (text === text.toLowerCase() || text === text.toUpperCase())) {
    text = text.toLowerCase().replace(/(^|[\s\-'\u2019.])([a-z\u00e0-\u024f])/g, function (m, before, ch) {
      return before + ch.toUpperCase();
    });
  }
  return issue ? { value: text, issue: issue } : { value: text };
}

function cleanEmail_(value) {
  var text = normalizeText_(value);
  if (text === '') return { value: '' };
  text = text.replace(/^mailto:/i, '').replace(/^<|>$/g, '');
  var lowered = text.toLowerCase();
  if (/[;,\s]/.test(lowered)) return { value: lowered, issue: 'more than one email, or stray characters' };
  if (!isEmail_(lowered)) return { value: lowered, issue: 'invalid email' };
  return { value: lowered };
}

/** Not libphonenumber: handles the common shapes and flags the rest. */
function cleanPhone_(value, countryCode) {
  var original = normalizeText_(value);
  if (original === '') return { value: '' };
  var text = original;
  var ext = '';
  var m = /^(.*?)\s*(?:ext\.?|extension|x|#)\s*(\d{1,6})$/i.exec(text);
  if (m) {
    text = m[1];
    ext = ' x' + m[2];
  }
  if (/[^\d\s().\-+\/]/.test(text) || text.indexOf('+', 1) !== -1) return { value: original, issue: 'invalid phone' };
  var plus = text.charAt(0) === '+';
  var digits = text.replace(/\D/g, '');
  if (digits === '') return { value: original, issue: 'invalid phone' };
  if (!plus && digits.slice(0, 2) === '00') {
    plus = true;
    digits = digits.slice(2);
  }
  if (plus) {
    if (digits.length < 8 || digits.length > 15) return { value: original, issue: 'phone length is not 8-15 digits' };
    return { value: '+' + digits + ext };
  }
  if (!countryCode) return { value: original, issue: 'phone has no country code (set DEFAULT_COUNTRY_CODE)' };
  if (countryCode === '1') {
    if (digits.length === 10) return { value: '+1' + digits + ext };
    if (digits.length === 11 && digits.charAt(0) === '1') return { value: '+' + digits + ext };
    return { value: original, issue: 'not a 10-digit US/Canada number' };
  }
  var national = digits.charAt(0) === '0' ? digits.slice(1) : digits;
  if (national.length < 6 || national.length > 11) return { value: original, issue: 'phone length unclear for country code ' + countryCode };
  return { value: '+' + countryCode + national + ext };
}

function monthNumber_(name) {
  var full = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
  var n = name.toLowerCase();
  for (var i = 0; i < 12; i++) {
    if (n === full[i] || n === full[i].slice(0, 3) || (i === 8 && n === 'sept')) return i + 1;
  }
  return 0;
}

function pad_(n, width) {
  var s = String(n);
  while (s.length < width) s = '0' + s;
  return s;
}

function makeIsoDate_(y, m, d) {
  if (y < 1900 || y > 2100) return null;
  var t = new Date(Date.UTC(y, m - 1, d));
  if (t.getUTCFullYear() !== y || t.getUTCMonth() !== m - 1 || t.getUTCDate() !== d) return null;
  return pad_(y, 4) + '-' + pad_(m, 2) + '-' + pad_(d, 2);
}

/** dateOrder: 'DMY', 'MDY' or empty (ambiguous dates are then flagged, never guessed). */
function cleanDate_(value, dateOrder, tz) {
  if (isBlankCell_(value)) return { value: '' };
  if (isDate_(value)) {
    if (isNaN(value.getTime())) return { value: '', issue: 'invalid date' };
    return { value: Utilities.formatDate(value, tz, 'yyyy-MM-dd') };
  }
  if (typeof value === 'number') return { value: String(value), issue: 'date is a bare number' };
  var text = normalizeText_(value);
  var m, y, mo, d, iso;

  m = /^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})(?:[T ]\d{1,2}:\d{2}.*)?$/.exec(text);
  if (m) {
    iso = makeIsoDate_(Number(m[1]), Number(m[2]), Number(m[3]));
    return iso ? { value: iso } : { value: text, issue: 'not a real calendar date' };
  }

  m = /^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})$/.exec(text);
  if (m) {
    var a = Number(m[1]), b = Number(m[2]);
    y = Number(m[3]);
    if (a > 12 && b <= 12) { d = a; mo = b; }
    else if (b > 12 && a <= 12) { d = b; mo = a; }
    else if (a === b) { d = a; mo = b; }
    else if (dateOrder === 'DMY') { d = a; mo = b; }
    else if (dateOrder === 'MDY') { mo = a; d = b; }
    else return { value: text, issue: 'ambiguous date (day/month order unknown; set DATE_ORDER)' };
    iso = makeIsoDate_(y, mo, d);
    return iso ? { value: iso } : { value: text, issue: 'not a real calendar date' };
  }

  m = /^(\d{1,2})(?:st|nd|rd|th)?[ \-]([A-Za-z]{3,9})\.?,?[ \-](\d{4})$/.exec(text);
  if (m && monthNumber_(m[2])) {
    iso = makeIsoDate_(Number(m[3]), monthNumber_(m[2]), Number(m[1]));
    return iso ? { value: iso } : { value: text, issue: 'not a real calendar date' };
  }

  m = /^([A-Za-z]{3,9})\.? (\d{1,2})(?:st|nd|rd|th)?,? (\d{4})$/.exec(text);
  if (m && monthNumber_(m[1])) {
    iso = makeIsoDate_(Number(m[3]), monthNumber_(m[1]), Number(m[2]));
    return iso ? { value: iso } : { value: text, issue: 'not a real calendar date' };
  }

  return { value: text, issue: 'unrecognised date format (4-digit year required)' };
}

function cleanText_(value) {
  if (typeof value === 'string') return { value: normalizeText_(value) };
  return { value: value === null || value === undefined ? '' : value };
}

/**
 * options: {outputName='Cleaned', rules, required:[headers], countryCode, dateOrder}.
 * Returns {rows, flagged, blankRowsSkipped, outputSheet, issues:[{row, issues:[...]}]}.
 */
function cleanSheet_(ss, sourceName, options) {
  options = options || {};
  var outName = options.outputName || 'Cleaned';
  if (outName === sourceName) throw new Error('The output sheet must have a different name from the source sheet.');
  var source = getSheetByName_(ss, sourceName);
  var table = readTable_(source);
  var result = { rows: 0, flagged: 0, blankRowsSkipped: 0, outputSheet: null, issues: [] };
  if (table.rows.length === 0) return result;

  var tz = ss.getSpreadsheetTimeZone();
  var rules = resolveCleanRules_(table, options.rules);
  var required = (options.required || []).map(function (h) {
    var idx = headerIndex_(table.headers, h);
    if (idx < 0) throw new Error('Required column "' + h + '" not found in "' + sourceName + '".');
    return idx;
  });

  var cleaned = [];
  var issueText = [];
  table.rows.forEach(function (row, i) {
    if (isBlankRow_(row)) {
      result.blankRowsSkipped++;
      return;
    }
    var issues = [];
    var out = row.map(function (cell, c) {
      var r;
      switch (rules[c]) {
        case 'name': r = cleanName_(cell); break;
        case 'email': r = cleanEmail_(cell); break;
        case 'phone': r = cleanPhone_(cell, options.countryCode || ''); break;
        case 'date': r = cleanDate_(cell, options.dateOrder || '', tz); break;
        case 'text': r = cleanText_(cell); break;
        default: r = { value: cell === null || cell === undefined ? '' : cell };
      }
      if (r.issue) issues.push(table.headers[c] + ': ' + r.issue);
      return r.value;
    });
    required.forEach(function (c) {
      if (isBlankCell_(out[c])) issues.push(table.headers[c] + ': missing');
    });
    cleaned.push(out);
    issueText.push(issues.join('; '));
    if (issues.length) {
      result.flagged++;
      result.issues.push({ row: i + 2, issues: issues });
    }
  });
  result.rows = cleaned.length;
  if (cleaned.length === 0) return result;

  var output = ss.getSheetByName(outName) || ss.insertSheet(outName);
  output.clear();
  var headers = table.headers.concat(['Clean Issues']);
  writeTextBlock_(output, 1, 1, [headers]);

  // A column is written as numbers only when every non-blank value in it is a number; otherwise it is
  // written as text, so IDs such as "007" or "+4479..." survive the trip through Sheets.
  for (var c = 0; c < table.headers.length; c++) {
    var column = cleaned.map(function (row) { return row[c]; });
    var numeric = column.every(function (v) { return v === '' || typeof v === 'number'; });
    if (numeric) {
      output.getRange(2, c + 1, column.length, 1).setValues(column.map(function (v) { return [v]; }));
    } else {
      writeTextBlock_(output, 2, c + 1, column.map(function (v) { return [toText_(v, tz)]; }));
    }
  }
  writeTextBlock_(output, 2, headers.length, issueText.map(function (t) { return [t]; }));
  result.outputSheet = outName;
  return result;
}

/** Menu action: cleans the open sheet into "Cleaned" and reports. */
function cleanActiveSheet() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rules = {};
  var rawRules = getProp_('CLEAN_RULES', null);
  if (rawRules) {
    try {
      rules = JSON.parse(rawRules);
    } catch (err) {
      throw new Error('CLEAN_RULES is not valid JSON: ' + err.message);
    }
  }
  var result = cleanSheet_(ss, ss.getActiveSheet().getName(), {
    rules: rules,
    required: getListProp_('CLEAN_REQUIRED'),
    countryCode: getProp_('DEFAULT_COUNTRY_CODE', ''),
    dateOrder: getProp_('DATE_ORDER', '').toUpperCase()
  });
  ui.alert(result.rows === 0
    ? 'Nothing to clean on this sheet.'
    : 'Cleaned ' + result.rows + ' row(s) into "' + result.outputSheet + '". ' + result.flagged +
      ' row(s) need a look (see the "Clean Issues" column). The original sheet was not changed.');
  return result;
}
