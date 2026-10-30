/**
 * Remove duplicate rows by a key column.
 *
 * Sheet layout: row 1 is the header; the key column is named by DEDUPE_KEY_HEADER (default "Email").
 * Keys are compared after trimming, collapsing spaces and lower-casing, so "A@x.com " and "a@x.com"
 * are the same key. Rows whose key is blank are never treated as duplicates.
 *
 * Safety: removed rows are copied to the "_Removed Rows" sheet before they are deleted, and the
 * menu action shows what would be removed and asks first.
 */

/**
 * options: {keyHeader, keep: 'first'|'last', dryRun (default true)}.
 * Returns {sheet, keyHeader, keep, dryRun, scanned, blankKeyRows, removed:[{row, key, duplicateOfRow}], archivedTo}.
 */
function dedupeRows_(sheet, options) {
  var keep = options.keep === 'last' ? 'last' : 'first';
  var dryRun = options.dryRun !== false;
  var ss = sheet.getParent();
  var tz = ss.getSpreadsheetTimeZone();
  var table = readTable_(sheet);
  var report = {
    sheet: sheet.getName(),
    keyHeader: options.keyHeader,
    keep: keep,
    dryRun: dryRun,
    scanned: table.rows.length,
    blankKeyRows: 0,
    removed: [],
    archivedTo: null
  };
  if (table.rows.length === 0) return report;

  var col = headerIndex_(table.headers, options.keyHeader);
  if (col < 0) {
    throw new Error('Key column "' + options.keyHeader + '" not found. Headers found: ' +
      table.headers.filter(function (h) { return h !== ''; }).join(', '));
  }

  var groups = new Map();
  table.rows.forEach(function (row, i) {
    var key = normalizeKey_(toText_(row[col], tz));
    if (key === '') {
      report.blankKeyRows++;
      return;
    }
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(i + 2);
  });

  groups.forEach(function (rowNumbers) {
    if (rowNumbers.length < 2) return;
    var survivor = keep === 'first' ? rowNumbers[0] : rowNumbers[rowNumbers.length - 1];
    rowNumbers.forEach(function (rowNumber) {
      if (rowNumber === survivor) return;
      report.removed.push({
        row: rowNumber,
        key: normalizeText_(toText_(table.rows[rowNumber - 2][col], tz)),
        duplicateOfRow: survivor
      });
    });
  });
  report.removed.sort(function (a, b) { return a.row - b.row; });

  if (dryRun || report.removed.length === 0) return report;

  archiveRemovedRows_(ss, sheet, table, report.removed, tz);
  report.archivedTo = removedRowsSheetName_();
  deleteRowNumbers_(sheet, report.removed.map(function (r) { return r.row; }));
  return report;
}

function removedRowsSheetName_() {
  return '_Removed Rows';
}

/** Appends a header row and the removed rows (as text) to the archive sheet. */
function archiveRemovedRows_(ss, sheet, table, removed, tz) {
  var name = removedRowsSheetName_();
  var archive = ss.getSheetByName(name) || ss.insertSheet(name);
  var width = 4 + table.headers.length;
  var stamp = new Date().toISOString();
  var matrix = [['Removed At', 'Source Sheet', 'Source Row', 'Key'].concat(table.headers)];
  removed.forEach(function (r) {
    var cells = table.rows[r.row - 2].map(function (v) { return toText_(v, tz); });
    matrix.push([stamp, sheet.getName(), String(r.row), r.key].concat(cells));
  });
  writeTextBlock_(archive, archive.getLastRow() + 1, 1, matrix.map(function (row) {
    while (row.length < width) row.push('');
    return row;
  }));
}

/** Deletes rows bottom-up, one deleteRows call per run of consecutive rows. */
function deleteRowNumbers_(sheet, rowNumbers) {
  var sorted = rowNumbers.slice().sort(function (a, b) { return a - b; });
  var runs = [];
  sorted.forEach(function (n) {
    var last = runs[runs.length - 1];
    if (last && n === last.start + last.count) last.count++;
    else runs.push({ start: n, count: 1 });
  });
  for (var i = runs.length - 1; i >= 0; i--) sheet.deleteRows(runs[i].start, runs[i].count);
}

/** Public, scriptable entry point. Dry run unless dryRun is explicitly false. */
function dedupeSheet(sheetName, keyHeader, keep, dryRun) {
  var sheet = getSheetByName_(getSpreadsheet_(), sheetName);
  return dedupeRows_(sheet, { keyHeader: keyHeader, keep: keep, dryRun: dryRun !== false });
}

/** Menu action: preview, ask, then remove. Uses the sheet that is open. */
function removeDuplicateRows() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var options = {
    keyHeader: getProp_('DEDUPE_KEY_HEADER', 'Email'),
    keep: getProp_('DEDUPE_KEEP', 'first'),
    dryRun: true
  };
  var preview = dedupeRows_(sheet, options);
  if (preview.removed.length === 0) {
    ui.alert('No duplicates found in "' + preview.sheet + '" by "' + preview.keyHeader + '" (' + preview.scanned + ' rows checked).');
    return preview;
  }
  var lines = preview.removed.slice(0, 10).map(function (r) {
    return 'Row ' + r.row + ' ("' + r.key + '") duplicates row ' + r.duplicateOfRow;
  });
  if (preview.removed.length > 10) lines.push('...and ' + (preview.removed.length - 10) + ' more');
  var answer = ui.alert(
    'Remove ' + preview.removed.length + ' duplicate row(s)?',
    lines.join('\n') + '\n\nA copy goes to the "' + removedRowsSheetName_() + '" sheet first.',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return preview;
  options.dryRun = false;
  var report = dedupeRows_(sheet, options);
  ui.alert('Removed ' + report.removed.length + ' row(s). A copy is in "' + removedRowsSheetName_() + '".');
  return report;
}
