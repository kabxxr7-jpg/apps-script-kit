/**
 * Scheduled digest email of new rows.
 *
 * DRY RUN IS THE DEFAULT. Nothing is sent to DIGEST_RECIPIENTS unless the script property DRY_RUN is
 * exactly "false" (any case). In a dry run no row is marked as digested; the run only logs what it
 * would send, and, if DIGEST_TEST_RECIPIENT holds ONE address, sends the preview to that address only.
 *
 * Sheet layout: row 1 headers; a row is "new" while its "Digested At" cell is blank. The column is
 * created by initDigestBaseline(), which also stamps the rows that already exist so the first live
 * run does not mail the whole history. A live run refuses to start until that column exists.
 *
 * Script properties: DIGEST_SHEET (default "Leads"), DIGEST_RECIPIENTS (comma list), DIGEST_SUBJECT
 * (default "New rows"), DIGEST_TEST_RECIPIENT, DIGEST_MAX_RECIPIENTS (10), DIGEST_MAX_ROWS (50 rows shown),
 * DIGEST_HOUR (0-23, default 8, used by installDigestTrigger), DRY_RUN.
 *
 * Delivery is at-least-once: rows are marked only after MailApp accepted the message, so a crash
 * between the send and the marking would send the same rows again on the next run.
 */

function isDryRun_() {
  return String(getProp_('DRY_RUN', 'true')).toLowerCase() !== 'false';
}

function readDigestConfig_() {
  return {
    sheetName: getProp_('DIGEST_SHEET', 'Leads'),
    dryRun: isDryRun_(),
    recipients: getListProp_('DIGEST_RECIPIENTS'),
    testRecipient: getProp_('DIGEST_TEST_RECIPIENT', null),
    maxRecipients: getIntProp_('DIGEST_MAX_RECIPIENTS', 10),
    maxRows: getIntProp_('DIGEST_MAX_ROWS', 50),
    subject: getProp_('DIGEST_SUBJECT', 'New rows')
  };
}

/** Trigger / menu entry point. */
function sendDigest() {
  var result = runDigest_(getSpreadsheet_(), readDigestConfig_());
  console.log('sendDigest: ' + JSON.stringify(result));
  return result;
}

/** Run once: adds the "Digested At" column and stamps every existing row "baseline". */
function initDigestBaseline() {
  var sheet = getSheetByName_(getSpreadsheet_(), readDigestConfig_().sheetName);
  var table = readTable_(sheet);
  if (table.headers.length === 0) throw new Error('The digest sheet is empty: add a header row first.');
  var existed = headerIndex_(table.headers, 'Digested At') >= 0;
  var cols = ensureColumns_(sheet, table, ['Digested At']);
  var tz = getSpreadsheet_().getSpreadsheetTimeZone();
  var stamped = 0;
  var values = table.rows.map(function (row) {
    if (isBlankRow_(row) || !isBlankCell_(row[cols['Digested At']])) return [toText_(row[cols['Digested At']], tz)];
    stamped++;
    return ['baseline'];
  });
  if (values.length > 0) writeTextBlock_(sheet, 2, cols['Digested At'] + 1, values);
  return { columnCreated: !existed, stamped: stamped };
}

function installDigestTrigger() {
  var hour = getIntProp_('DIGEST_HOUR', 8);
  if (hour > 23) throw new Error('DIGEST_HOUR must be 0-23.');
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendDigest') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendDigest').timeBased().everyDays(1).atHour(hour).create();
}

function buildDigest_(headers, rows, tz, cfg, dryRun) {
  var shown = rows.slice(0, cfg.maxRows);
  var cols = [];
  headers.forEach(function (h, c) {
    if (h !== '' && h.toLowerCase() !== 'digested at' && cols.length < 10) cols.push(c);
  });
  var more = rows.length - shown.length;
  var subject = (dryRun ? '[DRY RUN] ' : '') + cfg.subject + ' (' + rows.length + ')';

  var lines = shown.map(function (row) {
    return cols.map(function (c) { return headers[c] + ': ' + toText_(row[c], tz); }).join(' | ');
  });
  if (more > 0) lines.push('...and ' + more + ' more row(s) in the sheet.');

  var html = '<p>' + escapeHtml_(rows.length + ' new row(s).') + '</p><table border="1" cellpadding="4" cellspacing="0"><tr>' +
    cols.map(function (c) { return '<th>' + escapeHtml_(headers[c]) + '</th>'; }).join('') + '</tr>' +
    shown.map(function (row) {
      return '<tr>' + cols.map(function (c) { return '<td>' + escapeHtml_(toText_(row[c], tz)) + '</td>'; }).join('') + '</tr>';
    }).join('') + '</table>' +
    (more > 0 ? '<p>' + escapeHtml_('...and ' + more + ' more row(s) in the sheet.') + '</p>' : '');
  return { subject: subject, body: rows.length + ' new row(s).\n\n' + lines.join('\n'), htmlBody: html };
}

/**
 * Returns {status, ...}. status: 'nothing_new' | 'dry_run' | 'sent' | 'needs_baseline' | 'busy' |
 * 'no_recipients' | 'too_many_recipients' | 'invalid_recipient' | 'quota_too_low'.
 * Throws if MailApp fails (rows stay unmarked, so the next run tries again).
 */
function runDigest_(ss, cfg) {
  var result = null;
  var locked = withScriptLock_(1000, function () {
    var sheet = getSheetByName_(ss, cfg.sheetName);
    var tz = ss.getSpreadsheetTimeZone();
    var table = readTable_(sheet);
    var markerCol = headerIndex_(table.headers, 'Digested At');
    if (markerCol < 0 && !cfg.dryRun) {
      result = { status: 'needs_baseline', message: 'Run initDigestBaseline() once before the first live digest.' };
      return;
    }

    var fresh = [];
    var freshRowNumbers = [];
    table.rows.forEach(function (row, i) {
      if (isBlankRow_(row)) return;
      if (markerCol >= 0 && !isBlankCell_(row[markerCol])) return;
      fresh.push(row);
      freshRowNumbers.push(i + 2);
    });
    if (fresh.length === 0) {
      result = { status: 'nothing_new' };
      return;
    }

    var recipients = cfg.recipients;
    if (recipients.length > cfg.maxRecipients) {
      result = { status: 'too_many_recipients', count: recipients.length, max: cfg.maxRecipients };
      return;
    }
    var bad = recipients.filter(function (r) { return !isEmail_(r); });
    if (bad.length) {
      result = { status: 'invalid_recipient', invalid: bad };
      return;
    }
    if (!cfg.dryRun && recipients.length === 0) {
      result = { status: 'no_recipients' };
      return;
    }

    var digest = buildDigest_(table.headers, fresh, tz, cfg, cfg.dryRun);

    if (cfg.dryRun) {
      console.log('DRY RUN: would send "' + digest.subject + '" to ' + (recipients.join(', ') || '(no recipients set)') + '\n' + digest.body);
      var sentTo = null;
      if (cfg.testRecipient) {
        if (!isEmail_(cfg.testRecipient)) {
          result = { status: 'invalid_recipient', invalid: [cfg.testRecipient] };
          return;
        }
        MailApp.sendEmail({ to: cfg.testRecipient, subject: digest.subject, body: digest.body, htmlBody: digest.htmlBody });
        sentTo = cfg.testRecipient;
      }
      result = { status: 'dry_run', rows: fresh.length, wouldSendTo: recipients, testSentTo: sentTo, subject: digest.subject };
      return;
    }

    if (MailApp.getRemainingDailyQuota() < recipients.length) {
      result = { status: 'quota_too_low', remaining: MailApp.getRemainingDailyQuota(), needed: recipients.length };
      return;
    }
    MailApp.sendEmail({ to: recipients.join(','), subject: digest.subject, body: digest.body, htmlBody: digest.htmlBody });

    var stamp = new Date().toISOString();
    var marks = table.rows.map(function (row) { return [toText_(row[markerCol], tz)]; });
    freshRowNumbers.forEach(function (n) { marks[n - 2] = [stamp]; });
    writeTextBlock_(sheet, 2, markerCol + 1, marks);
    result = { status: 'sent', rows: fresh.length, sentTo: recipients, subject: digest.subject };
  });
  return locked.locked ? result : { status: 'busy' };
}
