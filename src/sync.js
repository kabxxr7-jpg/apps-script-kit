/**
 * Sheet-to-API sync: POSTs each row as JSON, retries transient failures with backoff, and marks
 * every row in the sheet so a run that stops half way can simply be run again.
 *
 * Sheet layout: row 1 headers; one row = one record. The script adds three columns if they are
 * missing: "Sync Status" (blank = pending, "synced", "failed", "rejected"), "Sync Error", "Synced At".
 *   synced   - done; never sent again
 *   failed   - gave up after the retries (network error, 429, 5xx); tried again on the next run
 *   rejected - the API answered with a permanent error (other 4xx); NOT retried until you clear the cell
 *
 * Request: POST API_URL, JSON object keyed by header text (blank cells -> null, dates -> yyyy-MM-dd),
 * `Authorization: Bearer <API_TOKEN>` when API_TOKEN is set, and an `Idempotency-Key` header taken
 * from the ID column (SYNC_ID_HEADER), so an API that honours the header can de-duplicate a retry
 * whose first response was lost. Without an ID column the key falls back to "<sheet>:row<N>",
 * which changes if rows are sorted or deleted.
 *
 * Script properties: API_URL (https only, required), API_TOKEN, SYNC_SHEET (default "Orders"),
 * SYNC_ID_HEADER (default "ID"), SYNC_FIELDS (comma list; default every column), SYNC_MAX_ATTEMPTS (4),
 * SYNC_BACKOFF_MS (1000, doubling, capped at 30000), SYNC_MAX_RUNTIME_MS (270000: stop cleanly before
 * the 6 minute execution limit).
 */

function readSyncConfig_() {
  var url = getProp_('API_URL', null);
  if (!url) throw new Error('Set the API_URL script property.');
  if (!/^https:\/\//i.test(url)) throw new Error('API_URL must start with https:// (the token would otherwise travel unencrypted).');
  var fields = getListProp_('SYNC_FIELDS');
  return {
    url: url,
    token: getSecret_('API_TOKEN'),
    sheetName: getProp_('SYNC_SHEET', 'Orders'),
    idHeader: getProp_('SYNC_ID_HEADER', 'ID'),
    fields: fields.length ? fields : null,
    maxAttempts: Math.max(1, getIntProp_('SYNC_MAX_ATTEMPTS', 4)),
    backoffMs: getIntProp_('SYNC_BACKOFF_MS', 1000),
    maxRuntimeMs: getIntProp_('SYNC_MAX_RUNTIME_MS', 270000)
  };
}

/** Trigger / menu entry point. */
function syncRows() {
  var result = syncRowsToApi_(getSpreadsheet_(), readSyncConfig_());
  console.log('syncRows: ' + JSON.stringify(result));
  return result;
}

function installSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncRows') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncRows').timeBased().everyMinutes(15).create();
}

function syncStatusHeaders_() {
  return ['Sync Status', 'Sync Error', 'Synced At'];
}

function syncRetryAfterSeconds_(response) {
  var headers = response && response.getHeaders ? response.getHeaders() : null;
  if (!headers) return null;
  var keys = Object.keys(headers);
  for (var i = 0; i < keys.length; i++) {
    if (keys[i].toLowerCase() === 'retry-after' && /^\d+$/.test(String(headers[keys[i]]).trim())) {
      return parseInt(String(headers[keys[i]]).trim(), 10);
    }
  }
  return null;
}

function syncBackoffMs_(cfg, attempt, retryAfterSeconds) {
  var wait = Math.min(cfg.backoffMs * Math.pow(2, attempt - 1), 30000);
  if (retryAfterSeconds !== null) wait = Math.min(Math.max(retryAfterSeconds * 1000, wait), 60000);
  return wait;
}

/**
 * Sends one record. Transient failures (network error, 408, 429, 5xx) are retried up to
 * cfg.maxAttempts times; any other non-2xx answer is permanent. Never retries past `deadlineMs`.
 * Returns {ok, permanent, error, attempts}.
 */
function postWithRetry_(cfg, record, idempotencyKey, deadlineMs) {
  var headers = { 'Idempotency-Key': idempotencyKey };
  if (cfg.token) headers.Authorization = 'Bearer ' + cfg.token;
  var lastError = '';
  for (var attempt = 1; attempt <= cfg.maxAttempts; attempt++) {
    var response = null;
    var code = 0;
    var networkError = null;
    try {
      response = UrlFetchApp.fetch(cfg.url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(record),
        headers: headers,
        muteHttpExceptions: true
      });
      code = response.getResponseCode();
    } catch (err) {
      networkError = oneLine_(err && err.message ? err.message : err, 150);
    }
    if (networkError === null && code >= 200 && code < 300) return { ok: true, permanent: false, error: '', attempts: attempt };

    lastError = networkError !== null
      ? 'Network error: ' + networkError
      : 'HTTP ' + code + ': ' + oneLine_(response.getContentText(), 150);
    var transient = networkError !== null || code === 408 || code === 429 || code >= 500;
    if (!transient) return { ok: false, permanent: true, error: lastError, attempts: attempt };

    if (attempt < cfg.maxAttempts) {
      var wait = syncBackoffMs_(cfg, attempt, code === 429 ? syncRetryAfterSeconds_(response) : null);
      if (new Date().getTime() + wait >= deadlineMs) {
        return { ok: false, permanent: false, error: lastError + ' (stopped: out of time)', attempts: attempt };
      }
      Utilities.sleep(wait);
    }
  }
  return { ok: false, permanent: false, error: lastError, attempts: cfg.maxAttempts };
}

function jsonCellValue_(value, tz) {
  if (isBlankCell_(value)) return null;
  if (isDate_(value)) return isNaN(value.getTime()) ? null : Utilities.formatDate(value, tz, 'yyyy-MM-dd');
  return value;
}

/** Writes status/error/time for one row; one call when the three columns are adjacent. */
function writeSyncStatus_(sheet, rowNumber, cols, status, error, syncedAt) {
  var values = [status, error, syncedAt];
  var c = [cols['Sync Status'], cols['Sync Error'], cols['Synced At']];
  if (c[1] === c[0] + 1 && c[2] === c[1] + 1) {
    sheet.getRange(rowNumber, c[0] + 1, 1, 3).setValues([values]);
  } else {
    for (var i = 0; i < 3; i++) sheet.getRange(rowNumber, c[i] + 1, 1, 1).setValues([[values[i]]]);
  }
}

/**
 * Returns {status:'done'|'stopped_early'|'busy'|'nothing_to_do', processed, synced, failed,
 * rejected, skipped, remaining}. `remaining` counts pending rows not attempted because time ran out.
 */
function syncRowsToApi_(ss, cfg) {
  var started = new Date().getTime();
  var deadline = started + cfg.maxRuntimeMs;
  var summary = { status: 'done', processed: 0, synced: 0, failed: 0, rejected: 0, skipped: 0, remaining: 0 };

  var locked = withScriptLock_(1000, function () {
    var sheet = getSheetByName_(ss, cfg.sheetName);
    var tz = ss.getSpreadsheetTimeZone();
    var table = readTable_(sheet);
    if (table.rows.length === 0) {
      summary.status = 'nothing_to_do';
      return;
    }
    var cols = ensureColumns_(sheet, table, syncStatusHeaders_());
    var statusCols = syncStatusHeaders_().map(function (h) { return cols[h]; });
    statusCols.forEach(function (c) {
      sheet.getRange(2, c + 1, table.rows.length, 1).setNumberFormat('@');
    });

    var fieldCols = [];
    if (cfg.fields) {
      cfg.fields.forEach(function (name) {
        var idx = headerIndex_(table.headers, name);
        if (idx < 0) throw new Error('SYNC_FIELDS names "' + name + '" but the sheet has no such column.');
        fieldCols.push(idx);
      });
    } else {
      table.headers.forEach(function (h, c) {
        if (h !== '' && statusCols.indexOf(c) < 0) fieldCols.push(c);
      });
    }
    var idCol = headerIndex_(table.headers, cfg.idHeader);

    for (var i = 0; i < table.rows.length; i++) {
      var row = table.rows[i];
      var rowNumber = i + 2;
      var current = normalizeKey_(row[cols['Sync Status']]);
      if (current === 'synced' || current === 'rejected') {
        summary.skipped++;
        continue;
      }
      var hasData = false;
      for (var c = 0; c < row.length; c++) {
        if (statusCols.indexOf(c) < 0 && !isBlankCell_(row[c])) hasData = true;
      }
      if (!hasData) {
        summary.skipped++;
        continue;
      }
      if (new Date().getTime() >= deadline) {
        summary.status = 'stopped_early';
        summary.remaining++;
        continue;
      }

      var record = {};
      fieldCols.forEach(function (fc) { record[table.headers[fc]] = jsonCellValue_(row[fc], tz); });
      var key = idCol >= 0 && !isBlankCell_(row[idCol]) ? String(row[idCol]) : sheet.getName() + ':row' + rowNumber;
      var outcome = postWithRetry_(cfg, record, key, deadline);
      summary.processed++;
      if (outcome.ok) {
        summary.synced++;
        writeSyncStatus_(sheet, rowNumber, cols, 'synced', '', new Date().toISOString());
      } else if (outcome.permanent) {
        summary.rejected++;
        writeSyncStatus_(sheet, rowNumber, cols, 'rejected', outcome.error, '');
      } else {
        summary.failed++;
        writeSyncStatus_(sheet, rowNumber, cols, 'failed', outcome.error + ' [' + outcome.attempts + ' attempts]', '');
      }
    }
  });
  if (!locked.locked) summary.status = 'busy';
  return summary;
}
