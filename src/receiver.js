/**
 * Signed webhook receiver (Apps Script web app).
 *
 * Same signing scheme as webhook-kit and Opsflow: HMAC-SHA256, hex, over `${timestamp}.${rawBody}`,
 * with a replay window in both directions. Two things differ because of the platform:
 *
 *  1. doPost cannot read request headers, so the signature, timestamp and (optionally) event id
 *     arrive as URL query parameters:  .../exec?ts=<ISO-8601>&sig=<hex or sha256=hex>&id=<event id>
 *  2. A web app cannot set the HTTP status code of its response, so the status is the `status`
 *     field of the JSON body (HTTP itself is 200). Senders must read that field.
 *
 * Script properties: WEBHOOK_SECRET (required), WEBHOOK_SECRET_PREVIOUS (optional, accepted during
 * rotation), WEBHOOK_TOLERANCE_SECONDS (default 300), WEBHOOK_MAX_BODY_CHARS (default 40000),
 * EVENTS_SHEET (default "Events"), SPREADSHEET_ID (if the script is not bound to the sheet).
 *
 * Response JSON: {ok, status, result} on success (status 200, result "ok" or "duplicate"), or
 * {ok:false, status, error} otherwise:
 *   400 missing_signature / malformed_signature / malformed_timestamp / empty_body / missing_event_id / invalid_event_id
 *   401 invalid_signature (wrong signature, tampered body, expired or future timestamp: one answer for all)
 *   405 method_not_allowed (GET) | 413 payload_too_large | 503 busy (retry later) | 500 not_configured / internal_error
 */

function jsonResponse_(status, body) {
  var payload = { ok: status >= 200 && status < 300, status: status };
  Object.keys(body || {}).forEach(function (k) { payload[k] = body[k]; });
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return jsonResponse_(405, { error: 'method_not_allowed' });
}

function doPost(e) {
  var result;
  try {
    result = handleWebhookRequest_(e, new Date().getTime());
  } catch (err) {
    // Detail goes to the execution log only; the caller learns nothing about the internals.
    console.error('webhook internal error: ' + (err && err.message ? err.message : err));
    result = { status: 500, body: { error: 'internal_error' } };
  }
  return jsonResponse_(result.status, result.body);
}

function webhookReject_(status, error, reason) {
  console.warn('webhook rejected: ' + status + ' ' + error + (reason ? ' (' + reason + ')' : ''));
  return { status: status, body: { error: error }, reason: reason || error };
}

function getWebhookSecrets_() {
  return [getSecret_('WEBHOOK_SECRET'), getSecret_('WEBHOOK_SECRET_PREVIOUS')].filter(function (s) { return s !== null; });
}

/** Pure request handling: returns {status, body, reason?}. `nowMs` is injectable for tests. */
function handleWebhookRequest_(e, nowMs) {
  var secrets = getWebhookSecrets_();
  if (secrets.length === 0) return webhookReject_(500, 'not_configured', 'no_secret');

  var params = (e && e.parameter) || {};
  var body = e && e.postData && typeof e.postData.contents === 'string' ? e.postData.contents : '';
  if (body.length > getIntProp_('WEBHOOK_MAX_BODY_CHARS', 40000)) return webhookReject_(413, 'payload_too_large', 'body_too_large');

  if (!params.sig || !params.ts) return webhookReject_(400, 'missing_signature', 'missing_param');
  var signature = String(params.sig).replace(/^sha256=/i, '');
  if (!/^[0-9a-fA-F]{64}$/.test(signature)) return webhookReject_(400, 'malformed_signature', 'malformed_signature');
  var timestamp = String(params.ts);
  var timestampMs = parseIsoTimestamp_(timestamp);
  if (timestampMs === null) return webhookReject_(400, 'malformed_timestamp', 'malformed_timestamp');

  // Compare against every secret without stopping at the first match.
  var matched = false;
  secrets.forEach(function (secret) {
    if (constantTimeEquals_(hmacSha256Hex_(timestamp + '.' + body, secret), signature.toLowerCase())) matched = true;
  });
  if (!matched) return webhookReject_(401, 'invalid_signature', 'signature_mismatch');

  // Same answer for every timestamp failure, so a caller cannot probe the window.
  var toleranceMs = getIntProp_('WEBHOOK_TOLERANCE_SECONDS', 300) * 1000;
  var age = nowMs - timestampMs;
  if (age > toleranceMs) return webhookReject_(401, 'invalid_signature', 'timestamp_too_old');
  if (-age > toleranceMs) return webhookReject_(401, 'invalid_signature', 'timestamp_in_future');

  if (body === '') return webhookReject_(400, 'empty_body', 'empty_body');

  var parsed = null;
  try {
    parsed = JSON.parse(body);
  } catch (err) {
    parsed = null;
  }
  var isObject = parsed !== null && typeof parsed === 'object' && !Array.isArray(parsed);
  var rawId = params.id !== undefined && params.id !== '' ? params.id : (isObject ? parsed.id : undefined);
  if (rawId === undefined || rawId === null || rawId === '') return webhookReject_(400, 'missing_event_id', 'missing_event_id');
  var eventId = String(rawId);
  if (eventId.length > 200 || /[\u0000-\u001f]/.test(eventId)) return webhookReject_(400, 'invalid_event_id', 'invalid_event_id');
  var eventType = params.type ? String(params.type) : (isObject && typeof parsed.type === 'string' ? parsed.type : '');

  var outcome = recordEvent_(eventId, oneLine_(eventType, 200), body, new Date(nowMs).toISOString());
  if (outcome === 'busy') return webhookReject_(503, 'busy', 'lock_timeout');
  return { status: 200, body: { result: outcome === 'duplicate' ? 'duplicate' : 'ok' } };
}

function eventsHeaders_() {
  return ['Received At', 'Event ID', 'Event Type', 'Payload'];
}

function getEventsSheet_() {
  var ss = getSpreadsheet_();
  var name = getProp_('EVENTS_SHEET', 'Events');
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    writeTextBlock_(sheet, 1, 1, [eventsHeaders_()]);
  }
  return sheet;
}

/**
 * De-duplicates on Event ID and appends one row. The sheet is the source of truth; the script
 * cache (6 hours) only answers quick retries without taking the lock or reading the sheet.
 * Returns 'ok', 'duplicate' or 'busy'.
 */
function recordEvent_(eventId, eventType, body, receivedAtIso) {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'evt:' + eventId;
  if (cache.get(cacheKey)) return 'duplicate';

  var locked = withScriptLock_(10000, function () {
    var sheet = getEventsSheet_();
    var headers = readHeaders_(sheet);
    var cols = {};
    eventsHeaders_().forEach(function (h) {
      cols[h] = headerIndex_(headers, h);
      if (cols[h] < 0) throw new Error('Events sheet "' + sheet.getName() + '" has no "' + h + '" column in row 1.');
    });

    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var ids = sheet.getRange(2, cols['Event ID'] + 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === eventId) {
          cache.put(cacheKey, '1', 21600);
          return 'duplicate';
        }
      }
    }

    var row = headers.map(function () { return ''; });
    row[cols['Received At']] = receivedAtIso;
    row[cols['Event ID']] = eventId;
    row[cols['Event Type']] = eventType;
    row[cols['Payload']] = body;
    writeTextBlock_(sheet, Math.max(lastRow, 1) + 1, 1, [row]);
    cache.put(cacheKey, '1', 21600);
    return 'ok';
  });
  return locked.locked ? locked.value : 'busy';
}
