/**
 * Several Google Forms -> one central sheet, tagged with where each answer came from, without ever
 * touching rows that are already there.
 *
 * How it works: one installable "on form submit" trigger per form (installFormTriggers creates them
 * from the FORM_IDS property). Each submission appends exactly one new row at the bottom of the
 * central sheet; existing rows are never read back, rewritten, re-sorted or deleted.
 *
 * Central sheet layout (row 1; columns may be in any order, extra columns are left blank):
 *   Timestamp | Form Origin | Response ID | <one column per question title> | Respondent Email | Other Answers
 *   - Form Origin: the form's title when the answer arrived
 *   - Response ID: "<formId>/<responseId>"; a submission whose Response ID is already in the sheet is skipped
 *   - a question's answer goes into the column with the same header (case/space-insensitive); answers
 *     with no matching column are written to "Other Answers" as "Question: answer" lines, or dropped
 *     (and logged) if that column does not exist. Headers are never added or changed by this script.
 *
 * Script properties: FORM_IDS (comma list), FORMS_SHEET (default "Submissions"), SPREADSHEET_ID (the
 * central spreadsheet; required unless the script is bound to it).
 */

function formsRequiredHeaders_() {
  return ['Timestamp', 'Form Origin', 'Response ID'];
}

function formOriginLabel_(form) {
  var title = oneLine_(form.getTitle() || '', 200);
  return title === '' ? '(untitled form)' : title;
}

function formAnswerText_(answer) {
  if (answer === null || answer === undefined) return '';
  if (Array.isArray(answer)) {
    return answer.map(function (a) { return Array.isArray(a) ? a.join(', ') : String(a); }).join('; ');
  }
  return String(answer);
}

function getSubmissionsSheet_(ss) {
  var name = getProp_('FORMS_SHEET', 'Submissions');
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    writeTextBlock_(sheet, 1, 1, [formsRequiredHeaders_().concat(['Respondent Email', 'Other Answers'])]);
  }
  return sheet;
}

/**
 * Appends one response to the central sheet. Returns {status:'appended'|'duplicate', unmapped:[question titles]}.
 * Throws if the sheet lacks Timestamp / Form Origin / Response ID, or the lock cannot be obtained
 * (a thrown error shows up as a failed execution; backfillFormResponses() recovers missed ones).
 */
function appendFormResponse_(ss, form, response) {
  var sheet = getSubmissionsSheet_(ss);
  var origin = formOriginLabel_(form);
  var responseKey = form.getId() + '/' + response.getId();
  var outcome = null;

  var locked = withScriptLock_(30000, function () {
    var headers = readHeaders_(sheet);
    var cols = {};
    formsRequiredHeaders_().forEach(function (h) {
      cols[h] = headerIndex_(headers, h);
      if (cols[h] < 0) throw new Error('Sheet "' + sheet.getName() + '" needs a "' + h + '" column in row 1.');
    });

    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var ids = sheet.getRange(2, cols['Response ID'] + 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === responseKey) {
          outcome = { status: 'duplicate', unmapped: [] };
          return;
        }
      }
    }

    var row = headers.map(function () { return ''; });
    row[cols['Timestamp']] = response.getTimestamp().toISOString();
    row[cols['Form Origin']] = origin;
    row[cols['Response ID']] = responseKey;

    var emailCol = headerIndex_(headers, 'Respondent Email');
    var email = response.getRespondentEmail();
    if (emailCol >= 0 && email) row[emailCol] = String(email);

    var otherCol = headerIndex_(headers, 'Other Answers');
    var unmapped = [];
    var otherLines = [];
    response.getItemResponses().forEach(function (itemResponse) {
      var title = normalizeText_(itemResponse.getItem().getTitle());
      var text = formAnswerText_(itemResponse.getResponse());
      var col = headerIndex_(headers, title);
      var reserved = formsRequiredHeaders_().concat(['Respondent Email', 'Other Answers']).map(normalizeKey_).indexOf(normalizeKey_(title)) >= 0;
      if (col >= 0 && !reserved) {
        row[col] = text;
      } else {
        unmapped.push(title);
        otherLines.push(title + ': ' + text);
      }
    });
    if (otherCol >= 0 && otherLines.length) row[otherCol] = otherLines.join('\n');
    if (otherCol < 0 && unmapped.length) {
      console.warn('Form "' + origin + '": no column for ' + unmapped.join(', ') + ' (add matching headers or an "Other Answers" column).');
    }

    writeTextBlock_(sheet, Math.max(lastRow, 1) + 1, 1, [row]);
    outcome = { status: 'appended', unmapped: unmapped };
  });
  if (!locked.locked) throw new Error('Could not get the script lock within 30 seconds; the submission was not written.');
  return outcome;
}

/** Installable on-form-submit handler. `e.source` is the form, `e.response` the submission. */
function appendFormResponse(e) {
  return appendFormResponse_(getSpreadsheet_(), e.source, e.response);
}

/** Adds any responses of one form that are missing from the central sheet (e.g. after a failed trigger). */
function backfillFormResponses(formId) {
  var ss = getSpreadsheet_();
  var form = FormApp.openById(formId);
  var added = 0;
  form.getResponses().forEach(function (response) {
    if (appendFormResponse_(ss, form, response).status === 'appended') added++;
  });
  return { formId: formId, added: added };
}

/** Creates one on-form-submit trigger per id in FORM_IDS (skips forms that already have one). */
function installFormTriggers() {
  var ids = getListProp_('FORM_IDS');
  if (ids.length === 0) throw new Error('Set the FORM_IDS script property (comma separated form IDs).');
  var existing = ScriptApp.getProjectTriggers()
    .filter(function (t) { return t.getHandlerFunction() === 'appendFormResponse'; })
    .map(function (t) { return t.getTriggerSourceId(); });
  var created = [];
  ids.forEach(function (id) {
    if (existing.indexOf(id) >= 0) return;
    ScriptApp.newTrigger('appendFormResponse').forForm(id).onFormSubmit().create();
    created.push(id);
  });
  return { created: created, alreadyInstalled: ids.filter(function (id) { return created.indexOf(id) < 0; }) };
}
