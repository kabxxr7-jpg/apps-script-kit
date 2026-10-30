# apps-script-kit

> **Status:** a proof project (161 tests, run locally on Node v25.2.1 against hand-written fakes of the Apps Script services). **It has never been run inside Google**, is not used in production, and no client has used it. See "Limits" below for exactly what was and was not verified.

Six small, tested Google Sheets / Apps Script building blocks for the usual freelance requests: "remove my duplicates", "clean this export", "receive a webhook into a sheet", "push my sheet rows to an API", "email me what's new", "connect several forms to one sheet without breaking it".

| # | Script | File | What it does |
|---|---|---|---|
| 1 | Remove duplicate rows | `src/dedupe.js` | By a key column (case/space-insensitive), keeps first or last, copies removed rows to `_Removed Rows` before deleting, preview + confirm from a menu |
| 2 | Clean a messy sheet | `src/clean.js` | Writes a new `Cleaned` sheet (source untouched): trims, names, emails, phones to `+country`, dates to `yyyy-mm-dd`; problems it cannot fix go in a `Clean Issues` column |
| 3 | Signed webhook receiver | `src/receiver.js` | `doPost`: HMAC-SHA256 signature + timestamp window (webhook-kit's scheme), event-id de-duplication, appends a row, JSON status in the body |
| 4 | Sheet-to-API sync | `src/sync.js` | POSTs each row, retries with backoff, marks every row `synced` / `failed` / `rejected`, stops before the time limit, resumable |
| 5 | Scheduled digest email | `src/digest.js` | Emails new rows on a time-driven trigger. **Dry run by default**; sends nothing to a list unless `DRY_RUN` is `false` |
| 6 | Form intake | `src/forms.js` | Several Google Forms append to one central sheet, tagged with a `Form Origin` (the form's title); existing rows are never touched |

Shared helpers are in `src/common.js`; `src/selftest.js` holds two checks you run once inside Google; `src/menu.js` adds a "Sheet tools" menu; `src/appsscript.json` is the manifest.

Runtime code is plain Apps Script JavaScript (V8): no dependencies, no build step. The only dev tooling is Node's built-in test runner.

```bash
npm test                  # node --test, 161 tests, no install needed (zero dependencies)
npm run mutation-check    # breaks a copy of src/ in 23 deliberate ways; the suite must catch each
```

## Sample sheet layouts and Script Properties

Everything configurable lives in **Project Settings > Script properties**, never in code. Row 1 is always the header row. Header matching ignores case and extra spaces.

### 1. Duplicates (`Leads`)

```
Name  | Email         | Plan
Ann   | ann@x.com     | Pro
Bo    | bo@x.com      | Free
Ann 2 | ANN@x.com     | Pro      <- removed (same key as row 2)
```

| Property | Default | |
|---|---|---|
| `DEDUPE_KEY_HEADER` | `Email` | column that identifies a record |
| `DEDUPE_KEEP` | `first` | `first` or `last` |

Run **Sheet tools > Remove duplicate rows** (the sheet that is open). It shows what would go and asks first. From code: `dedupeSheet(sheetName, keyHeader, keep, dryRun)` (dry run unless `dryRun` is `false`). Blank keys are never treated as duplicates.

### 2. Clean (`Raw` -> `Cleaned`)

```
Full Name    | Email           | Mobile         | Signup Date | Notes
  JANE  DOE  |  Jane@Example.COM | (555) 123-4567 | 5 Jan 2026  |   hello   there
bob smith    | bob@bad         | 555-12         | 2026-02-30  |
```
becomes
```
Full Name | Email            | Mobile       | Signup Date | Notes       | Clean Issues
Jane Doe  | jane@example.com | +15551234567 | 2026-01-05  | hello there |
Bob Smith | bob@bad          | 555-12       | 2026-02-30  |             | Email: invalid email; Mobile: not a 10-digit US/Canada number; Signup Date: not a real calendar date
```

The rule per column comes from the header (`email`, `phone/mobile/tel`, `date/dob`, `name/first name/...`; anything else is plain `text`), or from `CLEAN_RULES`.

| Property | Default | |
|---|---|---|
| `CLEAN_RULES` | none | JSON, e.g. `{"Mobile":"phone","Ref":"none"}`; rules: `name email phone date text none` |
| `CLEAN_REQUIRED` | none | comma list of headers that must not be blank |
| `DEFAULT_COUNTRY_CODE` | none | e.g. `1` or `91`; without it, phones lacking a `+` are flagged, not guessed |
| `DATE_ORDER` | none | `DMY` or `MDY`; without it, `01/02/2026` is flagged as ambiguous, not guessed |

It never invents data: a value it cannot place is copied through and listed in `Clean Issues`. IDs such as `007` and `+123` survive (text columns are written with the plain-text format). Run **Sheet tools > Clean this sheet**. This is not libphonenumber and not a full date parser (4-digit years only); review the flagged rows.

### 3. Webhook receiver (`Events`)

Deploy the project as a web app (below), then send signed requests to the `/exec` URL:

```
POST <exec URL>?ts=2026-06-01T12:00:00.000Z&sig=<hex or sha256=hex>&id=<event id, optional>
body: the raw JSON exactly as signed
```

Signed string is `${ts}.${rawBody}`, HMAC-SHA256, hex, with the shared secret: **the same scheme as webhook-kit** (a test pins a signature produced by webhook-kit's `signPayload`). Two differences come from the platform: a web app's `doPost` cannot read request headers, so `ts`/`sig`/`id` travel as query parameters; and a web app cannot set the HTTP status, so **read the `status` field of the JSON body** (HTTP itself is 200).

| Situation | Body |
|---|---|
| First delivery of an event id | `{"ok":true,"status":200,"result":"ok"}` and one row appended |
| Same event id again | `{"ok":true,"status":200,"result":"duplicate"}`, nothing written |
| `ts` or `sig` missing / malformed | `{"ok":false,"status":400,"error":"missing_signature" \| "malformed_signature" \| "malformed_timestamp"}` |
| Valid signature but empty body / no usable event id | `400` `empty_body` / `missing_event_id` / `invalid_event_id` |
| Wrong signature, tampered body, timestamp older or further ahead than the window | `401` `invalid_signature` (same body for all; the reason goes to the execution log only) |
| Body over `WEBHOOK_MAX_BODY_CHARS` | `413` `payload_too_large` |
| Another execution holds the lock | `503` `busy` (the sender should retry) |
| Not configured / unexpected error | `500` `not_configured` / `internal_error` (no detail in the response) |
| `GET` | `405` `method_not_allowed` |

Sheet `Events` (created if missing): `Received At | Event ID | Event Type | Payload`. Extra columns, in any order, are fine. The event id is `?id=` first, then `payload.id`. De-duplication reads the sheet (source of truth) under the script lock; the script cache (6 hours) only answers quick retries.

| Property | Default | |
|---|---|---|
| `WEBHOOK_SECRET` | required | the shared secret |
| `WEBHOOK_SECRET_PREVIOUS` | none | also accepted, for rotation |
| `WEBHOOK_TOLERANCE_SECONDS` | `300` | window in both directions |
| `WEBHOOK_MAX_BODY_CHARS` | `40000` | a Sheets cell holds at most 50,000 characters |
| `EVENTS_SHEET` | `Events` | |
| `SPREADSHEET_ID` | bound sheet | set it if the script is not bound to the spreadsheet |

To sign a test request (the secret comes from the environment so it stays out of shell history):

```bash
WEBHOOK_SECRET=... node tools/sign.js --body '{"id":"evt_1","type":"order.created"}' --id evt_1 --url 'https://script.google.com/macros/s/<deployment id>/exec'
```

It prints the query string and a `printf ... | curl -L ...` command. `-L` follows Google's redirect to the response; do not add `-X POST`.

### 4. Sheet-to-API sync (`Orders`)

```
ID    | Customer | Total | Ordered    | Sync Status | Sync Error | Synced At      <- last three are added for you
A-1   | Ann      | 10    | 2026-10-30 | synced      |            | 2026-11-01T...
A-2   | Bo       | 20.5  |            | failed      | HTTP 502: bad gateway [4 attempts] |
A-3   | Cy       | 30    |            |             |            |
```

Blank `Sync Status` = pending. `synced` is never sent again. `failed` (network error, 408, 429, 5xx after the retries) is tried again on the next run. `rejected` (any other non-2xx, e.g. 422) is **not** retried until you clear the cell. Each record is a JSON object keyed by header text; blank cells are `null`, dates are `yyyy-MM-dd`. Headers sent: `Authorization: Bearer <API_TOKEN>` (if set) and `Idempotency-Key` (the `ID` column, so an API that honours the header can drop a repeat whose first response was lost; without an ID column the key is `<sheet>:row<N>`, which changes if rows are sorted).

Retries: up to `SYNC_MAX_ATTEMPTS` per row per run, wait `SYNC_BACKOFF_MS` doubling, capped at 30 s; a numeric `Retry-After` on 429 is honoured (capped at 60 s). The run stops cleanly after `SYNC_MAX_RUNTIME_MS` and reports how many rows are left; run `syncRows` again, or install the trigger with `installSyncTrigger()` (every 15 minutes).

| Property | Default | |
|---|---|---|
| `API_URL` | required | must be `https://` |
| `API_TOKEN` | none | |
| `SYNC_SHEET` | `Orders` | |
| `SYNC_ID_HEADER` | `ID` | |
| `SYNC_FIELDS` | all columns | comma list of headers to send |
| `SYNC_MAX_ATTEMPTS` / `SYNC_BACKOFF_MS` / `SYNC_MAX_RUNTIME_MS` | `4` / `1000` / `270000` | |

One status write per row (so a stopped run loses nothing); with thousands of rows expect several runs.

### 5. Digest email (`Leads`)

```
Name | Email     | Source | Digested At
Ann  | ann@x.com | web    | baseline        <- already there before the digest was switched on
Bo   | bo@x.com  | ads    |                 <- new: goes in the next digest, then gets a timestamp
```

1. Set `DIGEST_RECIPIENTS` (comma list, at most `DIGEST_MAX_RECIPIENTS`, default 10).
2. Run `initDigestBaseline()` once: adds the `Digested At` column and stamps the existing rows `baseline`. A live run refuses to start without that column, so the history is never mailed by accident.
3. Run `sendDigest()` by hand. **It is a dry run until you set `DRY_RUN` to exactly `false`**: it logs what it would send and marks nothing. Optionally set `DIGEST_TEST_RECIPIENT` (one address) to receive the preview, prefixed `[DRY RUN]`, while the list receives nothing.
4. When the preview is right, set `DRY_RUN=false`, then `installDigestTrigger()` (daily at `DIGEST_HOUR`, default 8).

Rows are marked only after `MailApp` accepted the message, so delivery is at-least-once: a crash between the send and the marking would send the same rows again. Cell values are HTML-escaped; at most `DIGEST_MAX_ROWS` (50) rows are shown with a "...and N more" line. Also: `DIGEST_SHEET` (`Leads`), `DIGEST_SUBJECT` (`New rows`).

### 6. Forms into one sheet (`Submissions`)

```
Timestamp | Form Origin | Response ID | Full name | Email | Plan | Respondent Email | Other Answers
```

`Timestamp`, `Form Origin` and `Response ID` are required; the other columns are matched to question titles (case/space-insensitive) and may be in any order. Answers without a matching column go to `Other Answers` as `Question: answer` lines, or are logged and dropped if that column does not exist. Headers are never added or changed after the sheet is created. `Response ID` is `<formId>/<responseId>`; a response already in the sheet is skipped (a retried trigger cannot write twice).

**Existing rows are never touched:** each submission writes one new row below the last used row, under the script lock, and nothing else. A test compares the existing rows' values and formats before and after, and checks that every write lands below them.

Set `FORM_IDS` (comma list), run `installFormTriggers()` once (one on-submit trigger per form; safe to run again). If a trigger run fails (for example the lock was busy), `backfillFormResponses(formId)` adds the responses that are missing.

Everything written from a form (and from the webhook receiver) goes in with the plain-text format, so an answer such as `=HYPERLINK("http://...")`, `007` or `2026-01-05` stays exactly that.

## Deploying

Two ways, both standard Apps Script and neither run here.

**Paste:** open the spreadsheet, Extensions > Apps Script, create one file per `src/*.js` with the same names (keep `common.js`), and in Project Settings tick "Show appsscript.json" and paste `src/appsscript.json`.

**clasp:**

```bash
npm i -g @google/clasp
clasp login                                   # also enable the Apps Script API at script.google.com/home/usersettings
cp .clasp.json.example .clasp.json            # then put your script id in it (git-ignored)
clasp push                                    # pushes src/ including appsscript.json (this replaces the remote manifest)
```

Create the project bound to a spreadsheet (so `getActiveSpreadsheet()` and the menu work), or set `SPREADSHEET_ID`. Then set Script properties (never put secrets in code).

**Web app (script 3 only):** Deploy > New deployment > Web app, execute as **Me**, access **Anyone** (the sender is a server, not a Google user; the signature is the authentication). Use the `/exec` URL. Every code change needs a new deployment version to reach that URL.

The scopes in `appsscript.json` were written by hand; if Google asks for different ones, edit the list.

## Run these two checks first, inside Google

The local suite uses fakes, so these confirm the two platform behaviours the kit depends on. In the Apps Script editor, choose the function and press Run; read the Execution log.

- `selfTestHmac`: must log `PASS`. Confirms `Utilities.computeHmacSha256Signature` reproduces webhook-kit's signatures byte for byte, including a non-ASCII body.
- `selfTestTextCells`: must log `PASS`. Writes tricky strings into plain-text cells of a temporary sheet (deleted afterwards) and reads them back; it also logs, for information, what an unformatted cell does to `007`.

If either logs `FAIL`, do not use the affected scripts until it is understood.

## What the tests are

`test/harness.js` loads every file in `src/` into a bare `node:vm` context, all in one global scope like Apps Script, with only JavaScript built-ins plus the fakes in `test/fakes.js` (`SpreadsheetApp`, `Utilities`, `MailApp`, `UrlFetchApp`, `PropertiesService`, `LockService`, `CacheService`, `ScriptApp`, `ContentService`, `FormApp`). Nothing from Node (`require`, `process`, `Buffer`, `fetch`, timers) is visible to the scripts, and a static test also bans syntax newer engines might not have. Time is a fake clock that `Utilities.sleep` advances, so backoff and time-budget tests are instant and deterministic.

The fakes model a few documented Google behaviours that cause real bugs, so the tests can catch them: a string written to a non-text cell is reinterpreted (`=...` becomes a formula, `007` a number, `2026-01-05` a date); zero-row ranges and wrongly sized `setValues` throw; cells cap at 50,000 characters; HMAC bytes are signed; cache TTL is capped at 6 hours; `UrlFetchApp` throws on 4xx/5xx unless `muteHttpExceptions`; time-based triggers accept only certain minute values. These are modelled from documentation and experience, not checked against Google.

Edge cases covered include: empty and header-only sheets, blank keys, duplicates in every position, bad and ambiguous dates, bad/missing/tampered signatures, expired and future timestamps, replayed event ids (cache hit, cache lost, window passed), secret rotation, lock contention, API failure then success, permanent rejections, a run that hits its time budget and is resumed, dry-run default and every non-`false` spelling, quota and recipient guards, a failing `MailApp`, formula-looking input, and "existing rows untouched".

`npm run mutation-check` copies `src/` to a temp folder, breaks it 23 ways one at a time (signature check removed, timestamp window removed, dry-run default flipped, 5xx no longer retried, rows written over existing ones, ...) and runs the suite on each copy; it exits non-zero if any breakage goes unnoticed. `src/` itself is never edited.

## Limits (read before relying on it)

- **Never run inside Google.** No Google account was used; every result above is from fakes. The two self-tests are how you close that gap for HMAC and text cells. Everything else (menus and `getUi()`, `LockService` under real concurrency, `CacheService`, installable and form triggers, `FormApp` events, `MailApp`, `UrlFetchApp`) behaves here as I modelled it from the documentation, which may differ from Google in details.
- **Web app behaviour is assumed:** that `doPost` cannot read headers and cannot set an HTTP status (hence query parameters and a `status` field in the body), that `e.postData.contents` is the exact raw body that was signed, and that a POST to `/exec` is answered through a redirect (hence `curl -L`). Check the first end-to-end request with `tools/sign.js` before you rely on it.
- **Quotas and execution time are not exercised:** the 6-minute execution limit (the sync budget defaults to 4.5 minutes), `UrlFetchApp` and `MailApp` daily quotas, trigger frequency limits, and Sheets call latency at scale. The sync does one write per row by design.
- **Form triggers:** creating a trigger for a form from a standalone or sheet-bound project (`forForm(id)`), and the event's `source` / `response` objects, follow the documented API but were not exercised. The account that installs the triggers needs edit access to each form.
- **Cleaning is heuristic.** Phones: `+` numbers, US/Canada, and national numbers with `DEFAULT_COUNTRY_CODE` (6-11 national digits); anything else is flagged. Dates: 4-digit years only; ambiguous day/month is flagged unless `DATE_ORDER` is set; the date part of a timestamp is taken as written. All-caps and all-lower-case names are re-cased (`van der Berg` typed in lower case becomes `Van Der Berg`); mixed case is left alone.
- **Receiver:** the signature and timestamp are in the URL, so they can appear in logs along the way; a captured request can only be replayed inside the window, and then only as a duplicate of the same event id. The comparison avoids early exit but JavaScript cannot promise constant time. Senders must read the `status` field, because HTTP is always 200.
- **Digest delivery is at-least-once** (see above). Mail goes to all recipients in one message, so they see each other's addresses.
- Tested on Node v25.2.1 only; the tests use `node:test`, `node:vm` and `node:assert`.

## Design notes

- One shared global scope means name collisions are silent in Apps Script, so helpers end in `_`, only documented entry points are public, and a test fails on a duplicated function name or an undocumented public one.
- Destructive or outward-facing actions are guarded by default: dedupe is a dry run unless asked and archives before deleting, the digest is a dry run unless `DRY_RUN` is `false` and is capped in recipients, sync never retries a permanent rejection, cleaning writes a new sheet.
- Untrusted text (webhook payloads, form answers, server error messages) is always written with the `@` plain-text format.
- The receiver verifies the signature before looking at the timestamp, and gives one answer for every signature or timestamp failure.

## License

MIT, see `LICENSE`.
