#!/usr/bin/env node
'use strict';
/**
 * Sanity check for the test suite: copies src/ to a temp directory, breaks the copy in a set of
 * deliberate ways (one at a time), and runs the suite against each broken copy. A breakage that
 * leaves the suite green is reported as SURVIVED and the script exits 1. src/ itself is never edited.
 *
 *   npm run mutation-check
 */
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

const ROOT = path.join(__dirname, '..');
const SRC = path.join(ROOT, 'src');

const MUTATIONS = [
  ['dedupe.js', "var survivor = keep === 'first'", "var survivor = keep !== 'first'", 'keep first/last swapped'],
  ['dedupe.js', 'archiveRemovedRows_(ss, sheet, table, report.removed, tz);', '', 'rows deleted without an archive copy'],
  ['clean.js', "else return { value: text, issue: 'ambiguous date", "else if (true) { mo = a; d = b; } else return { value: text, issue: 'ambiguous date", 'ambiguous dates guessed instead of flagged'],
  ['clean.js', "if (dateOrder === 'DMY')", "if (dateOrder !== 'DMY')", 'DATE_ORDER inverted'],
  ['clean.js', '!issue && /[A-Za-z', '/[A-Za-z', 'names with digits get re-cased'],
  ['common.js', "range.setNumberFormat('@');\n  range.setValues(matrix);", 'range.setValues(matrix);', 'text cells written without the plain-text format'],
  ['common.js', 'var b = bytes[i] & 0xff;', 'var b = bytes[i];', 'signed HMAC bytes not masked'],
  ['common.js', 'if (d > new Date(Date.UTC(y, mo, 0)).getUTCDate()) return null;', '', 'impossible calendar dates accepted as timestamps'],
  ['receiver.js', 'if (!matched) return webhookReject_', 'if (false) return webhookReject_', 'signature never checked'],
  ['receiver.js', 'if (age > toleranceMs)', 'if (false)', 'expired timestamps accepted'],
  ['receiver.js', 'if (-age > toleranceMs)', 'if (false)', 'future timestamps accepted'],
  ['receiver.js', "if (cache.get(cacheKey)) return 'duplicate';", "if (false) return 'duplicate';", 'cache fast path removed'],
  ['receiver.js', 'if (String(ids[i][0]) === eventId)', 'if (false)', 'sheet duplicate check removed'],
  ['receiver.js', "hmacSha256Hex_(timestamp + '.' + body, secret)", 'hmacSha256Hex_(body, secret)', 'timestamp left out of the signed string'],
  ['sync.js', 'code === 408 || code === 429 || code >= 500', 'code === 408 || code === 429', '5xx not retried'],
  ['sync.js', 'if (!transient) return', 'if (false) return', '4xx retried as if transient'],
  ['sync.js', "current === 'synced' || current === 'rejected'", "current === 'synced'", 'rejected rows re-sent'],
  ['sync.js', 'if (new Date().getTime() >= deadline)', 'if (false)', 'time budget ignored'],
  ['digest.js', "getProp_('DRY_RUN', 'true')", "getProp_('DRY_RUN', 'false')", 'dry run no longer the default'],
  ['digest.js', 'if (recipients.length > cfg.maxRecipients)', 'if (false)', 'recipient cap removed'],
  ['digest.js', 'writeTextBlock_(sheet, 2, markerCol + 1, marks);', '', 'rows never marked as digested'],
  ['forms.js', 'if (String(ids[i][0]) === responseKey)', 'if (false)', 'duplicate form responses appended'],
  ['forms.js', 'Math.max(lastRow, 1) + 1, 1, [row]', '1, 1, [row]', 'new row written over existing rows'],
];

function runSuite(srcDir) {
  const result = spawnSync(process.execPath, ['--test', ...fs.readdirSync(path.join(ROOT, 'test')).filter((f) => f.endsWith('.test.js')).map((f) => path.join('test', f))], {
    cwd: ROOT,
    env: { ...process.env, KIT_SRC_DIR: srcDir },
    encoding: 'utf8',
  });
  const m = /# fail (\d+)/.exec(result.stdout) || /ℹ fail (\d+)/.exec(result.stdout);
  return { failed: m ? Number(m[1]) : -1, status: result.status };
}

const temp = fs.mkdtempSync(path.join(os.tmpdir(), 'apps-script-kit-mut-'));
let survived = 0;
try {
  const baseline = runSuite(SRC);
  if (baseline.status !== 0) {
    console.error('The unmodified suite is not green; fix that first.');
    process.exit(2);
  }
  MUTATIONS.forEach(([file, find, replace, label]) => {
    fs.rmSync(temp, { recursive: true, force: true });
    fs.cpSync(SRC, temp, { recursive: true });
    const target = path.join(temp, file);
    const code = fs.readFileSync(target, 'utf8');
    if (!code.includes(find)) {
      console.log(`STALE     ${file}: ${label} (text to break not found; update the mutation list)`);
      survived++;
      return;
    }
    fs.writeFileSync(target, code.replace(find, replace));
    const { failed } = runSuite(temp);
    const killed = failed > 0;
    if (!killed) survived++;
    console.log(`${killed ? 'caught   ' : 'SURVIVED '} ${file}: ${label}${killed ? ` (${failed} failing)` : ''}`);
  });
} finally {
  fs.rmSync(temp, { recursive: true, force: true });
}
console.log(`\n${MUTATIONS.length - survived}/${MUTATIONS.length} deliberate breakages caught by the suite`);
process.exit(survived === 0 ? 0 : 1);
