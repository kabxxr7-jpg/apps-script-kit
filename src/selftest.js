/**
 * Two checks to run ONCE in the real Apps Script editor (Run > selfTestHmac, Run > selfTestTextCells).
 * The local test suite uses hand-written fakes, so these are how you confirm the two platform
 * behaviours the kit depends on actually hold in Google's runtime. Both log a PASS/FAIL line.
 */

function hmacVectors_() {
  // Expected values were produced by webhook-kit's signPayload (HMAC-SHA256 over `${timestamp}.${body}`).
  return [
    { name: 'json body', secret: 'whsec_test_secret', ts: '2026-06-01T12:00:00.000Z', body: '{"id":"evt_1","type":"order.created"}', expected: 'b756eaf18c51802ee644074734a5a93743b6062da7e1c70f0c372f7570456ada' },
    { name: 'non-ASCII body', secret: 'whsec_test_secret', ts: '2026-06-01T12:00:00.000Z', body: '{"id":"evt_2","note":"caf\u00e9 \u2615 na\u00efve"}', expected: 'a45f29e16fd2af94218f232d76087e5362d3ac2087a9973afed76f6128c6f3b8' },
    { name: 'empty body', secret: 'k', ts: '2026-06-01T12:00:00.000Z', body: '', expected: '6d26b252feca52059e438a6b6f1800168bda5d1ab8fabcc460f3f4348e19322c' }
  ];
}

/** True when this runtime reproduces webhook-kit's signatures byte for byte. */
function selfTestHmac() {
  var results = hmacVectors_().map(function (v) {
    var actual = hmacSha256Hex_(v.ts + '.' + v.body, v.secret);
    return { name: v.name, pass: actual === v.expected, actual: actual };
  });
  var pass = results.every(function (r) { return r.pass; });
  console.log((pass ? 'PASS' : 'FAIL') + ' selfTestHmac ' + JSON.stringify(results));
  return pass;
}

/**
 * Creates a temporary sheet, writes tricky strings into '@' (plain text) cells, reads them back, and
 * deletes the sheet. Also reports (informational) what an unformatted cell does to "007".
 */
function selfTestTextCells() {
  var ss = getSpreadsheet_();
  var sheet = ss.insertSheet('_selftest_' + new Date().getTime());
  try {
    var samples = ['007', '=1+1', '2026-01-05', '+15551234567', 'TRUE'];
    var textRange = sheet.getRange(1, 1, 1, samples.length);
    textRange.setNumberFormat('@');
    textRange.setValues([samples]);
    var back = sheet.getRange(1, 1, 1, samples.length).getValues()[0];
    var plain = sheet.getRange(2, 1, 1, 1);
    plain.setValues([['007']]);
    var plainBack = plain.getValues()[0][0];
    var pass = back.every(function (v, i) { return v === samples[i]; });
    console.log((pass ? 'PASS' : 'FAIL') + ' selfTestTextCells textFormatBack=' + JSON.stringify(back) +
      ' (info) unformatted "007" reads back as ' + typeof plainBack + ' ' + JSON.stringify(plainBack));
    return pass;
  } finally {
    ss.deleteSheet(sheet);
  }
}
