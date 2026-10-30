#!/usr/bin/env node
'use strict';
/**
 * Signs a request body for the Apps Script receiver (src/receiver.js) and prints the query string.
 * The scheme is webhook-kit's: HMAC-SHA256 hex over `${timestamp}.${body}`.
 *
 *   WEBHOOK_SECRET=... node tools/sign.js --body '{"id":"evt_1"}' [--id evt_1] [--ts 2026-06-01T12:00:00.000Z]
 *   WEBHOOK_SECRET=... node tools/sign.js --body-file payload.json --url https://script.google.com/macros/s/.../exec
 *
 * The secret is read from the environment, never from an argument, so it does not land in shell history.
 */
const crypto = require('node:crypto');
const fs = require('node:fs');

function parseArgs(argv) {
  const out = {};
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (!a.startsWith('--')) throw new Error(`unexpected argument: ${a}`);
    const key = a.slice(2);
    if (i + 1 >= argv.length) throw new Error(`--${key} needs a value`);
    out[key] = argv[++i];
  }
  return out;
}

function sign(secret, timestamp, body) {
  return crypto.createHmac('sha256', secret).update(timestamp).update('.').update(body).digest('hex');
}

function build(args, secret, now) {
  const body = args['body-file'] !== undefined ? fs.readFileSync(args['body-file'], 'utf8') : args.body;
  if (body === undefined) throw new Error('give --body or --body-file');
  if (!secret) throw new Error('set the WEBHOOK_SECRET environment variable');
  const ts = args.ts || new Date(now).toISOString();
  const query = new URLSearchParams({ ts, sig: sign(secret, ts, body) });
  if (args.id) query.set('id', args.id);
  const url = args.url ? `${args.url}${args.url.includes('?') ? '&' : '?'}${query}` : null;
  return { ts, sig: query.get('sig'), query: query.toString(), url, body };
}

/** POSIX single-quote escaping, so the exact signed bytes reach curl (a heredoc would add a trailing newline). */
function shellQuote(text) {
  return `'${text.replace(/'/g, `'\\''`)}'`;
}

module.exports = { build, sign, parseArgs, shellQuote };

if (require.main === module) {
  try {
    const result = build(parseArgs(process.argv.slice(2)), process.env.WEBHOOK_SECRET, Date.now());
    console.log(`query: ${result.query}`);
    if (result.url) {
      console.log(`printf '%s' ${shellQuote(result.body)} | curl -L -H 'Content-Type: application/json' --data-binary @- ${shellQuote(result.url)}`);
    }
  } catch (err) {
    console.error(`error: ${err.message}`);
    process.exit(1);
  }
}
