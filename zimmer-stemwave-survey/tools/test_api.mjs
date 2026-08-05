/*
 * Executes the Pages Functions against a stub D1 binding.
 *
 *     node tools/test_api.mjs
 *
 * `node --check` only parses; it will happily pass a file that throws
 * ReferenceError on an undefined variable in a template literal. This actually
 * runs the handlers and asserts on what comes back.
 */

import { onRequestPost } from '../functions/api/submit.js';
import { onRequestGet } from '../functions/api/export.js';

let failed = 0;
const check = (label, cond, detail = '') => {
  console.log((cond ? '  PASS  ' : '  FAIL  ') + label + (!cond && detail ? '  -- ' + detail : ''));
  if (!cond) failed++;
};

// ---------------------------------------------------------------- stub D1
function stubDB() {
  const rows = [];
  return {
    rows,
    prepare(sql) {
      return {
        _args: [],
        bind(...args) { this._args = args; return this; },
        async run() {
          if (!/INSERT INTO responses/.test(sql)) throw new Error('unexpected sql: ' + sql);
          const [id, created_at, src, duration_ms, completed, vw_valid, path, payload] = this._args;
          rows.push({ id, created_at, src, duration_ms, completed, vw_valid, path, payload });
          return { success: true };
        },
        async all() {
          if (!/FROM responses/.test(sql)) throw new Error('unexpected sql: ' + sql);
          return { results: rows };
        },
      };
    },
  };
}

const post = (body) => new Request('https://x/api/submit', {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify(body),
});

const DB = stubDB();
const env = { DB, EXPORT_KEY: 'test-key-0123456789' };

// ---------------------------------------------------------------- submit
console.log('\n--- submit ---');

let res = await onRequestPost({
  request: post({
    src: 'counter', duration_ms: 182000, completed: 1, vw_valid: 1, path: 'full',
    answers: {
      a1: 'most_days', a2: ['knee', 'low_back'], a4: 7,
      vw: { vw_cheap: 300, vw_bargain: 900, vw_expensive: 1800, vw_tooexp: 3200 },
      e1: 9, e3: 'monthly', e4: { e4_monthly: 250 }, f2: 'Cost, mainly. Also "time".',
    },
    meta: { ua_mobile: 1, screen_w: 390 },
  }),
  env,
});
check('submit returns 200', res.status === 200, String(res.status));
const body = await res.json();
check('submit returns an id', typeof body.id === 'string' && body.id.length > 10);
check('row was inserted', DB.rows.length === 1);

res = await onRequestPost({ request: post({ answers: { name: 'Mary Smith' } }), env });
check('identifying key rejected', res.status === 400, String(res.status));

res = await onRequestPost({ request: post({ nope: true }), env });
check('missing answers rejected', res.status === 400, String(res.status));

res = await onRequestPost({
  request: new Request('https://x/api/submit', { method: 'POST', body: 'not json' }),
  env,
});
check('malformed json rejected', res.status === 400, String(res.status));

res = await onRequestPost({ request: post({ answers: { a1: 'none' } }), env: { EXPORT_KEY: 'x' } });
check('missing DB binding returns 500', res.status === 500, String(res.status));

// ---------------------------------------------------------------- export
console.log('\n--- export ---');

const get = (qs) => new Request('https://x/api/export' + qs);

res = await onRequestGet({ request: get('?key=wrong'), env });
check('wrong key returns 404', res.status === 404, String(res.status));

res = await onRequestGet({ request: get(''), env });
check('missing key returns 404', res.status === 404, String(res.status));

res = await onRequestGet({ request: get('?key=test-key-0123456789'), env: { DB } });
check('unset EXPORT_KEY refuses (does not default open)', res.status === 503, String(res.status));

res = await onRequestGet({ request: get('?key=test-key-0123456789'), env });
check('valid key returns 200', res.status === 200, String(res.status));

// Read the raw bytes, not text(). Response.text() performs a WHATWG "UTF-8
// decode", which strips a leading BOM — so the BOM has to be checked on the
// wire bytes, which is what Excel actually sees.
const raw = new Uint8Array(await res.clone().arrayBuffer());
check('csv starts with a UTF-8 BOM for Excel',
  raw[0] === 0xef && raw[1] === 0xbb && raw[2] === 0xbf,
  [...raw.slice(0, 3)].map((b) => b.toString(16)).join(' '));

const csv = await res.text();
const [header, ...lines] = csv.replace(/^﻿/, '').trim().split('\r\n');
const cols = header.split(',');
check('filename header is well-formed',
  /filename="stemwave-responses-\d{4}-\d{2}-\d{2}\.csv"/.test(
    res.headers.get('Content-Disposition') || ''),
  res.headers.get('Content-Disposition'));
check('meta columns pinned first',
  cols.slice(0, 4).join(',') === 'id,created_at,src,duration_ms', cols.slice(0, 4).join(','));
check('duration_s derived', cols.includes('duration_s') &&
  lines[0].split(',')[cols.indexOf('duration_s')] === '182');
check('currencyGroup flattened to sub-fields',
  ['vw_cheap', 'vw_bargain', 'vw_expensive', 'vw_tooexp', 'e4_monthly'].every((c) => cols.includes(c)),
  cols.join('|'));
check('multi-select pipe-joined',
  lines[0].includes('knee|low_back'), lines[0].slice(0, 120));
check('embedded quotes and commas escaped',
  lines[0].includes('"Cost, mainly. Also ""time""."'), lines[0].slice(-60));
check('no identifying column present',
  !['name', 'email', 'phone', 'contact'].some((c) => cols.includes(c)), cols.join('|'));

res = await onRequestGet({ request: get('?key=test-key-0123456789&format=json'), env });
const json = await res.json();
check('json format works', Array.isArray(json) && json.length === DB.rows.length);

// ---------------------------------------------------------------- done
console.log('\n' + '='.repeat(56));
if (failed) { console.log(`${failed} FAILURE(S)`); process.exit(1); }
console.log('All API checks passed.');
