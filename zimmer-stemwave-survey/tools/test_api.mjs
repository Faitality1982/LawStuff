/*
 * Executes the Netlify Functions against an in-memory blob store.
 *
 *     npm test
 *
 * `node --check` only parses; it will happily pass a file that throws
 * ReferenceError on an undefined variable in a template literal. This actually
 * runs the handlers and asserts on what comes back.
 */

import { __setStoreForTests, makeKey } from '../netlify/lib/records.mjs';
import submit from '../netlify/functions/submit.mjs';
import exportFn from '../netlify/functions/export.mjs';

let failed = 0;
const check = (label, cond, detail = '') => {
  console.log((cond ? '  PASS  ' : '  FAIL  ') + label + (!cond && detail ? '  -- ' + detail : ''));
  if (!cond) failed++;
};

// ------------------------------------------------------- in-memory blob store
const blobs = new Map();
let failNextWrite = false;

__setStoreForTests(() => ({
  async setJSON(key, value) {
    if (failNextWrite) { failNextWrite = false; throw new Error('simulated outage'); }
    blobs.set(key, JSON.parse(JSON.stringify(value)));
  },
  async get(key, opts) {
    const v = blobs.get(key);
    if (v === undefined) return null;
    return opts?.type === 'json' ? v : JSON.stringify(v);
  },
  async list() {
    // Netlify returns keys in arbitrary order; shuffle so the export can't
    // accidentally depend on insertion order.
    const keys = [...blobs.keys()].reverse();
    return { blobs: keys.map((key) => ({ key, etag: 'x' })), directories: [] };
  },
}));

const post = (body) => new Request('https://x/api/submit', {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: typeof body === 'string' ? body : JSON.stringify(body),
});

const FULL = {
  src: 'counter', duration_ms: 182000, completed: 1, vw_valid: 1, path: 'full',
  answers: {
    a1: 'most_days', a2: ['knee', 'low_back'], a4: 7,
    vw: { vw_cheap: 300, vw_bargain: 900, vw_expensive: 1800, vw_tooexp: 3200 },
    e1: 9, e3: 'monthly', e4: { e4_monthly: 250 }, f2: 'Cost, mainly. Also "time".',
  },
  meta: { ua_mobile: 1, screen_w: 390 },
};

// ------------------------------------------------------------------- submit
console.log('\n--- submit ---');

let res = await submit(post(FULL));
check('submit returns 200', res.status === 200, String(res.status));
const body = await res.json();
check('submit returns an id', typeof body.id === 'string' && body.id.length > 10);
check('record was stored', blobs.size === 1);

const [storedKey] = [...blobs.keys()];
check('key is ISO-timestamp prefixed (sorts chronologically)',
  /^\d{4}-\d{2}-\d{2}T[\d:.]+Z__[0-9a-f-]{36}$/.test(storedKey), storedKey);

const stored = blobs.get(storedKey);
check('no IP or user-agent stored',
  !('ip' in stored) && !('ua' in stored) && !JSON.stringify(stored).includes('Mozilla'));
check('answers stored verbatim', stored.answers.a2.join('|') === 'knee|low_back');

res = await submit(post({ answers: { name: 'Mary Smith' } }));
check('identifying key rejected', res.status === 400, String(res.status));
res = await submit(post({ answers: { EMAIL: 'x@y.z' } }));
check('identifying key rejected case-insensitively', res.status === 400, String(res.status));

res = await submit(post({ nope: true }));
check('missing answers rejected', res.status === 400, String(res.status));

res = await submit(post('not json'));
check('malformed json rejected', res.status === 400, String(res.status));

res = await submit(post({ answers: { a1: 'x' }, duration_ms: 'abc', completed: 'yes' }));
check('junk duration/completed coerced, not crashed', res.status === 200, String(res.status));

failNextWrite = true;
res = await submit(post({ answers: { a1: 'none' } }));
check('storage outage returns 500', res.status === 500, String(res.status));

// ------------------------------------------------------------------- export
console.log('\n--- export ---');

const get = (qs) => new Request('https://x/api/export' + qs);

process.env.EXPORT_KEY = '';
res = await exportFn(get('?key=anything'));
check('unset EXPORT_KEY refuses (does not default open)', res.status === 503, String(res.status));

process.env.EXPORT_KEY = 'test-key-0123456789';
res = await exportFn(get('?key=wrong-key-012345'));
check('wrong key returns 404', res.status === 404, String(res.status));
res = await exportFn(get(''));
check('missing key returns 404', res.status === 404, String(res.status));

res = await exportFn(get('?key=test-key-0123456789'));
check('valid key returns 200', res.status === 200, String(res.status));
check('record count surfaced in header',
  res.headers.get('X-Record-Count') === String(blobs.size),
  `${res.headers.get('X-Record-Count')} vs ${blobs.size}`);

// Read raw bytes, not text(). Response.text() performs a WHATWG "UTF-8 decode"
// which strips a leading BOM — so the BOM has to be checked on the wire bytes,
// which is what Excel actually sees.
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
check('duration_s derived',
  lines[0].split(',')[cols.indexOf('duration_s')] === '182');
check('currencyGroup flattened to sub-fields',
  ['vw_cheap', 'vw_bargain', 'vw_expensive', 'vw_tooexp', 'e4_monthly'].every((c) => cols.includes(c)),
  cols.join('|'));
check('multi-select pipe-joined', lines[0].includes('knee|low_back'), lines[0].slice(0, 100));
check('embedded quotes and commas escaped',
  lines[0].includes('"Cost, mainly. Also ""time""."'), lines[0].slice(-60));
check('no identifying column present',
  !['name', 'email', 'phone', 'contact', 'ip'].some((c) => cols.includes(c)), cols.join('|'));
check('one csv line per stored record', lines.length === blobs.size,
  `${lines.length} lines vs ${blobs.size} records`);

// Chronological ordering despite the store returning keys out of order.
const times = lines.map((l) => l.split(',')[cols.indexOf('created_at')]);
check('rows come out in chronological order',
  times.join('|') === [...times].sort().join('|'), times.join(' | ').slice(0, 90));

res = await exportFn(get('?key=test-key-0123456789&format=json'));
const asJson = await res.json();
check('json format works', Array.isArray(asJson) && asJson.length === blobs.size);

// Batching: more records than one batch (25) must all come through.
console.log('\n--- batching ---');
for (let i = 0; i < 60; i++) {
  await submit(post({ ...FULL, src: 'room1', answers: { ...FULL.answers, a4: i % 11 } }));
}
res = await exportFn(get('?key=test-key-0123456789&format=json'));
const all = await res.json();
check('all records returned across batches', all.length === blobs.size,
  `${all.length} vs ${blobs.size}`);
check('batched export stays ordered',
  all.map((r) => r.created_at).join('|') === all.map((r) => r.created_at).sort().join('|'));

// ---------------------------------------------------------------------- done
console.log('\n' + '='.repeat(56));
if (failed) { console.log(`${failed} FAILURE(S)`); process.exit(1); }
console.log('All API checks passed.');
