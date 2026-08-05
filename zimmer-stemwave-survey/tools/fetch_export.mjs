#!/usr/bin/env node
/*
 * Pull the responses CSV down to a local file.
 *
 *     EXPORT_KEY=... SITE_URL=https://zimmerstemwave.netlify.app npm run export
 *
 * Then:  python3 tools/analyze.py responses-YYYY-MM-DD.csv
 */

import { writeFile } from 'node:fs/promises';

const site = (process.env.SITE_URL || 'https://zimmerstemwave.netlify.app').replace(/\/+$/, '');
const key = process.env.EXPORT_KEY;

if (!key) {
  console.error('Set EXPORT_KEY (the same value configured in the Netlify UI).');
  process.exit(1);
}

const url = `${site}/api/export?key=${encodeURIComponent(key)}`;
const res = await fetch(url);

if (res.status === 404) {
  console.error('404 — wrong EXPORT_KEY, or the site URL is wrong.');
  process.exit(1);
}
if (res.status === 503) {
  console.error('503 — EXPORT_KEY is not set on the Netlify site itself.');
  process.exit(1);
}
if (!res.ok) {
  console.error(`${res.status} — ${(await res.text()).slice(0, 300)}`);
  process.exit(1);
}

const stamp = new Date().toISOString().slice(0, 10);
const out = `responses-${stamp}.csv`;
// Write bytes, not text: the response carries a UTF-8 BOM so Excel opens it
// correctly, and decoding to a string would drop it.
await writeFile(out, Buffer.from(await res.arrayBuffer()));

const count = res.headers.get('X-Record-Count');
console.log(`Wrote ${out}  (${count} responses)`);
if (res.headers.get('X-Truncated')) {
  console.warn('WARNING: the export was truncated — more records exist than were returned.');
}
if (Number(count) < 50) {
  console.log(`\nNote: ${count} responses. Van Westendorp needs ~50 to be readable,`);
  console.log('100+ to be solid. Below 30, analyze.py will tell you not to trust the curves.');
}
