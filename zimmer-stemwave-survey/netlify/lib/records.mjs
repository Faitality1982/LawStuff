/*
 * Shared record handling for the survey functions.
 *
 * Lives outside netlify/functions/ on purpose — anything inside that directory
 * is treated as a deployable function.
 */

import { getStore } from '@netlify/blobs';

export const STORE_NAME = 'stemwave-responses';

/*
 * Netlify Blobs is a key-value store, not SQL. Each response is one JSON blob.
 *
 * Keys are `<ISO timestamp>__<uuid>`:
 *   - blob listings come back lexicographically, and an ISO timestamp sorts
 *     chronologically as a string, so the export comes out in order for free
 *   - the uuid makes two submissions in the same millisecond impossible to
 *     collide
 *   - the separator is a double underscore because ISO timestamps already
 *     contain hyphens and colons
 */
export const makeKey = (iso, id) => `${iso}__${id}`;

/*
 * Test seam. `getStore` talks to Netlify's blob service and cannot run offline,
 * so tools/test_api.mjs swaps in an in-memory store to exercise the handlers
 * for real. Production never calls the setter, so the override stays null.
 */
let _storeOverride = null;
export const __setStoreForTests = (fn) => { _storeOverride = fn; };

export const store = (consistency = 'eventual') =>
  _storeOverride ? _storeOverride(consistency) : getStore({ name: STORE_NAME, consistency });

// ------------------------------------------------------------------ CSV

export function csvCell(v) {
  if (v === null || v === undefined) return '';
  const s = Array.isArray(v) ? v.join('|') : String(v);
  return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

export function toCsv(rows, columns) {
  const out = [columns.map(csvCell).join(',')];
  for (const r of rows) out.push(columns.map((c) => csvCell(r[c])).join(','));
  // Leading BOM so Excel opens UTF-8 correctly — this will be opened in Excel.
  return '﻿' + out.join('\r\n') + '\r\n';
}

/*
 * Flatten one stored record into a single CSV row. Nested currencyGroup answers
 * (vw, e4) expand to one column per sub-field; multi-select arrays collapse to
 * pipe-joined strings.
 */
export function flatten(rec) {
  const flat = {
    id: rec.id,
    created_at: rec.created_at,
    src: rec.src,
    duration_ms: rec.duration_ms,
    duration_s: rec.duration_ms == null ? null : Math.round(rec.duration_ms / 1000),
    completed: rec.completed,
    vw_valid: rec.vw_valid,
    path: rec.path,
  };

  const answers = rec.answers || {};
  for (const [k, v] of Object.entries(answers)) {
    if (v && typeof v === 'object' && !Array.isArray(v)) {
      for (const [sk, sv] of Object.entries(v)) flat[sk] = sv;
    } else {
      flat[k] = v;
    }
  }

  const meta = rec.meta || {};
  flat.ua_mobile = meta.ua_mobile ?? null;
  flat.screen_w = meta.screen_w ?? null;
  return flat;
}

export const PINNED_COLUMNS = [
  'id', 'created_at', 'src', 'duration_ms', 'duration_s',
  'completed', 'vw_valid', 'path',
];

// Union of keys across all rows, so a question added mid-run still exports.
export function columnsFor(rows) {
  const seen = new Set(PINNED_COLUMNS);
  const columns = PINNED_COLUMNS.slice();
  for (const r of rows) {
    for (const k of Object.keys(r)) {
      if (!seen.has(k)) { seen.add(k); columns.push(k); }
    }
  }
  return columns;
}

export const json = (obj, status = 200) =>
  new Response(JSON.stringify(obj), {
    status,
    headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
  });
