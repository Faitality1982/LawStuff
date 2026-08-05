/*
 * GET /api/export?key=…&format=csv|json
 *
 * Gated by the EXPORT_KEY environment variable, set in the Netlify UI under
 * Site configuration → Environment variables.
 *
 * If EXPORT_KEY is unset the endpoint refuses everything rather than defaulting
 * open — an unset variable is a deploy mistake, not permission to serve the data.
 */

import { store, flatten, toCsv, columnsFor } from '../lib/records.mjs';

// Blobs is key-value, so an export is one read per response. These run
// concurrently in batches to stay well inside the 10-second synchronous
// function timeout — sequentially, a few hundred responses would blow it.
const BATCH = 25;
const MAX_RECORDS = 20000;

// Constant-time-ish comparison. A plain === leaks key length and prefix through
// response timing, and this endpoint hands out the whole dataset.
function safeEqual(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string') return false;
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

export default async (request) => {
  const url = new URL(request.url);

  const expected = process.env.EXPORT_KEY;
  if (!expected) {
    return new Response('Export is not configured. Set the EXPORT_KEY environment variable.',
      { status: 503 });
  }
  // A wrong key gets a plain 404 so the endpoint doesn't advertise itself.
  if (!safeEqual(url.searchParams.get('key') || '', expected)) {
    return new Response('Not found', { status: 404 });
  }

  const format = url.searchParams.get('format') === 'json' ? 'json' : 'csv';
  const s = store('strong'); // a response submitted seconds ago must appear

  let keys = [];
  try {
    const { blobs } = await s.list();
    keys = blobs.map((b) => b.key).sort(); // ISO-prefixed keys sort chronologically
  } catch (e) {
    return new Response('list failed: ' + String(e).slice(0, 200), { status: 500 });
  }

  let truncated = false;
  if (keys.length > MAX_RECORDS) { keys = keys.slice(-MAX_RECORDS); truncated = true; }

  const records = [];
  for (let i = 0; i < keys.length; i += BATCH) {
    const chunk = await Promise.all(
      keys.slice(i, i + BATCH).map((k) => s.get(k, { type: 'json' }).catch(() => null))
    );
    for (const r of chunk) if (r) records.push(r);
  }

  const rows = records.map(flatten);

  if (format === 'json') {
    return new Response(JSON.stringify(rows, null, 2), {
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
    });
  }

  const stamp = new Date().toISOString().slice(0, 10);
  return new Response(toCsv(rows, columnsFor(rows)), {
    headers: {
      'Content-Type': 'text/csv; charset=utf-8',
      'Content-Disposition': `attachment; filename="stemwave-responses-${stamp}.csv"`,
      'Cache-Control': 'no-store',
      'X-Robots-Tag': 'noindex',
      // Surfaced as a header rather than swallowed: a silently truncated export
      // reads as "that's all the responses" when it isn't.
      'X-Record-Count': String(rows.length),
      ...(truncated ? { 'X-Truncated': 'true' } : {}),
    },
  });
};

export const config = {
  path: '/api/export',
  method: 'GET',
};
