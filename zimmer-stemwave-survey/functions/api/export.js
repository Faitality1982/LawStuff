/*
 * GET /api/export?key=…&table=responses|leads&format=csv|json
 *
 * Gated by the EXPORT_KEY secret:
 *   npx wrangler pages secret put EXPORT_KEY
 *
 * If EXPORT_KEY is unset the endpoint refuses everything rather than defaulting
 * open — an unset secret is a deploy mistake, not permission to serve the data.
 */

const MAX_ROWS = 20000;

// Constant-time-ish comparison. A plain === leaks key length and prefix through
// response timing, and this endpoint hands out the whole dataset.
function safeEqual(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string') return false;
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

function csvCell(v) {
  if (v === null || v === undefined) return '';
  const s = Array.isArray(v) ? v.join('|') : String(v);
  return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

function toCsv(rows, columns) {
  const out = [columns.map(csvCell).join(',')];
  for (const r of rows) out.push(columns.map((c) => csvCell(r[c])).join(','));
  // BOM so Excel opens UTF-8 correctly — Dr. Zimmer will open this in Excel.
  return '﻿' + out.join('\r\n') + '\r\n';
}

// Flatten one response row: meta columns, then every answer key. Nested
// currencyGroup objects (vw, e4) expand to one column per sub-field; arrays
// from multi-selects collapse to pipe-joined strings.
function flattenResponse(row) {
  const flat = {
    id: row.id,
    created_at: row.created_at,
    src: row.src,
    duration_ms: row.duration_ms,
    duration_s: row.duration_ms == null ? null : Math.round(row.duration_ms / 1000),
    completed: row.completed,
    vw_valid: row.vw_valid,
    path: row.path,
  };

  let parsed;
  try { parsed = JSON.parse(row.payload); } catch { parsed = {}; }
  const answers = parsed.answers || {};

  for (const [k, v] of Object.entries(answers)) {
    if (v && typeof v === 'object' && !Array.isArray(v)) {
      for (const [sk, sv] of Object.entries(v)) flat[sk] = sv;
    } else {
      flat[k] = v;
    }
  }

  const meta = parsed.meta || {};
  flat.ua_mobile = meta.ua_mobile ?? null;
  flat.screen_w = meta.screen_w ?? null;
  return flat;
}

export async function onRequestGet({ request, env }) {
  const url = new URL(request.url);

  const expected = env.EXPORT_KEY;
  if (!expected) {
    return new Response('Export is not configured. Set the EXPORT_KEY secret.', { status: 503 });
  }
  if (!safeEqual(url.searchParams.get('key') || '', expected)) {
    return new Response('Not found', { status: 404 });
  }
  if (!env.DB) return new Response('database not bound', { status: 500 });

  const table = url.searchParams.get('table') === 'leads' ? 'leads' : 'responses';
  const format = url.searchParams.get('format') === 'json' ? 'json' : 'csv';

  const sql = table === 'leads'
    ? `SELECT id, created_day, name, contact, best_time, src FROM leads
       ORDER BY created_day DESC LIMIT ${MAX_ROWS}`
    : `SELECT id, created_at, src, duration_ms, completed, vw_valid, path, payload
       FROM responses ORDER BY created_at DESC LIMIT ${MAX_ROWS}`;

  let rows;
  try {
    const res = await env.DB.prepare(sql).all();
    rows = res.results || [];
  } catch (e) {
    return new Response('query failed: ' + String(e).slice(0, 200), { status: 500 });
  }

  const flat = table === 'leads' ? rows : rows.map(flattenResponse);

  if (format === 'json') {
    return new Response(JSON.stringify(flat, null, 2), {
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
    });
  }

  // Union of keys across all rows, so a question added mid-run still exports.
  // Meta columns are pinned to the front in a fixed order for readability.
  const lead = table === 'leads'
    ? ['id', 'created_day', 'name', 'contact', 'best_time', 'src']
    : ['id', 'created_at', 'src', 'duration_ms', 'duration_s', 'completed', 'vw_valid', 'path'];

  const seen = new Set(lead);
  const columns = lead.slice();
  for (const r of flat) {
    for (const k of Object.keys(r)) {
      if (!seen.has(k)) { seen.add(k); columns.push(k); }
    }
  }

  const stamp = new Date().toISOString().slice(0, 10);
  return new Response(toCsv(flat, columns), {
    headers: {
      'Content-Type': 'text/csv; charset=utf-8',
      'Content-Disposition': `attachment; filename="stemwave-${table}-${stamp}.csv"`,
      'Cache-Control': 'no-store',
      'X-Robots-Tag': 'noindex',
    },
  });
}
