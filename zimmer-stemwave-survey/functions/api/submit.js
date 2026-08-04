/*
 * POST /api/submit — anonymous survey responses.
 *
 * Nothing identifying is accepted or stored here. No IP, no user agent string,
 * no name, no contact. If a future change adds an identifier to this table it
 * breaks the separation that lets this run on non-BAA hosting — see schema.sql.
 */

const MAX_BODY = 24 * 1024; // a full response is ~2KB; 24KB is generous

const json = (obj, status = 200) =>
  new Response(JSON.stringify(obj), {
    status,
    headers: {
      'Content-Type': 'application/json',
      'Cache-Control': 'no-store',
    },
  });

export async function onRequestPost({ request, env }) {
  if (!env.DB) return json({ error: 'database not bound' }, 500);

  const len = Number(request.headers.get('content-length') || 0);
  if (len > MAX_BODY) return json({ error: 'payload too large' }, 413);

  let body;
  try {
    const text = await request.text();
    if (text.length > MAX_BODY) return json({ error: 'payload too large' }, 413);
    body = JSON.parse(text);
  } catch {
    return json({ error: 'invalid json' }, 400);
  }

  if (!body || typeof body !== 'object' || !body.answers || typeof body.answers !== 'object') {
    return json({ error: 'missing answers' }, 400);
  }

  // Reject anything that looks like it carries contact details. This is a
  // guard against a future edit to questions.js accidentally putting
  // identifying fields into the anonymous table.
  const forbidden = ['name', 'email', 'phone', 'contact', 'dob', 'address'];
  for (const key of Object.keys(body.answers)) {
    if (forbidden.includes(key.toLowerCase())) {
      return json({ error: 'identifying field rejected' }, 400);
    }
  }

  const id = crypto.randomUUID();
  const src = typeof body.src === 'string' ? body.src.slice(0, 24) : null;
  const duration = Number.isFinite(body.duration_ms)
    ? Math.max(0, Math.min(Math.round(body.duration_ms), 86400000))
    : null;
  const completed = body.completed ? 1 : 0;
  const vwValid = body.vw_valid === 0 ? 0 : 1;
  const path = body.path === 'short' ? 'short' : 'full';

  try {
    await env.DB.prepare(
      `INSERT INTO responses (id, created_at, src, duration_ms, completed, vw_valid, path, payload)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
    ).bind(
      id,
      new Date().toISOString(),
      src,
      duration,
      completed,
      vwValid,
      path,
      JSON.stringify({ answers: body.answers, meta: body.meta || {} })
    ).run();
  } catch (e) {
    return json({ error: 'insert failed', detail: String(e).slice(0, 200) }, 500);
  }

  return json({ ok: true, id });
}

// Only onRequestPost is exported on purpose. Pages answers other methods with
// a 405 by itself; exporting a catch-all onRequest here would take precedence
// over this handler and swallow the POSTs.
