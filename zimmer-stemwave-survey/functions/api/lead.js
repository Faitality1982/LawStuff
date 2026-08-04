/*
 * POST /api/lead — optional callback request.
 *
 * This is the ONE identifiable table. It deliberately stores:
 *   - no survey answers
 *   - no response id
 *   - no clinical information
 *   - a DAY, not a timestamp
 *
 * The coarse date is not an oversight. A precise timestamp would let anyone
 * with both tables rejoin a lead to the survey submitted seconds earlier,
 * which reconstructs exactly the identifiable-health-data pairing this design
 * exists to avoid.
 */

const MAX_BODY = 4 * 1024;

const json = (obj, status = 200) =>
  new Response(JSON.stringify(obj), {
    status,
    headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
  });

// Strip control characters only. Spaces, hyphens, and apostrophes are real
// parts of real names and must survive.
const clean = (v, max) =>
  typeof v === 'string' ? v.replace(/[\x00-\x1F\x7F]/g, '').trim().slice(0, max) : '';

export async function onRequestPost({ request, env }) {
  if (!env.DB) return json({ error: 'database not bound' }, 500);

  let body;
  try {
    const text = await request.text();
    if (text.length > MAX_BODY) return json({ error: 'payload too large' }, 413);
    body = JSON.parse(text);
  } catch {
    return json({ error: 'invalid json' }, 400);
  }

  // Honeypot. Return 200 so a bot sees success and doesn't retry or adapt.
  if (body && clean(body.website, 100)) return json({ ok: true });

  const name = clean(body?.name, 120);
  const contact = clean(body?.contact, 160);
  if (!name || !contact) return json({ error: 'name and contact required' }, 400);

  const bestTime = ['morning', 'afternoon', 'evening'].includes(body?.best_time)
    ? body.best_time
    : null;
  const src = typeof body?.src === 'string' ? body.src.slice(0, 24) : null;

  try {
    await env.DB.prepare(
      `INSERT INTO leads (id, created_day, name, contact, best_time, src)
       VALUES (?, ?, ?, ?, ?, ?)`
    ).bind(
      crypto.randomUUID(),
      new Date().toISOString().slice(0, 10), // YYYY-MM-DD only — see header note
      name,
      contact,
      bestTime,
      src
    ).run();
  } catch (e) {
    return json({ error: 'insert failed', detail: String(e).slice(0, 200) }, 500);
  }

  return json({ ok: true });
}
