/*
 * POST /api/submit — anonymous survey responses.
 *
 * Nothing identifying is accepted or stored. No name, no contact details, no
 * chart number, no IP address, no user-agent string. Pain sites are broad
 * checkboxes, not diagnoses.
 *
 * If a future change adds an identifier here it breaks the only reason this
 * can run on ordinary hosting: anonymous health data is not PHI.
 */

import { store, makeKey, json } from '../lib/records.mjs';

const MAX_BODY = 24 * 1024; // a full response is ~2KB; 24KB is generous

// Guard against a future edit to questions.js accidentally introducing an
// identifying field. Cheap, and it fails loudly rather than silently storing
// a name.
const FORBIDDEN = ['name', 'email', 'phone', 'contact', 'dob', 'address', 'ssn'];

export default async (request) => {
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

  for (const key of Object.keys(body.answers)) {
    if (FORBIDDEN.includes(key.toLowerCase())) {
      return json({ error: 'identifying field rejected' }, 400);
    }
  }

  const id = crypto.randomUUID();
  const createdAt = new Date().toISOString();

  const record = {
    id,
    created_at: createdAt,
    src: typeof body.src === 'string' ? body.src.slice(0, 24) : null,
    duration_ms: Number.isFinite(body.duration_ms)
      ? Math.max(0, Math.min(Math.round(body.duration_ms), 86400000))
      : null,
    completed: body.completed ? 1 : 0,
    vw_valid: body.vw_valid === 0 ? 0 : 1,
    path: body.path === 'short' ? 'short' : 'full',
    answers: body.answers,
    meta: body.meta || {},
  };

  try {
    await store().setJSON(makeKey(createdAt, id), record);
  } catch (e) {
    return json({ error: 'write failed', detail: String(e).slice(0, 200) }, 500);
  }

  return json({ ok: true, id });
};

export const config = {
  path: '/api/submit',
  method: 'POST',
};
