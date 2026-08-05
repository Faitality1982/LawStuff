-- Zimmer Chiropractic & Integrative Health — Stem Wave interest survey
--
-- DESIGN CONSTRAINT (read before altering):
-- There is exactly one table and it holds NO identifying information. No name,
-- no contact details, no chart number, no IP address, no user-agent string.
-- Pain sites are recorded as broad checkboxes, not diagnoses.
--
-- This is deliberate and it is the reason the survey can run on ordinary
-- non-BAA hosting: anonymous health data is not PHI. An earlier revision had a
-- second `leads` table for optional callback requests; it was removed rather
-- than disabled, because a live endpoint accepting names is not made safe by a
-- front-end flag that hides the form.
--
-- Do not add a column that identifies a respondent.

DROP TABLE IF EXISTS responses;
CREATE TABLE responses (
  id          TEXT PRIMARY KEY,           -- uuid v4, generated server-side
  created_at  TEXT NOT NULL,              -- ISO8601 UTC
  src         TEXT,                       -- ?src= placement tag: counter, room1, ...
  duration_ms INTEGER,                    -- time from first paint to submit
  completed   INTEGER NOT NULL DEFAULT 0, -- 1 if the respondent reached the end
  vw_valid    INTEGER NOT NULL DEFAULT 1, -- 0 if Van Westendorp answers were non-monotonic
  path        TEXT,                       -- 'full' or 'short' (no-current-pain branch)
  payload     TEXT NOT NULL               -- JSON object, all answers keyed by question id
);

CREATE INDEX idx_responses_created ON responses (created_at);
CREATE INDEX idx_responses_src     ON responses (src);
