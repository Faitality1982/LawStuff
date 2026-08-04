-- Zimmer Chiropractic — Stem Wave interest survey
--
-- DESIGN CONSTRAINT (read before altering):
-- `responses` is anonymous and `leads` is identifiable. There is deliberately
-- NO key, no session id, and no shared high-resolution timestamp linking them.
-- That separation is the reason this can run on non-BAA hosting. Do not add a
-- foreign key, do not store a lead id on a response, and do not upgrade
-- leads.created_day to a full timestamp.

DROP TABLE IF EXISTS responses;
CREATE TABLE responses (
  id          TEXT PRIMARY KEY,           -- uuid v4, generated server-side
  created_at  TEXT NOT NULL,              -- ISO8601 UTC
  src         TEXT,                       -- ?src= placement tag: counter, room1, email, ...
  duration_ms INTEGER,                    -- time from first paint to submit
  completed   INTEGER NOT NULL DEFAULT 0, -- 1 if the respondent reached the end screen
  vw_valid    INTEGER NOT NULL DEFAULT 1, -- 0 if Van Westendorp answers were non-monotonic
  path        TEXT,                        -- 'full' or 'short' (no-current-pain branch)
  payload     TEXT NOT NULL               -- JSON object, all answers keyed by question id
);

CREATE INDEX idx_responses_created ON responses (created_at);
CREATE INDEX idx_responses_src     ON responses (src);

DROP TABLE IF EXISTS leads;
CREATE TABLE leads (
  id          TEXT PRIMARY KEY,
  created_day TEXT NOT NULL,   -- YYYY-MM-DD ONLY. Deliberately coarse; see note above.
  name        TEXT NOT NULL,
  contact     TEXT NOT NULL,
  best_time   TEXT,
  src         TEXT
);

CREATE INDEX idx_leads_day ON leads (created_day);
