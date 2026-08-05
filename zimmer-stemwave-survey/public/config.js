/*
 * Survey configuration.
 *
 * Everything Dr. Zimmer is likely to want changed lives here or in questions.js.
 * Neither file requires touching survey.js.
 */
window.SURVEY_CONFIG = {
  // ---- Branding -----------------------------------------------------------
  practiceName: 'Zimmer Chiropractic',
  practiceTagline: 'Chiropractic & Integrative Health',
  // Optional: drop a file in public/ and set the filename, e.g. 'logo.png'.
  // Leave null to render the practice name as text.
  logoFile: null,

  // ---- Pricing (single source of truth; questions.js interpolates these) ---
  price: {
    discovery: 49,
    packageOnly: 2760,
    total: 2809,
    prepayDiscountPct: 5,
    sessions: 9,
    weeks: 6,
  },

  // ---- Feature flags ------------------------------------------------------
  // Tests appetite for a monthly payment plan that is NOT currently offered.
  // Highest-value question in the survey. Turn off only if Dr. Zimmer would
  // never offer financing under any circumstances.
  enableMonthlyPlanQuestion: true,

  // Cloudflare Turnstile. Leave null to disable (the honeypot still applies).
  // Set to the site key from the Turnstile dashboard to enable. Turnstile is
  // the one anti-abuse measure that costs nothing in anonymity — it identifies
  // browsers, not people, and stores nothing about the respondent.
  turnstileSiteKey: null,

  // ---- Behaviour ----------------------------------------------------------
  // Resume a partially completed survey from localStorage.
  enableResume: true,
  storageKey: 'zc_stemwave_v1',

  // Valid ?src= values. Anything else is recorded as 'unknown' so a mistyped
  // or guessed QR param can't pollute the placement analysis.
  allowedSrc: ['counter', 'room1', 'room2', 'room3', 'email', 'sms', 'poster', 'card', 'direct'],

  // Progress-counter projection.
  //
  // Branching means the real screen count isn't known until the branching
  // question is answered. Without this, the counter reads "2 of 9" and then
  // jumps to "of 21" the moment they say they're in pain — a survey that
  // visibly grows is one people abandon, and at a busy front desk that is
  // exactly the wrong signal.
  //
  // These are the answers to ASSUME while the question is still unanswered,
  // so the denominator starts at the longest path and can only shrink.
  progressAssumes: { a1: 'most_days' },
};
