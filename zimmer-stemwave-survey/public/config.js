/*
 * Survey configuration.
 *
 * Everything Dr. Zimmer is likely to want changed lives here or in questions.js.
 * Neither file requires touching survey.js.
 */
window.SURVEY_CONFIG = {
  // ---- Branding -----------------------------------------------------------
  practiceName: 'Zimmer Chiropractic',
  // Optional: drop a file in public/ and set the filename, e.g. 'logo.png'.
  // Leave null to show the practice name as text.
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
  // Optional contact-capture screen at the end. Posts to a SEPARATE endpoint
  // and a SEPARATE table with no linkage to survey answers. If Dr. Zimmer
  // would rather the front desk handle follow-up, set this to false.
  enableLeadCapture: true,

  // Tests appetite for a monthly payment plan that is NOT currently offered.
  // Highest-value question in the survey. Turn off only if he would never
  // offer financing under any circumstances.
  enableMonthlyPlanQuestion: true,

  // Cloudflare Turnstile. Leave null to disable (honeypot still applies).
  // Set to the site key from the Turnstile dashboard to enable.
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
