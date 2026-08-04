/*
 * Question bank.
 *
 * This is DATA, not markup. Dr. Zimmer's edits should never require touching
 * survey.js. Reorder, reword, add, or delete entries here.
 *
 * Screen shape:
 *   id        unique key; becomes the JSON payload key
 *   type      info | single | multi | scale | currencyGroup | text | lead | end
 *   title     the question itself
 *   help      optional sub-line
 *   options   [{value, label}] for single/multi
 *   max       multi only: max selections
 *   required  default true (info/end screens are never required)
 *   showIf    fn(answers) -> bool; screen is skipped when false
 *   section   grouping label, used for the progress readout
 *
 * ORDERING RULE THAT MUST NOT BE BROKEN:
 * The Van Westendorp screen (vw) has to come BEFORE the price reveal (reveal).
 * Showing $2,809 first anchors every dollar answer to $2,809 and the pricing
 * data becomes worthless. If you reorder anything, keep vw < reveal.
 */
(function () {
  var C = window.SURVEY_CONFIG;
  var P = C.price;

  var money = function (n) { return '$' + n.toLocaleString('en-US'); };

  // Respondent said they have no current or recent pain. They still get the
  // concept question (useful for future positioning) but the pricing block
  // would be noise, so it is skipped.
  var hasPain = function (a) { return a.a1 && a.a1 !== 'none'; };

  window.SURVEY_QUESTIONS = [
    // ---------------------------------------------------------------- intro
    {
      id: 'intro',
      type: 'info',
      section: 'Welcome',
      title: 'Help us decide whether to bring a new therapy to the practice.',
      body: [
        'We’re considering adding a non-invasive treatment for chronic pain and injury, and we want to know what our patients actually think before we commit.',
        'This survey is <strong>completely anonymous</strong>. We don’t ask for your name, and your answers are not added to your chart.',
        'It takes about 3 minutes. If you get interrupted, you can close this and come back on the same phone — your answers are saved.',
      ],
      cta: 'Start',
    },

    // ------------------------------------------------------------ section A
    {
      id: 'a1',
      type: 'single',
      section: 'About you',
      title: 'Are you currently dealing with pain, stiffness, or an injury that limits what you’d like to do?',
      options: [
        { value: 'most_days', label: 'Yes, most days' },
        { value: 'on_off', label: 'Yes, on and off' },
        { value: 'past_year', label: 'Not right now, but I have in the past year' },
        { value: 'none', label: 'No' },
      ],
    },
    {
      id: 'a2',
      type: 'multi',
      section: 'About you',
      title: 'Where?',
      help: 'Select all that apply.',
      showIf: hasPain,
      options: [
        { value: 'low_back', label: 'Low back' },
        { value: 'neck', label: 'Neck' },
        { value: 'shoulder', label: 'Shoulder' },
        { value: 'knee', label: 'Knee' },
        { value: 'hip', label: 'Hip' },
        { value: 'foot', label: 'Foot, heel, or plantar fascia' },
        { value: 'elbow', label: 'Elbow' },
        { value: 'hand_wrist', label: 'Hand or wrist' },
        { value: 'neuropathy', label: 'Numbness or burning in hands or feet' },
        { value: 'other', label: 'Somewhere else' },
      ],
    },
    {
      id: 'a3',
      type: 'single',
      section: 'About you',
      title: 'How long has it been going on?',
      showIf: hasPain,
      options: [
        { value: 'lt1m', label: 'Less than 1 month' },
        { value: '1_6m', label: '1 to 6 months' },
        { value: '6_12m', label: '6 to 12 months' },
        { value: '1_3y', label: '1 to 3 years' },
        { value: 'gt3y', label: 'More than 3 years' },
      ],
    },
    {
      id: 'a4',
      type: 'scale',
      section: 'About you',
      title: 'On a typical day, how much does it interfere with what you want to do?',
      showIf: hasPain,
      min: 0,
      max: 10,
      minLabel: 'Not at all',
      maxLabel: 'Completely',
    },

    // ------------------------------------------------------------ section B
    {
      id: 'b1',
      type: 'multi',
      section: 'What you’ve tried',
      title: 'What have you already tried for it?',
      help: 'Select all that apply.',
      showIf: hasPain,
      options: [
        { value: 'chiro', label: 'Chiropractic adjustments' },
        { value: 'pt', label: 'Physical therapy' },
        { value: 'massage', label: 'Massage' },
        { value: 'injections', label: 'Cortisone or other injections' },
        { value: 'rx', label: 'Prescription pain medication' },
        { value: 'otc', label: 'Over-the-counter pain relievers' },
        { value: 'surgery', label: 'Surgery' },
        { value: 'needling', label: 'Dry needling' },
        { value: 'laser', label: 'Laser or red light therapy' },
        { value: 'nothing', label: 'Nothing yet', exclusive: true },
      ],
    },
    {
      id: 'b2',
      type: 'single',
      section: 'What you’ve tried',
      title: 'How satisfied are you with the results so far?',
      showIf: function (a) {
        return hasPain(a) && !(a.b1 && a.b1.length === 1 && a.b1[0] === 'nothing');
      },
      options: [
        { value: 'very_sat', label: 'Very satisfied' },
        { value: 'somewhat_sat', label: 'Somewhat satisfied' },
        { value: 'neutral', label: 'Neutral' },
        { value: 'somewhat_dis', label: 'Somewhat dissatisfied' },
        { value: 'very_dis', label: 'Very dissatisfied' },
      ],
    },
    {
      id: 'b3',
      type: 'single',
      section: 'What you’ve tried',
      title: 'Has surgery been discussed as an option for this?',
      showIf: hasPain,
      options: [
        { value: 'had_it', label: 'I’ve already had surgery for it' },
        { value: 'recommended', label: 'It’s been recommended, I haven’t done it' },
        { value: 'mentioned', label: 'It’s been mentioned as a possibility' },
        { value: 'no', label: 'No' },
        { value: 'unsure', label: 'Not sure' },
      ],
    },

    // ------------------------------------------------- section C — concept
    {
      id: 'concept',
      type: 'info',
      section: 'The idea',
      title: 'Here’s what we’re considering.',
      // ---------------------------------------------------------------------
      // COPY REQUIRING SIGN-OFF. Deliberately conservative: describes the
      // mechanism and the logistics, promises nothing. Do not add
      // regenerative or stem-cell claims without Dr. Zimmer's explicit
      // approval -- that language draws FTC and state board attention.
      // ---------------------------------------------------------------------
      body: [
        'We’re considering adding <strong>Stem Wave therapy</strong> (also called soft wave or shockwave therapy).',
        'A handheld device sends acoustic pressure waves into the injured area. It’s non-invasive — no needles, no incisions, no anesthesia, no medication. Each session takes about 10 to 15 minutes and there’s typically no downtime, so you can go straight back to work.',
        'A full course is <strong>' + P.sessions + ' sessions over about ' + P.weeks + ' weeks</strong> — twice a week for two weeks, then once a week for four weeks, with a re-evaluation a month after your eighth visit. It’s used for muscle, tendon, ligament, and joint pain.',
        '<em>Results vary from person to person, and no treatment works for everyone.</em>',
      ],
      cta: 'Got it',
    },
    {
      id: 'c1',
      type: 'single',
      section: 'The idea',
      title: 'Had you heard of this type of therapy before today?',
      options: [
        { value: 'had_elsewhere', label: 'I’ve had it done elsewhere' },
        { value: 'heard', label: 'I’ve heard of it, never tried it' },
        { value: 'new', label: 'No, this is new to me' },
      ],
    },
    {
      id: 'c2',
      type: 'single',
      section: 'The idea',
      title: 'How appealing does this sound for your situation?',
      options: [
        { value: 'very', label: 'Very appealing' },
        { value: 'somewhat', label: 'Somewhat appealing' },
        { value: 'neutral', label: 'Neutral' },
        { value: 'not_very', label: 'Not very appealing' },
        { value: 'not_at_all', label: 'Not at all appealing' },
      ],
    },
    {
      id: 'c4',
      type: 'multi',
      section: 'The idea',
      title: 'What’s your biggest hesitation?',
      help: 'Choose up to 2.',
      max: 2,
      options: [
        { value: 'cost', label: 'Cost' },
        { value: 'insurance', label: 'Not covered by insurance' },
        { value: 'efficacy', label: 'Might not work for me' },
        { value: 'time', label: 'Time commitment (' + P.sessions + ' visits)' },
        { value: 'pain', label: 'Might be painful' },
        { value: 'understanding', label: 'I don’t understand how it works' },
        { value: 'evidence', label: 'I’d want to see more evidence' },
        { value: 'none', label: 'No hesitations', exclusive: true },
      ],
    },

    // ------------------------------- section D — pricing, NO price shown yet
    {
      id: 'vw',
      type: 'currencyGroup',
      section: 'What it’s worth',
      showIf: hasPain,
      title: 'What would you expect something like this to cost?',
      help: 'Thinking about a <strong>complete course of care — ' + P.sessions + ' sessions over about ' +
            P.weeks + ' weeks</strong>, and knowing it would <strong>not be covered by insurance</strong>. ' +
            'There are no wrong answers, we genuinely don’t know yet.',
      fields: [
        { id: 'vw_cheap',     label: 'So inexpensive you’d question whether it works' },
        { id: 'vw_bargain',   label: 'A good value — a price you’d be pleased to pay' },
        { id: 'vw_expensive', label: 'Starting to feel expensive, but you’d still consider it' },
        { id: 'vw_tooexp',    label: 'So expensive you wouldn’t consider it at all' },
      ],
      // Enforced ascending. Violations warn once, then allow through with the
      // record flagged vw_valid = 0 so analysis can exclude it.
      monotonic: true,
    },

    // --------------------------------- section E — reveal, then intent
    {
      id: 'reveal',
      type: 'info',
      section: 'What we’re considering charging',
      showIf: hasPain,
      title: 'Here’s the pricing we’re considering.',
      body: [
        '<strong>Discovery Visit — ' + money(P.discovery) + '</strong><br>A consultation plus one treatment, so you can try it before committing to anything.',
        '<strong>Full course — ' + money(P.packageOnly) + '</strong><br>' + P.sessions +
          ' soft wave treatments plus ' + P.sessions + ' PEMF treatments over ' + P.weeks + ' weeks.',
        '<strong>Total — ' + money(P.total) + '</strong>',
        '<em>This would not be covered by insurance.</em>',
      ],
      cta: 'Continue',
    },
    {
      id: 'e1',
      type: 'scale',
      section: 'What we’re considering charging',
      showIf: hasPain,
      title: 'If this were available next month, how likely would you be to book the ' + money(P.discovery) + ' Discovery Visit?',
      min: 0,
      max: 10,
      minLabel: 'No chance',
      midLabel: 'About even',
      maxLabel: 'Certain',
    },
    {
      id: 'e2',
      type: 'scale',
      section: 'What we’re considering charging',
      showIf: hasPain,
      title: 'And if the Discovery Visit went well, how likely would you be to purchase the full ' + money(P.packageOnly) + ' course?',
      min: 0,
      max: 10,
      minLabel: 'No chance',
      midLabel: 'About even',
      maxLabel: 'Certain',
    },
    {
      id: 'e3',
      type: 'single',
      section: 'What we’re considering charging',
      showIf: hasPain,
      title: 'Which way of paying would be most attractive to you?',
      options: (function () {
        var opts = [
          { value: 'full', label: 'Pay in full up front (' + P.prepayDiscountPct + '% discount)' },
          { value: 'two', label: 'Two payments — first visit and fifth visit' },
          { value: 'per_visit', label: 'Pay for each treatment as you go' },
        ];
        if (C.enableMonthlyPlanQuestion) {
          opts.push({ value: 'monthly', label: 'Monthly payments over 6 to 12 months' });
        }
        opts.push({ value: 'none', label: 'None of these — the price is the problem, not the structure' });
        return opts;
      })(),
    },
    {
      id: 'e4',
      type: 'currencyGroup',
      section: 'What we’re considering charging',
      showIf: function (a) { return hasPain(a) && a.e3 === 'monthly'; },
      title: 'What monthly payment would feel manageable?',
      fields: [{ id: 'e4_monthly', label: 'Per month' }],
    },
    {
      id: 'e5',
      type: 'single',
      section: 'What we’re considering charging',
      showIf: hasPain,
      title: 'Would it change things if you could pay with an HSA or FSA?',
      options: [
        { value: 'significant', label: 'Yes, significantly' },
        { value: 'somewhat', label: 'Yes, somewhat' },
        { value: 'no_diff', label: 'No difference' },
        { value: 'no_account', label: 'I don’t have one' },
      ],
    },

    // ------------------------------------------------------------ section F
    {
      id: 'f3',
      type: 'single',
      section: 'Last thing',
      title: 'Are you currently a patient at ' + C.practiceName + '?',
      options: [
        { value: 'current', label: 'Yes, currently' },
        { value: 'former', label: 'I have been in the past' },
        { value: 'no', label: 'No' },
      ],
    },
    {
      id: 'f4',
      type: 'single',
      section: 'Last thing',
      title: 'Your age range',
      help: 'Optional. It helps us understand who this would serve.',
      required: false,
      options: [
        { value: '18_34', label: '18 to 34' },
        { value: '35_49', label: '35 to 49' },
        { value: '50_64', label: '50 to 64' },
        { value: '65p', label: '65 or older' },
        { value: 'decline', label: 'Prefer not to say' },
      ],
    },
    {
      id: 'f2',
      type: 'text',
      section: 'Last thing',
      title: 'Anything else Dr. Zimmer should know before deciding whether to offer this?',
      help: 'Optional.',
      required: false,
      maxLength: 400,
    },
  ];
})();
