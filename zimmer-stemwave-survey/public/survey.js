/*
 * Survey engine.
 *
 * Plain ES5-ish JS on purpose — no build step, no framework, no dependencies.
 * Loads directly off Pages and will still run in five years.
 *
 * Screens are recomputed from questions.js on every render, so `showIf`
 * branching stays correct when an earlier answer is changed on the way back.
 * Position is tracked by screen ID and a visit stack, never by array index —
 * indices shift when a branch opens or closes.
 */
(function () {
  'use strict';

  var C  = window.SURVEY_CONFIG;
  var QS = window.SURVEY_QUESTIONS;

  var el = {
    screen:   document.getElementById('screen'),
    back:     document.getElementById('back'),
    next:     document.getElementById('next'),
    counter:  document.getElementById('counter'),
    fill:     document.getElementById('progressfill'),
    bar:      document.getElementById('progressbar'),
    brand:    document.getElementById('brand'),
    announce: document.getElementById('announce'),
    main:     document.getElementById('main'),
  };

  // --------------------------------------------------------------- state
  var state = {
    answers: {},
    currentId: QS[0].id,
    stack: [],          // visited screen ids, for Back
    activeMs: 0,        // accumulated on-screen time, idle-capped
    vwWarned: false,    // Van Westendorp ordering warned once already
    vwValid: 1,
    phase: 'survey',    // survey | done
    submittedId: null,
  };

  var tScreen = Date.now();
  var SCREEN_IDLE_CAP = 120000; // don't count more than 2 min on one screen

  // ------------------------------------------------------------ utilities
  function esc(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  function srcParam() {
    var v = '';
    try { v = new URLSearchParams(location.search).get('src') || ''; } catch (e) { v = ''; }
    v = v.toLowerCase().replace(/[^a-z0-9_-]/g, '').slice(0, 24);
    return C.allowedSrc.indexOf(v) >= 0 ? v : (v ? 'unknown' : 'direct');
  }

  function parseMoney(raw) {
    if (raw === null || raw === undefined) return null;
    var cleaned = String(raw).replace(/[^0-9.]/g, '');
    if (!cleaned) return null;
    var n = parseFloat(cleaned);
    return isFinite(n) && n >= 0 ? Math.round(n) : null;
  }

  // Screens whose showIf currently passes.
  function visible() {
    return QS.filter(function (q) {
      return typeof q.showIf !== 'function' || q.showIf(state.answers);
    });
  }

  // Screen count for the progress denominator, assuming any not-yet-answered
  // branching question will open its branch (see config.progressAssumes).
  // Keeps the counter from growing mid-survey.
  function projectedTotal(visNow) {
    var assumes = C.progressAssumes || {};
    var pending = Object.keys(assumes).filter(function (k) {
      return state.answers[k] === undefined;
    });
    if (!pending.length) return visNow.length;

    pending.forEach(function (k) { state.answers[k] = assumes[k]; });
    var n = visible().length;
    pending.forEach(function (k) { delete state.answers[k]; });
    return Math.max(n, visNow.length);
  }

  function isAnswered(q) {
    if (q.type === 'info') return true;
    if (q.required === false) return true;
    var a = state.answers[q.id];
    if (q.type === 'multi') return Array.isArray(a) && a.length > 0;
    if (q.type === 'scale') return typeof a === 'number';
    if (q.type === 'currencyGroup') {
      if (!a) return false;
      for (var i = 0; i < q.fields.length; i++) {
        if (typeof a[q.fields[i].id] !== 'number') return false;
      }
      return true;
    }
    return a !== undefined && a !== null && a !== '';
  }

  // ------------------------------------------------------------- persistence
  function save() {
    if (!C.enableResume) return;
    // Never persist a post-submit state, or the visibilitychange handler
    // writes a finished survey back into localStorage when the tab is
    // backgrounded, blocking the next scan on that phone.
    if (state.phase !== 'survey') return;
    try {
      localStorage.setItem(C.storageKey, JSON.stringify({
        answers: state.answers,
        currentId: state.currentId,
        stack: state.stack,
        activeMs: state.activeMs,
        vwWarned: state.vwWarned,
        vwValid: state.vwValid,
        phase: state.phase,
        t: Date.now(),
      }));
    } catch (e) { /* private browsing, quota — not worth failing over */ }
  }

  function restore() {
    if (!C.enableResume) return;
    try {
      var raw = localStorage.getItem(C.storageKey);
      if (!raw) return;
      var d = JSON.parse(raw);
      // Stale partials are worse than none — a week-old resume is a different person.
      if (!d || !d.t || Date.now() - d.t > 7 * 24 * 3600 * 1000) { clearSaved(); return; }
      // An already-submitted survey must never resurrect: the next person to
      // scan this phone at the counter would be blocked from taking it.
      if (d.phase !== 'survey') { clearSaved(); return; }
      state.answers  = d.answers || {};
      state.currentId= d.currentId || QS[0].id;
      state.stack    = d.stack || [];
      state.activeMs = d.activeMs || 0;
      state.vwWarned = !!d.vwWarned;
      state.vwValid  = typeof d.vwValid === 'number' ? d.vwValid : 1;
      state.phase    = 'survey'; // only an in-progress survey is ever resumable
    } catch (e) { /* ignore */ }
  }

  function clearSaved() {
    try { localStorage.removeItem(C.storageKey); } catch (e) {}
  }

  // ------------------------------------------------------------- rendering
  function render() {
    var vis = state.phase === 'survey' ? visible() : [];
    var idx = -1, q = null;

    if (state.phase === 'survey') {
      for (var i = 0; i < vis.length; i++) {
        if (vis[i].id === state.currentId) { idx = i; q = vis[i]; break; }
      }
      // The current screen was branched away by a changed earlier answer.
      // Fall forward to the first unanswered screen rather than dead-ending.
      if (!q) {
        q = vis.find(function (s) { return !isAnswered(s); }) || vis[vis.length - 1];
        idx = vis.indexOf(q);
        state.currentId = q.id;
      }
    }

    el.screen.innerHTML = '';

    if (state.phase === 'done') { renderDone(); progress(1, 1, true); return; }

    var node;
    switch (q.type) {
      case 'info':          node = viewInfo(q); break;
      case 'single':        node = viewSingle(q); break;
      case 'multi':         node = viewMulti(q); break;
      case 'scale':         node = viewScale(q); break;
      case 'currencyGroup': node = viewCurrency(q); break;
      case 'text':          node = viewText(q); break;
      default:              node = viewInfo(q);
    }
    el.screen.appendChild(node);

    el.back.hidden = state.stack.length === 0;
    el.next.textContent = q.cta || (idx === vis.length - 1 ? 'Finish' : 'Next');
    el.next.disabled = false;

    // Counter is hidden on the very first screen only — "1 of 21" is the first
    // thing someone sees after scanning at the counter, and it reads as a
    // chore. From screen two on, knowing where you are helps.
    var total = projectedTotal(vis);
    progress(idx + 1, total, idx === 0);
    el.announce.textContent = (q.title || '') + '. Question ' + (idx + 1) + ' of ' + total + '.';
    el.main.focus();
    window.scrollTo(0, 0);
  }

  function progress(cur, total, hideCounter) {
    var pct = total ? Math.round((cur / total) * 100) : 0;
    el.fill.style.width = pct + '%';
    el.bar.setAttribute('aria-valuenow', String(pct));
    el.counter.textContent = hideCounter ? '' : cur + ' of ' + total;
  }

  function h(tag, cls, html) {
    var n = document.createElement(tag);
    if (cls) n.className = cls;
    if (html !== undefined) n.innerHTML = html;
    return n;
  }

  function head(q) {
    var frag = document.createDocumentFragment();
    frag.appendChild(h('h1', null, esc(q.title)));
    if (q.help) frag.appendChild(h('p', 'help', q.help)); // help may carry markup
    return frag;
  }

  function viewInfo(q) {
    var wrap = h('div');
    wrap.appendChild(h('h1', null, esc(q.title)));
    if (q.body) {
      var b = h('div', 'body-copy');
      q.body.forEach(function (p) { b.appendChild(h('p', null, p)); });
      wrap.appendChild(b);
    }
    return wrap;
  }

  function viewSingle(q) {
    var wrap = h('div');
    wrap.appendChild(head(q));
    var list = h('div', 'choices');
    list.setAttribute('role', 'radiogroup');
    list.setAttribute('aria-label', q.title);

    q.options.forEach(function (o) {
      var lab = h('label', 'choice');
      var inp = document.createElement('input');
      inp.type = 'radio'; inp.name = q.id; inp.value = o.value;
      inp.checked = state.answers[q.id] === o.value;
      if (inp.checked) lab.classList.add('on');
      inp.addEventListener('change', function () {
        state.answers[q.id] = o.value;
        Array.prototype.forEach.call(list.querySelectorAll('.choice'), function (c) {
          c.classList.remove('on');
        });
        lab.classList.add('on');
        save();
      });
      lab.appendChild(inp);
      lab.appendChild(h('span', null, esc(o.label)));
      list.appendChild(lab);
    });

    wrap.appendChild(list);
    return wrap;
  }

  function viewMulti(q) {
    var wrap = h('div');
    wrap.appendChild(head(q));
    var list = h('div', 'choices');
    list.setAttribute('role', 'group');
    list.setAttribute('aria-label', q.title);
    if (!Array.isArray(state.answers[q.id])) state.answers[q.id] = [];

    function sync() {
      var sel = state.answers[q.id];
      Array.prototype.forEach.call(list.querySelectorAll('.choice'), function (lab) {
        var inp = lab.querySelector('input');
        var on = sel.indexOf(inp.value) >= 0;
        inp.checked = on;
        lab.classList.toggle('on', on);
        // Grey out the rest once a cap is hit, but never the ones already picked.
        var capped = q.max && sel.length >= q.max && !on;
        inp.disabled = !!capped;
        lab.classList.toggle('disabled', !!capped);
      });
    }

    q.options.forEach(function (o) {
      var lab = h('label', 'choice');
      var inp = document.createElement('input');
      inp.type = 'checkbox'; inp.name = q.id; inp.value = o.value;
      inp.addEventListener('change', function () {
        var sel = state.answers[q.id];
        if (inp.checked) {
          // "Nothing yet" / "No hesitations" clear everything else, and vice versa.
          if (o.exclusive) sel = [o.value];
          else {
            sel = sel.filter(function (v) {
              var opt = q.options.find(function (x) { return x.value === v; });
              return !(opt && opt.exclusive);
            });
            if (sel.indexOf(o.value) < 0) sel.push(o.value);
          }
        } else {
          sel = sel.filter(function (v) { return v !== o.value; });
        }
        state.answers[q.id] = sel;
        sync(); save();
      });
      lab.appendChild(inp);
      lab.appendChild(h('span', null, esc(o.label)));
      list.appendChild(lab);
    });

    wrap.appendChild(list);
    sync();
    return wrap;
  }

  function viewScale(q) {
    var wrap = h('div');
    wrap.appendChild(head(q));
    var sw = h('div', 'scale-wrap');
    var row = h('div', 'scale');
    row.setAttribute('role', 'group');
    row.setAttribute('aria-label', q.title);

    for (var n = q.min; n <= q.max; n++) {
      (function (val) {
        var b = document.createElement('button');
        b.type = 'button';
        b.textContent = String(val);
        b.setAttribute('aria-pressed', state.answers[q.id] === val ? 'true' : 'false');
        b.setAttribute('aria-label', String(val));
        b.addEventListener('click', function () {
          state.answers[q.id] = val;
          Array.prototype.forEach.call(row.querySelectorAll('button'), function (x) {
            x.setAttribute('aria-pressed', 'false');
          });
          b.setAttribute('aria-pressed', 'true');
          save();
        });
        row.appendChild(b);
      })(n);
    }
    sw.appendChild(row);

    var labs = h('div', 'scale-labels');
    labs.appendChild(h('span', null, esc(q.minLabel || '')));
    labs.appendChild(h('span', null, esc(q.midLabel || '')));
    labs.appendChild(h('span', null, esc(q.maxLabel || '')));
    sw.appendChild(labs);

    wrap.appendChild(sw);
    return wrap;
  }

  function viewCurrency(q) {
    var wrap = h('div');
    wrap.appendChild(head(q));
    if (!state.answers[q.id]) state.answers[q.id] = {};

    q.fields.forEach(function (f) {
      var box = h('div', 'money-field');
      var lab = document.createElement('label');
      lab.setAttribute('for', 'in_' + f.id);
      lab.textContent = f.label;
      box.appendChild(lab);

      var mi = h('div', 'money-input');
      mi.appendChild(h('span', 'sigil', '$'));
      var inp = document.createElement('input');
      inp.id = 'in_' + f.id;
      inp.type = 'text';
      // inputmode over type=number: no spinners, no locale decimal weirdness,
      // and it still raises the numeric keypad on iOS and Android.
      inp.inputMode = 'numeric';
      inp.autocomplete = 'off';
      inp.setAttribute('enterkeyhint', 'next');
      var cur = state.answers[q.id][f.id];
      inp.value = typeof cur === 'number' ? cur.toLocaleString('en-US') : '';
      inp.addEventListener('input', function () {
        var v = parseMoney(inp.value);
        if (v === null) delete state.answers[q.id][f.id];
        else state.answers[q.id][f.id] = v;
        save();
      });
      inp.addEventListener('blur', function () {
        var v = state.answers[q.id][f.id];
        inp.value = typeof v === 'number' ? v.toLocaleString('en-US') : '';
      });
      mi.appendChild(inp);
      box.appendChild(mi);
      wrap.appendChild(box);
    });

    var slot = h('div');
    slot.id = 'vw-alert';
    wrap.appendChild(slot);
    return wrap;
  }

  function viewText(q) {
    var wrap = h('div');
    wrap.appendChild(head(q));
    var ta = document.createElement('textarea');
    ta.maxLength = q.maxLength || 400;
    ta.value = state.answers[q.id] || '';
    ta.setAttribute('aria-label', q.title);
    ta.addEventListener('input', function () {
      state.answers[q.id] = ta.value;
      save();
    });
    wrap.appendChild(ta);
    return wrap;
  }

  // --------------------------------------------------------- lead & done
  function renderDone() {
    el.back.hidden = true;
    el.next.hidden = true;
    var wrap = h('div', 'done');
    wrap.innerHTML =
      '<div class="tick"><svg viewBox="0 0 24 24" fill="none" stroke-width="2.5" ' +
      'stroke-linecap="round" stroke-linejoin="round"><path d="M20 6L9 17l-5-5"/></svg></div>' +
      '<h1>Thank you.</h1>' +
      '<div class="body-copy"><p>Your answers help us decide whether this is worth ' +
      'bringing to the practice. We appreciate the two minutes.</p></div>';
    el.screen.appendChild(wrap);
    el.announce.textContent = 'Survey complete. Thank you.';
  }

  // ------------------------------------------------------------ validation
  function validateCurrent(q) {
    if (!isAnswered(q)) {
      var msg = q.type === 'currencyGroup'
        ? 'Please put a number in each box.'
        : 'Please choose an answer to continue.';
      flash(msg);
      return false;
    }

    // Van Westendorp ordering. Warn once, then let them through with the
    // record flagged — a hard block here costs more responses than the bad
    // rows are worth, and analyze.py excludes vw_valid = 0 from the curves.
    if (q.type === 'currencyGroup' && q.monotonic) {
      var a = state.answers[q.id];
      var seq = q.fields.map(function (f) { return a[f.id]; });
      var ok = true;
      for (var i = 1; i < seq.length; i++) if (seq[i] < seq[i - 1]) ok = false;
      if (!ok && !state.vwWarned) {
        state.vwWarned = true;
        var slot = document.getElementById('vw-alert');
        if (slot) {
          slot.innerHTML = '<div class="alert"><strong>These look out of order.</strong> ' +
            'Each amount should be the same or higher than the one above it. ' +
            'Mind taking another look? If they’re right as-is, just tap Next again.</div>';
          slot.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        }
        save();
        return false;
      }
      if (!ok) state.vwValid = 0;
    }
    return true;
  }

  function flash(msg) {
    el.announce.textContent = msg;
    var slot = document.getElementById('vw-alert');
    if (!slot) {
      slot = h('div');
      slot.id = 'vw-alert';
      el.screen.firstChild.appendChild(slot);
    }
    slot.innerHTML = '<div class="alert">' + esc(msg) + '</div>';
    slot.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }

  // ---------------------------------------------------------------- submit
  function buildPayload(completed) {
    return {
      src: srcParam(),
      duration_ms: state.activeMs,
      completed: completed ? 1 : 0,
      vw_valid: state.vwValid,
      path: (state.answers.a1 && state.answers.a1 !== 'none') ? 'full' : 'short',
      answers: state.answers,
      meta: {
        screen_w: window.screen && window.screen.width || null,
        ua_mobile: /Mobi|Android|iPhone|iPad/i.test(navigator.userAgent) ? 1 : 0,
        submitted_client_at: new Date().toISOString(),
      },
    };
  }

  function submitSurvey() {
    el.next.disabled = true;
    el.next.textContent = 'Sending…';

    // Honeypot tripped: show the thank-you, send nothing. Returning a normal
    // success keeps a bot from noticing it was caught and adapting.
    var hp = document.getElementById('hp_website');
    if (hp && hp.value) {
      state.phase = 'done';
      clearSaved();
      render();
      return Promise.resolve();
    }

    return fetch('/api/submit', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(buildPayload(true)),
    }).then(function (r) {
      if (!r.ok) throw new Error('http ' + r.status);
      return r.json();
    }).then(function (j) {
      state.submittedId = j && j.id || null;
      state.phase = 'done';
      // The response is on the server now, so the local copy has no value and
      // keeping it only risks a stale resume when the next patient scans the
      // same phone at the counter.
      clearSaved();
      render();
    }).catch(function () {
      el.next.disabled = false;
      el.next.textContent = 'Finish';
      flash('We couldn’t send that — check your signal and tap Finish again. ' +
            'Your answers are saved on this phone either way.');
    });
  }

  // ------------------------------------------------------------ navigation
  function tickTime() {
    var d = Date.now() - tScreen;
    state.activeMs += Math.min(d, SCREEN_IDLE_CAP);
    tScreen = Date.now();
  }

  el.next.addEventListener('click', function () {
    if (state.phase === 'done') return;

    var vis = visible();
    var q = vis.find(function (s) { return s.id === state.currentId; });
    if (!q) return;
    if (!validateCurrent(q)) return;

    tickTime();

    // Recompute AFTER answering — this screen's answer may have opened or
    // closed a branch (e.g. e3 = monthly reveals e4).
    var after = visible();
    var i = after.findIndex(function (s) { return s.id === q.id; });

    if (i >= 0 && i < after.length - 1) {
      state.stack.push(q.id);
      state.currentId = after[i + 1].id;
      save();
      render();
    } else {
      submitSurvey();
    }
  });

  el.back.addEventListener('click', function () {
    if (!state.stack.length) return;
    tickTime();
    state.currentId = state.stack.pop();
    save();
    render();
  });

  // ------------------------------------------------------------------ boot
  function boot() {
    if (C.logoFile) {
      var img = document.createElement('img');
      img.src = '/' + C.logoFile;
      img.alt = C.practiceName + (C.practiceTagline ? ' — ' + C.practiceTagline : '');
      el.brand.innerHTML = '';
      el.brand.appendChild(img);
    } else {
      el.brand.innerHTML = '';
      el.brand.appendChild(h('span', 'brand-name', esc(C.practiceName)));
      if (C.practiceTagline) {
        el.brand.appendChild(h('span', 'brand-tag', esc(C.practiceTagline)));
      }
    }
    document.title = C.practiceName + ' — Patient Survey';

    restore();
    render();

    // Best-effort save if the tab is backgrounded or closed at the counter.
    document.addEventListener('visibilitychange', function () {
      if (document.visibilityState === 'hidden') { tickTime(); save(); }
    });
  }

  boot();
})();
