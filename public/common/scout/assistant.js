/* VibeOffice — the assistant: Scout (scout.js), the beagle on his mat, on every app's window; later the home
 * of an AI assistant. He comes in when an app starts, except on iOS or Android and in a browser driven by a
 * test (there only once the user has shown him), and not again on a day the user hid him. The first time each
 * day he is shown, by himself or from the menu, he says good morning and who he is, and offers a few things to
 * do; after that Help ▸ Show the Assistant brings him in with a wave. A click on him offers the things to do, and a right-click gives Animate!, Search commands, the app's Help and Hide
 * (he waves goodbye). Where he sits is kept for the whole suite. While the user works he acts out what they do
 * (saving, printing, searching, …), from the suite's own signals; the apps don't call him. His idling follows
 * them: a small thing now and then while they work, more once they are quiet, a nap after five minutes, from
 * which they wake him. While the window is in the background he stands still and nothing is drawn, to save
 * energy; he carries on when it comes back.
 *
 *   L.assistant.show(asked) / hide() / shown() / scout()
 *   command showAssistant: Show the Assistant / Hide the Assistant (each app's Help menu)
 */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {});
  if (typeof document === 'undefined') return;
  const ui = L.ui;
  const KEY = 'vibeoffice.assistant';     // suite-wide, as the look: { on, r, b }: shown by the user, his distance from the window's right and bottom
  const GREETED = 'scout.lastGreetedDay'; // the last day he said good morning
  const HIDDEN = 'scout.lastHiddenDay';   // the last day the user hid him: he doesn't come in by himself that day
  const SIZE = 154, H = Math.round(SIZE * 235 / 260);
  const AWAY_MS = 250;   // a blur shorter than this (a quick switch) is not "in the background"
  /* the suite's settings (VO.settings, common/look.js); what can't be kept is forgotten with the page */
  const get = (k) => (root.VO && root.VO.settings ? root.VO.settings.get(k) : null);
  const put = (k, v) => { if (root.VO && root.VO.settings) root.VO.settings.set(k, v); };
  const load = () => { try { return JSON.parse(get(KEY)) || {}; } catch (e) { return {}; } };
  const save = (o) => put(KEY, JSON.stringify(Object.assign(load(), o)));
  /** the local date, 2026-10-11; days are only ever compared for equality, so a clock set back breaks nothing */
  const today = () => { const d = new Date(); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); };

  /* in when the app starts: not on a day the user hid him; otherwise once they have shown him, or by default
     (but not on a phone or tablet, in an automated browser (UI checks) or in the lite edition, for its small screen
     and memory) */
  const wanted = () => get(HIDDEN) !== today() && (load().on === true || (!L.isMobile && !navigator.webdriver && !(window.VO && VO.host.lite)));
  let scout = null, on = wanted(), away = false, awayTimer = 0;   // away: the window is in the background

  const searchInput = () => ui.searchInput && ui.searchInput();
  const can = (id) => !!ui.cmds[id] && ui.enabled(id);
  /* points at Search commands and puts the cursor in it */
  function toSearch() {
    const inp = searchInput();
    if (!inp) return;
    const r = inp.getBoundingClientRect();
    scout.stop();
    scout.gestureAt(r.left + r.width / 2, r.top + r.height / 2);
    scout.speak(L.t('Type what you want to do in Search commands, up there.'));
    inp.focus();
  }
  /* the welcome's choices; a command's only where the app has it */
  const CHOICES = [
    { label: 'Find a command by name', when: () => !!searchInput(), run: toSearch },
    { label: 'Look something up in Help', cmd: 'help' },
    { label: 'Get started', cmd: 'gettingStarted' },
    { label: 'See the keyboard shortcuts', cmd: 'keyboardHelp' },
    { label: 'Show me a trick', run: () => scout.animate() },
    { label: 'Nothing for now' },
  ];

  function place() {
    const st = load();
    scout.placeAt(innerWidth - SIZE - (st.r != null ? st.r : 24), innerHeight - H - (st.b != null ? st.b : 56));
  }
  function make() {
    scout = new L.Scout({
      size: SIZE, label: L.t('Scout, the assistant. Click him for help; right-click for more.'),
      onclick: () => { if (scout.hidden) return; if (balloonOpen()) scout.stop(); else ask('What would you like to do?'); },
      onmenu: (x, y) => ui.openMenu([
        { label: '&Animate!', run: () => { scout.stop(); scout.animate(); } },
        '-',
        { label: '&Search commands', key: 'Alt+Q', run: toSearch, enabled: () => !!searchInput() },
        'help',
        '-',
        { label: '&Hide', run: () => hide() },
      ], { x, y }),
      onmoved: (x, y) => save({ r: Math.round(innerWidth - SIZE - x), b: Math.round(innerHeight - H - y) }),
    });
    scout.freeze(away);
    place();
  }
  const balloonOpen = () => !!(scout && scout.balloon && scout.balloon.classList.contains('is-on'));

  /** The balloon with the choices; the one picked runs once it closes. */
  function ask(text, wait) {
    if (!wait) scout.stop();
    const list = CHOICES.filter((c) => (c.cmd ? can(c.cmd) : !c.when || c.when()));
    scout.ask(L.t(text), list.map((c) => ({ label: L.t(c.label), value: c }))).then((c) => {
      if (!c) return;
      if (c.cmd) ui.exec(c.cmd);
      else if (c.run) c.run();
    });
  }

  /** the first time he is shown each day: good morning (afternoon, evening) and who he is; '' after that */
  function hello() {
    if (get(GREETED) === today()) return '';
    put(GREETED, today());
    const h = new Date().getHours();
    return h >= 5 && h < 12 ? L.t('Good morning! I\'m Scout, your soon-to-be-AI assistant. For now, I mostly fetch.')
      : h >= 12 && h < 18 ? L.t('Good afternoon! I\'m Scout, your soon-to-be-AI assistant. For now, I mostly fetch.') : L.t('Good evening! I\'m Scout, your soon-to-be-AI assistant. For now, I mostly fetch.');
  }

  /** Bring him in. The first time today he waves (Greeting), introduces himself and asks what the user would like
   *  to do; after that, asked (the user showed him) he waves, nothing said, and in by himself he just pops in (Show). */
  function show(asked) {
    if (!on) { on = true; save({ on: true }); put(HIDDEN, null); }
    if (!scout) make();
    scout.stop();
    const intro = hello();
    scout.play(asked || intro ? 'Greeting' : 'Show', { interrupt: true });
    if (!intro) return;
    scout.delay(900);
    ask(L.t('{0} What would you like to do?', intro), true);
  }
  /** He waves goodbye and drops through his mat, and stays away today; quiet: gone at once (hidden in another window). */
  function hide(quiet) {
    if (on) { on = false; save({ on: false }); put(HIDDEN, today()); }
    if (!scout || scout.hidden) return;
    scout.stop();
    if (quiet === true) scout.hide(true); else scout.play('GoodBye');
  }

  ui.cmd('showAssistant', {
    get label() { return on ? 'Hide the &Assistant' : 'Show the &Assistant'; },
    tip: 'Scout, the assistant',
    run: () => (on ? hide() : show(true)),
  });

  /* ---------- while the user works ----------
     He acts out what the user asks for, as they ask (the command, or the dialog it opens), from the suite's own
     signals (common/ui.js): the commands they run (vo-command), message boxes (vo-message) and the browser
     printing; the work that keeps them waiting (vo-busy, the busy box) keeps him at it. Saving is the
     exception: he saves once the file is written (vo-saved, from L.saveFile), not when the user asks. The pairings are
     Office's (msoAnimationSaving, msoAnimationPrinting, …). Manners: only while he is on screen with the window
     in front; never while he is asking something or coming and going; an idle gives way at once, and a nap ends
     with his waking up; an act under way plays out, and one act waits for it (a newer one takes its place); the
     act he is at already, nothing new; nothing with reduced motion. */
  const ON_COMMAND = {
    Print: 'print printQuick printPreview',
    Fetch: 'open slidesFromFiles insertFile importText',
    Searching: 'find replace goTo goToSpecial fileSearch research thesaurus',
    CheckingSomething: 'spelling errorChecking wordCount compareDocs evaluateFormula tracePrecedents traceDependents traceError circleInvalid',
    SendMail: 'sendMail',
    Writing: 'insertComment insertFootnote insertFootnoteNow insertEndnoteNow speakerNotes',
    EmptyTrash: 'deleteSlide deleteSheet deleteAllComments clearAll',
    GetArtsy: 'insertPicture insertClipArt insertWordArt insertDiagram insertChart photoAlbum insertMedia slideDesign background themeDlg',
    GetTechy: 'autoSum insertFunction formulaDlg insertField macros calcNow calcSheet goalSeek scenarios pivotTable dataTable subtotals consolidate',
    GetWizardy: 'mailMerge letterWizard envelopes chartWizard autoSummarize autoFormat tableAutoFormat animSchemes',
  };
  const ACT_OF = {};
  for (const [act, ids] of Object.entries(ON_COMMAND)) for (const id of ids.split(' ')) ACT_OF[id] = act;
  const SAVES = /^(save|saveAs|saveWeb|exportPDF|exportCSV)$/;   // a file written after these is a save (after Print, a PDF is not)
  const WORK_MS = 400;   // the busy box shorter than this is no work to act out
  const ENTRANCES = /^(Show|Greeting|Hide|GoodBye)$/;
  let work = null, workDone = null, lastCmd = '';   // work: a promise that settles when the busy box goes

  const free = () => !!scout && !scout.hidden && !scout.frozen && !balloonOpen() && !matchMedia('(prefers-reduced-motion: reduce)').matches;
  /** Act out what the user asks for; if work keeps them waiting once it is over, he carries on at it (Processing). */
  function react(act) {
    if (!free()) return;
    const cur = scout.current;
    if (cur && (ENTRANCES.test(cur.name) || cur.name === act)) return;
    scout.play(act, { next: true }).then((done) => { if (done && work) scout.play('Processing', { until: work }); });
  }

  addEventListener('vo-command', (e) => { lastCmd = e.detail.id; if (ACT_OF[lastCmd]) react(ACT_OF[lastCmd]); });
  addEventListener('vo-saved', () => { if (SAVES.test(lastCmd)) react('Save'); });
  /* work: what he is doing carries on until it is over; with nothing to do, he works along (Processing) */
  addEventListener('vo-busy', (e) => {
    if (!e.detail.on) { if (workDone) workDone(); work = workDone = null; return; }
    const w = work = new Promise((r) => { workDone = r; });
    setTimeout(() => { if (work === w && free() && (!scout.current || scout.current.idle)) scout.play('Processing', { next: true, until: w }); }, WORK_MS);
  });
  addEventListener('vo-message', (e) => { if (e.detail.icon === 'warn' || e.detail.icon === 'error') react('Alert'); });
  addEventListener('beforeprint', () => react('Print'));
  /* the user at work: what he idles at follows them (keys, presses, the wheel, the window coming back; the mouse
     merely moving is no work) */
  const stir = () => { if (scout) scout.userActive(); };
  for (const type of ['keydown', 'pointerdown', 'wheel']) addEventListener(type, stir, { capture: true, passive: true });

  /* the window in the background: he stands still, nothing drawn until it comes back (a hidden page draws
     nothing anyway) */
  addEventListener('blur', () => {
    clearTimeout(awayTimer);
    awayTimer = setTimeout(() => {
      if (document.hasFocus()) return;
      away = true;
      if (scout) scout.freeze(true);
    }, AWAY_MS);
  });
  addEventListener('focus', () => {
    clearTimeout(awayTimer);
    away = false;
    if (scout) scout.freeze(false).userActive();
  });
  addEventListener('resize', () => { if (scout) place(); });
  /* shown or hidden in another window */
  if (root.VO && root.VO.settings) root.VO.settings.on((key) => {
    if ((key !== KEY && key !== HIDDEN && key !== null) || wanted() === on) return;
    if (on) hide(true); else show(false);
  });
  /* in when the app has started */
  if (on) addEventListener('load', () => setTimeout(() => { if (on && !scout) show(false); }, 400));

  L.assistant = { show, hide, shown: () => on, scout: () => scout };
})(typeof window !== 'undefined' ? window : globalThis);
