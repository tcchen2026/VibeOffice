# Testing

## Serving

    python3 tools/serve.py          # http://127.0.0.1:8760/, no caching

Every response carries `Cache-Control: no-store`, so edited HTML and scripts are always re-fetched. Serve rather than open from disk: a `file://` page cannot fetch the proofing dictionaries, and password-protected files need WebCrypto, which wants https or localhost.

## Screenshots and console errors: tools/shot.mjs

Drives a page in the Chromium listening on CDP 127.0.0.1:9222 (start one with `chromium --remote-debugging-port=9222 --remote-allow-origins=*`, or `--headless=new` on a machine without a display). It starts `tools/serve.py` if nothing listens on :8760. It prints console errors, exceptions and requests that leave 127.0.0.1, and fails (exit code 1) on a refusal of the pages' Content-Security-Policy, a window the page opens or a page of another site in the tab (docs/suite.md, Security); the UI tests' `t.clean()` fails on the same.

    node tools/shot.mjs ledger/ '[{"wait":1000,"shot":"ledger.png"}]'
    node tools/shot.mjs quire/ steps.json --w 1440 --h 900 --out /tmp/shots

Each step is `{ "js": "...", "wait": ms, "shot": "name.png|jpg" }`; `js` is an async function body run in the page, and its return value is printed. The app's namespace is `window.L`. Its pages say they are automated (`navigator.webdriver`), so Scout, the assistant, does not come in by himself; a step that wants him calls `L.assistant.show()` (and sets the setting back at the end, see Looks). At the end the tool lists console errors, exceptions and every request that left 127.0.0.1 (there should be none). It opens and closes its own tab; never close the browser's other tabs.

The install screenshots in the manifest (`public/<app>/preview.jpg`) are 1280×720:

    node tools/shot.mjs quire/ '[{"wait":1500,"shot":"preview.jpg"}]' --h 720 --out public/quire

The PWA icons are drawn from `public/icons/icon.svg` and `icon-maskable.svg` (see docs/suite.md).

## The Start Center and app launch

The browser profile keeps IndexedDB between runs of `shot.mjs`, so a test can hand a file to an app and then look at the Start Center:

    node tools/shot.mjs quire/ '[{"js":"const b = await L.docx.write(L.D.doc, { stats: L.app.docStats() }); setTimeout(() => VO.handoff(new File([b], \"t.docx\")), 50)"},{"wait":3500,"js":"return document.title"}]'
    node tools/shot.mjs "" '[{"wait":800,"shot":"start.png"}]'

Check after a change to `suite.js` or to an app's start, open or save code: `<app>/?new` starts blank; a handed-over file opens in place of the blank document (one window); the file and a thumbnail appear under Recent Files; clicking it reopens it. Clear the list with `await VO.recent.clear()`.

Plain http: `node --test tests/insecure.test.mjs` checks that only the shared helpers call `crypto.randomUUID` and `crypto.subtle`, which a page served over plain http lacks, and that their fallbacks work: random UUIDs, and a password-protected file opening without WebCrypto (docs/suite.md, Plain http). To see such a page, serve the site on an address that isn't loopback, for example the Docker image (docs/suite.md, Self-hosting with Docker) published on a port and opened at the machine's LAN address: the Start Center shows its bar.

Templates: `node tools/templates.mjs` opens every template (`<app>/?template=<id>`) and fails on a script error or a missing preview; look at the new previews in the Create File view.

Preset shapes: `common/presets.js` (written by `python3 -I tools/shapes/presets.py` from Apache OpenOffice's copy of the standard's definitions, three slips there mended) is drawn by `common/geometry.js`. `node tools/shapes/deck.mjs` writes `tmp/shapes/shapes.pptx`, every preset (176 without the lines and connectors) at its default adjustments in a wide and a tall cell; save it as PDF from PowerPoint, then `python3 -I tools/shapes/masks.py office.pdf` writes where PowerPoint paints each one (`tests/shapes/powerpoint.json`, for `node --test tests/shapes.test.mjs`), and `python3 -I tools/shapes/compare.py office.pdf ours/` scores screenshots of our slides (`ours/shapes__N.png`, 1920 px wide) against it. PowerPoint 2016 (2026-11-07): every preset within 1.3% of PowerPoint's painted pixels in both cells (Corner Tabs the most); the hand-made drawings before differed by more than 5% in 51 presets. Text warps likewise: `python3 -I tools/shapes/warp-deck.py` writes `tmp/shapes/warps.pptx` over the shapes deck (each of the 40 warps in one, two and three lines, small and long text, and variants of insets, anchoring and alignment; cell names in `tmp/shapes/warps.txt`), and `python3 -I tools/shapes/compare.py office.pdf ours/ - warps` scores `ours/warps__N.png` (`W=1920 node tools/lectern-pictures.mjs`) against PowerPoint's PDF of it: 15% on average (2026-11-07; `node --test tests/text-warps.test.mjs`).

Pattern fills: the preset bitmaps in `common/geometry.js` are PowerPoint's. `node tools/patterns/deck.mjs` writes `tmp/patterns/patterns.pptx` (each of the 54 presets in a labelled rectangle); save it as PDF from PowerPoint and `python3 tools/patterns/read-pdf.py patterns.pdf` prints the tiles, the form of `tests/patterns.powerpoint.json` (`node --test tests/patterns.test.mjs`). Ledger's cell patterns (`js/render.js`) likewise: `node tools/patterns/sheet.mjs` writes `tmp/patterns/cellpatterns.xlsx` (the 17 patterns), Excel saves it as PDF, and `read-pdf.py --cells` prints the tiles, `tests/patterns.excel-mac.json` (from Excel for Mac, which draws them at double size; the test applies the rule in `render.js` that brings them to single pixels).

Line grids: Quire's line height on a document line grid (`R.gridLine` in `quire/js/render.js`) follows Word. `python3 tools/linegrid.py make` writes `tmp/linegrid/linegrid.docx`, one labelled paragraph per case (English and Chinese text at sizes either side of the pitch, several line spacings, 15.6pt and 18pt grids); save it as PDF from Word and `python3 tools/linegrid.py read linegrid.pdf` prints the distance between the lines of each case, in points, and the line each page starts with: `tests/from-issues/11/word.json`, from Word for Mac's PDF of `linegrid.docx` there. `grid-spacing.test.mjs` checks `R.gridLine` against it and `render.test.mjs` the document as Quire draws it, line spacing and page breaks (a second). LibreOffice draws these differently from Word (issue #11), so its PDF only checks the tool.

## Looks

`node --test tests/looks.test.mjs` checks `public/common/looks.css`: every look sets every setting, every setting the chrome uses exists, `ui.css` has no colours of its own, and each page picks the look before its stylesheets load. A change to chrome CSS is also checked in every look: the browser profile keeps `localStorage`, so set the look in a step and set it back at the end (other runs would otherwise start in it):

    node tools/shot.mjs ledger/ '[{"js":"VO.look.set(\"paper\")","wait":800,"shot":"paper.png"},{"js":"VO.look.set(\"classic\")"}]'

Classic must not change when only the theme machinery changes: compare screenshots of the same scenes before and after, pixel for pixel.

`node --test tests/ribbon.test.mjs` loads each app in the Paper 2016 look (through `tools/shot.mjs`, so Chromium must be running) and checks its ribbon: every command id exists, every button has an icon, and every menu command is reachable from the ribbon or File, apart from the exceptions listed in the test. Run it after adding or renaming a command. A ribbon change is also looked at in a narrow window (`--w 900`, where groups fold) and with KeyTips shown (`L.app.menuBar.openIndex(0)`).

`node --test tests/search.test.mjs` checks Search commands the same way: in each app and both layouts, a table of everyday queries ("image", "margins", "hyperlnk") must find the intended command first, with its location. Add a line there when a query finds the wrong thing.

`node --test tests/scout.test.mjs` checks the assistant in each app (docs/suite.md, The assistant): that he comes in by himself on a desktop (waving the first time each day, popping in after) and not on an iPhone or on a day the user hid him (scripts in `tests/scout/` run before the page's own), Help ▸ Show the Assistant (a wave; the introduction and choices the first time each day), a click on him, acting out Print, Find and saving (once the file is written; the test holds back the download), working along with a long wait, the welcome's choices, the right-click menu, no frames drawn while the window is in the background or once he is hidden. Showing or hiding him writes `localStorage['vibeoffice.assistant']`, which other open app pages follow, `scout.lastGreetedDay` and `scout.lastHiddenDay`; a UI check that shows him sets them back at the end, as the test does.

## Languages

Tests open the apps in English (`lang=en`), whatever the browser's language (docs/i18n.md). `node --test tests/i18n.test.mjs` checks that the English catalog is current (`node tools/i18n.mjs gather`) and that text which can't be translated only gets less; `node tools/i18n.mjs harvest` opens each app in the pseudo-locale and lists the English still on screen. `node --test tests/fonts.test.mjs` checks the East Asian font table, and `node --test tests/locales.test.mjs` (Chromium) saves a blank file of each app in each language and checks its defaults. Neither reads the translations.

## Mouse, selection, drag and focus: tests/ui/

Two sizes. After each change, `node tools/quick.mjs` (the uncommitted changes, or the last commit; `node tools/quick.mjs
<rev>` for everything since a commit; `--dry` lists the plan): it maps the files changed to the tests that cover them
and runs those, the Node tests first. The smoke test, `tests/ui/smoke.test.mjs` (about 20 s), runs whenever the apps
changed: the Start Center and each app load with no errors and nothing from outside; in each app a click puts the
keyboard in the document and typing lands there, the File menu and a dialog from its shortcut open and close with Esc
and give the keyboard back, a drop-down inside a dialog opens above it (Ledger's Format Cells), and Undo takes the
typing back. A touched app adds its Node tests and its own UI file (2–3 minutes), shared event, menu and dialog code
(`common/ui.js`) the focus tests in the Classic look and the drag tests, and a changed test file runs itself; the rules
are at the top of `tools/quick.mjs`. Run the change's own test on the code before it too (`tools/older.mjs`).

The whole of `tests/ui/` (about 25 minutes) runs before a release: `node tools/ui-full.mjs` runs it and records the
result in `.ui-full-tested.json` (ignored by Git) for a hash of `public/`'s content, committed or not.
`tools/ui-tested.sh` compares the `public/` about to go out with that record; `tools/web-release.sh` and
`../cf-vibeoffice/sync-cf-pages.sh` warn when the suite has not run on it, ran with failures, or `public/` has changed
since. A partial run (extra arguments, e.g. `--test-name-pattern`) or one during which `public/` changed records nothing.

`node --test --test-concurrency=1 tests/ui/*.test.mjs` drives each app in the local Chromium with real input: CDP Input events
paced like a hand (`tests/lib/input.mjs`), because synthetic DOM events hide what an app does in the middle of a drag.
The helper starts the static server, and a headless Chromium (profile `tmp/chromium-profile`, left running) when
nothing listens on CDP 127.0.0.1:9222. One file at a time (`--test-concurrency=1`): some tests move the focus to
another tab and back. Two runs at once each need a Chromium of their own: `CDP=http://127.0.0.1:9223` starts one there
(profile `tmp/chromium-profile-9223`), so neither takes the other's focus.

- `drag.test.mjs`: `ui.drag` itself, on a pad in each app's page: a click and a wobble move nothing, Esc, Ctrl+Z, a
  second button, a release the page never saw, a release in another tab and `ui.dragCancel()` cancel, and the key, the
  button and its menu go no further (a right-click right after a middle-button cancel still opens its menu); the
  press's cursor stays.
- `gestures.test.mjs`: drags in each app (a floating shape in Quire, a shape and the fill handle in Ledger, moving and
  resizing a shape in Lectern) run through the lists in `tests/lib/gestures.mjs`: each state it starts from (nothing,
  text or a range selected, the shape selected, text being edited), each key in the middle (an arrow, a letter,
  Delete, Ctrl+Z) and each way out (the release, Esc, a right-click, a release the page never saw, a release in
  another tab), then, briefly, in each mode the app can be in (Track Changes and 150 % zoom in Quire; frozen panes, the
  second sheet and 150 % in Ledger; a slide with a transition, the second slide, the Outline tab and 150 % in
  Lectern). After any of them no drag is left and no cursor held; the release made one undo step that brings back
  exactly what was there, any other ending changed and recorded nothing. A new drag adds its `matrix(…)` there, and a
  new mode goes into the lists of both files: what broke in review broke only in a mode no test had on.
- `focus.test.mjs`: the keyboard back to the document, in each app and in the Classic and Paper 2016 looks: dialogs
  from a shortcut, a toolbar or ribbon button and the context menu, closed with Esc, Cancel and OK; a dialog's access
  key with Alt (Zoom's 200%); menu-bar commands;
  the ribbon's tabs; presses on chrome without a control (toolbars, grips, the status bar, the ribbon's tab strip);
  the status bar's zoom slider dragged, dragged and cancelled with Esc, and its + clicked (a view control: no gestures
  test, as a zoom makes no undo step).
  After each, the element that had the keyboard has it again, with the same selection. Then the same in each app's
  modes (Classic): Quire with Track Changes, a header being edited, Outline view; Ledger with frozen panes, the second
  sheet, 150 %; Lectern with the speaker notes being edited, the Slides pane holding the keyboard, a slide with a
  transition.
- `textmap.test.mjs`: every piece of editable text maps to the document, and a click on it puts the caret there.
  Check 1 sweeps the page without input (`tests/lib/textmap.mjs`): in Quire every editable text node is in an
  indexed paragraph and maps to a position and back with the model's text; in Lectern the text of every placeholder,
  text box, shape, table cell and group or diagram node, each edited in turn, does the same. Check 2 clicks into text
  of each kind with real input: the caret is at the clicked character, and a letter typed goes there. Generated
  documents with every kind of text; `node tools/textmap-corpus.mjs quire|lectern DIR N OUT.json [--hf]` runs check 1
  over corpus files (`--hf`: Quire's headers and footers, opened for editing).
- `painter.test.mjs`: the Format Painter in each app (docs/suite.md, Brushes): it paints the dragged text, cells or
  object once the mouse is let go, a clicked word; a double-clicked painter stays on until Esc; links stay; in Lectern
  the text being edited stays edited, and text in another box can be painted.
- `quire.test.mjs`, `ledger.test.mjs`, `lectern.test.mjs`: each app's mouse, selection, drag and focus behaviour. A
  test named after an item of todos/UI_MOUSE_LOGIC_IMPROVE.md (Q1…, S1…, P1…) checks what Office does there and is
  marked `todo` until the item is fixed: it runs and is reported, but doesn't fail the run. The fix takes the mark off,
  so the test becomes its regression test, and the todo count in the summary is what is left of the plan.

The tests assert what the user sees (the cursor, the text and its formatting, the selection, where typing goes), so
they also run against older code, and a fix's test must fail there. `node tools/older.mjs [REV] [-- <node --test
arguments>]` extracts `public/` as it was at REV into `tmp/older/<commit>/`, serves it on a port of its own and runs
today's tests against it (`BASE`); REV is `HEAD` while `public/` has uncommitted changes, else `HEAD~1`. For example
`node tools/older.mjs 1035247~1 -- tests/ui/painter.test.mjs` (the Format Painter before its fix: all three fail), or
`node tools/older.mjs -- --test-name-pattern '^Q1 ' tests/ui/quire.test.mjs` for a fix not committed yet.

## SmartArt

`node --test tests/sw.test.mjs` checks the service worker (docs/suite.md, Service worker) in the Chromium at :9222, on copies of `public/` staged in `./tmp/sw-test` by `tools/web-release.sh` and served on 127.0.0.1 by a server of the test's own (it redirects `index.html` to its folder, like the published site): the first install and its progress (bytes too), the cached files' `Content-Type`, the Start Center's bar with its megabytes and its *Just updated successfully* notice (a button that reloads the open page in the new version), the Start Center and Quire offline, an update that fetches the new version's zip and nothing else while an open page keeps its version and a new page gets the new one, the old cache going once its page has closed, a damaged zip (one byte changed) failing the install, and no worker on any other origin (the same server as `localhost`). It clears the test origin's storage at the end. `SW_TEST_LOG=1` prints the tabs it opens and the pages it loads.

To try the worker by hand, serve a published copy on another 127.0.0.1 port (the dev server's own files are unstamped, so the worker caches nothing on :8760):

    rm -rf tmp/site && cp -r public tmp/site && tools/web-release.sh tmp/site v2026.10.11-000001
    python3 tools/serve.py 8770 tmp/site

and open http://127.0.0.1:8770/. For an update, stop the server, stage again with a later version (`…-000002`, after an edit to see it change) and serve that: the open Start Center shows *Updating cached resources* and then *Just updated successfully, Click here to reload*. With the server stopped the pages still open (offline). Afterwards clear the site's data (DevTools ▸ Application ▸ Storage ▸ Clear site data), or it goes on serving that copy.

`node --test tests/smartart.test.mjs` runs the layout engine (`public/common/smartart.js`) in Node: every layout, for 1 to 7 items, gives finite shapes inside the box, one shape per item it shows, and one text size within the limits. A change to the SmartArt writer (`public/lectern/js/smartart-io.js`, `diagramFrameXML` in `pptx-write.js`) is checked on a deck with one diagram per layout: save it from Lectern, run the Open XML SDK validator and `tools/ooxml/package.py check` on it, convert it with LibreOffice to PDF (does every diagram draw, with no word broken) and to PPTX, and reopen both copies in Lectern (every diagram editable again). Keep the decks and renders in `./tmp`.

`python3 tests/lectern/smartart-office.py OUT [name]` checks SmartArt that PowerPoint made, in the corpus decks whose layouts Lectern has: unedited saves keep every diagram part byte-identical, an edit through the items gives a save with no new `package.py` findings, and that save reopens with the diagram editable (`OUT/results.json`). It removes the Recent Files entries it makes; run the SDK validator on `OUT/*-edited.pptx` too.

## Quire: tests/quire/

| | |
|---|---|
| `node tests/quire/mdspec.js DIR` | The Markdown reader against the CommonMark spec examples and the GFM extension examples (in Node). DIR holds `spec.json` (https://spec.commonmark.org/0.31.2/spec.json), cmark-gfm's `test/spec.txt` saved as `gfm-spec.txt`, and the WHATWG `entities.json` |
| `node tests/quire/mdroundtrip.mjs DIR` | Markdown round trips in Quire (through `shot.mjs`): every `.md` in DIR saved unchanged (must be byte for byte the same), with one paragraph edited (only that block may change), and fully rewritten (same HTML). `--spec spec.json` runs the rewrite check on the spec examples |
| `python3 tests/quire/loss-audit.py originals/ saved/ out.json` | What a save loses: each original `.docx` against Quire's saved copy, feature by feature and word by word (standard library only); results in docs/quire.md |

Run both after a change to `markdown.js` or `mdio.js`; the numbers are in docs/quire.md.

Line heights (`quire/js/portions.js`): `node --test tests/quire/lines.test.mjs` (a few seconds; `REPORT=1` prints what it
measured) checks two documents in Quire against another application's PDF, read by `python3 tools/wordlines.py PDF DOCX
OUT.json` into rows (the text at one height on a page, matched by a hash of its last characters, so a document whose
text isn't ours is compared without storing it):

- a Word-saved 19-page walkthrough against Word's PDF of it (`tests/quire/lines/walkthrough-word.json`): pages start
  where Word's do, and two lines that follow each other in both are as far apart as in Word, within 0.75pt. The
  document and the PDF stay out of git, in `/work/corpora/word-ref/` (`fluent-walkthrough.docx`, `fluent-walkthrough.word.pdf`;
  or the folder named in `QUIRE_WORD_REF`); without them that part is skipped;
- `tests/quire/lines/linecases.docx` (`python3 tools/linecases.py make`), one case per paragraph between marker lines,
  against LibreOffice's PDF of it (`soffice --headless --convert-to pdf`, then `python3 tools/linecases.py read PDF
  tests/quire/lines/linecases-lo.json`): the cases LibreOffice draws as Word does are asserted, the others recorded.

Page counts across the corpus: `node tools/quire-pages.mjs OUT.jsonl [--jobs 4]` opens every Word-saved `.docx` in
`/work/corpora/word` (Microsoft Office Word with `<Pages>` in docProps/app.xml) and records Quire's page count beside
Word's; `--compare BEFORE.jsonl AFTER.jsonl` lists what a change moved, and every file it moved away from Word's count.
For the "before", serve the last commit's site on a port of its own (`git archive HEAD public | tar -x -C
tmp/pages-before`, `python3 tools/serve.py 8771 tmp/pages-before/public`) and run with `BASE=http://127.0.0.1:8771`.

## From issues: tests/from-issues/

Regression tests for a reported issue live in `tests/from-issues/N/` (N is the issue number), so a later change that breaks one shows which report it reopens. Each fails on the commit before its fix.

`node tests/from-issues/1/regressions.mjs OUTPUT [fill|fields|marks|breaks]` runs the issue #1 browser regressions (Ledger and Quire)
in isolated tabs without persistence. It checks Ledger's real Ctrl+D/Ctrl+R events, formulas,
formatting and undo/redo; Quire's 9/10 and 99/100 page totals, section counts and tab positions;
unchanged visible-text geometry when toggling marks in Print and Web Layout; and encoded
carriage returns through edits, history and DOCX save/reopen. It writes screenshots, a JSON report
and generated DOCX fixtures for independent package/SDK and LibreOffice comparisons. Set `BASE`
to an isolated previous checkout to demonstrate the failures before a fix.

`node tests/from-issues/1/review.mjs OUTPUT [protection|breaks|pagination|marks|print]` covers the follow-up
review: locked blank and distant fill destinations; CRLF across formatting, bookmarks and
inline content controls, including history and save/reopen; centered/right-aligned forced-line
pagination; mark placement in rotated, RTL and vertical text at 75%, 100% and 150% zoom; and
identical exported page pixels with marks off/on, plus native print styling.
Its generated DOCX originals, saves and edits can be compared independently in LibreOffice.

## Hostile documents: tests/hostile/

Documents written to attack the suite (todos/HARDENING.md, phase 3), and checks that the tests would see an attack.
`helpers.test.mjs`: `t.clean()` fails on a refusal of the pages' policy, a window the page opens and a page of
another site in the tab. `tests/csp.test.mjs` (Node, in a second) checks that every page carries the policy and no
inline script (`node tools/csp.mjs --check`). `export-html.test.mjs` (about 6 s, in the quick
check when an HTML writer changes) gives each app a document whose font names, colours, notes and link addresses
are markup, writes Save as Web Page (and Quire's and Ledger's clipboard HTML), and checks the result as a document
of its own: no element, attribute or script the writer didn't make, the names still there as text, the page's own
policy (`L.exportCSP`) in place; then it opens each page from a local server, where nothing of the document runs
and Lectern's own script still does.

## Ledger: tests/ledger/

The engine and the file readers and writers load in Node through `load.js` (the shared ones from `public/common/`; a missing file is an error), so most checks run without a browser.

| | |
|---|---|
| `node --test tests/ledger/*.test.js` | Unit tests: calculation, formula parser, number formats, SHA |
| `roundtrip.js file.xlsx [outdir]` | read → write → read, compare every cell's value, formula and style and the sheet settings |
| `xmlss-rt.js file.xlsx [outdir]` | the same through XML Spreadsheet 2003 |
| `recalc1.js file.xlsx` | recalculate every formula and compare with the value stored in the file |
| `csvtest.js file.csv` | parse, compare with Python's `csv` module (`csvref.py`), write back, re-read |
| `lo-compare.js ourDir loDir` | compare our saved files with LibreOffice's re-saved copies cell by cell |
| `loss-audit.py originals/ saved/ out.json` | what a save loses: features per workbook, and every cell's value, formula and style read with openpyxl from both files (Python) |
| `pool.js script.js dir out.jsonl`, `recalc-all.js`, `summ.js` | run a per-file script over a corpus with a process pool, summarise the results |
| `ui/*.js` | Playwright checks in Chromium against http://127.0.0.1:8760/ledger/: `smoke`, `flows` (44 end-to-end flows), `allcmds` (every menu and toolbar command), `password`, `perf`, `shots`. Need Playwright (`npm install --no-save playwright`) |

## Lectern: tests/lectern/

`corpus.js` opens real decks in Chromium (open, render every slide, save, reopen, compare), `summarize.js` sums up its results, `ui-features.js` drives the password dialog, Tools ▸ Options ▸ Security and the chart dialog, `chartedit.js` writes every chart for validation, `decrypt.py` checks encrypted saves independently, `loss-audit.py` counts what a save loses (features and words, original against saved copy; standard-library Python), `viscompare.py` / `triptych.py` / `stack.py` compare renderings. Against PowerPoint: `node tools/lectern-pictures.mjs decks.json out/` takes a picture of each slide as Lectern draws it, and `python3 -I tools/lectern-vs-powerpoint.py officeDir out/ cmp/` scores each against PowerPoint 2016's PDF of the deck (from the Office test API) and writes the pairs side by side, worst first. They need Playwright; how to run them is in docs/lectern.md.

## Corpora

The compatibility numbers in each app's doc come from public test corpora (LibreOffice, Apache POI, Open XML SDK, XlsxWriter, ClosedXML and others; see the app docs for the list). They are not checked in; download them outside the repo and point the per-file scripts at the folder.

The local corpus root is `/work/corpora`. The test and tool defaults use it; existing overrides
(`VO_CORPORA`, `CORPORA`, `EXCEL_CORPUS`, `LEDGER_FEATURES`, `QUIRE_WORD_REF` and command-line paths)
still take precedence.

## Lossless same-format save: reproducible baseline

### Office 2016 VM: HTTP open/repair and PDF checks

`tools/win_test_server/` is a standalone Go test server for the `win10` VM, working in `C:\vo`.
Build from the repository root with
`GOOS=windows GOARCH=amd64 go -C tools/win_test_server build -o ../../tmp/win-test-server/win_test_server.exe .`
(create `tmp/win-test-server/` first). Copy the executable and `install.ps1` there with `scp`, then run
`ssh win10 powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:/vo/install.ps1`.
The installer registers an interactive logon task; keep the VM signed in and unlocked. The server binds
to VM loopback `127.0.0.1:8765`; reach it through `ssh -N -L 18765:127.0.0.1:8765 win10`.

```sh
curl -fsS --data-binary @saved.docx \
  'http://127.0.0.1:18765/convert?filename=saved.docx&timeout=120' -o tmp/office-result.json
```

Raw DOCX/XLSX/PPTX uploads only, up to 100 MiB. `POST /convert` waits for the JSON result;
`POST /jobs` queues it and returns an ID (`202`); `GET /jobs/<id>` reads its result. `GET /health`
reports active/pending counts. The result's `pdf` and `dialogs[].screenshot` are downloadable relative URLs.
Office jobs run serially, with a 120 s default deadline (5–600 s), fresh processes, alerts enabled,
macros disabled, read-only normal opens and no repair retries. Existing Office processes cause `busy`.
The Word/Excel/PPT native exporters produce the PDFs; Excel respects print areas, PPT includes hidden slides.

`canOpenWithoutRepair` is `true` only after an observed successful normal open, `false` for an open rejection
or repair request, and `null` when inconclusive (including unknown dialogs and timeouts). A COM rejection may
have no visible prompt; keep its error. `pdfStatus` separately reports export success. Completed jobs return
HTTP 200 even on failure: inspect the JSON. Results retain source/PDF hashes, unchanged-input verification,
Office build, dialog text/images and timings in `C:\vo\jobs\<id>`. Opening/exporting is evidence for the
lossless-save review, not proof that objects remain editable or content is unchanged: test original and saved
files, compare their PDFs, and retain the feature checklist and package/SDK checks.

`node --test tests/win-test-server.test.mjs` tests the HTTP protocol, serial queue, failure evidence and
Windows cross-build. The opt-in `python3 tests/win_test_server/live.py --out tmp/win-test-server/live`
uses local public-corpus files, historical confirmed repair cases, password and timeout controls on the VM.
The VM was checked on 2026-11-02 with Office build `16.0.4266.1003`. Before replacing its executable, wait for
an idle queue and stop the task `VibeOffice Office 2016 Test API`; run `install.ps1` again after copying.

### Corpus drivers

`sh tools/corpora.sh /work/corpora` fetches the pinned LibreOffice, Apache POI and Open XML SDK revisions
listed in `tools/ooxml/corpora.py`. It resumes interrupted fetches, refuses dirty or unrelated checkouts,
and never deletes an existing corpus. `manifest.json` records source, commit, original path, size and
SHA-256. `word/`, `excel/` and `powerpoint/` contain collision-resistant symlinks. Fixtures stay outside
the repository. A source name after the directory selects just that source.
`sh tools/corpora.sh /work/corpora legacy` adds the Office 97-2003 files (`.doc`, `.xls`, `.ppt` and their
templates) from the same pinned POI and LibreOffice trees into `legacy-word/`, `legacy-excel/` and
`legacy-powerpoint/`, each checked against its Git blob hash; `node tools/legacy/measure.mjs --out DIR
[--hostile]` runs the legacy importers over them (and truncated and bit-flipped copies) in bounded workers.

The corpus drivers record every file and scenario as `ok`, `failed` or `excluded`, with the input hash
and an error/reason. Here, `ok` means the requested operation and reread completed; it does **not** mean
the file was lossless. Independent comparisons and Office acceptance are separate checks.

```sh
node tests/quire/corpus.mjs /work/corpora/word /tmp/quire.jsonl --save /tmp/lossless/quire
node tests/ledger/corpus.js /work/corpora/excel /tmp/ledger.jsonl --save /tmp/lossless/ledger
node tests/lectern/corpus.js /work/corpora/powerpoint /tmp/lectern.jsonl --lossless --save /tmp/lossless/lectern
```

Default scenarios are `save` and `repeat` (save, reopen, save again). `--edit` selects separate text,
formatting, geometry, duplicate and delete-object scenarios, plus row/column/sheet operations in Ledger
or slide insertion/deletion in Lectern. Each successful mutation also saves its undo and redo states.
`--scenarios text` runs just text editing; `--scenarios draft` in the browser drivers exercises the app's
actual snapshot and recovery hooks. For Ledger drafts, use `node tests/ledger/browser-corpus.mjs`
with the same arguments; the Node engine driver rejects that scenario. `--limit N` limits files; `--resume` skips recorded input hashes and
scenarios. A single Office file or a file list can replace the directory. Saved artifacts go into
`save/`, `repeat/`, `repeat-second/`, `text/`, `text-undo/`, etc. Files retain their source extensions so
variant errors remain observable. The browser drivers use the existing CDP browser and close only their
own tabs. Ledger uses isolated Node workers with a per-file timeout. The older Lectern rendering driver
without these switches remains available and still requires Playwright.

Quire history regressions run with `node --test tests/quire/history.test.js`. The real-file check
`node tests/quire/history.mjs /work/corpora/word/libreoffice__9b899b6707be__chart-size.docx`
checks original chart bytes after editing, undo, redo, copying and whole-document snapshot restore.

`tools/ooxml/check.mjs FILE...` uses an independent strict Python XML/ZIP reader to check content types,
relationship targets and relationship attributes, scoped drawing/slide/timing IDs, connector and timing
references, bookmark/permission endpoints, and pivot cache references. AlternateContent branches are
checked without falsely counting mutually exclusive IDs as duplicates. Output is JSONL; a failure exits
nonzero. `node --test tools/ooxml/*.test.mjs` exercises malformed packages, cycles, shared dependencies,
alternate branches, target swaps and ordered text.
`audit.test.mjs` also checks that a cell-reader timeout remains a failed, timed-out attempt instead of
being mislabeled as an unreadable style; genuine style errors stay separate.

### Office conventions: tools/ooxml/conventions.py

`node tools/ooxml/office-batch.mjs MANIFEST.json OUTPUT_DIR` makes a review batch of at most 15 Office
files, its `CHECKLIST.md` and a manifest of hashes, edits and recorded losses. Each input names its app,
authored source, output filename, edit, feature and inspection instructions. `BASE` can point at an
isolated previous writer to reproduce the same edits for the required save-to-save comparison.
Use an empty output directory. Original sources stay outside the batch. A generated batch is pending
Office review even when SDK and LibreOffice checks pass.

`python3 tools/ooxml/conventions.py /work/corpora/powerpoint SAVED_DIR [--out report.json]` lists markup our
saved files contain that no Office-authored original in the corpus does: attributes, child elements,
child order, enumerated values, preset adjustment sets, content-type forms, relationship types. Markup
already in a file's own original is not counted. It finds Office-rejected output the SDK accepts (the
slide-6 repair was such a case); run it with every full corpus run. Open findings: docs/LOSSLESS_SAVE.md.

### Shared preservation core

`node --test tests/ooxml/opc.test.mjs` tests the Node-safe `L.opc` API: immutable bytes, full relationship
types, cyclic/shared graphs, reserved names and IDs, duplicate reference remapping (including GUID
definitions, shared-person references, copies of copies and unchanged GUID-looking text), namespace/QName
scope, original nested AlternateContent emitted once, cross-document dependency transport, loss
acknowledgement, property merging, Strict-to-Transitional conversion and DrawingML/Word/spreadsheet/VML
geometry adapters. These tests do not replace real-file comparisons or Office acceptance. API contracts
and edit-preservation rules are in [suite.md](suite.md#preservation).

Load `common/xml.js`, `common/opc.js`, then `common/opc-order.js`. `L.xmlTree` supplies a portable tree;
Ledger's small `js/xml.js` bridge exposes it as `L.xml`. `L.opc.open(zip)` keeps a separate immutable
baseline; `attach(model, pkg)` makes `model.pkg` non-enumerable. `fragment(element, {pkg, part})` produces
JSON-safe XML and dependency/identity references. Use an original baseline element when exact lexical
XML is needed: `raw` reads its source range, not later mutations. `captureAC(dom, rawText)` makes its
own original parse before the reader normalises the working DOM.

Each app supplies part ownership to `L.opc.Writer`: opaque bytes are carried, merged parts replace
owned properties in schema order, and regenerated content parts emit model-attached fragments.
`emit(fragment, owner, {box})` remaps relationships and IDs and optionally changes geometry. `slice`
extracts a tokenised property while retaining only its dependencies and identities. ZIP entries,
content types and internal relationship URIs share one decoded identity; output URIs are escaped.
`duplicate`
remaps copied definitions and references together; `export`/`import` transport dependency bytes for a
cross-document clipboard. `finish()` checks relationship targets and returns `{files, dropped}`.
Loss acknowledgement is per entry; autosave must neither show a dialog nor acknowledge entries.
Package ownership, object preservation and save-dialog integration are implemented in all three apps.
Current corpus limits and pending Office acceptance are listed in each app’s save table.
`office-batch.mjs` also supports list-alignment edits, independent list copies (optionally after
an alignment edit), two-column index insertion, repeated control copies followed by a sheet copy,
master/layout decoration edits, and ungrouping charts or retained frames. A frame selector can target
an opaque preview without relying on XML text inside its alternatives. Use the same manifest against
the frozen current and previous writers before handing these files to Office.

`opc-order.js` contains ordering and ID-bound facts for 27 types, generated from the official
[ECMA-376 Part 4 schemas](https://ecma-international.org/publications-and-standards/standards/ecma-376/), and each preset shape's adjustment values (names, order, defaults) from Part 1's `presetShapeDefinitions.xml`.
The archive stays outside git; its SHA-256 is recorded in the generated file. Regenerate with:

```sh
python3 tools/ooxml/schema-order.py /work/corpora/ecma-376/part4.zip public/common/opc-order.js /work/corpora/ecma-376/part1.zip
```

`tools/ooxml/compare.mjs ORIGINAL SAVED [--policy policy.json]` defaults to exact part-byte preservation.
A policy can designate individual parts as `merged` or `regenerated`, map their target names, allow
specific ElementTree paths/attributes to change, and specify removed parts and expected slide count.
Merged XML compares expanded namespace names and QName-valued compatibility attributes. Regenerated
content compares ordered text, run/paragraph properties, cells, drawing properties and feature counts;
adjacent runs with identical properties may be split. Relationships compare their resolved targets,
including the case where unchanged XML's rIds were redirected. This is a deliberately conservative
comparison, not a complete Word layout or Excel calculation oracle. The separate loss audits retain
their historical feature inventories; their new `accounting` field exposes missing saved files and
failed comparisons rather than silently excluding them.

`node tools/ooxml/namespace-size.mjs ORIGINALS SAVED OUTPUT [LIMIT=400]` isolates namespace compaction
on existing Word saves. It reports original/before/after main-part sizes and compacts writer-owned
parts while keeping opaque dependency bytes. Descendant `Ignorable` declarations required by local
`PreserveAttributes`, `PreserveElements` or `ProcessContent` are counted as `scopedIgnorable`;
only `unexpectedIgnorable` declarations fail the measurement. It is a serialization benchmark, not an app round trip;
use fresh browser saves, package/SDK comparisons and Office samples for integration checks. Run
`conventions.py` on its output and every full corpus run, recording newly introduced signatures.
`tests/quire/comment-office-copies.py BATCH PREVIOUS_BATCH OUTPUT` recreates the bounded first
Word comment-repair isolation round. Package checks now reject excess comment references beyond
their matching definitions; comparison also rejects automatic pre-release thread-namespace promotion.
They also reject a shape ID repeated within a PowerPoint VML part or across spreadsheet VML parts (`duplicate-vml-id`), two
embedded objects sharing one preview (`shared-vml-preview`), two spreadsheet VML parts sharing an
`o:idmap` block (`duplicate-vml-block`) and diagram/chart/OLE part animations whose target is not a
graphic frame (`animation-part-target`). None occurs in the 4,697 pinned originals.

Ledger's `pivots.test.js` and `threads.test.js` use the pinned public inputs (`VO_CORPORA` overrides
`/work/corpora`). `PIVOT_RESULTS=DIR` / `THREAD_RESULTS=DIR` retain the emitted edit/history states and
their source paths for package/SDK comparisons. Pivot source fixtures are generated automatically from
the pinned input into `PIVOT_RESULTS/fixtures` (or a temporary directory); `PIVOT_FIXTURES` overrides
that destination. They do not depend on an earlier results directory. The thread test also authors a mention/no-legacy-note
fixture missing from the public corpus. `node tests/ledger/threads-ui.mjs INPUT.xlsx OUTPUT`
checks the real Paste Special, undo/redo and suite draft/recovery hooks on a thread in the first sheet.
Add `--edit` to first edit its root through the actual comment editor; replies and untouched mentions
must survive, including edits around both ends of a mentioned name and later clipboard/draft recovery.
`objects.test.js` uses `OBJECT_RESULTS=DIR` to retain controls/OLE/VML edit states from the same corpus.
`node tests/ledger/objects-ui.mjs INPUT.xlsx OUTPUT` checks object paste into another workbook,
undo/redo and actual draft recovery, including the dependent payload bytes.
`sh tools/corpora.sh /work/corpora ledger-features` fetches three supplemental slicer/timeline workbooks
at fixed commits and SHA-256 hashes without changing the main corpus manifest. `slicers.test.js`
uses them (`LEDGER_FEATURES` overrides their directory); `SLICER_RESULTS=DIR` retains edit states.
`node tests/ledger/slicers-ui.mjs INPUT.xlsx OUTPUT` exercises real drawing copy/paste, undo/redo
and draft recovery. The supplemental sources and hashes are recorded in `ledger-features/manifest.json`.
`extensions.test.js` uses `EXTENSION_RESULTS=DIR` for extended bars, validation, sparklines and unknown
worksheet extension edit/history saves (`EXCEL_CORPUS` overrides `/work/corpora/excel`).
`node tests/ledger/extensions-ui.mjs INPUT.xlsx OUTPUT` checks an extended bar through edits,
sheet copy, undo/redo and actual draft recovery.
`tables.test.js` uses `QUERY_RESULTS=DIR` for query-table edits, column identities, copies and
history saves (`EXCEL_CORPUS` overrides `/work/corpora/excel`).
`node tests/ledger/tables-ui.mjs INPUT.xlsx OUTPUT` checks inserted query columns, sheet
copy, scoped names, undo/redo and actual draft recovery.
`recovery.test.js` covers readable cells in damaged packages and unchanged mixed date/text filters.
`sheet-properties.test.js` uses `SHEET_PROPERTY_RESULTS=DIR` to retain automatic row-height,
dimension-edit, sheet-copy and history saves from public workbooks. It also checks literal boolean
view flags and compatibility alternatives around worksheet-format properties.
`page-setup.test.js` uses `PAGE_SETUP_RESULTS=DIR` to retain absent-default, paper/quality-edit,
printer-dependency, chart-sheet, copy and history checks from public workbooks.
For the independent cell audit after `corpus.js --scenarios text`, set `SCENARIO=text` when running
`tests/ledger/loss-audit.py`: only A1 on the first ordinary worksheet is excluded as the intended
edit. Dependent array-result recalculation still appears in the report. ZIP directory markers are
excluded from feature-part counts.

`node tests/ledger/scroll-ui.mjs` scrolls a sheet with the mouse wheel in all four directions
and with the arrow keys, with and without frozen panes, letting the page repaint between steps.

Build the SDK validator and install the independent workbook reader in a test environment:

```sh
dotnet build tools/ooxml/oxval -c Release
python3 -m venv /tmp/vo-ooxml-venv
/tmp/vo-ooxml-venv/bin/python -m pip install -r tools/ooxml/requirements.txt
dotnet tools/ooxml/oxval/bin/Release/net8.0/oxval.dll --baseline original.docm saved.docm
```

The validator is pinned to SDK 3.1.1, defaults to Office 2019, handles all document/macro/template/slideshow
extensions and reports uncapped diagnostics by part, XPath and diagnostic ID. `--baseline` reports new
diagnostic identities, not merely a change in count. Style diagnostics use `styleId` rather than an
unstable style-list index. Unmatched diagnostics can also match identical namespace-expanded node
content within the same part; each original occurrence is consumed once, so duplicating an invalid
property still adds a diagnostic. The actual XPath and content fingerprint remain in the report.
`--comparison-self-test` checks movement, duplication, changed content and changed errors.
Use `--version Microsoft365` before other arguments to validate modern presentation comments;
Office 2019 does not validate those newer parts. The batch wrapper accepts `--sdk-version Microsoft365`.
Package diagnostics normalize Strict/
Transitional namespace aliases and compare duplicate IDs by scope, identity and multiplicity.
`--stdin-pairs` accepts JSON lines containing
`original` and `saved` paths for a warm batch process.

`python3 tools/ooxml/validate.py RESULTS.jsonl ORIGINALS SAVED_ROOT OUT.jsonl --sdk PATH_TO_DLL`
combines package checks and SDK comparison, accounting for unsuccessful driver attempts as well.
It validates every emitted artifact, including before-edit, undo, redo, second-save and recovered
draft states. A diagnostic comparison requires a validatable original; standalone validation of
repaired output is reported separately when its original cannot be opened by the SDK.
Reusing an earlier SDK comparison requires an unchanged SDK/version, the same original SHA-256,
and byte-identical content for every uncompressed member of the saved package. Record the source
report and equivalence proof, rerun package checks if their rules changed, and validate all other
states normally. Preserve fresh failures/timeouts in the raw reports even when an identical artifact
has earlier successful independent evidence; disclose that reuse in the run README.
`--dotnet PATH` selects the SDK runtime. `--lo PDF_DIRECTORY` additionally converts original and saved
files using LibreOffice with its own temporary profile, checks conversion success and extracts page
counts with `pdfinfo`. Unedited saves must keep the page count; edit scenarios can record
`expectedPages` or `expectedSlides`. LibreOffice and the validator cannot certify acceptance by Office itself.

For a full rendering pass over already saved artifacts, use
`python3 tools/ooxml/render.py MANIFEST.jsonl OUTPUT --jobs 4`. Each manifest row names `app`, `file`,
`scenario`, `original`, `saved` and the driver `status`; `originalSha256`/`savedSha256` pin the inputs.
The optional `expectedPages` overrides comparison with the original PDF count. The tool uses warm,
isolated LibreOffice processes through the distribution's Python UNO bindings (`--uno-python`, default
`/usr/bin/python3`), with read-only loading, no macro/link updates and a timeout per document.
It stores PDFs, page counts and every failed/excluded attempt outside the repo. Exact member-byte
fingerprints plus filenames and rendering-environment identities deduplicate history/draft states;
changed members cannot reuse a render. Raw hashes identify malformed or oversized packages.
`--resume` requires the same manifest/environment and verifies pinned input hashes before using
completed rows. Rendering failures and page-count differences remain separate from driver and SDK
results; a successful PDF export alone is not a lossless-save claim.
Failed conversions retire their worker. A lost worker connection gets one fresh-process retry,
with the first failure retained in `workerAttempts`; timeouts and document errors are not retried.
After reviewing a harness-only change that leaves loading/export unchanged, a fresh output may use
`--reuse-successes PRIOR_OUTPUT --reuse-reason "reviewed change and rationale"`. This imports only
successful PDFs with verified source/package/PDF hashes and identical LibreOffice, fonts, limits
and rendering policy. Each reused record identifies the old engine and cache entry. Failed attempts
stay in the original report and are attempted again; never publish worker-cascade errors as app failures.

### Lectern preservation

`node tools/ooxml/conversions.mjs OUTPUT_DIR [SCENARIO_REGEX]` checks conversion notices on public
Word/Excel/PowerPoint files and verifies that reverting an edit removes its save-time notice.
`node tools/ooxml/save-matrix.mjs OUTPUT_DIR --checker-only` checks silent drafts, recovery,
cancellation, download failure, successful acknowledgement, newly introduced notices and recorded-only
entries that must save without the dialog, in all apps.
Feature inventories decode UTF-16 parts and count retained SmartArt text on both sides.

`node tests/lectern/properties.mjs OUTPUT_DIR [SCENARIO_REGEX]` checks per-property edits,
tags, actions, transitions, animation changes, clipboard dependencies, undo/redo and actual drafts.
The corpus driver also records original/opened/saved slide counts so a skipped source slide cannot
pass merely because the remaining package is valid.

`tests/lectern/designs.mjs OUTPUT [SCENARIO_REGEX]` checks imported design edits, independent
copies, unused-master retention, history and draft recovery. It checks that unrelated layouts
retain their exact bytes, theme effects retain their expanded XML, and new layouts allocate unique
shape IDs including the structural group. It also covers interleaved unread master/layout shapes,
stacking order, deletion, addition and edited default run properties. Reports follow the same
package/SDK and LibreOffice comparison workflow.

`node tests/lectern/frames.mjs OUTPUT_DIR [SCENARIO_REGEX]` exercises opaque SmartArt, OLE,
ink, 3-D and chartEx previews with geometry/content edits, copies, deletion, ungrouping, clipboard
renditions and drafts. Every saved state must have unique VML preview IDs and part animations aimed
at graphic frames (an embedded object copied twice; edited SmartArt animated by parts). Run
`python3 tests/lectern/frame-fixtures.py FIXTURE_DIR` first for its
multiple-member fallback case; `FRAME_FIXTURES` overrides the default sibling `frames-fixtures/`
directory. The fixture is derived from a public SDK chartEx input, with explicit provenance.

`python3 tests/lectern/comment-fixtures.py FIXTURE_DIR` creates a documented schema-authored
modern thread/task/anchor sample on a pinned public deck; it is not an Office-authored input.
`node tests/lectern/comments.mjs OUTPUT_DIR FIXTURE_DIR [SCENARIO_REGEX]` checks original
comment bytes, slide/shape ownership, author/index and GUID references, detached clipboard packages,
Strict copies, undo/redo and suite drafts. Validate its `results.jsonl` using the Microsoft365 target.
The scenario filter retains other recorded cases so corrections need not repeat unaffected work.

`node tests/lectern/membership.mjs OUTPUT_DIR` uses pinned presentation inputs for section and
custom-show insertion, copy, move, deletion, large IDs, empty sections, undo/redo and actual suite
draft recovery (`VO_CORPORA` overrides `/work/corpora`). It retains every saved state and a draft screenshot.
Use its `results.jsonl` with `tools/ooxml/validate.py` for package/SDK and LibreOffice comparison.

`node tests/lectern/media.mjs FILE_OR_LIST OUTDIR` opens authored video/audio decks in Chromium.
It saves unedited, moves/resizes, changes shadow and crop, duplicates, deletes, replaces, duplicates
the slide, recovers an actual suite draft and copies through keyboard/native clipboard events into
another tab. Every mutation that uses history emits its pre-edit, undo and redo state.

`python3 tests/lectern/media-check.py ORIGINALS OUTDIR` independently compares media frames,
bytes, reference targets, playback and exact history restoration (core timestamps excepted).
`python3 tests/lectern/media-check.py --sounds ORIGINALS RESULTS.jsonl SAVED_ROOT` compares
transition/click sounds and playback through ordinary save/repeat/text/draft corpus scenarios,
including decks without a media picture. `media-variant.py SOURCE.pptx OUT.pptx` derives an
AlternateContent/effects fixture from an authored media deck (`--strict` adds a Strict case with an
existing shadow to replace); keep it outside the repository and
validate the fixture itself before testing. Feed each driver's report to `validate.py` for all-state
package and SDK checks.

### Package preservation and save flows

`node tools/ooxml/save-matrix.mjs OUTDIR` exercises the actual app Save and Save As dialogs,
download cancellation, all 14 variants, MIME/main-part types, encrypted draft recovery and the
Compatibility Checker. It intercepts the final download handoff and disables recent-file bookkeeping
in its own tabs. Fixtures for the format matrix are synthetic blank documents; authored-file fidelity
is measured separately. `--macro-files FILE.json` accepts an app-to-file mapping for authored VBA
samples and checks cancelled and approved macro-free Save As. `--checker-only` runs just the checker
and pending-loss recovery checks.

`python3 tools/ooxml/package-carry.py ORIGINALS SAVED RESULTS.jsonl OUT.jsonl` independently verifies
selected package features, including dependency bytes, content types, original rIds and intended
targets. It stops at model-owned boundaries and does not treat a preserved pivot part as proof that
pivot editing works. Driver failures and malformed originals remain explicit failed/excluded rows.
Corpus writers select the source main-part variant, including inputs with a misleading extension. The presentation
inventory parses transition/run-effect elements and text as XML: namespace hoisting, self-closing
properties and escaped apostrophes must not appear as feature or word loss. Edited SmartArt previews
are intentional conversions and must be matched with their recorded loss entries.

`python3 tests/ledger/pivot-fixtures.py /work/corpora/excel OUTDIR` authors the two pivot source kinds
absent from the pinned inventory: a defined-name source and a two-area consolidation source. It keeps
all unchanged package parts byte-identical, records source/output hashes and changed parts, and
requires an empty destination. Consolidation input cells, cached records and displayed output agree.
These are structural fixtures; package/SDK validity does not establish Office acceptance or Ledger's
edit behavior. Use them alongside the public range, table, external-source and shared-cache workbooks.

### Quire controls, properties and objects

`node --test tests/quire/*.test.js` covers control boundaries and locks, independent property
replacement, alternatives, range/object identity and history. `node tests/quire/controls.mjs
/work/corpora/word OUTDIR` edits authored bound-text, date and checkbox controls, emitting before/edit/undo/
copy/draft/recovered files. `node tests/quire/objects.mjs /work/corpora/word OUTDIR` transfers authored
OLE/SmartArt objects between browser tabs, edits converted content, undoes it and recovers a draft.
`node tests/quire/sections.mjs /work/corpora/word OUTDIR` checks section XML, absent defaults,
header relationships and alternatives through actual saves, margin edits, undo/redo and draft recovery.
Its generated fixture also exercises the page-number dialog, section insertion, page setup/columns
forward at new and existing boundaries, multi-column index insertion and merged letters. Add
`--fixture` to run only these focused cases, including copied revision IDs through history and drafts.
It writes the original and each saved state for package/SDK and external rendering comparisons.

`node tests/quire/picture-bullets.mjs /work/corpora/word OUTDIR` exercises the public SDK picture-bullet
fixture through the list dialog, symbol conversion, undo/redo, actual draft recovery and an import
with colliding source IDs. It compares retained definition XML and image bytes and checks unused
definitions after all list references have been removed.

`node tests/quire/numbering.mjs /work/corpora/word OUTDIR` exercises the public custom-format and
style-linked numbering fixtures through list commands, independent copies, repeated imports,
undo/redo and actual draft recovery. An optional final `custom` or `style` selects one fixture.
The loss audit also inventories picture bullets, list identities/templates, legacy settings and
cleanup metadata; the numbering harness compares metadata values and original level counts.
It also checks fresh `nsid` identities for independent copies. `index-sections.mjs OUTDIR` exercises
multi-column index insertion in final and non-final sections, including page starts, undo/redo and
actual draft recovery.

`tools/ooxml/prescan-bench.mjs` measures the shared writer's identity scan on a generated large sheet
containing ordinary words such as “del”; it reports elapsed time and heap growth separately from
whole-file save time. `tests/ooxml/lectern-review.test.mjs` checks chart rotation policy, placeholder
flips, single-member alternative orientation and comment-extension order. The browser harness
`tests/lectern/graphic-rotation.mjs OUTDIR` exercises actual chart/OLE commands and ungrouping.
`tests/ledger/pictures.test.js` checks retained picture properties through geometry/crop edits,
copying, deletion, history and recovery. The package checker additionally rejects orientation on
PowerPoint graphic-frame transforms and inconsistent ActiveX worksheet/VML names.

`node tests/quire/recovered-breaks.mjs /work/corpora/word OUTDIR` checks the public malformed-break
fixture through text edits, deletion, undo/redo and actual draft recovery. It verifies valid saved
break placement and that the Compatibility Checker follows surviving recovered content.

`node tests/quire/drawing-properties.mjs /work/corpora/word OUTDIR` checks authored picture/shape
properties and watermarks, plus generated row/cell alternatives with populated and empty choices.
It writes originals and every edit/history/draft state, checking that deletion cannot restore hidden
table content and that original watermark elements remain byte-identical before editing.
Use the resulting JSONL state rows with `validate.py`; it validates every emitted state.
The Office batch generator supports Quire edits `picture-properties`, `shape-shadow`, `drawing-anchor`
and `watermark`, so review copies and previous-writer comparisons use the same scripted edits.

Current results belong in the app docs. Preserve complete run commands, hashes, previous-writer
comparisons, initial failures, corrections and historical tables in `/work/corpora/results/<run>/README.md`.
A full run should use a frozen writer revision, served on a separate port while development continues.
The browser harness disables persistence in its own tabs, captures loss metadata after snapshot, and
checks the recovered document. Locked content-control edits are explicit exclusions, not undo failures.

`python3 tools/ooxml/samples.py /work/corpora /tmp/lossless tmp/lossless/baseline`
selects small independently authored files for each carried-feature test, copies original/saved/edited
variants where available, and writes a manifest. It refuses to replace differing handoff files.
Use a new destination for each implementation stage. Record actual Word/Excel/PowerPoint repair
results; an unperformed Office check remains pending.

`node tests/lectern/timing-repair.mjs OUTDIR --sdk PATH_TO_DLL` reproduces the slide-6
diagnostic copies from the current sample. `--source SAVED_SAMPLE.pptx` instead uses the exact bytes
of a reported failure; `--dotnet PATH` selects the runtime. The output directory must be empty.
The variants isolate removal of slide-6 timing, its build list, its teeter effect, and ordering of
timing IDs; from the built-in sample it also writes the deck without slide 6 and slide 6 alone. The LibreOffice resave is a separate control and does not fail the run. The script asserts build-pair identity and
text targets, validates each package, and writes hashes and a pending PowerPoint result per copy.
Only slide6.xml changes in the four isolated variants. PowerPoint results are a manual gate:
schema validity and LibreOffice opening do not prove that the repair problem has been fixed.
