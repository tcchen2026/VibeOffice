# VibeOffice

A static web office suite in the style of Office 2003 that opens and saves real Office Open XML files, installable as a PWA. `public/` is the site root: one folder per app, and `public/index.html`, the Start Center (as in LibreOffice: Open File, Recent Files with thumbnails, Create File from Blank or a template, Apps: Quire, Ledger, Lectern). Its About text stays general, never naming another company's product.

| App | Folder | Kind | Files |
|---|---|---|---|
| Quire | `public/quire/` | Word processor | .docx/.docm/.dotx/.dotm, .doc/.dot in (converted), RTF, HTML, text, Markdown (GitHub flavour; .zip with pictures); PDF out |
| Ledger | `public/ledger/` | Spreadsheet | .xlsx/.xlsm/.xltx/.xltm, .xls/.xlt in (converted), CSV/text, XML Spreadsheet 2003, HTML; PDF out |
| Lectern | `public/lectern/` | Presentations | .pptx/.pptm/.ppsx/.ppsm/.potx/.potm, .ppt/.pps/.pot in (converted); PDF, PNG, HTML out |

Each app is classic `<script>` files (no ES modules) that attach to one global namespace, `window.L` (also `window.Quire` / `window.Ledger` / `window.Lectern`), loaded in the order listed in its `index.html`. `public/common/` holds what the apps share: the libraries and the stylesheets (one copy each: core, ui, icons, zip, sha, crypto, geometry, metafile, charts, numfmt, dml, spell engine, clipart, SmartArt layouts, ribbon, search, `looks.css`, `ui.css`), the assistant (Scout, `common/scout/`), `suite.js` (`window.VO`: launching apps, Recent Files, the service worker) and the proofing dictionaries (`common/dict/`). See docs/suite.md.

## Docs

| | |
|---|---|
| [docs/quire.md](docs/quire.md) | Quire: features, proofing, Word layout rules, compatibility results, source layout |
| [docs/ledger.md](docs/ledger.md) | Ledger: features, formulas and calculation, compatibility results, source layout |
| [docs/lectern.md](docs/lectern.md) | Lectern: source layout, opening, saving, known limits |
| [docs/suite.md](docs/suite.md) | Start Center, `common/suite.js` (app launch, Recent Files), PWA manifest, service worker |
| [docs/desktop.md](docs/desktop.md) | Desktop app (Tauri; Windows with mingw, macOS: the DMG and the Mac App Store, Linux): host interface (`VO.host`), building, installer and file types, signing, CI, size, to do |
| [docs/testing.md](docs/testing.md) | Test workflow and tools |
| [docs/i18n.md](docs/i18n.md) | The UI in other languages: choosing it, `L.t`, the English catalog, the pseudo-locale, translations, defaults and fonts per language, adding a language |
| [docs/LOSSLESS_SAVE.md](docs/LOSSLESS_SAVE.md) | Same-format save: design summary, what was done, working rules, to do |
| [docs/smartart.md](docs/smartart.md) | IntelliArt (SmartArt): how PowerPoint draws our definitions, what was learned, to do |

Keep them current: a change to an app's features, formats or source layout updates its `docs/<app>.md`. No per-app READMEs. User-facing docs (README, `docs/<app>.md`) send people to https://vibeoffice.work; running it locally on 127.0.0.1 belongs in docs/testing.md. Screenshots for public docs show no local recent files or drafts (close the task panes).

## Design rules

- **Important: interoperability with Microsoft Office.** A file saved here opens in Word, Excel and PowerPoint without a repair prompt, with everything the user did not change intact and their edits as they made them. The UI uses Office's names (fonts, features, file types), so it matches what Office shows, except for brand-like feature names, which get our own: **IntelliArt** for SmartArt and **LettersArt** for WordArt. The code keeps Office's names everywhere (identifiers, file names, comments, command ids, model types such as `wordart` and `smartart.js`), so it maps directly onto the file format; docs say SmartArt or WordArt when they mean the file format.
- Static files, vanilla JavaScript, no frameworks and no build step at run time. No inline scripts, `on…=` handlers, `eval` or `new Function`: every page's Content-Security-Policy runs only our own script files (`tools/csp.mjs`; docs/suite.md, Security). Ledger's `index.html` is generated: edit `tools/ledger/build/` and run `python3 tools/ledger/build/make.py`, never the output.
- Everything runs in the browser; documents never leave the machine. No back end, no accounts.
- Chromium is the tier-1 browser: our layout engine always aligns to Chromium first (pages, lines and positions measured there against Office); other browsers just need to be functional.
- Nothing from third-party servers: no CDNs, web fonts or analytics; vendored files go in `public/common/`, fonts in `public/fonts/` (`tools/fonts.py`), word lists in `public/common/dict/` (`tools/dict.py` for British English), each with their licences.
- Code from other projects: no MPL, GPL, LGPL or AGPL code (LibreOffice, OnlyOffice, Gnumeric, Hunspell, …), never copied, ported or rewritten in JavaScript; LibreOffice is still run as a program to check interoperability. An LGPL library may be used as a dynamically linked shared library (.so, .dll, .dylib), unmodified and replaceable, with its licence listed (GTK and WebKitGTK on Linux), never linked statically or compiled into the app. Apache 2.0 code may be read, reused and rewritten in JavaScript: Apache OpenOffice is the reference, a shallow clone in `refs/openoffice/` (`git clone --depth 1 https://github.com/apache/openoffice.git refs/openoffice`; ignored), and LibreOffice's changes to the same code after the fork are MPL. Permissive code (MIT, BSD, zlib) may be kept with its notice, and CC BY material (for example Microsoft's documentation) with credit. No CC BY-SA material either (Stack Overflow answers, Wikipedia text and code), nor anything under the Numerical Recipes licence; an idea or algorithm described there may be implemented in our own words from the maths or a standard. Scan for restrictive licences such as MPL, GPL, LGPL, AGPL and CC BY-SA. Derived code says so in the file (`Derived from <project> <path> (<licence>)`) and is credited in `tools/oss-licenses.py`; todos/provenance.md (private, not committed) records each source.
- Compatibility is measured against files other people wrote (public test corpora), compared with outside references (the page or cell values the authoring application stored, LibreOffice, openpyxl, python-pptx), and the result is written down in the app's doc with numbers.
- Saved files must open in the authoring application's newer versions and in LibreOffice; what an app cannot edit it keeps and writes back unchanged.
- Follow the [Preservation contracts](docs/suite.md#preservation) for OOXML ownership, identities, edits, history/copy, drafts and loss reporting.
- UI text is written in English in the code; its translations into Japanese, Korean, French, German, etc. are in `public/common/i18n/<lang>.json` (docs/i18n.md). Translating is a task of its own, done only when the user asks for it; in any other work agents never read, write, edit or check those files or their wording, and don't translate UI text. A new or changed UI string goes through `L.t()` (or a helper that calls it) and into the English catalog with `node tools/i18n.mjs gather` (`node --test tests/i18n.test.mjs`); a string with no translation shows in English. Check layout and missed strings with the pseudo-locale (`?lang=qps`), not with real translations.
- Chrome colours come only from the look settings in `public/common/looks.css` (Classic, Paper, Paper 2016; docs/suite.md, Looks), never hard-coded; document content keeps its own colours in every look. A new setting goes into every look (`node --test tests/looks.test.mjs`).
- A new command also goes into the app's ribbon (`js/ribbon.js`, the Paper 2016 look), where Office 2016 has it (`node --test tests/ribbon.test.mjs`).
- Shared code lives once, in `public/common/`; never copy it into an app. A change there is checked in all three apps (load each, and screenshot what it touches); `common/zip.js` also runs in the service worker (`public/sw.js`), so a change to it is checked with `node --test tests/sw.test.mjs` too. What stays app-specific goes through the hooks in docs/suite.md (app config, `ui.hooks`, `L.icons.app`, …), not through app names inside common code.

## Office acceptance (lessons from the slide-6 repair)

- Schema-valid is not Office-valid: the Open XML SDK and LibreOffice passed decks PowerPoint repaired (a preset's partial adjustment list; a master and layout sharing an ID). Write what Office writes, not merely what the schema allows; when unsure, look at how Office-authored files in the corpus do it.
- Every Office-found defect becomes a `tools/ooxml/package.py` rule with a test, plus a corpus scan for how widespread it is.
- Before handing files to the user for Office, diff the new save against the previous writer's save of the same input; unexplained differences are suspects. Two overlapping faults made one Office round misleading.
- New IDs and names (list, shape, VML, slicer, control, creation IDs) follow what Office-authored files in the corpus do, not a local max+1 or a made-up suffix.
- A fix's regression test must fail on the previous commit; run it there before committing.
- A notice says exactly what happened: "converted" only if the value was converted, otherwise name what was dropped. Show it to the user (`notify: true`) only when they lose content or a working feature, in plain words about their content with a place they recognise; everything else is recorded silently for tests (docs/suite.md, What the user is told).
- Narrow Office failures by bisection: copies that each change one thing, with whole-file controls (without the suspect part, the part alone), at most about 8 per round, each with a checklist.

## Tools (details in docs/testing.md)

- `tools/serve.py`: no-cache server for `public/` on 127.0.0.1:8760.
- Chromium with CDP at 127.0.0.1:9222. `tools/shot.mjs <page> [steps]` drives a page over CDP, takes screenshots and reports console errors and any request that leaves 127.0.0.1; a refusal of the pages' policy, a window opened or a page of another site fails it. **Never close the browser's other tabs** (Chromium quits with its last tab).
- `tools/quick.mjs [REV]`: the quick check after a change (the tests covering the files changed); `tools/ui-full.mjs`: the whole UI suite before a release, recorded for `tools/ui-tested.sh`, which `tools/web-release.sh` and `../cf-vibeoffice/sync-cf-pages.sh` run.
- `tools/older.mjs [REV] -- <tests>`: runs today's tests against `public/` as it was at REV (default: before the uncommitted or last change), so a fix's regression test can be seen failing there.
- `tools/oss-licenses.py`: regenerates `public/oss-licenses.html` (Open source licenses, linked from About) from the license files in `public/`, the desktop app's direct crates and its DLL; run it after changing a font, a word list or a desktop dependency.
- `tools/templates.mjs`: regenerates the Start Center's template gallery (`public/common/templates.json` and previews) from the apps; run it after changing a template.
- Tests live in `tests/`, not in `tools/`: `tests/*.test.mjs` (looks, ribbon, search, SmartArt, Scout, the service worker), `tests/ui/` (real mouse and keyboard input in each app, `smoke.test.mjs` the quick check: drags, selection and focus, run through the starting states, keys in the middle, ways out and modes of `tests/lib/gestures.mjs`; an item not fixed yet stays a `todo` test), `tests/ooxml/`, one folder per app, `tests/from-issues/N/` for the regression tests of reported issue N, and `tests/hostile/`, documents written to attack the suite (what they put into exported HTML).
- `tests/ledger/`: Ledger's Node harnesses (unit tests, `.xlsx` round trip, recalculation, CSV) and Playwright UI checks.
- `tests/lectern/`: Lectern's real-deck harness (open, render, save, reopen in Chromium), UI checks for passwords and charts, encrypted-save check (see docs/lectern.md).

## Working with other sessions

Several agents work in this repository at once. Commit only your own files and hunks (stage hunks when a file also holds someone else's edits); never commit, revert or rewrite another session's uncommitted changes. Leave a finding for the owner in the relevant plan's To do (preservation: docs/LOSSLESS_SAVE.md) instead of fixing inside their in-progress files.

Always work in the repository's own working directory: no git worktrees or subtrees, no copies of the repository to work in.

Don't look up this project's GitHub issues (no `gh`, no fetching issue pages) unless the user pastes an issue URL; then read only that issue.

## Testing

1. Ledger engine changes: `node --test tests/ledger/*.test.js`.
2. File-format changes: round-trip real files (`node tests/ledger/roundtrip.js file.xlsx`) and open the saved file in LibreOffice.
3. Look: `node tools/shot.mjs <app>/` with steps for the change; check every app still loads without console errors or external requests. UI checks pause between steps so the page repaints, as with real input: bursts of synthetic events hid a Ledger scroll bug. After each change, `node tools/quick.mjs`: the tests that cover the files changed (always the smoke test `tests/ui/smoke.test.mjs` when the apps changed, under a minute; a touched app's Node and UI tests; seconds to a few minutes), with the change's own test also run on the code before it (`tools/older.mjs`). The whole of `tests/ui/` (real input, about 25 minutes), `node tools/ui-full.mjs`, before a release: no new failures, and the `todo` count only goes down; it records its result, and the release scripts warn when it has not passed on the `public/` they publish (`tools/ui-tested.sh`).
4. Shared save code (`public/common/opc.js` and the writers) runs on every part of every file: time a change on a large file (a 30 MB sheet) as well as on the corpus. A regex pre-scan or a full parse there costs seconds and gigabytes.
5. Don't overtest: use your own judgment on whether a change needs a test at all. Add one where a fault could come back unnoticed or the logic is subtle; a small, obvious fix (a key cancelled, a guard added) is checked once and needs no test of its own. UI tests are slow: prefer a Node test or a one-off `tools/shot.mjs` check, and add a UI test only for what needs real input (focus, drags, keys the browser acts on). Run what the change touches once, not again to reconfirm; the full corpus once per feature group and before each Office batch, not after every fix.
