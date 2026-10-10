# VibeOffice

A static web office suite in the style of Office 2003 that opens and saves real Office Open XML files, installable as a PWA. `public/` is the site root: one folder per app, and `public/index.html`, the Start Center (as in LibreOffice: Open File, Recent Files with thumbnails, Create File from Blank or a template, Apps: Quire, Ledger, Lectern). Its About text stays general, never naming another company's product.

| App | Folder | Kind | Files |
|---|---|---|---|
| Quire | `public/quire/` | Word processor | .docx/.docm/.dotx/.dotm, RTF, HTML, text, Markdown (GitHub flavour; .zip with pictures); PDF out |
| Ledger | `public/ledger/` | Spreadsheet | .xlsx/.xlsm/.xltx/.xltm, CSV/text, XML Spreadsheet 2003, HTML; PDF out |
| Lectern | `public/lectern/` | Presentations | .pptx/.pptm/.ppsx/.ppsm/.potx/.potm; PDF, PNG, HTML out |

Each app is classic `<script>` files (no ES modules) that attach to one global namespace, `window.L` (also `window.Quire` / `window.Ledger` / `window.Lectern`), loaded in the order listed in its `index.html`. `public/common/` holds what the apps share: the libraries and the stylesheets (one copy each: core, ui, icons, zip, sha, crypto, geometry, metafile, charts, numfmt, dml, spell engine, clipart, SmartArt layouts, ribbon, search, `looks.css`, `ui.css`), `suite.js` (`window.VO`: launching apps, Recent Files, the service worker) and the proofing dictionaries (`common/dict/`). See docs/suite.md.

## Docs

| | |
|---|---|
| [docs/quire.md](docs/quire.md) | Quire: features, proofing, Word layout rules, compatibility results, source layout |
| [docs/ledger.md](docs/ledger.md) | Ledger: features, formulas and calculation, compatibility results, source layout |
| [docs/lectern.md](docs/lectern.md) | Lectern: source layout, opening, saving, known limits |
| [docs/suite.md](docs/suite.md) | Start Center, `common/suite.js` (app launch, Recent Files), PWA manifest, service worker |
| [docs/testing.md](docs/testing.md) | Test workflow and tools |
| [docs/LOSSLESS_SAVE.md](docs/LOSSLESS_SAVE.md) | Same-format save: design summary, what was done, working rules, to do |

Keep them current: a change to an app's features, formats or source layout updates its `docs/<app>.md`. No per-app READMEs. User-facing docs (README, `docs/<app>.md`) send people to https://vibeoffice.work; running it locally on 127.0.0.1 belongs in docs/testing.md. Screenshots for public docs show no local recent files or drafts (close the task panes).

## Design rules

- **Important: interoperability with Microsoft Office.** A file saved here opens in Word, Excel and PowerPoint without a repair prompt, with everything the user did not change intact and their edits as they made them. The UI uses Office's names (fonts, features, file types), so it matches what Office shows, except for brand-like feature names, which get our own: **IntelliArt** for SmartArt and **LettersArt** for WordArt. The code keeps Office's names everywhere (identifiers, file names, comments, command ids, model types such as `wordart` and `smartart.js`), so it maps directly onto the file format; docs say SmartArt or WordArt when they mean the file format.
- Static files, vanilla JavaScript, no frameworks and no build step at run time. Ledger's `index.html` is generated: edit `tools/ledger/build/` and run `python3 tools/ledger/build/make.py`, never the output.
- Everything runs in the browser; documents never leave the machine. No back end, no accounts.
- Nothing from third-party servers: no CDNs, web fonts or analytics; vendored files go in `public/common/`, fonts in `public/fonts/` (`tools/fonts.py`), word lists in `public/common/dict/` (`tools/dict.py` for British English), each with their licences.
- Compatibility is measured against files other people wrote (public test corpora), compared with outside references (the page or cell values the authoring application stored, LibreOffice, openpyxl, python-pptx), and the result is written down in the app's doc with numbers.
- Saved files must open in the authoring application's newer versions and in LibreOffice; what an app cannot edit it keeps and writes back unchanged.
- Follow the [Preservation contracts](docs/suite.md#preservation) for OOXML ownership, identities, edits, history/copy, drafts and loss reporting.
- Chrome colours come only from the look settings in `public/common/looks.css` (Classic, Paper, Paper 2016; docs/suite.md, Looks), never hard-coded; document content keeps its own colours in every look. A new setting goes into every look (`node --test tools/looks.test.mjs`).
- A new command also goes into the app's ribbon (`js/ribbon.js`, the Paper 2016 look), where Office 2016 has it (`node --test tools/ribbon.test.mjs`).
- Shared code lives once, in `public/common/`; never copy it into an app. A change there is checked in all three apps (load each, and screenshot what it touches). What stays app-specific goes through the hooks in docs/suite.md (app config, `ui.hooks`, `L.icons.app`, …), not through app names inside common code.

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
- Chromium with CDP at 127.0.0.1:9222. `tools/shot.mjs <page> [steps]` drives a page over CDP, takes screenshots and reports console errors and any request that leaves 127.0.0.1. **Never close the browser's other tabs** (Chromium quits with its last tab).
- `tools/templates.mjs`: regenerates the Start Center's template gallery (`public/common/templates.json` and previews) from the apps; run it after changing a template.
- `tools/ledger/test/`: Ledger's Node harnesses (unit tests, `.xlsx` round trip, recalculation, CSV) and Playwright UI checks.
- `tools/lectern/test/`: Lectern's real-deck harness (open, render, save, reopen in Chromium), UI checks for passwords and charts, encrypted-save check (see docs/lectern.md).

## Working with other sessions

Several agents work in this repository at once. Commit only your own files and hunks (stage hunks when a file also holds someone else's edits); never commit, revert or rewrite another session's uncommitted changes. Leave a finding for the owner in the relevant plan's To do (preservation: docs/LOSSLESS_SAVE.md) instead of fixing inside their in-progress files.

## Testing

1. Ledger engine changes: `node --test tools/ledger/test/*.test.js`.
2. File-format changes: round-trip real files (`node tools/ledger/test/roundtrip.js file.xlsx`) and open the saved file in LibreOffice.
3. Look: `node tools/shot.mjs <app>/` with steps for the change; check every app still loads without console errors or external requests. UI checks pause between steps so the page repaints, as with real input: bursts of synthetic events hid a Ledger scroll bug.
4. Shared save code (`public/common/opc.js` and the writers) runs on every part of every file: time a change on a large file (a 30 MB sheet) as well as on the corpus. A regex pre-scan or a full parse there costs seconds and gigabytes.
5. Prefer a focused regression test per fix; run the full corpus once per feature group and before each Office batch, not after every fix.
