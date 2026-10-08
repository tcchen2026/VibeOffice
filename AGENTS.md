# VibeOffice

A static web office suite in the style of Office 2003 (Luna Blue) that opens and saves real Office Open XML files, installable as a PWA. `public/` is the site root: one folder per app, and `public/index.html`, the Start Center (as in LibreOffice: Open File, Recent Files with thumbnails, Create File from Blank or a template, Apps: Quire, Ledger, Lectern). Its About text stays general, never naming another company's product.

| App | Folder | Kind | Files |
|---|---|---|---|
| Quire | `public/quire/` | Word processor | .docx/.docm/.dotx/.dotm, RTF, HTML, text, Markdown (GitHub flavour; .zip with pictures); PDF out |
| Ledger | `public/ledger/` | Spreadsheet | .xlsx/.xlsm/.xltx/.xltm, CSV/text, XML Spreadsheet 2003, HTML; PDF out |
| Lectern | `public/lectern/` | Presentations | .pptx/.pptm/.ppsx/.ppsm/.potx/.potm; PDF, PNG, HTML out |

Each app is classic `<script>` files (no ES modules) that attach to one global namespace, `window.L` (also `window.Quire` / `window.Ledger` / `window.Lectern`), loaded in the order listed in its `index.html`. `public/common/` holds what the apps share: the libraries and the Luna stylesheet (one copy each: core, ui, icons, zip, sha, crypto, geometry, metafile, charts, numfmt, dml, spell engine, clipart, `luna.css`), `suite.js` (`window.VO`: launching apps, Recent Files, the service worker) and the proofing dictionaries (`common/dict/`). See docs/suite.md.

## Docs

| | |
|---|---|
| [docs/quire.md](docs/quire.md) | Quire: features, proofing, Word layout rules, compatibility results, source layout |
| [docs/ledger.md](docs/ledger.md) | Ledger: features, formulas and calculation, compatibility results, source layout |
| [docs/lectern.md](docs/lectern.md) | Lectern: source layout, opening, saving, known limits |
| [docs/suite.md](docs/suite.md) | Start Center, `common/suite.js` (app launch, Recent Files), PWA manifest, service worker |
| [docs/testing.md](docs/testing.md) | Test workflow and tools |

Keep them current: a change to an app's features, formats or source layout updates its `docs/<app>.md`. No per-app READMEs.

## Design rules

- Static files, vanilla JavaScript, no frameworks and no build step at run time. Ledger's `index.html` is generated: edit `tools/ledger/build/` and run `python3 tools/ledger/build/make.py`, never the output.
- Everything runs in the browser; documents never leave the machine. No back end, no accounts.
- Nothing from third-party servers: no CDNs, web fonts or analytics; vendored files go in `public/common/`. (Open gap: the three apps still load the metric-compatible fonts from Google Fonts.)
- Compatibility is measured against files other people wrote (public test corpora), compared with outside references (the page or cell values the authoring application stored, LibreOffice, openpyxl, python-pptx), and the result is written down in the app's doc with numbers.
- Saved files must open in the authoring application's newer versions and in LibreOffice; what an app cannot edit it keeps and writes back unchanged.
- Follow the [Preservation contracts](docs/suite.md#preservation) for OOXML ownership, identities, edits, history/copy, drafts and loss reporting.
- Shared code lives once, in `public/common/`; never copy it into an app. A change there is checked in all three apps (load each, and screenshot what it touches). What stays app-specific goes through the hooks in docs/suite.md (app config, `ui.hooks`, `L.icons.app`, …), not through app names inside common code.

## Office acceptance (lessons from the slide-6 repair)

- Schema-valid is not Office-valid: the Open XML SDK and LibreOffice passed decks PowerPoint repaired (a preset's partial adjustment list; a master and layout sharing an ID). Write what Office writes, not merely what the schema allows; when unsure, look at how Office-authored files in the corpus do it.
- Every Office-found defect becomes a `tools/ooxml/package.py` rule with a test, plus a corpus scan for how widespread it is.
- Before handing files to the user for Office, diff the new save against the previous writer's save of the same input; unexplained differences are suspects. Two overlapping faults made one Office round misleading.
- Narrow Office failures by bisection: copies that each change one thing, with whole-file controls (without the suspect part, the part alone), at most about 8 per round, each with a checklist.

## Tools (details in docs/testing.md)

- `tools/serve.py`: no-cache server for `public/` on 127.0.0.1:8760.
- Chromium with CDP at 127.0.0.1:9222. `tools/shot.mjs <page> [steps]` drives a page over CDP, takes screenshots and reports console errors and any request that leaves 127.0.0.1. **Never close the browser's other tabs** (Chromium quits with its last tab).
- `tools/templates.mjs`: regenerates the Start Center's template gallery (`public/common/templates.json` and previews) from the apps; run it after changing a template.
- `tools/ledger/test/`: Ledger's Node harnesses (unit tests, `.xlsx` round trip, recalculation, CSV) and Playwright UI checks.
- `tools/lectern/test/`: Lectern's real-deck harness (open, render, save, reopen in Chromium), UI checks for passwords and charts, encrypted-save check (see docs/lectern.md).

## Testing

1. Ledger engine changes: `node --test tools/ledger/test/*.test.js`.
2. File-format changes: round-trip real files (`node tools/ledger/test/roundtrip.js file.xlsx`) and open the saved file in LibreOffice.
3. Look: `node tools/shot.mjs <app>/` with steps for the change; check every app still loads without console errors or external requests (fonts aside, until they are vendored).
