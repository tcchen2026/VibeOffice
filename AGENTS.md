# VibeOffice

A static web office suite in the style of Office 2003 (Luna Blue) that opens and saves real Office Open XML files. `public/` is the site root: one folder per app, listed on `public/index.html` (card: the app name, a plain subtitle, at most three short sentences, + 1280×720 `preview.jpg`; the disclaimer stays general, never naming another company's product).

| App | Folder | Kind | Files |
|---|---|---|---|
| Quire | `public/quire/` | Word processor | .docx/.dotx, RTF, HTML, text; PDF out |
| Ledger | `public/ledger/` | Spreadsheet | .xlsx/.xlsm/.xltx, CSV/text, XML Spreadsheet 2003, HTML; PDF out |
| Lectern | `public/lectern/` | Presentations | .pptx/.ppsx/.potx; PDF, PNG, HTML out |

Each app is classic `<script>` files (no ES modules) that attach to one global namespace, `window.L` (also `window.Quire` / `window.Ledger` / `window.Lectern`), loaded in the order listed in its `index.html`. `public/common/` holds what the apps share; today that is the offline proofing dictionaries (`common/dict/`).

## Docs

| | |
|---|---|
| [docs/quire.md](docs/quire.md) | Quire: features, proofing, Word layout rules, compatibility results, source layout |
| [docs/ledger.md](docs/ledger.md) | Ledger: features, formulas and calculation, compatibility results, source layout |
| [docs/lectern.md](docs/lectern.md) | Lectern: source layout, opening, saving, known limits |
| [docs/testing.md](docs/testing.md) | Test workflow and tools |

Keep them current: a change to an app's features, formats or source layout updates its `docs/<app>.md`. No per-app READMEs.

## Design rules

- Static files, vanilla JavaScript, no frameworks and no build step at run time. Ledger's `index.html` is generated: edit `tools/ledger/build/` and run `python3 tools/ledger/build/make.py`, never the output.
- Everything runs in the browser; documents never leave the machine. No back end, no accounts.
- Nothing from third-party servers: no CDNs, web fonts or analytics; vendored files go in `public/common/`. (Open gap: the three apps still load the metric-compatible fonts from Google Fonts.)
- Compatibility is measured against files other people wrote (public test corpora), compared with outside references (the page or cell values the authoring application stored, LibreOffice, openpyxl, python-pptx), and the result is written down in the app's doc with numbers.
- Saved files must open in the authoring application's newer versions and in LibreOffice; what an app cannot edit it keeps and writes back unchanged.
- Shared modules (`core`, `zip`, `geometry`, `metafile`, `charts`, `ui`, `icons`, `dml`, `crypto`, `spell`, …) exist as one copy per app and have drifted slightly. A fix to one copy is checked against the others; the direction is to merge them into `public/common/`.

## Tools (details in docs/testing.md)

- `tools/serve.py`: no-cache server for `public/` on 127.0.0.1:8760.
- Chromium with CDP at 127.0.0.1:9222. `tools/shot.mjs <page> [steps]` drives a page over CDP, takes screenshots and reports console errors and any request that leaves 127.0.0.1. **Never close the browser's other tabs** (Chromium quits with its last tab).
- `tools/ledger/test/`: Ledger's Node harnesses (unit tests, `.xlsx` round trip, recalculation, CSV) and Playwright UI checks.

## Testing

1. Ledger engine changes: `node --test tools/ledger/test/*.test.js`.
2. File-format changes: round-trip real files (`node tools/ledger/test/roundtrip.js file.xlsx`) and open the saved file in LibreOffice.
3. Look: `node tools/shot.mjs <app>/` with steps for the change; check every app still loads without console errors or external requests (fonts aside, until they are vendored).
