# Testing

## Serving

    python3 tools/serve.py          # http://127.0.0.1:8760/, no caching

Every response carries `Cache-Control: no-store`, so edited HTML and scripts are always re-fetched. Serve rather than open from disk: a `file://` page cannot fetch the proofing dictionaries, and password-protected files need WebCrypto, which wants https or localhost.

## Screenshots and console errors: tools/shot.mjs

Drives a page in the Chromium listening on CDP 127.0.0.1:9222 (start one with `chromium --remote-debugging-port=9222 --remote-allow-origins=*`, or `--headless=new` on a machine without a display). It starts `tools/serve.py` if nothing listens on :8760.

    node tools/shot.mjs ledger/ '[{"wait":1000,"shot":"ledger.png"}]'
    node tools/shot.mjs quire/ steps.json --w 1440 --h 900 --out /tmp/shots

Each step is `{ "js": "...", "wait": ms, "shot": "name.png|jpg" }`; `js` is an async function body run in the page, and its return value is printed. The app's namespace is `window.L`. At the end the tool lists console errors, exceptions and every request that left 127.0.0.1 (today: the Google Fonts requests). It opens and closes its own tab; never close the browser's other tabs.

The install screenshots in the manifest (`public/<app>/preview.jpg`) are 1280×720:

    node tools/shot.mjs quire/ '[{"wait":1500,"shot":"preview.jpg"}]' --h 720 --out public/quire

The PWA icons are drawn from `public/icons/icon.svg` and `icon-maskable.svg` (see docs/suite.md).

## The Start Center and app launch

The browser profile keeps IndexedDB between runs of `shot.mjs`, so a test can hand a file to an app and then look at the Start Center:

    node tools/shot.mjs quire/ '[{"js":"const b = await L.docx.write(L.D.doc, { stats: L.app.docStats() }); setTimeout(() => VO.handoff(new File([b], \"t.docx\")), 50)"},{"wait":3500,"js":"return document.title"}]'
    node tools/shot.mjs "" '[{"wait":800,"shot":"start.png"}]'

Check after a change to `suite.js` or to an app's start, open or save code: `<app>/?new` starts blank; a handed-over file opens in place of the blank document (one window); the file and a thumbnail appear under Recent Files; clicking it reopens it. Clear the list with `await VO.recent.clear()`.

Templates: `node tools/templates.mjs` opens every template (`<app>/?template=<id>`) and fails on a script error or a missing preview; look at the new previews in the Create File view.

## Ledger: tools/ledger/test/

The engine and the file readers and writers load in Node through `load.js` (the shared ones from `public/common/`; a missing file is an error), so most checks run without a browser.

| | |
|---|---|
| `node --test tools/ledger/test/*.test.js` | Unit tests: calculation, formula parser, number formats, SHA |
| `roundtrip.js file.xlsx [outdir]` | read → write → read, compare every cell's value, formula and style and the sheet settings |
| `xmlss-rt.js file.xlsx [outdir]` | the same through XML Spreadsheet 2003 |
| `recalc1.js file.xlsx` | recalculate every formula and compare with the value stored in the file |
| `csvtest.js file.csv` | parse, compare with Python's `csv` module (`csvref.py`), write back, re-read |
| `lo-compare.js ourDir loDir` | compare our saved files with LibreOffice's re-saved copies cell by cell |
| `pool.js script.js dir out.jsonl`, `recalc-all.js`, `summ.js` | run a per-file script over a corpus with a process pool, summarise the results |
| `ui/*.js` | Playwright checks in Chromium against http://127.0.0.1:8760/ledger/: `smoke`, `flows` (44 end-to-end flows), `allcmds` (every menu and toolbar command), `password`, `perf`, `shots`. Need Playwright (`npm install --no-save playwright`) |

## Lectern: tools/lectern/test/

`corpus.js` opens real decks in Chromium (open, render every slide, save, reopen, compare), `summarize.js` sums up its results, `ui-features.js` drives the password dialog, Tools ▸ Options ▸ Security and the chart dialog, `chartedit.js` writes every chart for validation, `decrypt.py` checks encrypted saves independently, `viscompare.py` / `triptych.py` / `stack.py` compare renderings. They need Playwright; how to run them is in docs/lectern.md.

## Corpora

The compatibility numbers in each app's doc come from public test corpora (LibreOffice, Apache POI, Open XML SDK, XlsxWriter, ClosedXML and others; see the app docs for the list). They are not checked in; download them outside the repo and point the per-file scripts at the folder.
