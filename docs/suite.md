# The suite: Start Center, app launch, Recent Files, PWA

| File | What it does |
|---|---|
| `public/index.html` | The Start Center |
| `public/common/suite.js` | `window.VO`, loaded by the Start Center and before every app's own scripts |
| `public/manifest.webmanifest` | Install metadata: icons, shortcuts, file handlers, screenshots |
| `public/sw.js` | Service worker (dev mode: no caching) |
| `public/icons/` | `icon.svg` (rounded, transparent corners) and `icon-maskable.svg` (full bleed, artwork inside the 40 % safe circle), and the PNGs made from them |

## Start Center

Modeled on LibreOffice's Start Center, drawn in the apps' Luna Blue. No title bar of its own (installed, the window has the system's). Left: Open File…, Recent Files, Create File, *Apps*: Quire (word processor), Ledger (spreadsheet), Lectern (presentations), each opening with its demo file, Install (when the browser offers it), About. Right, one of two views:

- **Recent Files** (the default): cards (thumbnail, name, app, when), each with a remove button; *Clear Recent Documents*; when the list is empty, the welcome text and the drop hint.
- **Create File** (`#create`, so Back returns to Recent Files): Documents, Spreadsheets and Presentations, each starting with Blank (`<app>/?new`) followed by the app's templates (`<app>/?template=<id>`), with a preview of each.

 A file dropped anywhere on the page, picked with Open File… or passed by the operating system (installed app, `launchQueue`) goes to the app for its extension (`VO.appFor`); other files get a message in the bar.

Installed (display mode standalone), app pages open in a new window and the Start Center stays, as in LibreOffice; in a browser tab they open in the same tab (Back returns).

## Windows

No page draws a title bar of its own: the browser tab, or the system title bar of the installed app, shows `document.title` (`<document> - Quire 2003`). Each app sets its own icon as the page icon with `VO.setIcon(L.icons.app(32))`.

## App launch: `VO.launch` and `VO.attach`

| URL | The app starts with |
|---|---|
| `<app>/` | its sample document, as before |
| `<app>/?new` | a blank document (`Document1`, `Book1`, `Presentation1`) |
| `<app>/?template=<id>` | a new document from that template (`Document1`, `TaskTracker1`, `Presentation1`) |
| `<app>/?open=<id>` | a blank document, then the file the Start Center stored under `id` in the `handoff` store (read once, then deleted; entries older than a day are swept) |
| `<app>/?recent=<key>` | a blank document, then that Recent Files entry |

The query is removed from the address once read, so a reload starts clean. Opening the file replaces the untouched blank document (each app already does this for its own File ▸ Open), so it ends up in one window.

Each app, in its `A.init`, picks template, blank or sample from `VO.launch` and finishes with

    VO.attach('<app>', { open(file), thumb() -> canvas | null, message(text), templates() -> [{ id, name }] })

and reports `VO.opened(file)` after any successful open (File ▸ Open, drop, Start Center) and `VO.saved(name, blob)` after a save in its own format (Quire `.docx/.dotx`; Ledger everything but web pages and PDF; Lectern `.pptx/.ppsx/.potx`). Every call is guarded with `if (window.VO)`, so an app page still runs on its own.

## Templates

Each app owns its templates; the Start Center only shows them.

| App | Where | Templates |
|---|---|---|
| Quire | `quire/js/templates.js` (`L.templates`, also File ▸ New) | Contemporary Letter, Professional Memo, Elegant Fax, Contemporary Résumé, Professional Report, Meeting Agenda, Meeting Minutes, To-Do Checklist, Event Flyer, Recipe |
| Ledger | `ledger/js/panes.js` (`TPL`, also the New Workbook task pane) | Loan Amortization, Expense Statement, Sales Invoice, Timecard, Personal Monthly Budget, Task Tracker, Shopping List, Weekly Schedule |
| Lectern | `lectern/js/templates.js` (`L.templates`): content decks, each on one of the 14 designs | Project Status Report, Team Meeting, Lesson, Event Announcement |

Inside the apps, File ▸ New… (and Ctrl+N) shows the same gallery, Blank first, with the same previews: `VO.newDialog(app, create)` builds it in the app's own dialog (`L.ui.dialog`) from `common/templates.json`, and `create(id)` makes the document in that app (`''` is Blank). The New toolbar button still makes a blank document straight away, as in Office.

Placeholder text says what goes where (Quire's are click-to-replace fields). The gallery's list and pictures are generated, never edited by hand: after adding or changing a template run

    node tools/templates.mjs

which opens every app and template in Chromium and writes `public/common/templates.json` and `public/<app>/templates/<id>.jpg` (`blank.jpg` for Blank).

## Recent Files

IndexedDB database `vibeoffice`, store `recent`, keyed `<app>:<file name>`: `{ key, app, name, size, time, file: Blob, thumb: Blob }`. The browser has no file paths, so the list keeps a **copy of the file**: the one opened, then the one last saved. Two files with the same name in different folders share an entry. At most 30 entries; files over 50 MB are not kept. All other pages are told about changes over the `BroadcastChannel` `vibeoffice`.

Thumbnails are JPEG, at most 320 px, made 0.6 s after an open or save from `hooks.thumb()`: Quire rasterizes its first page as for PDF (print layout only), Ledger crops the top left of its grid canvas (4:3, at most 600 px wide), Lectern draws slide 1. Without one, the card shows the app's page picture and the extension.

## PWA

The manifest's `id`, `start_url` and `scope` are `./`, so the suite installs from any folder it is served from. Shortcuts create a new document, spreadsheet or presentation (`<app>/?new`). File handlers (`.docx .dotx .docm .rtf`, `.xlsx .xltx .xlsm .csv`, `.pptx .ppsx .potx`) all go to the Start Center, which hands the file on; `launch_handler` `focus-existing` reuses an open Start Center window. App pages link the same manifest, so installing from inside an app installs the suite.

Icons: edit the SVGs, then render the PNGs in Chromium:

    R() { node tools/shot.mjs icons/$1.svg "[{\"js\":\"const s=document.documentElement; s.setAttribute('width',$2); s.setAttribute('height',$2);\",\"wait\":200,\"shot\":\"$3\"}]" --w $2 --h $2 --transparent --out public/icons; }
    R icon 512 icon-512.png; R icon 192 icon-192.png; R icon-maskable 512 icon-maskable-512.png; R icon-maskable 180 apple-touch-icon.png

## Service worker

**Dev mode (now):** `sw.js` has no fetch handler, so requests go to the server untouched (with `tools/serve.py`'s no-store headers an edit shows on reload), and activating it deletes any cache an earlier worker left.

**Release mode (later), for working offline.** The plan, and what it fixes in the usual cache-first worker:

- A `VERSION` stamped at publish time names the cache (`vibeoffice-<commit>`); install fetches every asset with `{ cache: 'no-cache' }` (and `?v=VERSION`) into it, all or nothing; activate deletes other `vibeoffice-*` caches. Keep these parts of the common pattern.
- The asset list is not hand-kept: the publish step (or install itself) reads the `<script src>` list from each app's `index.html`, plus the Start Center, `common/suite.js`, manifest and icons. The dictionaries (4 MB) are cached when first used.
- Every `cache.put` from the fetch handler sits inside `event.waitUntil`, or the worker may stop before it is written.
- Only `status === 200` same-origin (`type === 'basic'`) responses are stored: `response.ok` also lets 206 partial responses through, and `cache.put` rejects them.
- Navigations match with `ignoreSearch` (the apps are opened as `?new`, `?template=`, `?open=`, `?recent=`), and an offline navigation that misses falls back to the cached Start Center instead of the browser's error page.
- No automatic `skipWaiting`: an app page that is already open keeps loading scripts of its own version (the apps load about 20–45 classic scripts that must match) instead of a mix of old and new. The Start Center shows *A new version is ready — Reload*, which posts `skipWaiting` to the waiting worker.
- Paths relative to the worker's scope, never `/…`, so the site also works from a sub-folder (GitHub Pages).
