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

    VO.attach('<app>', { open(file, { draft, name, saved }), thumb(), message(text), templates(),
                         snapshot(doc), docInfo(doc), isDirty(doc), current(), autosave(), showRecovery() })

and reports `VO.opened(file, doc)` after any successful open (File ▸ Open, drop, Start Center), `VO.saved(name, blob, doc)` after a save in its own format, `VO.changed(doc)` after an edit and `VO.discard(doc)` when the user drops the changes (Quire `.docx/.dotx`; Ledger everything but web pages and PDF; Lectern `.pptx/.ppsx/.potx`). Every call is guarded with `if (window.VO)`, so an app page still runs on its own.

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

## Shared code (public/common/)

One copy of everything the three apps share; a change here is a change in every app, so check all three. Classic scripts like the apps' own, attaching to `window.L`, with the Node-safe wrapper `(function (root) { const L = root.L || (root.L = {}); … })(typeof window !== 'undefined' ? window : globalThis)` (Ledger's tests load some of them in Node).

| File | What | Loaded by |
|---|---|---|
| `luna.css` | the Office 2003 Luna Blue look: palette, chrome, menus, toolbars, panes, dialogs, controls; each app's CSS follows and wins | all |
| `core.js` | DOM helpers, colours, units, events, storage, file saving, PDF writer, media store, fonts and substitutes, symbol bullets | all |
| `ui.js` | commands, menus, toolbars, combo boxes, dialogs, colour menus, password and prompt boxes | all |
| `icons.js` | the 16×16 icon sets (common, document, spreadsheet), `L.icons.add`, `L.icons.kit` | all |
| `zip.js`, `sha.js`, `crypto.js` | ZIP read/write; SHA; ECMA-376 encryption (password-protected files) | all |
| `xml.js`, `opc.js`, `opc-order.js` | Portable XML tree, immutable OOXML packages, dependency and identity maps, fragment capture and geometry adapters; generated schema order and preset adjustment tables. | all |
| `geometry.js`, `metafile.js` | AutoShape presets; WMF/EMF player | all |
| `charts.js` | chart model, SVG renderer, chart data dialog (Lectern's `chart-draw.js` adds the v2 renderer) | all |
| `numfmt.js` | Excel number formats | Ledger, Lectern |
| `dml.js` | DrawingML helpers | Quire, Ledger |
| `spell.js` | spelling engine: word lists (`dict/`), checking, suggestions, custom dictionary, Ignore All | Quire, Ledger |
| `clipart.js` | built-in clip art | Quire, Lectern |
| `suite.js`, `dict/`, `templates.json` | the suite: Start Center services; word lists; template gallery | — |

What each app provides:

- before `core.js`: `<script>window.L = { APP_ID: 'quire', APP_NAME: 'Quire', APP: 'Quire 2003' };</script>`. `APP_ID` prefixes the app's settings in localStorage (`quire.…`), so it must never change.
- in its `app.js`, at load: `L.icons.app` (its icon), `ui.hooks = { beforeExec, refocus, commitEdit }` (its editor around commands, toolbar boxes and dialogs); Lectern also `ui.schemeColors` / `ui.currentDesign` (theme colours in colour menus) and PowerPoint's `L.SIZE_LIST`.
- a spelling UI sets `L.spell.refreshSoon` if it draws underlines (Quire).

Script order: `suite.js`, the app config, `core.js`, then the libraries and the app's own scripts as listed in each `index.html` (Ledger's from `tools/ledger/build/make.py`, where `'common/x'` means `public/common/x.js`).

## Preservation

Same-format Office saves carry what the model cannot edit. These contracts apply to all three apps;
the current feature and edit-behaviour tables are in [Quire](quire.md#what-a-save-keeps-converts-and-drops),
[Ledger](ledger.md#what-a-save-keeps-converts-and-drops) and
[Lectern](lectern.md#what-a-save-keeps-converts-and-drops). Conversions to Markdown, CSV, PDF and other
formats have their own rules.

- OOXML preservation uses `common/opc.js`: immutable package bytes stay outside JSON history; model fragments carry source identities and dependency references. Classify parts as opaque, merged or regenerated. Preserve original identities unless every referrer is regenerated; duplicates remap definitions and references together. Merge settings-like parts only; regenerate content parts with model-attached fragments. Record conversions and drops in the document's loss ledger.
- Save and drafts keep the main-part variant and encryption. Save As changes the filename/type only after a successful download. Show unacknowledged loss entries in the Compatibility Checker after preparing the file and before download; acknowledge only a successful user save. Drafts neither show nor acknowledge the checker. Recompute writer losses on every save so cancelling a conversion does not poison the next save.
- Preserved content-control boundaries and binding updates participate in undo. Keep boundaries balanced through edits and paste, remap copied identities together, and roll back the whole transaction when a content/deletion lock forbids it. Typed value commands retain their binding; ordinary typing that converts a typed control records that conversion.
- Retained formatting belongs to individual properties: compare the owning model fields, replace only the property the user changed, and keep its siblings. Capture property alternatives before normalisation and emit changed children in Office-compatible schema order. Copy range endpoints together with fresh identities; report incomplete ranges at save.
- Opaque frames belong to their displayed objects. A compatibility wrapper with several model members is emitted once only while all members remain intact, ordered and contiguous; never resurrect deleted content. Geometry updates preserve the payload. Clipboard transfer includes preview bytes, dependent parts and implicit references, with copied definitions and references remapped together.
- Preservation changes are measured against the pinned corpus (`tools/corpora.sh`), including edits, undo/redo and drafts. Record failed and excluded attempts, compare validator diagnostics with each original, and keep Office acceptance pending until it has actually been checked. The measurement tools are in `tools/ooxml/`.

| Part ownership | Writer contract |
|---|---|
| Opaque | Carry original bytes and original relationships; follow their dependency graph, including cycles and shared targets |
| Merged | Start with the original settings-like part and replace only model-owned properties in schema order |
| Regenerated | Write model content and attached fragments; preserve identities used by any retained referrer |

Strict inputs save as Transitional OOXML after the approved compatibility notice. This explicitly
converts retained XML namespaces and vocabulary; binary dependencies remain unchanged. It preserves
the document/template/slideshow and macro variant. A cancelled save does not acknowledge the notice,
and draft recovery keeps it pending until a successful user save.

`L.opc.open` reads the immutable baseline; `attach` keeps it out of JSON snapshots. `fragment` records
namespace context, relationship attributes and identity definitions/references. `captureAC` runs on
the original XML before branch selection, and `emit` writes an outer compatibility record once.
Inherited compatibility attributes are merged by namespace URI and local name, even when an ancestor
and a fragment use different prefixes. Repeated saves must not add a second `Ignorable` attribute
under an alias; the fragment keeps its own spelling and the inherited value list.
Fragments stay self-contained in history and clipboard data. Once a regenerated or merged part is
assembled, `output.put` hoists missing namespace declarations and the union of `mc:Ignorable` onto
its root. It removes redundant inner declarations, keeps actual prefix rebindings and default-namespace
resets, and leaves other compatibility properties scoped as written. Opaque parts bypass this pass
and keep their bytes. This prevents every preserved paragraph/run property from repeating the full
document namespace list.
`duplicate` assigns a shared copy identity to the selected definitions and references; `export` and
`import` transport their dependency bytes between documents. The writer reserves existing part names,
relationship IDs and identity spaces before allocating new values. Master and layout IDs share one
space; non-empty preset adjustment lists include every preset value in Office's order.

The loss ledger distinguishes conversion from deletion and keeps stable entry identities. Digital
signatures are removed and reported because a changed package cannot retain their validity. Missing
dependencies and unavoidable conversions are reported by the writer; draft recovery retains pending
entries even when the original content has already been converted in the draft.

An Office review batch contains at most 15 files and a checklist. Before handoff, compare the new save
with the previous writer's save of the same input and explain each difference. Automated validators
and LibreOffice do not establish Office acceptance. Keep run history outside the repository in
`~/corpora/results/<run>/README.md`, and keep app docs focused on current behavior and measurements.

## Recent Files and unsaved versions

One list for everything a user worked on, shown in four places: the Start Center (cards), each app's File menu (1–4), its Getting Started pane (Open) and its Document Recovery pane. Every entry really reopens.

IndexedDB database `vibeoffice`, store `recent`: `{ key, app, name, size, time, file, thumb, draft }`, keyed `<app>:<file name>`, or `<app>:~<id>` for a document that was never saved. The browser has no file paths, so the list keeps a **copy of the file**: `file` is the one opened or last saved; `draft` (`{ file, name, time, lossState }`) is the **unsaved version**. Two files with the same name in different folders share an entry. At most 30 entries, but an entry with unsaved changes is never dropped; files over 50 MB are not kept. Every page hears about changes over the `BroadcastChannel` `vibeoffice` (`VO.onRecentChange`); `VO.recent.cache` is the list for menus.

**Autosave.** The app calls `VO.changed(doc)` when a document changes; 15 s later (longer for documents that are slow to write: 30 × the time the last write took) `hooks.snapshot(doc)` writes it in the app's own format (including macro, template and slideshow variants; password-protected documents stay encrypted) into the entry's `draft`. `VO.opened(file, doc)` and `VO.saved(name, blob, doc)` tie the document to its entry; saving clears the draft (and retires the `~` entry of a never-saved document); No in "Do you want to save the changes…?" calls `VO.discard(doc)`, which throws the draft away. Closing the tab or a crash leaves it. Draft metadata also keeps pending and acknowledged compatibility entries; recovery restores them silently so the next user save can report an earlier conversion. Quire's Tools ▸ Options ▸ Save ▸ "Save AutoRecover info" turns it off.

**Recovery.** When an app starts and other documents of it have unsaved versions, it opens the **Document Recovery** task pane (as in Office 2003): each one with *Open the unsaved version* and *Discard the changes*. Opening one (`VO.openRecent(key, { draft: true })`, or `<app>/?recent=<key>&draft` from the Start Center) restores the changes as unsaved, so the document still has to be saved; a never-saved one then asks for a name. On the Start Center such a card is marked *Unsaved changes* (or *Never saved*), opens the unsaved version, offers *Saved version* when there is one, and × discards the changes.

Thumbnails are JPEG, at most 320 px, made 0.6 s after an open, a save or an autosave of the current document from `hooks.thumb()`: Quire rasterizes its first page as for PDF (print layout only), Ledger crops the top left of its grid canvas (4:3, at most 600 px wide), Lectern draws slide 1. Without one, the card shows the app's page picture and the extension.

## PWA

The manifest's `id`, `start_url` and `scope` are `./`, so the suite installs from any folder it is served from. Shortcuts create a new document, spreadsheet or presentation (`<app>/?new`). File handlers (`.docx .dotx .docm .rtf .md .markdown`, `.xlsx .xltx .xlsm .csv`, `.pptx .ppsx .potx`) all go to the Start Center, which hands the file on; `launch_handler` `focus-existing` reuses an open Start Center window. App pages link the same manifest, so installing from inside an app installs the suite.

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
