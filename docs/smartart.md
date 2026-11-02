# IntelliArt (SmartArt): status and to do

IntelliArt is Lectern's name for its diagrams; in files they are Office SmartArt. Features and file behaviour:
[lectern.md](lectern.md#intelliart-smartart). This page tracks how far PowerPoint agrees with us, and what is left.

## Where it stands (2026-10-10)

- **In Lectern:**
  - 26 layouts (`public/common/smartart.js`), with a gallery, text pane, Diagram toolbar and IntelliArt Tools;
  - colours (two from the theme, 13 fixed palettes of our own) and styles in three families (filled, outlined,
    tinted), pictured in the IntelliArt pane;
  - a shape's own formatting is kept;
  - the sample deck and four templates use IntelliArt;
  - PowerPoint's SmartArt in 21 layouts we share is editable.
- **Saved as SmartArt** (`public/lectern/js/smartart-io.js`):
  - the data model with presentation points;
  - our own layout, style and colour definitions (`urn:vibeoffice.work/diagram/…`);
  - the drawing of our shapes.
- **Checks that pass:**
  - Open XML SDK and `package.py` on every layout;
  - LibreOffice draws every layout from the stored drawing;
  - Lectern reopens every layout editable;
  - `tests/lectern/smartart-office.py`: 43/43 corpus decks;
  - `frames.mjs`: 25/25.
- **PowerPoint:**
  - every test file opened without a repair prompt;
  - **PowerPoint ignores the stored drawing and always lays IntelliArt out from our layout definition**, so its look in PowerPoint depends only on `smartart-io.js`.

How PowerPoint draws each category from our definitions (PDF exports, batches 1–3 and rounds 1–4):

| Layouts | In PowerPoint |
|---|---|
| Lists (6) | Right content, sub-item cards pale. Stacked, Lined and Vertical Box List look like Vertical Bullet List (they share one definition). Block List colours differ |
| Processes (6) | Boxes too tall; Chevron, Step Up, Continuous Arrow and Timeline lose their shapes; arrows faint |
| Cycles (3) | Circles overlap, no arrows; Basic Radial: all satellites, but against the centre and without visible spokes |
| Hierarchies (3) | Levels, children and gaps right; lines drawn between the wrong boxes (to the last child, or between siblings) |
| Relationship, Matrix, Pyramids (8) | Not yet compared (the final batch drew none of our definitions; see below) |

Learned from PowerPoint (bisection, and PowerPoint's own org chart in the corpus):

- **Spacing constraints:** a branch that refers to its own descendants makes PowerPoint reject the whole definition and fall back to a block list. Set spacing once, at the top.
- **Tree connectors:** a child's connector is its parent transition, selected with `axis="precedSib" ptType="parTrans" st="-1" cnt="1"`. Without `st`/`cnt`, children after the first disappear.
- **Connector settings:** `dim 1D`, `endSty noArr`, `bendPt end`; top items have no connector.
- **Text in its own shape:** a shape whose sub-items have shapes of their own shows `presOf axis="self"`; `desOrSelf` repeats them.
- **Finding the definition:** the data's `loTypeId` must equal the layout part's uniqueId, or PowerPoint draws its default list.
- **Presentation links:** a node with an empty `presOf` (other than the root) gets no `presOf` link in the data.
- **Lines:** a line ignores `rot` in a definition; use `lineInv` for lines running down to the left.
- **One algorithm type per node:** every branch of a `choose` must give a layout node the same algorithm type; a fallback needs a child node for a different algorithm.

**Current definitions (uncommitted):** each saved diagram gets its own layout definition (uniqueId
`…/layout/<id>#<hash>`). It places every shape exactly where our engine draws it (a composite, one `if` branch per
item count 1–10, Lectern's shapes sampled at the frame's aspect), with a plain row (or, for trees, the recursive tree
above) as the `else`. Items are bound with `forEach axis="ch" ptType="node" st=k cnt=1` (nested for tree paths); a
shape's text sits in a transparent text node on top. Trees choose their branch by total items (`cnt axis="des"`) or
by top items (`axis="ch"`); LibreOffice handles only `ch`. In LibreOffice (without the drawing) this matches Lectern
for every flat layout.

**Final batch result (`tmp/smartart-final`, 14 decks, 2026-10-10):** no repair prompts, but PowerPoint drew its
default block list for every diagram, even where the `else` branch applied, so it rejects the whole definition.
The data part's `loTypeId` lacked the `#<hash>`. The writer now copies the layout's uniqueId, and `package.py` reports
`diagram-layout-id`: 133 of 134 corpus files with SmartArt match, and the other is a LibreOffice bug sample.
**The rerun with only that fix (`tmp/smartart-final2`) was rejected the same way**, so the mismatch was not the
(only) cause.

**Round 6 (`tmp/smartart-r6`, 100 decks; scripts in `tmp/smartart-r6-tools`) found the cause.** PowerPoint drew our
definition in exactly the 15 decks where no choice switched the algorithm type, and its default list in the 85 where
one did. Our `else` (more items than the branches cover) used `lin` or the tree algorithms under a `composite`
`if`. None of the corpus's 303 choices in 60 Office definitions switches type. Everything else tested was accepted:
`#` ids, separate text layers, arrows, equal font sizes, ctr constraints, a count choice without `else`. The fallback
branches now stay `composite` and hold one child node that runs the row or the tree. `package.py` reports
`diagram-mixed-algorithm` (with a test); it matches all 100 PowerPoint results and flags none of the corpus's 134
SmartArt files.

**Final4 (`tmp/smartart-final4`, 2026-10-11): PowerPoint draws our definitions.** Chevron, Step Up, Continuous
Arrow, Timeline, Radial (4), Continuous Cycle, the trees, Venn, Target, Linear Venn, Funnel and the pyramids look
like Lectern. Both tree conditions worked; we keep the total item count (`des`), which falls back to the tree
algorithms after Add Shape. Still wrong, fixed for final5:
- one shape at full diagram size (Process with 4 and 7 items, Vertical Process, Cycle with 4, Radial with 2). Our
  presentation points linked every shape of an item to it, also those whose `presOf` is empty; Office links none of
  those (473 of 476 in the corpus);
- lines down to the left drawn wrong: PowerPoint ignores `rot` on a line in a definition (shapes turn), so they use
  the `lineInv` preset;
- text larger than Lectern's: each text node now has our size (`primFontSz` per node); Block List lost its bullets,
  left-aligned text gets `parTxLTRAlign`.

**Colours and styles (2026-10-11, not yet seen in PowerPoint):** quick styles refer to the theme's style matrix by
index, so the themes Lectern writes now have one (`L.model.FMT`: solid, soft and rich gradient fills; 0.75, 1.5 and
2.5 pt lines; three drop shadows). Ranges are two-colour `fillClrLst` lists without `meth` (span, as Office writes
Colorful Range); our palettes are `srgbClr`, the theme's choices `schemeClr`; outlined and
tinted styles get colours of their own, saved under their own id (`…/colors/<id>/outline`, `/tint`), with the
item's colour in `linClrLst` and in the text. The text sits on the transparent `trAlignAcc1` nodes, so that label
carries it. LibreOffice without the drawing showed two things: it takes text colour from the style's `fontRef` (so
outlined and tinted styles name `tx1` there), and it drops transparency in a span list (so a tint is a colour mix,
not an alpha). Checked: `package.py`; LibreOffice with and without the drawing; Lectern and LibreOffice re-saves reopen
with every diagram's colours and style.

## Method

1. Build a small batch in `./tmp`. Each file must pass the SDK and `package.py`.
2. The user opens the batch in PowerPoint and exports it to PDF into the same folder, or `tmp/office/office.py run`
   sends it to the Office 2016 test VM for PowerPoint 2016's PDFs (newer PowerPoint can still differ).
3. Compare each page with Lectern's drawing (`tmp/smartart-hier/sheet.py`).
4. Narrow failures by bisection: one change per copy, at most about 8 copies.

LibreOffice is a partial stand-in. Without the stored drawing (`tmp/smartart-iter/loop.sh CATEGORY`) it lays our definitions out, but it draws no tree connectors and handles trees poorly.

## To do

**Final5 (`tmp/smartart-final5`, 2026-10-11):** text sizes and alignment now match Lectern, and the trees are right,
including Hierarchy's slanted lines (`lineInv`). One shape is still drawn across the whole diagram in a few counts:
the middle arrow of a 4-item Process or Vertical Process (in every frame shape), an arrow of a 4-item Cycle, box 5 of
a 7-item Process, the centre of a 2-satellite Radial. The copies without presentation points and drawing show the
same, so the definition causes it, not the data link. PowerPoint ignores all four constraints of that one node,
although they look like its neighbours'.

PowerPoint's re-saves of final5 (`*-saved.pptx`, `*-added.pptx` there) record the size it gave each shape
(`tmp/smartart-r6-tools/ppsaved.py`): the stretched shapes are exactly the diagram's size, so their four
constraints were ignored. They are the same nodes at every frame shape, so the cause is the structure, not the values.
**Add Shape works:** a fifth process item took our 5-item layout at Lectern's spacing; a new org-chart box took the
tree fallback with its connectors. `powerpoint-own.pptx` (Basic Process, Name and Title Organization Chart, Labeled
Hierarchy made in PowerPoint) is a reference; its Basic Process boxes are 123 × 74 pt to our 125 × 75.

**PowerPoint 2016 (the Office 2016 test VM, build 16.0.4266, 2026-11-02; PDF exports through
`tmp/office/office.py`):** final5's decks showed nothing at all, and the hierarchies deck closed PowerPoint, while
newer PowerPoint draws them. Our presentation links (`presOf`, `presParOf`) named `…/layout/<id>` without the
layout's `#<hash>`; with `presId` equal to the uniqueId (as in all 233 corpus data parts with such links) every
layout draws, the stored drawing or not. `package.py` reports `diagram-pres-id`. Round 7 then ran there too: the
stretched shape went only when the definition had a single count branch or when every layout node was renamed
(`n000`, `n001`, … through the whole definition); renaming one bad node fixed the radial but not the process, whose
other arrow stretched instead, and moving its values changed nothing. The cause in PowerPoint is not known; names
numbered once per definition fixed every stretched shape in all 63 slides of final5's decks (`tmp/office/final6`),
so the writer now names them that way. `tests/lectern/smartart-definitions.test.mjs` covers both.

1. **Newer PowerPoint:** confirm the two fixes there (final6's decks); round 7's re-saves are no longer needed.
   Also: one slide per new style family (filled with gradients and shadows, outlined, tinted) and a palette range
   and cycle, to see that PowerPoint steps through a range item by item as Lectern does (Lectern mixes in RGB).
   In PowerPoint 2016, still to fix:
   - more items than the count branches cover (11 in a process): the plain-row fallback draws full-height boxes;
   - a square or tall frame: the process and chevron rows come out much smaller than the frame (the branches keep
     the aspect ratio they were sampled at); compare with Lectern's drawing in the same frame.
   Already right in final5, to keep an eye on:
2. **Processes:** box proportions; Chevron, Step Up, Continuous Arrow and Basic Timeline shapes; visible arrows.
3. **Cycles:** spacing on the ring, and arrows between items.
4. **Lists:** Stacked, Lined and Vertical Box List differ from Vertical Bullet; Block List colours.
5. **Relationship, Matrix and Pyramids:** compare in PowerPoint.
6. Tie arrows, connectors and cards in the stored drawing to their presentation points (only item shapes are tied now).
7. An item holds one paragraph: PowerPoint items with several paragraphs are joined, and org-chart assistants become ordinary members.
8. Later: IntelliArt in Quire (Word) and Ledger (Excel), which today keep SmartArt unchanged or (Ledger) drop it.
