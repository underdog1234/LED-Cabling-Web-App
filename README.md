# LED Cabling Web App

Version `0.23.0`

Standalone React web app for planning LED wall layouts, signal port mapping, power outlet assignment, stock checks, deployment hardware, and PDF/settings/video exports.

## What It Does

- Build LED walls by rows and columns, or place panels freely (non-uniform layouts) with drag, edge-snap and joining
- Switch between `MG9` and `MT` panel profiles, plus `MG12` triangle and `MG13` curved variants
- Group panels into named **Sub-Screens**, edit/patch each one in isolation (other sub-screens stay visible but dimmed and locked), and view the complete assembled layout in Canvas View
- Position the whole layout or individual sub-screens within a configurable-resolution **Output Canvas** (drag, numeric entry, align/snap tools, boundary/overlap warnings) for multi-processor / media-server mapping
- Select a **NovaStar processor model** (`VX1000 Pro` / `VX2000 Pro`), see its live pixel/port capacity vs. current usage, assign a video input per sub-screen or one input for the whole canvas, and generate a real, importable `.uprj` processor configuration file - validated against the selected processor's port count, per-port and total pixel limits, and canvas size, with a summary of what will be exported and any blocking errors or warnings before download
- Import projects from the Creative Layout Tool
- Patch signal and power manually or with auto-snake / automatic letter-patching routing, scoped to the active sub-screen when one is selected
- Flip the panel layout between `Back View` and `Front View`
- Export a PDF report with portrait detail pages, a per-sub-screen summary page, plus both layout views in landscape
- Export a native-resolution PNG test pattern, a full-screen canvas-only live animated test pattern, or a downloadable looping WebM video of it
- Save and reopen settings as JSON (v5 format adds NovaStar processor/input selection; v3 sub-screens and output-canvas positioning, v2 free-panel and legacy grid formats still open)
- Check stock levels, shortfalls, and deployment hardware requirements
- Collapse any section of the UI to reduce clutter on long projects

## Recent Changes In v0.23.0

- Added **NovaStar Processor Configuration File Export**: pick `VX1000 Pro` or `VX2000 Pro` in LED Wall Setup (with a live capacity-vs-usage readout beside the selector), assign a video input per sub-screen or one input for the whole canvas via a mode toggle in Output Canvas, then generate a real `.uprj` file from a new "NovaStar Processor Configuration" section showing a full pre-download summary (processor, surface, canvas/screen resolution, panel/sub-screen counts, Ethernet outputs used, pixel load per output, input assignments) plus any validation errors (which block download) or warnings. The file format - a custom envelope wrapping an embedded SQLite project database, including its checksum algorithm - was reverse engineered byte-for-byte against real NovaStar-exported project files rather than guessed; cabinet/panel positions and signal-port patching order are written into the LED screen's own native pixel grid exactly as NovaStar itself represents it, verified byte-for-byte against real reference exports covering gapped/irregular walls, multiple Ethernet outputs, and multi-sub-screen layouts. Output-canvas/sub-screen placement is not yet reflected in the exported cabinet positions (only in the video-routing side) - a known gap for a future pass. Settings JSON bumps to v5 to carry the new processor/input fields; older saves load with no processor selected, same as always
- Output Canvas and Stock Calculations are now collapsed by default
- Fixed the PNG and animated/video test patterns not locking panels to their true positions on walls with a missing/removed panel in the middle of a row: panels were packed tightly left-to-right in array order (silently closing up any gap), shifting every panel after the gap out of position. Both now place each panel at its real relative pixel offset, so a physical gap shows as empty space instead of squeezing later panels together - uniform, gap-free walls render identically to before. The PNG export also had its own separate copy of this same (still-buggy) positioning logic despite the live/video test pattern already having been fixed for it previously; it now shares the one, corrected implementation
- Fixed the test pattern's per-panel row/column corner labels reading from the panel's raw back-view position instead of the front (mirrored) view the pattern always renders in - column numbers were backwards versus what's on screen. The rendered top-left panel now always reads row 1, column 1, and the centred info-text block is smaller
- The Panel Layout workspace's (and PDF's) height ruler now reads bottom-up - 0m at the wall's base, increasing upward - matching how a physical wall is measured and built; tick positions are unchanged, only the printed labels. Added compact direction arrows to the Columns/Rows field labels
- Added an automated test suite (Vitest, `npm test`) covering the NovaStar export pipeline - envelope/checksum round-trips, both processor models, irregular and multi-sub-screen walls, patching-order preservation, and golden-file comparisons against real NovaStar-exported project files

## Recent Changes In v0.22.2

- Fixed the PNG and animated/video test patterns rendering MT panels at the wrong resolution: both always sized the whole canvas and every panel using MG9's mm-to-pixel ratio (168px per 500mm), which happened to be correct for MG9 but gave MT panels 336x168px instead of their true native 256x64px. Both exports now place each panel using its own native `pixW`/`pixH` from its panel type, accumulated per row band - the same algorithm already used for the "Resolution" stat in Wall Summary - so the exported canvas size and every panel's pixel footprint always match what the app reports on screen. MG9-only walls are unaffected

## Recent Changes In v0.22.1

- Fixed the animated/video test pattern's RGB checkerboard: the colour tile was always sized to an MG9 panel's pixel width, so on MT walls each panel (physically twice as wide) showed two colours split down its middle instead of one solid colour. The tile now matches the actual panel footprint on the wall (derived from the most common panel type present), so MT panels correctly show one colour across their full width and MG9 walls are unaffected
- Reworked the test pattern's text legibility and sizing: the centre info block (project name/stats) now has a black stroke behind its white fill so it stays readable where the alignment cross-hatch overlay crosses it, is much larger, and the per-panel row/column corner labels are now half their previous size. The alignment cross-hatch overlay itself is now more transparent

## Recent Changes In v0.22.0

- Added **Sub-Screens**: create/rename/delete named panel groupings, assign/reassign/remove panels, select-all-in-screen, and a Canvas View showing the whole layout at once. Selecting a sub-screen dims and locks every panel outside it and scopes all calculations (panel count, resolution, weight, power, stock, port usage), manual/auto patching, and PDF/PNG/video exports to just that screen. A labelled boundary outline is drawn in the workspace for each sub-screen. Projects with no sub-screens behave exactly as before, and old save files load unchanged
- Added **Output Canvas Positioning**: configurable output resolution (common presets plus custom width/height), per-sub-screen (or whole-layout) placement via numeric X/Y entry, dragging, arrow-key nudging, or align/centre/snap-to-edge tools, a scaled live preview, and warnings for out-of-bounds or overlapping screens. Each panel's final canvas-space X/Y is computed from its sub-screen's position plus its own placement, independently of the physical millimetre layout
- Auto-patching, manual patching, and the clear/match-power actions now operate only on the active sub-screen's panels (shared signal/power port pools are respected, and other sub-screens' patching is never touched or reset)
- PDF export gains a per-sub-screen summary page (name, resolution, physical size, canvas position, panel count, ports in use) whenever sub-screens exist
- Fixed sub-screen creation jumping straight into the new (empty) sub-screen, which visually collapsed the workspace and dimmed every existing panel until something was assigned to it - creating a sub-screen now leaves the current view untouched
- Reorganised the layout: removed the legacy "LED Surface / Sub-Screen Name" field, moved the Sub-Screens panel below Wall Summary, moved Output Canvas below Power Outputs (and made it always visible, no longer behind a toggle), and moved panel-to-sub-screen assignment controls next to Undo/Redo with a live selection count
- Every major section of the UI is now collapsible via its header

## Recent Changes In v0.21.2

- Quick Panel Layout: added an `Export PDF` button - a one-page landscape summary with the panel type/grid/wall size/resolution/aspect ratio/16:9 content-area stats plus a to-scale diagram of the wall (grid lines, metre rulers and the 16:9 overlay), matching the on-screen preview

## Recent Changes In v0.21.1

- Quick Panel Layout: added a `Clear` button (resets to 1×1 MG9), metre rulers along the top and left of the preview, and up/down buttons to nudge the centred 16:9 content-area box vertically within any available slack
- Quick Panel Layout: added Width (m) / Height (m) inputs alongside the Columns/Rows counters, kept in sync both ways - width steps in 0.5m increments for MG9 / 1m for MT, height always steps in 0.5m (both panel types are 0.5m tall)
- Quick Panel Layout: reworded the preview caption from "Dashed box = centred 16:9 content area" to "Dashed box = 16:9"

## Recent Changes In v0.21.0

- Added **Quick Panel Layout**: a standalone panel-count calculator, opened in its own browser tab (`Quick Panel Layout` toolbar button), independent of any open project. Pick MG9/MT and columns/rows and see live wall size, pixel resolution, panel count, aspect ratio, a centred 16:9 content-area overlay, and a warning if the wall (or the 16:9 area) is below 1920×1080. A `Send to Main Layout Tool` button hands the chosen grid off and jumps to the main planner, which applies it on load (replacing the canvas if it already has panels, after a Replace/Add to canvas/Cancel prompt)
- The main LED Cabling Planner now starts with an **empty canvas** instead of an auto-generated 24×8 grid - build a layout via `Apply Grid Size`, import, open a saved project, or send one in from Quick Panel Layout

## Recent Changes In v0.20.2

- Fixed the on-screen signal/power chain-start ring indicator rendering as a square on MT (or any non-square) panels instead of following the panel's true rectangle. The ring's SVG had no explicit width/height, so browsers fell back to its viewBox's 1:1 intrinsic aspect ratio when sizing it as an absolutely-positioned element, overriding the intended stretch-to-fill. Now explicit, always matches the panel's actual shape

## Recent Changes In v0.20.1

- Removed the signal/power chain-start ring indicators from the PNG test pattern - it's a clean per-panel pixel map now, not a patching diagram. PDF and on-screen views still show them

## Recent Changes In v0.20.0

Consolidates and confirms the PNG shape-rendering fix: verified consistent between the PNG test pattern and the PDF's front view across triangle panels at all four rotations (0/90/180/270) and the corner panel's hatch pattern, in addition to the curved-panel fix already in v0.19.2.

## Recent Changes In v0.19.2

- Fixed a bug where curved (MG13) panels rendered with the wrong corner cut in the PNG test pattern - the PNG traced a curve path geometrically opposite to the one used on-screen and in the PDF, making rotated curved panels look incorrectly oriented. The PNG now always shares the exact same shape-tracing logic as the PDF/screen, so per-panel rotation matches the PDF's front view exactly
- Removed the rotate icon (🔄) and all per-panel signal/power port info from the PNG test pattern; panels now show only their row/column label and shape symbol

## Recent Changes In v0.19.1

- The outer-extremity outline is now a single, thicker white line (3px) instead of two separate 1px lines with a gap between them

## Recent Changes In v0.19.0

- Panel alignment outlines are now pixel-snapped and drawn as a crisp true 1px line (rect/MT/corner panels get an exact strokeRect fast path; shaped panels keep their straight legs crisp)
- Removed the black outline around the wall info text; removed all per-panel signal/power port labels from the test pattern
- Added a "LED Surface / Sub-Screen Name" field (alongside Project Name), saved with the project; both names are shown centred on the wall when defined, with no placeholder when empty
- Panel location labels moved to the top-left corner of each panel as two lines (`↓row` / `→col`), consistently positioned regardless of shape or rotation
- The full-screen live view now defaults to true 1:1 pixel mapping (centred if smaller than the window, scrollable if larger) instead of stretching to fit; any keypress toggles an optional scaled-to-fit preview
- Added a double 1px white outline around the true outer extremity of the whole assembled LED surface (not per-panel), accurately following triangular/curved/irregular outlines and ignoring internal panel-to-panel seams

## Recent Changes In v0.18.0

Animated test pattern tweaks:

- Split into two dedicated buttons: **Video Test Pattern** opens a pure full-screen canvas in a new tab - no header, no buttons, no text outside the LED canvas itself; **Download Video Test Pattern** records and downloads the WebM directly from the main app, no tab required
- The moving greyscale gradient is now a single large sweep spanning the whole wall corner-to-corner, instead of several smaller repeating bands
- Removed the info panel's background box and the "Test: ..." description line; the remaining wall info (resolution, physical size, panel count, grid) is now centred on the wall
- Added a corner-to-corner alignment cross and a centre circle (diameter equal to the wall's height) as a geometry reference for spotting warped, offset or stretched panels
- Fixed washed-out/blocky WebM exports by giving the recorder a much higher, resolution-scaled video bitrate instead of the codec's low default

## Recent Changes In v0.17.0

Added an animated LED wall test pattern for spotting orientation, patching and alignment errors that a static swatch can't reveal:

- New "Animated Test Pattern" button opens a live, looping canvas view in its own tab, rendered at the wall's exact configured pixel resolution
- RGB checkerboard: every panel shows a solid red, green or blue test colour in a diagonally staggered arrangement (never a blended rainbow), sliding smoothly left-to-right and cycling red -> green -> blue -> red
- A moving diagonal greyscale brightness sweep plays across the whole wall at the same time, continuous across every panel boundary (not restarting inside each panel), without introducing colour or making panels hard to identify
- 1px white outlines follow each panel's true shape (rectangle/triangle/curve) and rotation
- Every panel is labelled (row/column, signal port, power port) in white, correctly positioned even on rotated or shaped panels
- A small on-canvas info panel shows resolution, physical size, panel count, grid size and the active test description
- The whole animation loops seamlessly every 20 seconds (verified bit-for-bit identical at the loop boundary) and always renders Front View, matching the PNG test pattern's convention
- "Download Video (WebM)" records exactly one loop as a native WebM file (no extra dependencies - browser MediaRecorder/canvas.captureStream) that plays back looped with no visible seam
- Works for uniform grids and freely placed/imported non-uniform layouts, including mixed MG9/MT and rotated/shaped panels, by defining the animation in wall pixel-coordinate space and revealing it through each panel's own clip mask

## Recent Changes In v0.16.0

Non-uniform layout overhaul (Stages 2-4) plus a round of fixes and new editing features, delivered as staged local commits:

**Free panel placement + import**
- Panels are no longer a fixed grid: place, drag, rotate, snap, join and multi-select panels freely, with overlap warnings and a live snap/join guide
- Import projects from the Creative Layout Tool, with a preview (name, panel mix, wall size) before replacing or adding as a new project
- Imported projects are interpreted and displayed as **Front View**, matching the original Creative Layout Tool design exactly (position, shape and rotation), instead of the app's default back/wiring view
- New save format v2 (free mm-positioned panel list); legacy grid-format settings files still open and migrate automatically
- Automatic letter-shaped patching (bottom-up, fork-aware) for text/logo-shaped layouts

**Editing and safety**
- Deleting panels now prompts with **Remove Panel**, **Mark as Inactive**, or **Cancel** — inactive panels stay visible (dashed) in place but are excluded from totals, patching and exports
- Keyboard shortcuts `S` (Select), `M` (Move), `P` (Patch) documented in Help, alongside the existing shortcuts

**Signal/power cable rendering**
- Cable lines now draw behind panels with a thin black outline; arrowheads draw in front, also black-outlined, and always point in the true signal/power direction (including when adjacent panels touch edge-to-edge)
- A selected panel is brought to the front, above cable lines, so its info stays readable
- Orthogonal (90°) cable routing everywhere: on-screen, PDF and PNG test pattern
- Snap/join logic ported from the Creative Layout Tool (connector-anchor based, shape/rotation aware)
- Signal/power chain-start indicator outlines now follow the true panel shape (triangle/curve/rect) at any rotation

**PNG test pattern export**
- Always renders Front View regardless of the on-screen toggle, matching what an observer sees standing in front of the finished wall
- No longer includes cable-routing lines or arrowheads
- Fixed panel alignment (true mm positions, no band-packing offset) and rotation accuracy
- Excludes inactive panels

**UI**
- New design-system `Button` component with clear active/selected states across all toolbar controls
- Panel Type control moved above Apply Grid Size; added Clear All Panels
- Renamed the import button to "Import Project from Creative Layout Tool"
- Dashed, wall-aligned background grid (1m major / 0.5m minor lines)

## Recent Changes In v0.15.0

Stage 1 of the non-uniform overhaul: interface refresh (data model unchanged).

- New shared design-system `Button` with consistent intents (primary / secondary / ghost / danger / success) and a clear active/selected state (bright fill + ring), replacing ad-hoc per-button colours
- Tools and modes now show an unmistakable active highlight: Signal / Power patch mode, Select mode, and view flip
- Controls grouped into labelled sections (Patch mode, Auto patching, Documentation & exports, Import & save, Selection & editing) with status chips for the active mode
- Cleaner cards, spacing and typography; extracted UI primitives into `src/components/ui.tsx`
- Cleanup: fixed the long-standing `useState` type warnings (typecheck now clean), tightened `patchMode`/`snakeDirection` types, removed a shadowed variable

## Recent Changes In v0.14.0

- Added chain-start indicators drawn alongside (never replacing) the existing panel outlines
- Blue outline on the first panel of each signal chain; orange outline on the first panel of each power chain
- A panel that starts both chains shows both outlines as clearly separated concentric rings (blue outer, orange inner)
- With "Do backup signal loop" enabled, the blue outline is also added to the last panel of each signal chain to show where the backup loop connects
- Indicators appear everywhere the layout is drawn: the live editor, printed page, PDF export, and the PNG test pattern

## Recent Changes In v0.13.0

- You can now mix `MG9` and `MT` panels in one wall. The layout is a 0.5m module grid: MG9 fills one module, MT spans two side-by-side modules
- Select Mode has a panel-type dropdown (MG9 / MT) to convert the selected panels; MT takes the module to its right and is blocked at the right edge
- Live stats are per panel type: panel count, weight, power/amps, and pixel resolution all use each panel's own profile
- Signal/power patching, auto-snake, and Match Power To Signal Pattern all treat an MT as a single panel and route cabling to its true edges
- Stock table is split by type and combined: MG9 (panels + variants) and MT each use their own catalog, spare ratio, box size, and hanging bar; MG9-only hardware (reinforcement, corner connectors, ground/floor frames) counts MG9 panels only
- PDF layout draws MT as a wide `(MT)`-labelled panel, with a mixed panel-type summary; PNG test pattern places each panel at its native pixel size (MG9 168x168, MT 256x64)
- Opening an older all-MT settings file migrates it onto the new module grid automatically

## Recent Changes In v0.12.0

- Added a `Match Power To Signal Pattern` button next to `Power Patch Mode`
- Power patching can now follow the existing signal patch: panels are powered in signal order (signal port, then sequence)
- Power plugs line up with the signal ports - each signal port starts on a fresh plug, giving a clean 1:1 plug-to-port mapping when a port fits in one plug
- Large signal ports spill onto consecutive plugs in order, and the tool still respects the power panel-count and 16A-per-plug limits

## Recent Changes In v0.11.0

- `MT` panels now render to their true 1m x 0.5m shape (2:1 wide rectangles) in the live panel layout instead of as squares
- PDF layout pages now draw `MT` panels as the same wide rectangles, with patching and power arrows routed correctly between them
- Cell width now scales from each panel profile's real-world width/height, so `MG9` stays square and `MT` is twice as wide as tall
- Confirmed the PNG test pattern exports at the correct native pixel ratio (256 x 64 per `MT` panel)

## Recent Changes In v0.10.1

- Restored patching arrows in the live panel layout and PDF layout pages
- Switched the test-pattern export from JPG to lossless PNG at true wall pixel dimensions
- Kept patching arrows and first-power markers out of the PNG test pattern only
- Improved Select Mode so panel editing does not accidentally patch panels
- Added undo/redo controls and shortcuts for layout edits
- Added a Help button with shortcut and workflow guidance
- Linked the version badge to this changelog
- Improved MG12 triangle, MG13 curved, and MG9 corner-panel drawing
- Changed corner-panel text to `Corner` and reduced corner hatching in the PNG export
- Updated connector stock values for `12260` and `12258`
- Reworked the PDF first page to show the full stock summary table before adding overflow stock pages

## Recent Changes In v0.10.0

- Added MG9-compatible special panel variants: MG12 Triangle, MG13 1/4 Curved, and MG9 LED Corner Panel
- Added persisted per-panel variant and rotation data in settings files
- Added drag selection for editing multiple panels at once
- Added multi-panel actions for changing panel type, rotating, clearing patching, deleting, and restoring panels
- Added keyboard shortcuts: `Delete` removes selected panels, `R` rotates, `C` clears selected patching, and `Escape` clears selection
- Removed the visible `Removed` label from deleted panels so holes stay visually blank
- Added a selected-port clear action for clearing the active signal port or power plug
- Backup signal loop now doubles the effective processor signal-port count while keeping the visible primary patch path readable
- Added corner-panel stock logic for corner panels, flat connectors, and corner connectors
- PDF Stock Summary now includes item names, item codes, required quantities, spare stock, rounded quantities, stock, and net stock
- Added JPG test-pattern export using the front view, true wall pixel dimensions, existing panel labels, port colors, panel shapes, hatches, and a 1px white border

## Recent Changes In v0.9.0

- Back view is now the default panel-layout view
- Panel layout clearly shows the current and alternate view
- PDF export now includes both `Back View` and `Front View` layout pages
- Load bars now stay orange when near the limit and only turn red when overloaded
- Added `LED Wall Deployment Settings`
- Added `Do backup signal loop`, enabled by default
- Added deployment types: `Flown`, `Ground`, `No Support`, `Floor`
- Backup signal loop now doubles `15m Signal Cable`
- Backup signal loop now adds `SEETRONIC SE8FF-05 F/M - F/M Joiner` per signal port with fallback to `SEETRONIC F/M - F/M Cable`
- Added MG9 ground and floor deployment stock calculations
- Settings export/import now includes deployment type and backup signal loop
- Added `Loop together` auto-snake preset
- Updated signal-port colors to match the NovaStar Unico look more closely
- Added stock CSV export and simplified stock table columns
- Added `12317 LED Prod Case` to every project stock list
- Added orange power-run start outlines on panel layout views
- Removed the dark background from panel-layout PDF exports
- Added removable and restorable panels for non-grid wall shapes
- Removed panels now skip patching, counts, stock, power, and support math
- Added per-panel `Clear Power And Signal Patching`, `Delete Panel`, and `Restore Panel` actions

## Local Development

```bash
npm install
npm run dev
```

Or double-click:

```text
start-local.bat
```

## Testing

```bash
npm test
```

Runs the Vitest suite (`src/**/*.test.ts`), including the NovaStar export tests, which compare generated files byte-for-byte against real NovaStar-exported reference projects checked into `src/novastar/__fixtures__/`.

## Production Build

```bash
npm run build
npm run preview
```

## GitHub Pages Deployment

This repo includes [`.github/workflows/deploy-pages.yml`](./.github/workflows/deploy-pages.yml).

To publish:

1. Push this folder to the GitHub repository.
2. In GitHub, open `Settings` -> `Pages`.
3. Set the source to `GitHub Actions`.
4. Push to `main` or rerun the Pages workflow.
5. Wait for the `Deploy GitHub Pages` workflow to finish.

The site uses a relative Vite base path so it works on repository Pages URLs.
