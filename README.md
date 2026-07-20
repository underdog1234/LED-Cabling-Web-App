# LED Cabling Web App

Version `0.19.1`

Standalone React web app for planning LED wall layouts, signal port mapping, power outlet assignment, stock checks, deployment hardware, and PDF/settings/video exports.

## What It Does

- Build LED walls by rows and columns, or place panels freely (non-uniform layouts) with drag, edge-snap and joining
- Switch between `MG9` and `MT` panel profiles, plus `MG12` triangle and `MG13` curved variants
- Import projects from the Creative Layout Tool
- Patch signal and power manually or with auto-snake / automatic letter-patching routing
- Flip the panel layout between `Back View` and `Front View`
- Export a PDF report with portrait detail pages plus both layout views in landscape
- Export a native-resolution PNG test pattern, a full-screen canvas-only live animated test pattern, or a downloadable looping WebM video of it
- Save and reopen settings as JSON (v2 free-panel format, with legacy grid migration)
- Check stock levels, shortfalls, and deployment hardware requirements

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
