# LED Cabling Web App

Version `0.33.0`

Standalone React web app for planning LED wall layouts, signal port mapping, power outlet assignment, stock checks, deployment hardware, and PDF/settings/video exports.

## What It Does

- Build LED walls by rows and columns, or place panels freely (non-uniform layouts) with drag, edge-snap and joining
- Switch between `MG9` and `MT` panel profiles, plus `MG12` triangle and `MG13` curved variants
- Group panels into named **Sub-Screens** and edit/patch each one in isolation - every other panel (including unassigned ones) is fully hidden and un-interactive while a sub-screen is active, with its name clearly shown; **All Screens** returns to the complete layout instantly, with no panel data ever altered
- Position the whole layout or individual sub-screens within a configurable-resolution **Output Canvas** (drag, numeric entry, align/snap tools, boundary/overlap warnings) for multi-processor / media-server mapping
- Select a **NovaStar processor model** (`VX1000 Pro` / `VX2000 Pro`), see its live pixel/port capacity vs. current usage, assign a video input per sub-screen or one input for the whole canvas, and generate a real, importable `.uprj` processor configuration file - validated against the selected processor's port count, per-port and total pixel limits, and canvas size, with a summary of what will be exported and any blocking errors or warnings before download
- Import projects from the Creative Layout Tool
- Patch signal and power manually or with auto-snake / automatic letter-patching routing, scoped to the active sub-screen when one is selected. The first panel of each signal chain shows its port number in a blue circle (top-left) and the first panel of each power chain shows its port number in an orange circle (top-right), in both the Panel Layout and the PDF Report. With **Do backup signal loop** enabled, the chain's last panel also shows the backup port number (the second half of the available signal ports, e.g. port 11 backs up port 1 on a 20-port setup) - the number of signal ports itself follows the selected NovaStar processor (10 for VX1000 Pro, 20 for VX2000 Pro, 20 if none is selected), and the backup half is hatched and unselectable in the Signal Patching panel
- Flip the panel layout between `Back View` and `Front View`
- Export a PDF report with portrait detail pages, a per-sub-screen summary page, plus both layout views in landscape
- Export a native-resolution Test Pattern image, a full-screen canvas-only live Moving Test Pattern (top-left anchored, scaled to fit the browser window without stretching/cropping, with a small bouncing logo browser-side only), or a downloadable looping WebM or MP4 video of it
- Rotate panels by 45°, 90° or any custom angle, individually or as a multi-selected group (spacing/arrangement preserved); copy and paste panel groups with Ctrl/Cmd+C/V, with a cursor-following placement preview that snaps to the grid and nearby panels
- Toggle a vertical centre-line indicator on the Panel Layout (accounts for rotated panels' true outer bounds), optionally included in the PDF export
- Save and reopen settings as JSON (v5 format adds NovaStar processor/input selection; v3 sub-screens and output-canvas positioning, v2 free-panel and legacy grid formats still open)
- Check stock levels, shortfalls, and deployment hardware requirements, optionally overridden with live on-hand stock counts and date-range availability pulled from **Rentman** (see [Rentman Integration](#rentman-integration))
- Collapse any section of the UI to reduce clutter on long projects

## Recent Changes In v0.33.0

- Added a **Rentman Integration** card to Stock Calculations: map each stock item to its matching Rentman equipment record (searchable picker), then click Refresh to pull live on-hand stock counts, overriding the built-in catalog numbers everywhere they're used (on-screen table, CSV export, PDF report, Shortfalls card)
- Set a date range in the same card to also show an **Available (range)** column - stock minus whatever's already booked on other Rentman projects overlapping those dates - in the on-screen table and PDF stock table
- Equipment mappings are saved in the browser (per-machine, not per-project); the date range is saved with the project file (v6 format - older saves still open fine)
- Requires a small separate one-time deployment (a Cloudflare Worker that holds the Rentman API token server-side, since this app is a static site with no backend of its own) - see [Rentman Integration](#rentman-integration) below. Without it, the card just shows "Not configured" and everything else works exactly as before
- Investigated pushing planned items to an existing Rentman project by number, as originally requested, but Rentman's public API has no way to add equipment to an existing project today (their own roadmap lists it as not yet shipped) - dropped from scope, may revisit if Rentman adds it

## Recent Changes In v0.32.1

- Reverted MT's spare ratio back to 0% (a deliberate catalog choice, not an oversight - v0.32.0 had mistakenly changed it to match MG9's 7%)
- Reworded the PDF's "Boxes: X (Y spare in boxes)" line to "Boxes: X (Y additional spare)" - the old wording used "spare" twice in a confusing way

## Recent Changes In v0.32.0

- Reworked how spare panels are shown: a new "Spare Panels by Surface" breakdown lists panels used, spare (7% of used), and spare rounded up to a full box, for each surface (each Sub-Screen plus "Unassigned" if any panels aren't in one, or just "Whole Layout" with no Sub-Screens) and each panel type (MG9 Standard, MG9 Corner, MG9 Triangle, MG9 Curved, MT) - with a subtotal per surface and a grand total when there's more than one. MG9 Standard/Corner round up to boxes of 10 and MT rounds up to boxes of 6; shaped panels (Triangle/Curved) are bought individually so their spare is added as-is, unrounded. MT panels now also get the same 7% spare allowance as MG9 (previously 0%). Shown in both the web app's Stock Calculations card and its own page in the PDF report

## Recent Changes In v0.31.0

- Fixed the Quick Panel Layout -> Main Layout Tool hand-off: when sending into a tab that already has an existing project (the Replace/Add prompt), the panel type now actually switches to match, and Columns/Rows now update too - previously only the fresh-project hand-off applied these correctly
- Cut the full PDF report's file size dramatically (a large wall could reach ~20MB) by rendering the embedded Panel Layout images at a fixed 300 DPI for their actual printed size on the page, instead of a flat pixel multiplier that scaled with the wall's real-world size - a huge wall always prints at the same page-sized image regardless of how big it is, so the old approach wasted enormous, invisible resolution on big projects. Typical/small projects are unaffected (the same ~300 DPI they already got); a 1200-panel test wall dropped from a projected ~20MB+ to well under 1MB with no visible quality loss. Also enabled PDF stream compression on both PDF exports

## Recent Changes In v0.30.1

- Relabelled the stock table's "Rounded" column to "Rounded + Spare" (on-screen and PDF) to make clear it's the order quantity, spare included
- Added Spare Panels Needed and Rounded To Full Boxes to the standalone Quick Panel Layout tool (on-screen and its PDF export), using the same per-panel-type spare ratio and box size as the main tool's own stock maths

## Recent Changes In v0.30.0

- Reworked the required-stock calculation so **Required** is always the raw quantity needed to build the wall (no spare folded in), and **Rounded** consistently adds each item's spare (plus packaging rounding, e.g. boxes of 10 for MG9 panels) - fixing several rows (MG9/MT/corner/shaped panels, power cable, signal cable) that previously showed an already-spared number in "Required". Stock shortfalls are now checked against the real order quantity (Rounded), not the bare required count. Items whose final order quantity comes out to 0 are hidden from the on-screen table, the PDF stock table, and the CSV export (which now exports the Rounded order quantity, not the raw required count) - the on-screen table also gained Spare/Rounded/Stock columns to match the PDF
- Fixed the PDF's "Signal Ports In Use" and "Power Outputs In Use" boxes silently dropping any ports past about 7 with no indication (easy to hit - VX2000 Pro alone offers up to 20 signal ports). Those boxes now show a compact in-use count, and a new dedicated "Signal & Power Ports In Use" PDF page lists every used signal port and power outlet in full, with no truncation

## Recent Changes In v0.29.0

- Correctly handle MT's non-square LED pixel pitch (3.9mm horizontal x 7.8mm vertical - every second LED row is physically missing on this transparent panel) everywhere a wall's resolution or aspect ratio is shown. For an MT-only wall, Quick Panel Layout and the main tool's Wall Summary now show **LED Wall Resolution** (the panel's real 256x64 pixel grid), **Recommended Content Resolution** (LED height doubled, so content authored at this resolution has the correct proportions), and a corrected **Physical Aspect Ratio** (from the wall's true physical size, not its raw pixel grid) - previously the "Aspect Ratio" stat was silently wrong for MT walls (e.g. showing 8:1 for a wall that's physically 4:1), and Quick Panel Layout's preview diagram and 16:9 content-area overlay were the wrong shape too. MG9 walls are unaffected, since its pixel and physical aspect ratios already match. The PDF exports from both tools, and the on-screen info text baked into the Test Pattern image/video, all reflect the same distinction for MT.

## Recent Changes In v0.28.2

- Added a "Show 16:9 content area" checkbox to the standalone Quick Panel Layout tool, controlling the dashed overlay box, its up/down nudge controls, its resolution stat, and its Full-HD warning, on both the web page and the PDF export - when shown, the PDF now also includes a short explanation of what the dashed box means
- Added an optional Project Name field to Quick Panel Layout, shown as a header on its PDF export and carried forward to become the main tool's Project Name when using **Send to Main Layout Tool**

## Recent Changes In v0.28.1

- Tidied the Quick Panel Layout PDF export: stats are now grouped under clear "Panel", "Power" and "Weight" section headers in a compact multi-column layout, replacing the old flat list that ran off the bottom of the page

## Recent Changes In v0.28.0

- Moved the orange power-port badge to sit directly beside the blue signal-port badge(s), all in the panel's top-left corner (was previously in the opposite corner) - neatly spaced, non-overlapping, in both the Panel Layout and the PDF Report
- Added an automatic weight estimate to the standalone Quick Panel Layout tool: panel weight, flying hardware (fly bars + slings/shackles, assuming a Flown deployment), estimated cable weight (power + signal, assuming cables snake left-to-right and alternate direction each row), and total flown weight - no manual rigging or cable input needed, clearly labelled as an indicative estimate rather than a certified rigging calculation, and included in its PDF export

## Recent Changes In v0.27.1

- Restored the shape-hugging signal/power chain-start ring indicators (removed in v0.27.0) - they're now shown together with the new port-number badges, not instead of them
- Moved each panel's info text (row/column label, assigned signal/power port, shape symbol) to the bottom of the panel, in both the Panel Layout and the PDF Report, to keep the top corners clear for the port-number badges
- The NovaStar processor model now defaults to VX2000 Pro for a new project instead of none selected

## Recent Changes In v0.27.0

- Added port-number badges to the first panel of every signal and power chain, in both the Panel Layout and the PDF Report: a blue circle (top-left) with the signal port number, and an orange circle (top-right) with the power port number - replacing the old shape-hugging "chain start" ring outlines. Badges are drawn outside the panel's own rotate transform so the digit stays upright and legible even on a rotated panel, and follow the panel when it's moved, rotated, or the layout is exported
- Added backup signal port numbering: with **Do backup signal loop** enabled, the chain's last panel also gets a badge with the backup port number - the second half of the available signal ports backs up the first half (e.g. port 11 backs up port 1 on a 20-port setup; port 6 backs up port 1 on a 10-port setup)
- The number of selectable signal ports now follows the selected NovaStar processor model (10 for VX1000 Pro, 20 for VX2000 Pro) instead of always offering 20 regardless of processor; falls back to 20 when no processor is selected
- With the backup signal loop enabled, the second half of the port range is reserved for backups: hatched and unselectable in the Signal Patching panel, and automatically excluded from Auto Snake / manual port assignment, so primary signal-port capacity updates automatically

## Recent Changes In v0.26.0

- Added a **Power** summary to the standalone Quick Panel Layout tool: total power draw (max/avg W and A) plus, for both a 32A and a 63A distro, the circuits needed, distro units needed, and percentage of safe capacity used - reusing the same per-panel power spec and safe-panels-per-outlet defaults as the main Layout Tool, also included in its PDF export
- Added +/- stepper buttons next to the Width (m) and Height (m) fields in Quick Panel Layout, matching the existing Columns/Rows steppers, for easier use on mobile/touch

## Recent Changes In v0.25.1

- Fixed a real regression from v0.25.0's DPI-aware live Moving Test Pattern rendering: the RGB checkerboard tiles no longer lined up with panel boundaries (each flat-coloured square could span multiple panels). Root cause was `drawTestPatternFrame` resetting the canvas transform to a hard-coded identity to draw its pre-rendered pattern layer, which only stays correct when the canvas is 1:1 with its content (true for the recorded video/PNG/PDF exports) - once the live view started rendering at a devicePixelRatio/fit-to-window scale, that hard reset drew the pattern at the wrong size. It now resets to whichever transform the caller had active instead of a literal identity
- Regrouped the Panel Layout toolbar: **Patch**, **Select**, **Move** and **Clear Patching** now live together in one "Panel tools" group (previously split across two groups), with the Snap/Move joined group/Allow overlaps options staying alongside Move

## Recent Changes In v0.25.0

- Doubled the bouncing MMS logo's size (now ~1/2 a standard panel's native pixel width, up from 1/4) and halved its movement speed in the browser-only live Moving Test Pattern
- The "Include Centre Line" PDF export option now defaults to enabled and sits at the end of its toolbar row
- Replaced the separate WebM/MP4 download buttons with a single "Download Moving Test Pattern" button that opens a format-choice dialog (WebM or MP4, with a clear warning that MP4 takes significantly longer to encode) with Download/Cancel
- Fixed blurry/sub-pixel text, lines and arrows in the live Moving Test Pattern tab: the canvas's backing store is now sized to the actual physical device pixels it's displayed at (CSS size x devicePixelRatio), not just the wall's native resolution, so high-DPI displays and non-1:1 window scaling no longer blur or alias thin strokes
- Fixed rotated-panel snapping: panels rotated to a custom angle (not just 0/90/180/270) now snap and join along their own true rotated edges instead of silently snapping as if they were unrotated. This was a real bug in `panelWorldAnchors` (it rounded rotation to the nearest 90deg before computing connector anchor positions) affecting individual panels, multi-selected groups, custom angles and imported rotated panels alike; covered by a new regression test
- Added a **Fit to View** button that zooms the Panel Layout workspace so the entire project - including a wide or tall imported layout - is visible at once, with the scroll position reset to the origin; confirmed the existing scrollable workspace already expands and scrolls correctly for any project size
- Reorganised the Panel Layout toolbar into labelled groups (Selection & editing, Move/align & snap, Rotation & transforms, View/zoom & navigation, Overlays & display) instead of one long unlabelled row - all existing controls preserved, including moving the Front/Back View and Centre Line toggles down from the card header into their matching groups

## Recent Changes In v0.24.0

- Rounded the measurements shown in the Panel Layout header (and the Wall Summary/PDF size lines) to at most 2 decimal places with trailing zeros trimmed, instead of the raw unrounded floating-point value
- Expanded panel rotation: dedicated 45° and 90° buttons plus a custom-angle input, on top of the existing keyboard shortcut. Rotating a multi-selection spins every selected panel in place by the same amount, so their arrangement and spacing relative to each other never changes
- Fixed layout imports from the Creative Layout Tool silently snapping every panel's rotation to the nearest 90° and mis-mirroring square panels' rotation under the front-view flip - both bugs only showed up once panels could be rotated to non-cardinal angles. Imported panels now keep their exact source rotation and orientation
- The NovaStar Processor Configuration section is now collapsed by default and moved to the very bottom of the sidebar, out of the way during normal layout work
- Added copy/paste for selected panels: Ctrl/Cmd+C copies, Ctrl/Cmd+V enters a paste-placement mode with a dashed preview that follows the cursor and snaps to the grid/nearby panels, click to place, Escape or right-click to cancel. Pasted panels are auto-selected and keep their source spacing, rotation, panel type and arrangement (patching is left unassigned, same as importing a layout)
- Applying a new grid size over an existing layout now asks for confirmation ("Remove Panels and Apply Grid" / "Cancel") instead of silently wiping every panel
- The Moving Test Pattern's per-panel row/column direction indicators are now drawn as vector arrows (explicit stroke width with a floor, arrowhead scaled to match) instead of relying on a font glyph's own internal strokes, which could shrink below a visible pixel width at small panel sizes or when the output is scaled
- Added a small bouncing MMS logo (DVD-screensaver style) to the browser-only live Moving Test Pattern view - about a quarter of a panel's native width, aspect-ratio preserved, stays fully inside the canvas, never appears in the recorded video or PNG/PDF exports
- Added a toggleable vertical centre-line indicator to the Panel Layout workspace, with a "Centre" label and an "Include Centre Line" option for the PDF export. The centre is computed from every active panel's true rotated outer bounds, not just the axis-aligned wall bounding box, so a panel spun to a non-cardinal angle is still accounted for correctly
- Selecting a sub-screen now fully hides every panel not assigned to it (previously they stayed visible, dimmed and locked) - the active sub-screen's name is shown clearly above the workspace, and an "All Screens" option (renamed from "Canvas View") returns to the complete layout instantly. No panel data is ever changed by switching the visible screen
- Renamed **Video Test Pattern** to **Moving Test Pattern** and dropped "PNG" from **PNG Test Pattern** (now just **Test Pattern**) everywhere the labels appear - buttons, tab titles, help text and the README
- The browser's live Moving Test Pattern tab now anchors the LED canvas to the top-left corner and scales it to fit the window on both axes (never stretched, cropped, centred, or auto-rotated between portrait/landscape), recalculating on resize; unused space fills with a plain black background instead of centring the canvas
- Added a **Download MP4** option next to the existing WebM download. MediaRecorder can't produce MP4 directly in most browsers, so this records the same WebM as today and then transcodes it to H.264 MP4 in the browser via a lazily-loaded ffmpeg.wasm (only fetched the first time this button is used - roughly 30MB, entirely separate from the app's normal bundle)

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

## Rentman Integration

Optional. Lets Stock Calculations pull live on-hand stock counts and date-range availability from [Rentman](https://www.rentman.io/). This app is a static site with no backend, and the Rentman API token must never end up in anything shipped to the browser - so this works through a small separate Cloudflare Worker that holds the token server-side and proxies read-only requests to Rentman.

Setup (one-time):

1. Deploy the Worker - see [`rentman-proxy/README.md`](./rentman-proxy/README.md) for the full steps (`wrangler secret put`, `wrangler deploy`).
2. Copy [`.env.example`](./.env.example) to `.env` for local dev, and/or add a `RENTMAN_PROXY_URL` repository **variable** (not secret) under `Settings -> Secrets and variables -> Actions -> Variables`) so the GitHub Pages build picks it up via [`deploy-pages.yml`](./.github/workflows/deploy-pages.yml).
3. Rebuild/redeploy. The "Rentman Integration" card (in Stock Calculations) will show as configured; map each stock item to its Rentman equipment and click Refresh.

Left unset, the card just shows "Not configured" and Stock Calculations keeps using its built-in numbers - nothing else changes.
