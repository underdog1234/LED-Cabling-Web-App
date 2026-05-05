# LED Cabling Web App

Version `0.10.0`

Standalone React web app for planning LED wall layouts, signal port mapping, power outlet assignment, stock checks, deployment hardware, and PDF/settings exports.

## What It Does

- Build LED walls by rows and columns
- Switch between `MG9` and `MT` panel profiles
- Patch signal and power manually or with auto-snake routing
- Flip the panel layout between `Back View` and `Front View`
- Export a PDF report with portrait detail pages plus both layout views in landscape
- Save and reopen settings as JSON
- Check stock levels, shortfalls, and deployment hardware requirements

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
