# LED Cabling Web App

Standalone React web app for planning LED wall layouts, signal port mapping, power outlet assignment, stock checks, and PDF/JSON exports.

## Features

- Build wall layouts by rows and columns
- Switch between `MG9` and `MT` panel profiles
- Patch signal and power manually or with automatic snake routing
- Track per-port load, phase load, wall resolution, and weight
- Export plans to JSON
- Generate printable PDF reports
- Deploy to GitHub Pages with the included Actions workflow

## Local development

```bash
npm install
npm run dev
```

## Production build

```bash
npm run build
npm run preview
```

## GitHub Pages deployment

This repo includes [`.github/workflows/deploy-pages.yml`](./.github/workflows/deploy-pages.yml).

To publish:

1. Push this folder to a GitHub repository.
2. In the repository on GitHub, open `Settings` -> `Pages`.
3. Set the source to `GitHub Actions`.
4. Push to the `main` branch.
5. Wait for the `Deploy GitHub Pages` workflow to finish.

The site build uses a relative Vite base path, so it works for both repository Pages and a user/site root.
