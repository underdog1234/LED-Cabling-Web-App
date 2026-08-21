# rentman-proxy

A small read-only [Cloudflare Worker](https://workers.cloudflare.com/) that sits between the `led-cabling-web` static site and the [Rentman](https://www.rentman.io/) API.

`led-cabling-web` is a 100% static, GitHub-Pages-deployed site with no backend of its own - everything under its `src/` ships as plain text to every visitor's browser. The Rentman API token can't live there. This Worker holds the token as an encrypted secret and exposes three narrow, read-only endpoints the frontend calls instead of Rentman directly. It has no write access to Rentman at all.

This is a separate deployable from the main site - it is **not** built or deployed by `npm run build` / the GitHub Pages workflow. Deploy it once (and again whenever `src/index.ts` changes) with `wrangler deploy` from this folder.

## Endpoints

- `GET /equipment-stock?codes=a,b,c` -> `{ [code]: { name, currentQuantity } | null }`
- `GET /equipment-availability?codes=a,b,c&from=YYYY-MM-DD&to=YYYY-MM-DD` -> `{ [code]: number | null }` (on-hand stock minus quantity already booked on other Rentman projects whose plan period overlaps the given range)
- `GET /equipment-search?query=...` -> `[{ id, code, name }]` (used by the mapping picker in the app)

All three require a `codes`/`query` param, are CORS-restricted to the `ALLOWED_ORIGIN` configured below, and return `502` with an error message if the Rentman API call itself fails.

## One-time setup

Requires a [Cloudflare account](https://dash.cloudflare.com/sign-up) (the free tier is enough) and a Rentman API token (Rentman -> Settings -> API).

```bash
cd rentman-proxy
npm install
npx wrangler login
```

Set the Rentman token as an encrypted Worker secret - **never** put it in `wrangler.toml` or any other tracked file:

```bash
npm run secret:token
```

This prompts for the token value and stores it encrypted server-side.

Open [`wrangler.toml`](./wrangler.toml) and set `ALLOWED_ORIGIN` to the exact origin (scheme + host, no path, no trailing slash) the site is deployed at, e.g. `https://<your-github-username>.github.io`. This isn't secret, just a `[vars]` entry - it controls the `Access-Control-Allow-Origin` response header so only your deployed app (not `*`) can call this Worker from a browser.

## Deploy

```bash
npm run deploy
```

Prints the Worker's URL (`https://led-cabling-rentman-proxy.<your-subdomain>.workers.dev` by default - rename via `name` in `wrangler.toml` if you want something else). Configure the main site to use it:

- **Local dev**: copy `../.env.example` to `../.env` and set `VITE_RENTMAN_PROXY_URL` to that URL.
- **Deployed site**: add a `RENTMAN_PROXY_URL` repository **variable** (not a secret - it's just a public URL) under the GitHub repo's `Settings -> Secrets and variables -> Actions -> Variables`, matching what [`deploy-pages.yml`](../.github/workflows/deploy-pages.yml) reads. Rerun the Pages workflow (or push) to rebuild with it.

Redeploy (`npm run deploy`) any time `src/index.ts` changes. Rotate the token at any time by rerunning `npm run secret:token` - no redeploy needed.

## Local development

```bash
npm run dev
```

Runs `wrangler dev`, which reads secrets/vars from a local `.dev.vars` file (gitignored - create it with `RENTMAN_API_TOKEN=...`, one `KEY=value` per line; see `wrangler.toml`'s comments for what's needed). Note `[vars]` in `wrangler.toml` takes precedence over `.dev.vars` for the same key, so to test CORS locally, temporarily point `ALLOWED_ORIGIN` in `wrangler.toml` at `http://localhost:5173` and revert it before deploying.
