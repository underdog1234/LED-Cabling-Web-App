# rentman-proxy — step-by-step setup

## What is this, in plain terms?

The main site (`led-cabling-web`) is just static files with no server behind it — anything placed in its code is visible to anyone who visits the site (view source / browser dev tools). Your Rentman API token can't go there.

So instead, this folder is a *tiny separate program* ("a Worker") that lives on Cloudflare, a free hosting service for exactly this kind of thing. It holds your Rentman token privately and answers three simple questions when the main site asks:

1. "What's the current stock for these items?"
2. "What's available between these two dates?"
3. "Search Rentman equipment matching this text" (for the mapping picker)

It never writes anything to Rentman — read-only, one-way.

You only need to set this up **once**. After that, it just runs quietly in the background. Nothing here costs money (Cloudflare's free tier is far more than this needs).

Total time: about 10–15 minutes.

---

## Before you start

You'll need:

- A free Cloudflare account (if you don't already have one — step 1 below covers it)
- Your Rentman API token. You already gave me one earlier in this chat, and it's saved on your computer at `rentman-proxy/.dev.vars` — open that file in any text editor to copy it back out when step 5 asks for it. (If you'd rather use a fresh one, Rentman → Settings → API/Integrations has it.)
- A terminal open in the `rentman-proxy` folder — if you're reading this via Claude, just ask Claude to run each command for you instead of typing it yourself; that's the easiest path.

---

## Step 1 — Create a free Cloudflare account

Skip this if you already have one.

1. Go to [dash.cloudflare.com/sign-up](https://dash.cloudflare.com/sign-up)
2. Enter your email and a password, verify your email
3. That's it — no card details needed for the free tier

## Step 2 — Install the tools

This step is likely already done (Claude ran it while building this feature), but for completeness, from inside the `rentman-proxy` folder:

```bash
npm install
```

You'll see a line like `added 39 packages` — that means it worked.

## Step 3 — Connect your terminal to your new Cloudflare account

```bash
npx wrangler login
```

This opens a browser tab asking you to log in to Cloudflare and click **Allow**. Once you see "Successfully logged in" in the terminal, come back here.

*(If you'd rather have Claude do everything from Step 4 onward for you: do this step yourself since it's your personal account login, then tell Claude "I've logged in" and it can run the rest.)*

## Step 4 — Store your Rentman token securely

```bash
npm run secret:token
```

This will ask:

```
Enter a secret value:
```

Paste your Rentman API token (from `.dev.vars`, or a fresh one from Rentman) and press Enter. **Nothing will appear on screen as you paste/type — that's normal**, it's hidden on purpose like a password field. You should then see something like `✨ Success! Uploaded secret RENTMAN_API_TOKEN`.

This value is now stored encrypted on Cloudflare's servers — it is never written into any file in this project.

## Step 5 — Check who's allowed to use this (already done, just a check)

Open [`wrangler.toml`](./wrangler.toml) in this folder. It should already contain:

```toml
[vars]
ALLOWED_ORIGIN = "https://underdog1234.github.io"
```

This is already set correctly for your site (`github.com/underdog1234/LED-Cabling-Web-App`, published at `underdog1234.github.io`) — you don't need to change anything here. (It just means: only your published site is allowed to call this Worker from a browser, not just anyone. If you ever move the site to a different address, this is the line you'd update.)

## Step 6 — Publish (deploy) the Worker

```bash
npm run deploy
```

After a few seconds you'll see output ending with something like:

```
Uploaded led-cabling-rentman-proxy
Deployed led-cabling-rentman-proxy triggers
  https://led-cabling-rentman-proxy.<your-subdomain>.workers.dev
```

**Copy that `https://...workers.dev` address** — that's your Worker's web address, and you'll need it in the next step. It's not a secret (it's just a URL), so it's fine to paste it anywhere.

## Step 7 — Tell your GitHub-published site about that address

The main site needs to know that address so it can call it. This is set as a GitHub **repository variable** (not a secret — it's just a public URL):

1. Go to your repository's variables page: [github.com/underdog1234/LED-Cabling-Web-App/settings/variables/actions](https://github.com/underdog1234/LED-Cabling-Web-App/settings/variables/actions)
2. Click **New repository variable**
3. Name: `RENTMAN_PROXY_URL`
4. Value: paste the `https://...workers.dev` address from Step 6
5. Click **Add variable**

## Step 8 — Rebuild the published site

The site only picks up that address when it's rebuilt. Either:

- Push any small change to `main` (this triggers it automatically), **or**
- Go to the [**Actions** tab](https://github.com/underdog1234/LED-Cabling-Web-App/actions) → **Deploy GitHub Pages** (left sidebar) → **Run workflow** button → **Run workflow**

Wait for it to finish (green checkmark, usually under a minute).

## Step 9 — Check it worked

Open your live site, expand **Stock Calculations**, then expand the **Rentman Integration** card. The amber "Not configured" message should be gone. Map an item to a Rentman equipment record and click **Refresh Stock** — you should see real numbers appear.

You're done! This was all one-time setup — you won't need to repeat any of it unless you rotate your token or move the site to a new address.

---

## If something's not working

- **"Not configured" banner still shows after Step 8** — the site only reads the address at build time, so it needs an actual rebuild *after* the variable was added (Step 8 again). Double-check the variable name is exactly `RENTMAN_PROXY_URL`, no typos.
- **A red error message in the Rentman Integration card mentioning "502" or "Rentman proxy request failed"** — the stored Rentman token is likely wrong or expired. Repeat Step 4 with a fresh token from Rentman → Settings → API.
- **A browser error mentioning "CORS" or "blocked"** — the site's address doesn't match `ALLOWED_ORIGIN` in `wrangler.toml`. Fix the value (Step 5) and redeploy (Step 6).
- **`wrangler` says "You are not authenticated"** — run Step 3 (`npx wrangler login`) again; login sessions occasionally expire.

---

## Maintenance (only if needed later)

- **Rotate your Rentman token**: just repeat Step 4 (`npm run secret:token`) with the new value — no redeploy or rebuild needed.
- **Changed the Worker's code** (`src/index.ts`): repeat Step 6 (`npm run deploy`) to publish the change. The address stays the same, so no further steps needed.
- **Moved the site to a new address**: update `ALLOWED_ORIGIN` in `wrangler.toml` (Step 5) and redeploy (Step 6).

---

## Technical reference

For anyone editing the Worker's code rather than just deploying it:

This Worker is a separate deployable from the main site — it is **not** built or deployed by `npm run build` / the GitHub Pages workflow. It has no write access to Rentman at all; all three endpoints are read-only GETs:

- `GET /equipment-stock?codes=a,b,c` → `{ [code]: { name, currentQuantity } | null }`
- `GET /equipment-availability?codes=a,b,c&from=YYYY-MM-DD&to=YYYY-MM-DD` → `{ [code]: number | null }` (on-hand stock minus quantity already booked on other Rentman projects whose plan period overlaps the given range)
- `GET /equipment-search?query=...` → `[{ id, code, name }]` (used by the mapping picker in the app)

All three require a `codes`/`query` param, are CORS-restricted to `ALLOWED_ORIGIN`, and return `502` with an error message if the underlying Rentman API call fails.

### Local development

```bash
npm run dev
```

Runs `wrangler dev`, which reads secrets/vars from a local `.dev.vars` file (gitignored — create it with `RENTMAN_API_TOKEN=...`, one `KEY=value` per line). Note `[vars]` in `wrangler.toml` takes precedence over `.dev.vars` for the same key, so to test CORS locally, temporarily point `ALLOWED_ORIGIN` in `wrangler.toml` at `http://localhost:5173` and revert it before deploying. For the main site to reach this locally-running Worker instead of a deployed one, copy `../.env.example` to `../.env` and set `VITE_RENTMAN_PROXY_URL` to the local address `wrangler dev` prints (typically `http://localhost:8787`).
