// Read-only proxy in front of the Rentman API (https://api.rentman.net).
// Exists solely so RENTMAN_API_TOKEN never has to live in led-cabling-web's
// source (that's a public static site - anything in src/ ships to every
// visitor's browser). This Worker holds the token as an encrypted secret and
// exposes only the three narrow, purpose-built routes the app actually
// needs - never a generic Rentman passthrough.
//
// Confirmed against Rentman's real OpenAPI spec (linked from
// https://api.rentman.net/), not guessed:
//   - GET /equipment supports ?fields=..., ?code=a,b,c (comma = OR, verified
//     live), and GENERATED fields like current_quantity need to be listed in
//     ?fields explicitly or they're omitted.
//   - GET /projectequipment rows carry planperiod_start/planperiod_end
//     (ISO 8601) and an `equipment` field shaped "/equipment/<id>" (a link
//     string, not expanded by default). Filtering planperiod_start[lte]=X&
//     planperiod_end[gte]=Y correctly implements a date-range overlap query
//     (verified live against a known booking).
//   - There is NO write endpoint anywhere in the spec for adding equipment
//     to an existing project (/projectequipment has no POST/PUT/PATCH) - so
//     this Worker deliberately has zero write routes. See the plan doc for
//     the full investigation.

export interface Env {
  /** Secret - set via `wrangler secret put RENTMAN_API_TOKEN`, never a var. */
  RENTMAN_API_TOKEN: string;
  /** The led-cabling-web deployed origin, e.g. "https://user.github.io". */
  ALLOWED_ORIGIN: string;
}

const RENTMAN_BASE = "https://api.rentman.net";
// Rentman's own documented max; the whole catalog fits in one page.
const MAX_LIMIT = 1500;

function corsHeaders(env: Env): Record<string, string> {
  return {
    "Access-Control-Allow-Origin": env.ALLOWED_ORIGIN,
    "Access-Control-Allow-Methods": "GET, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Access-Control-Max-Age": "86400",
  };
}

function json(body: unknown, env: Env, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json", ...corsHeaders(env) },
  });
}

async function rentmanGet(env: Env, path: string): Promise<any> {
  const res = await fetch(`${RENTMAN_BASE}${path}`, {
    headers: {
      Authorization: `Bearer ${env.RENTMAN_API_TOKEN}`,
      Accept: "application/json",
    },
  });
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`Rentman ${path} -> HTTP ${res.status}${body ? `: ${body.slice(0, 300)}` : ""}`);
  }
  return res.json();
}

function parseListParam(url: URL, name: string): string[] {
  const raw = url.searchParams.get(name) || "";
  return raw
    .split(",")
    .map((v) => v.trim())
    .filter(Boolean);
}

/** Every equipment code the app cares about, resolved to its Rentman id/name/stock in one call. */
async function resolveEquipmentByCode(
  env: Env,
  codes: string[],
): Promise<Record<string, { id: number; name: string; currentQuantity: number }>> {
  const data = await rentmanGet(
    env,
    `/equipment?code=${encodeURIComponent(codes.join(","))}&fields=id,code,name,current_quantity&limit=${MAX_LIMIT}`,
  );
  const byCode: Record<string, { id: number; name: string; currentQuantity: number }> = {};
  for (const item of data.data || []) {
    byCode[item.code] = { id: item.id, name: item.name, currentQuantity: item.current_quantity ?? 0 };
  }
  return byCode;
}

async function handleEquipmentStock(env: Env, url: URL): Promise<Response> {
  const codes = parseListParam(url, "codes");
  if (!codes.length) return json({ error: "codes query param is required (comma-separated)" }, env, 400);

  const byCode = await resolveEquipmentByCode(env, codes);
  const result: Record<string, { name: string; currentQuantity: number } | null> = {};
  codes.forEach((code) => {
    const eq = byCode[code];
    result[code] = eq ? { name: eq.name, currentQuantity: eq.currentQuantity } : null;
  });
  return json(result, env);
}

async function handleEquipmentAvailability(env: Env, url: URL): Promise<Response> {
  const codes = parseListParam(url, "codes");
  const from = url.searchParams.get("from");
  const to = url.searchParams.get("to");
  if (!codes.length || !from || !to) {
    return json({ error: "codes, from, and to query params are all required" }, env, 400);
  }

  const byCode = await resolveEquipmentByCode(env, codes);
  const ids = Object.values(byCode).map((eq) => eq.id);

  const result: Record<string, number | null> = {};
  if (!ids.length) {
    codes.forEach((code) => (result[code] = null));
    return json(result, env);
  }

  // Every projectequipment row for these equipment ids whose plan period
  // overlaps [from, to] - classic interval-overlap filter, verified live.
  const linkValues = ids.map((id) => `/equipment/${id}`).join(",");
  const peData = await rentmanGet(
    env,
    `/projectequipment?equipment=${encodeURIComponent(linkValues)}` +
      `&planperiod_start[lte]=${encodeURIComponent(to)}&planperiod_end[gte]=${encodeURIComponent(from)}` +
      `&fields=equipment,quantity_total&limit=${MAX_LIMIT}`,
  );

  const bookedById: Record<number, number> = {};
  for (const row of peData.data || []) {
    const match = /\/equipment\/(\d+)/.exec(row.equipment || "");
    if (!match) continue;
    const id = Number(match[1]);
    bookedById[id] = (bookedById[id] || 0) + (Number(row.quantity_total) || 0);
  }

  codes.forEach((code) => {
    const eq = byCode[code];
    if (!eq) {
      result[code] = null;
      return;
    }
    const booked = bookedById[eq.id] || 0;
    result[code] = Math.max(0, eq.currentQuantity - booked);
  });
  return json(result, env);
}

/**
 * Rentman's confirmed filter operators (lt/gt/lte/gte/neq/isnull) don't
 * include a documented fuzzy "contains" match, so this fetches a page of
 * the catalog (up to MAX_LIMIT) and filters in-Worker rather than relying
 * on unconfirmed server-side text search. Fine at this scale - the whole
 * point is a one-time equipment-mapping picker, not a live product search.
 */
async function handleEquipmentSearch(env: Env, url: URL): Promise<Response> {
  const query = (url.searchParams.get("query") || "").trim().toLowerCase();
  const data = await rentmanGet(env, `/equipment?fields=id,code,name&limit=${MAX_LIMIT}`);
  // Rentman always includes baseline fields (created/modified/creator/custom/
  // etc.) regardless of ?fields - map down to exactly what the mapping UI
  // needs rather than forwarding Rentman's internal metadata verbatim.
  const all: Array<{ id: number; code: string; name: string }> = (data.data || []).map((eq: any) => ({
    id: eq.id,
    code: eq.code,
    name: eq.name,
  }));
  const filtered = query ? all.filter((eq) => eq.name.toLowerCase().includes(query) || eq.code.toLowerCase().includes(query)) : all;
  return json(filtered.slice(0, 50), env);
}

export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    if (request.method === "OPTIONS") {
      return new Response(null, { status: 204, headers: corsHeaders(env) });
    }
    if (request.method !== "GET") {
      return json({ error: "Only GET is supported - this proxy is read-only." }, env, 405);
    }

    const url = new URL(request.url);
    try {
      switch (url.pathname) {
        case "/equipment-stock":
          return await handleEquipmentStock(env, url);
        case "/equipment-availability":
          return await handleEquipmentAvailability(env, url);
        case "/equipment-search":
          return await handleEquipmentSearch(env, url);
        default:
          return json({ error: "Not found" }, env, 404);
      }
    } catch (err) {
      return json({ error: err instanceof Error ? err.message : "Unknown error" }, env, 502);
    }
  },
};
