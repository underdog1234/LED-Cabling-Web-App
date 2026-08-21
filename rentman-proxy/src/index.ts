// Read-only proxy in front of the Rentman API (https://api.rentman.net).
// Exists solely so RENTMAN_API_TOKEN never has to live in led-cabling-web's
// source (that's a public static site - anything in src/ ships to every
// visitor's browser). This Worker holds the token as an encrypted secret and
// exposes only the two narrow, purpose-built routes the app actually
// needs - never a generic Rentman passthrough.
//
// Confirmed against Rentman's real OpenAPI spec (linked from
// https://api.rentman.net/), not guessed:
//   - GET /equipment supports ?fields=..., ?code=a,b,c (comma = OR, verified
//     live), and GENERATED fields like current_quantity need to be listed in
//     ?fields explicitly or they're omitted. The app sends its own stock
//     codes directly as Rentman equipment codes - no separate mapping step,
//     since this catalog's codes already are Rentman's real equipment codes.
//   - GET /projectequipment rows carry planperiod_start/planperiod_end
//     (ISO 8601) and an `equipment` field shaped "/equipment/<id>" (a link
//     string, not expanded by default). Filtering planperiod_start[lte]=X&
//     planperiod_end[gte]=Y correctly implements a date-range overlap query
//     (verified live against a known booking), and equipment[[in]]-style
//     comma-separated `equipment=` filtering works as OR, same as `code=`
//     on /equipment.
//   - /projectequipment has NO direct link to its project. The real chain,
//     verified live: `equipment_group` (link) -> /projectequipmentgroup/{id}
//     has both `project` (link) and `subproject` (link) -> /subprojects/{id}
//     has `status` (link) -> /statuses/{id} has the actual status name (e.g.
//     "Confirmed"). Project itself has NO status field at all (confirmed:
//     ?fields=status on /projects 400s with "Unknown field"). Nested expand
//     works up to 3 levels and combines fine with filters:
//     `expand=equipment_group.project,equipment_group.subproject.status`
//     returns project name/number and subproject status inline in one
//     request. `fields` dot-paths cap at 3 segments too, so
//     `equipment_group.subproject.status` (the whole small status object)
//     is requested rather than the 4-segment `...status.name`.
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

/** Every equipment code the app cares about, resolved to its Rentman id/name/stock in one call - the app's own stock codes ARE Rentman's equipment codes, so this is a direct lookup, not a fuzzy match. */
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

type AvailabilityProject = {
  projectName: string;
  projectNumber: number | string;
  status: string | null;
  quantity: number;
  planPeriodStart: string;
  planPeriodEnd: string;
};

async function handleEquipmentAvailability(env: Env, url: URL): Promise<Response> {
  const codes = parseListParam(url, "codes");
  const from = url.searchParams.get("from");
  const to = url.searchParams.get("to");
  if (!codes.length || !from || !to) {
    return json({ error: "codes, from, and to query params are all required" }, env, 400);
  }

  const byCode = await resolveEquipmentByCode(env, codes);
  const ids = Object.values(byCode).map((eq) => eq.id);

  const byEquipmentId: Record<number, AvailabilityProject[]> = {};
  if (ids.length) {
    // Every projectequipment row for these equipment ids whose plan period
    // overlaps [from, to] - classic interval-overlap filter, verified live -
    // with the project name/number and subproject status expanded inline
    // (see the header comment for the confirmed equipment_group -> project /
    // equipment_group -> subproject -> status chain).
    const linkValues = ids.map((id) => `/equipment/${id}`).join(",");
    const peData = await rentmanGet(
      env,
      `/projectequipment?equipment=${encodeURIComponent(linkValues)}` +
        `&planperiod_start[lte]=${encodeURIComponent(to)}&planperiod_end[gte]=${encodeURIComponent(from)}` +
        `&expand=${encodeURIComponent("equipment_group.project,equipment_group.subproject.status")}` +
        `&fields=${encodeURIComponent(
          "equipment,quantity_total,planperiod_start,planperiod_end,equipment_group.project.name,equipment_group.project.number,equipment_group.subproject.status",
        )}&limit=${MAX_LIMIT}`,
    );

    for (const row of peData.data || []) {
      const match = /\/equipment\/(\d+)/.exec(row.equipment || "");
      if (!match) continue;
      const id = Number(match[1]);
      const project = row.equipment_group?.project;
      const status = row.equipment_group?.subproject?.status;
      (byEquipmentId[id] ||= []).push({
        projectName: project?.name ?? "Unknown project",
        projectNumber: project?.number ?? "?",
        status: status?.name ?? null,
        quantity: Number(row.quantity_total) || 0,
        planPeriodStart: row.planperiod_start,
        planPeriodEnd: row.planperiod_end,
      });
    }
  }

  const result: Record<string, { totalStock: number; totalRequired: number; remaining: number; projects: AvailabilityProject[] } | null> = {};
  codes.forEach((code) => {
    const eq = byCode[code];
    if (!eq) {
      result[code] = null;
      return;
    }
    const projects = byEquipmentId[eq.id] || [];
    const totalRequired = projects.reduce((sum, p) => sum + p.quantity, 0);
    // Deliberately not clamped at 0 - a negative "remaining" IS the shortage
    // this feature exists to surface.
    result[code] = { totalStock: eq.currentQuantity, totalRequired, remaining: eq.currentQuantity - totalRequired, projects };
  });
  return json(result, env);
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
        default:
          return json({ error: "Not found" }, env, 404);
      }
    } catch (err) {
      return json({ error: err instanceof Error ? err.message : "Unknown error" }, env, 502);
    }
  },
};
