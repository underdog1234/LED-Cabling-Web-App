// Thin fetch wrapper - calls the rentman-proxy Worker's endpoints, never
// Rentman directly (see rentman-proxy/src/index.ts for why: the Rentman API
// token can't safely live in this public static site's bundle). The
// Worker's own URL isn't secret - it's set at build time via the
// VITE_RENTMAN_PROXY_URL env var (see led-cabling-web/.env.example) rather
// than hardcoded, since it depends on the user's own Cloudflare deployment.
export type LiveStockEntry = { name: string; currentQuantity: number };

export type EquipmentAvailabilityProject = {
  projectName: string;
  projectNumber: number | string;
  /** The Rentman subproject's status (e.g. "Confirmed") - Rentman has no status field on the project itself, only its subproject(s). Null if the booking's subproject has no status set. */
  status: string | null;
  quantity: number;
  planPeriodStart: string;
  planPeriodEnd: string;
};

export type EquipmentAvailability = {
  totalStock: number;
  /** Sum of quantity required across every project overlapping the requested range. */
  totalRequired: number;
  /** totalStock - totalRequired. Deliberately NOT clamped at 0 - negative means overbooked. */
  remaining: number;
  projects: EquipmentAvailabilityProject[];
};

const PROXY_URL = (import.meta.env.VITE_RENTMAN_PROXY_URL ?? "").replace(/\/$/, "");

export const isRentmanProxyConfigured = (): boolean => PROXY_URL.length > 0;

async function proxyGet<T>(path: string, params: Record<string, string>): Promise<T> {
  if (!PROXY_URL) {
    throw new Error("Rentman proxy isn't configured - set VITE_RENTMAN_PROXY_URL and rebuild (see .env.example).");
  }
  const url = new URL(PROXY_URL + path);
  Object.entries(params).forEach(([key, value]) => url.searchParams.set(key, value));
  let res: Response;
  try {
    res = await fetch(url.toString());
  } catch {
    throw new Error("Couldn't reach the Rentman proxy - check VITE_RENTMAN_PROXY_URL and that the Worker is deployed.");
  }
  if (!res.ok) {
    const body = await res.json().catch(() => null);
    throw new Error((body && typeof body.error === "string" && body.error) || `Rentman proxy request failed (HTTP ${res.status})`);
  }
  return res.json();
}

export async function fetchEquipmentStock(codes: string[]): Promise<Record<string, LiveStockEntry | null>> {
  if (!codes.length) return {};
  return proxyGet("/equipment-stock", { codes: codes.join(",") });
}

export async function fetchEquipmentAvailability(
  codes: string[],
  from: string,
  to: string,
): Promise<Record<string, EquipmentAvailability | null>> {
  if (!codes.length || !from || !to) return {};
  return proxyGet("/equipment-availability", { codes: codes.join(","), from, to });
}
