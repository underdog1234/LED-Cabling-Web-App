// Thin fetch wrapper - calls the rentman-proxy Worker's endpoints, never
// Rentman directly (see rentman-proxy/src/index.ts for why: the Rentman API
// token can't safely live in this public static site's bundle). The
// Worker's own URL isn't secret - it's set at build time via the
// VITE_RENTMAN_PROXY_URL env var (see led-cabling-web/.env.example) rather
// than hardcoded, since it depends on the user's own Cloudflare deployment.
import type { LiveStockEntry } from "./applyLiveRentmanData";

export type { LiveStockEntry };
export type EquipmentSearchResult = { id: number; code: string; name: string };

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

export async function fetchEquipmentAvailability(codes: string[], from: string, to: string): Promise<Record<string, number | null>> {
  if (!codes.length || !from || !to) return {};
  return proxyGet("/equipment-availability", { codes: codes.join(","), from, to });
}

export async function searchEquipment(query: string): Promise<EquipmentSearchResult[]> {
  return proxyGet("/equipment-search", { query });
}
