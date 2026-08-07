// Loads the embedded blank single-processor NovaStar project (see
// novaDb.ts's header comment for why every export starts from this file
// rather than a hand-built database).
//
// Two branches instead of one clever abstraction: the browser bundle needs
// Vite's `new URL(..., import.meta.url)` asset convention to get a fetchable
// URL in both dev and production builds, while Vitest/Node has no bundler
// asset pipeline to fetch from - reading the file straight off disk is both
// simpler and more reliable there.
export async function loadBlankTemplateBytes(): Promise<Uint8Array> {
  if (typeof window !== "undefined" && typeof fetch === "function") {
    const url = new URL("./assets/blank-template.uprj", import.meta.url);
    const res = await fetch(url);
    return new Uint8Array(await res.arrayBuffer());
  }
  const { readFileSync } = await import("node:fs");
  const { fileURLToPath } = await import("node:url");
  const path = fileURLToPath(new URL("./assets/blank-template.uprj", import.meta.url));
  return new Uint8Array(readFileSync(path));
}
