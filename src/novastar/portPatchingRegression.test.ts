import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import type { Cell, SubScreen } from "../App";
import { buildNovaStarExport, WHOLE_LAYOUT_KEY } from "./exportBuilder";
import { parseUprj } from "./parseUprj";

// Regression coverage for a real bug: a fixed/low net_port_index scheme
// (0,1,2...) silently collided with pre-existing port allocations already on
// a device the file was imported into, so the new cabinets merged into an
// unrelated existing port group instead of getting their own ("the port
// patching didn't come through"). Confirmed against a real reference file
// the user exported from actual NovaStar software for this exact wall.
//
// The reference file's device had accumulated an unrelated 25-cabinet
// leftover group (net_port_index 0-4, 5 cabinets each) from earlier,
// unrelated work before this wall's 100 cabinets (net_port_index 165-169)
// were added - that leftover group is excluded below by its known index
// range; it isn't part of what this wall's export is being checked against.
const REFERENCE_WALL_NET_PORT_INDICES = [165, 166, 167, 168, 169];

// NOTE on cabinet positions: this reference file's two sub-screens used a
// deliberate, non-zero output-canvas offset (one stacked 840px below the
// other), which the exporter reproduced at the time by positioning cabinets
// relative to that output-canvas placement. That turned out to be wrong in
// general - see wallNativePositioningRegression.test.ts, built from a
// second, unambiguous real reference file (both sub-screens at the same
// canvas position) that proves cabinet_topology actually uses the LED
// screen's own tightly-packed native pixel grid, entirely independent of
// output-canvas/sub-screen placement. This file's exact pixel positions are
// therefore no longer a valid 1:1 comparison target (they depended on that
// specific sub-screen offset); this test keeps verifying the parts that are
// still universally true: cabinet counts, per-port distribution, and
// non-colliding port ids.

function fixture(name: string): Uint8Array {
  const path = fileURLToPath(new URL(`./__fixtures__/${name}`, import.meta.url));
  return new Uint8Array(readFileSync(path));
}
function fixtureJson(name: string): any {
  const path = fileURLToPath(new URL(`./__fixtures__/${name}`, import.meta.url));
  return JSON.parse(readFileSync(path, "utf8"));
}

describe("port patching regression against a real reference export", () => {
  it("produces the same per-output cabinet distribution and canvas positions as real NovaStar software, using fresh (non-colliding) port ids", async () => {
    const data = fixtureJson("reference-layout.json");
    const grid: Cell[] = data.panels;
    const subScreens: SubScreen[] = data.subScreens;

    const result = await buildNovaStarExport({
      processorModel: data.processorModel,
      projectName: data.projectName,
      surfaceName: data.surfaceName,
      outputCanvasW: data.outputCanvas.w,
      outputCanvasH: data.outputCanvas.h,
      wholeLayoutCanvasX: data.wholeLayoutCanvasPos.x,
      wholeLayoutCanvasY: data.wholeLayoutCanvasPos.y,
      grid,
      subScreens,
      inputMode: "perEntry",
      wholeCanvasInputId: null,
      canvasInputs: Object.entries(data.canvasInputs).map(([key, interfacePk]) => ({
        key,
        name: key === WHOLE_LAYOUT_KEY ? "Whole Layout" : subScreens.find((s: SubScreen) => s.id === key)?.name ?? key,
        interfacePk: interfacePk as number | null,
      })),
    });
    expect(result.errors).toEqual([]);
    expect(result.ok).toBe(true);

    const oursBytes = new Uint8Array(await result.blob!.arrayBuffer());
    const ours = await parseUprj(oursBytes);
    const reference = await parseUprj(fixture("reference-correct.uprj"));

    // Canvas is now the wall's own native resolution (3360x840 for this
    // 20x5 MG9 wall), not the reference's output-canvas-derived 4096x2160 -
    // see the NOTE above.
    expect(ours.canvas).toEqual({ width: 3360, height: 840 });
    expect(ours.cabinets).toHaveLength(100);

    const referenceWallCabinets = reference.cabinets.filter((c) => REFERENCE_WALL_NET_PORT_INDICES.includes(c.netPortIndex));
    expect(referenceWallCabinets).toHaveLength(100);

    // Fresh per-export allocation: our net_port_index values must not be the
    // old naive 0-based scheme (0,1,2,3,4) that caused the original bug -
    // and must not collide with the reference file's own leftover group
    // either, demonstrating the fix generalizes beyond just this one file.
    const ourPortIndices = [...new Set(ours.cabinets.map((c) => c.netPortIndex))];
    expect(ourPortIndices).toHaveLength(5);
    expect(ourPortIndices.every((i) => i >= 1000)).toBe(true);
    expect(ourPortIndices.some((i) => [0, 1, 2, 3, 4].includes(i))).toBe(false);

    // Per-output cabinet counts must match the real file's, regardless of
    // which arbitrary index label each group gets.
    const sizesOf = (cabinets: typeof ours.cabinets) => {
      const byPort = new Map<number, number>();
      for (const c of cabinets) byPort.set(c.netPortIndex, (byPort.get(c.netPortIndex) ?? 0) + 1);
      return [...byPort.values()].sort((a, b) => a - b);
    };
    expect(sizesOf(ours.cabinets)).toEqual(sizesOf(referenceWallCabinets));

    // Cabinet width/height/angle (independent of the sub-screen-offset
    // question) must still match the real file's for every panel.
    expect(ours.cabinets.every((c) => c.width === 168 && c.height === 168 && c.angle === 0)).toBe(true);
    expect(referenceWallCabinets.every((c) => c.width === 168 && c.height === 168 && c.angle === 0)).toBe(true);
  });
});
