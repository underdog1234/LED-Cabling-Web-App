import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import type { Cell, SubScreen } from "../App";
import { buildNovaStarExport, WHOLE_LAYOUT_KEY } from "./exportBuilder";
import { parseUprj } from "./parseUprj";

// Regression coverage for a real bug: cabinet_topology was positioned using
// the sub-screen's output-canvas placement (finalCanvasPositionOf), but a
// real NovaStar-generated file for this exact wall (both sub-screens left at
// the same, default canvas position) proved the LED screen's cabinet canvas
// is actually the wall's own tightly-packed NATIVE pixel grid, entirely
// independent of output-canvas/sub-screen placement - confirmed via a
// NovaStar-exported .uscr cabinet-topology JSON (a much more direct source
// than reverse-engineering the SQLite blob) cross-checked against the .uprj
// itself. Positioning cabinets by output-canvas placement meant two
// sub-screens left at the same canvas position (as here) would collide onto
// the exact same pixel coordinates, which is very likely why "no panels/
// patching" showed up when the file was imported - NovaStar had nowhere
// valid to put half the cabinets.
function fixture(name: string): Uint8Array {
  const path = fileURLToPath(new URL(`./__fixtures__/${name}`, import.meta.url));
  return new Uint8Array(readFileSync(path));
}
function fixtureJson(name: string): any {
  const path = fileURLToPath(new URL(`./__fixtures__/${name}`, import.meta.url));
  return JSON.parse(readFileSync(path, "utf8"));
}

describe("wall-native cabinet positioning against a real reference export", () => {
  it("matches a real NovaStar file's canvas size, cabinet count and exact positions for a 2-sub-screen wall with a gap", async () => {
    const data = fixtureJson("reference-10x5-layout.json");
    const grid: Cell[] = data.panels;
    const subScreens: SubScreen[] = data.subScreens;
    // Both sub-screens sit at the same (0,0) output-canvas position in this
    // real project - the exact scenario that broke under the old
    // output-canvas-relative positioning.
    expect(subScreens.every((s) => s.canvasX === 0 && s.canvasY === 0)).toBe(true);

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
      canvasInputs: Object.entries(data.canvasInputs as Record<string, number | null>).map(([key, interfacePk]) => ({
        key,
        name: key === WHOLE_LAYOUT_KEY ? "Whole Layout" : (subScreens.find((s) => s.id === key)?.name ?? key),
        interfacePk,
      })),
    });
    expect(result.errors).toEqual([]);
    expect(result.ok).toBe(true);

    const oursBytes = new Uint8Array(await result.blob!.arrayBuffer());
    const ours = await parseUprj(oursBytes);
    const reference = await parseUprj(fixture("reference-10x5.uprj"));

    // Wall is 10 cols x 5 rows of MG9 (168px), tightly packed -> 1680x840,
    // matching the real file's t_canvas exactly (not the 1920x1080 output
    // canvas from the settings JSON).
    expect(reference.canvas).toEqual({ width: 1680, height: 840 });
    expect(ours.canvas).toEqual({ width: 1680, height: 840 });

    expect(ours.cabinets).toHaveLength(50);
    expect(reference.cabinets).toHaveLength(50);

    const posKey = (c: { x: number; y: number }) => `${c.x},${c.y}`;
    const ourPositions = new Set(ours.cabinets.map(posKey));
    const referencePositions = new Set(reference.cabinets.map(posKey));
    expect(ourPositions).toEqual(referencePositions);

    // Every cabinet must fall strictly within the wall-native canvas - no
    // stray output-canvas offset leaking through.
    for (const c of ours.cabinets) {
      expect(c.x).toBeGreaterThanOrEqual(0);
      expect(c.y).toBeGreaterThanOrEqual(0);
      expect(c.x + c.width).toBeLessThanOrEqual(1680);
      expect(c.y + c.height).toBeLessThanOrEqual(840);
    }
  });

  it("packs the two sub-screen groups side by side across the wall's gap, matching hand-computed positions", async () => {
    const data = fixtureJson("reference-10x5-layout.json");
    const grid: Cell[] = data.panels;
    const subScreens: SubScreen[] = data.subScreens;

    const result = await buildNovaStarExport({
      processorModel: data.processorModel,
      projectName: data.projectName,
      surfaceName: data.surfaceName,
      outputCanvasW: data.outputCanvas.w,
      outputCanvasH: data.outputCanvas.h,
      wholeLayoutCanvasX: 0,
      wholeLayoutCanvasY: 0,
      grid,
      subScreens,
      inputMode: "perEntry",
      wholeCanvasInputId: null,
      canvasInputs: [],
    });
    expect(result.ok).toBe(true);

    const bytes = new Uint8Array(await result.blob!.arrayBuffer());
    const ours = await parseUprj(bytes);

    // Row y=0 (mm) -> pixel y=0: "right" sub-screen (mm x -1500..500) packs
    // into wall-native columns 1-5 (pixel x 0,168,336,504,672); "left"
    // sub-screen (mm x 2500..4500) continues immediately after the gap at
    // columns 6-10 (pixel x 840,1008,1176,1344,1512).
    const row0 = ours.cabinets.filter((c) => c.y === 0).map((c) => c.x).sort((a, b) => a - b);
    expect(row0).toEqual([0, 168, 336, 504, 672, 840, 1008, 1176, 1344, 1512]);
  });
});
