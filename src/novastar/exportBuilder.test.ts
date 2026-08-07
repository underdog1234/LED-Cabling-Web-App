import { describe, it, expect } from "vitest";
import { makeGridPanels, type Cell, type SubScreen } from "../App";
import { buildNovaStarExport, buildExportSummaryAndCabinets, WHOLE_LAYOUT_KEY, type ExportBuilderInput } from "./exportBuilder";
import { parseUprj } from "./parseUprj";
import { PROCESSOR_SPECS } from "./processorModels";

// MG9 panels are 168x168px / 500x500mm each (see App.tsx PANEL_TYPES).
const PANEL_PX = 168;

function patchSequentially(grid: Cell[], port: number): Cell[] {
  return grid.map((cell, i) => ({ ...cell, assignedPort: port, sequence: i + 1 }));
}

function baseInput(overrides: Partial<ExportBuilderInput>): ExportBuilderInput {
  return {
    processorModel: "VX1000_PRO",
    projectName: "Test Project",
    surfaceName: "Front Wall",
    outputCanvasW: 672,
    outputCanvasH: 672,
    wholeLayoutCanvasX: 0,
    wholeLayoutCanvasY: 0,
    grid: [],
    subScreens: [],
    inputMode: "perEntry",
    wholeCanvasInputId: null,
    canvasInputs: [{ key: WHOLE_LAYOUT_KEY, name: "Whole Layout", interfacePk: 4 }],
    ...overrides,
  };
}

describe("buildNovaStarExport - basic rectangular wall", () => {
  it("generates a valid VX1000 Pro config for a 4x4 MG9 wall on one port", async () => {
    const grid = patchSequentially(makeGridPanels(4, 4, "MG9"), 1);
    const input = baseInput({ grid });

    const result = await buildNovaStarExport(input);
    expect(result.errors).toEqual([]);
    expect(result.ok).toBe(true);
    expect(result.blob).toBeDefined();
    expect(result.fileName).toBe("Test-Project_Front-Wall_VX1000Pro_Config.uprj");
    expect(result.summary.panelCount).toBe(16);
    expect(result.summary.ethernetOutputsUsed).toBe(1);
    expect(result.summary.outputPixelLoads).toEqual([{ signalPort: 1, pixels: 16 * PANEL_PX * PANEL_PX }]);

    const bytes = new Uint8Array(await result.blob!.arrayBuffer());
    const parsed = await parseUprj(bytes);
    expect(parsed.checksumValid).toBe(true);
    expect(parsed.deviceCount).toBe(1);
    expect(parsed.canvas).toEqual({ width: 672, height: 672 });
    expect(parsed.mainLogicScreen).toEqual({ width: 672, height: 672, name: "Front Wall" });
    expect(parsed.cabinets).toHaveLength(16);
  });

  it("generates a valid VX2000 Pro config for the same wall", async () => {
    const grid = patchSequentially(makeGridPanels(4, 4, "MG9"), 1);
    const input = baseInput({ processorModel: "VX2000_PRO", grid });

    const result = await buildNovaStarExport(input);
    expect(result.ok).toBe(true);
    expect(result.fileName).toBe("Test-Project_Front-Wall_VX2000Pro_Config.uprj");

    const bytes = new Uint8Array(await result.blob!.arrayBuffer());
    const parsed = await parseUprj(bytes);
    expect(parsed.checksumValid).toBe(true);
    expect(parsed.cabinets).toHaveLength(16);
  });
});

describe("buildNovaStarExport - irregular wall", () => {
  it("excludes removed panels and preserves the remaining layout", async () => {
    let grid = makeGridPanels(3, 3, "MG9");
    // Knock out the corners to make an irregular (plus-shaped) wall.
    const corners = new Set([grid[0].id, grid[2].id, grid[6].id, grid[8].id]);
    grid = grid.map((c) => (corners.has(c.id) ? { ...c, isRemoved: true } : c));
    grid = patchSequentially(
      grid.filter((c) => !c.isRemoved),
      1,
    ).concat(grid.filter((c) => c.isRemoved));

    const input = baseInput({ grid, outputCanvasW: 504, outputCanvasH: 504 });
    const result = await buildNovaStarExport(input);

    expect(result.ok).toBe(true);
    expect(result.summary.panelCount).toBe(5);

    const bytes = new Uint8Array(await result.blob!.arrayBuffer());
    const parsed = await parseUprj(bytes);
    expect(parsed.cabinets).toHaveLength(5);
  });
});

describe("buildNovaStarExport - multiple sub-screens", () => {
  it("tight-packs cabinets across both sub-screens in the wall's own native grid, ignoring output-canvas placement, while still assigning per-entry inputs", async () => {
    const leftGrid = makeGridPanels(2, 2, "MG9", "left");
    const rightGrid = makeGridPanels(2, 2, "MG9", "right").map((c) => ({ ...c, x: c.x + 5000 }));
    let grid = [...leftGrid, ...rightGrid];
    grid = grid.map((c, i) => ({ ...c, assignedPort: 1, sequence: i + 1 }));

    // Deliberately a large, arbitrary output-canvas offset for "right" (and
    // a real mm gap between the two groups' panel positions) - cabinet
    // positions must come out identical regardless of this value, since
    // cabinet_topology uses the wall's own tightly-packed native pixel grid,
    // not output-canvas/sub-screen placement (see
    // wallNativePositioningRegression.test.ts for the real-file proof).
    const subScreens: SubScreen[] = [
      { id: "left", name: "Left Screen", canvasX: 0, canvasY: 0, createdAt: 1 },
      { id: "right", name: "Right Screen", canvasX: 9999, canvasY: 9999, createdAt: 2 },
    ];

    const input = baseInput({
      grid,
      subScreens,
      outputCanvasW: 672,
      outputCanvasH: 336,
      canvasInputs: [
        { key: "left", name: "Left Screen", interfacePk: 4 },
        { key: "right", name: "Right Screen", interfacePk: 5 },
      ],
    });

    const result = await buildNovaStarExport(input);
    expect(result.ok).toBe(true);
    expect(result.summary.subScreenCount).toBe(2);
    expect(result.summary.inputAssignments).toEqual([
      { name: "Left Screen", inputLabel: "HDMI 1" },
      { name: "Right Screen", inputLabel: "HDMI 2" },
    ]);

    const bytes = new Uint8Array(await result.blob!.arrayBuffer());
    const parsed = await parseUprj(bytes);
    const xs = parsed.cabinets.map((c) => c.x).sort((a, b) => a - b);
    // 2 left columns (0, 168) immediately followed by 2 right columns
    // (336, 504) in the wall's own native grid - the mm gap between the two
    // groups and "right"'s 9999 canvas offset both collapse away, exactly
    // like a missing/removed panel does (see drawTestPattern's own gap
    // handling for the visual-preservation counterpart of this).
    expect(Math.min(...xs)).toBe(0);
    expect(Math.max(...xs)).toBe(336 + PANEL_PX);
    expect(parsed.inputLayers).toHaveLength(2);
    expect(parsed.inputLayers.map((l) => l.interfacePk).sort()).toEqual([4, 5]);
  });
});

describe("buildNovaStarExport - multiple Ethernet outputs and patching order", () => {
  it("preserves per-port cabinet_index matching the app's signal-port sequence", async () => {
    const grid = makeGridPanels(4, 1, "MG9").map((cell, i) => ({
      ...cell,
      assignedPort: i < 2 ? 1 : 2,
      sequence: i < 2 ? i + 1 : i - 2 + 1,
    }));
    const input = baseInput({ grid, outputCanvasW: 672, outputCanvasH: 168 });

    const result = await buildNovaStarExport(input);
    expect(result.ok).toBe(true);
    expect(result.summary.ethernetOutputsUsed).toBe(2);

    const bytes = new Uint8Array(await result.blob!.arrayBuffer());
    const parsed = await parseUprj(bytes);
    const byPort = new Map<number, typeof parsed.cabinets>();
    for (const c of parsed.cabinets) {
      byPort.set(c.netPortIndex, [...(byPort.get(c.netPortIndex) ?? []), c]);
    }
    // net_port_index is allocated fresh per export, not a stable "port 0/1"
    // label (see novaDb.ts's replaceCabinetTopology) - there must be exactly
    // 2 distinct groups, each with 2 cabinets in daisy-chain order 0,1.
    expect(byPort.size).toBe(2);
    for (const cabinetsOnPort of byPort.values()) {
      expect(cabinetsOnPort.map((c) => c.cabinetIndex).sort()).toEqual([0, 1]);
    }
    // The starting cabinet (cabinetIndex 0) of whichever port the first grid
    // cell landed on must be that same first cell's canvas position.
    const startingCabinets = [...byPort.values()].map((list) => list.find((c) => c.cabinetIndex === 0)!);
    expect(startingCabinets.some((c) => c.x === 0 && c.y === 0)).toBe(true);
  });
});

describe("buildNovaStarExport - whole-canvas input mode", () => {
  it("assigns a single input covering the full output canvas, ignoring sub-screen boundaries", async () => {
    const leftGrid = makeGridPanels(2, 2, "MG9", "left");
    const rightGrid = makeGridPanels(2, 2, "MG9", "right").map((c) => ({ ...c, x: c.x + 5000 }));
    let grid = [...leftGrid, ...rightGrid];
    grid = grid.map((c, i) => ({ ...c, assignedPort: 1, sequence: i + 1 }));
    const subScreens: SubScreen[] = [
      { id: "left", name: "Left Screen", canvasX: 0, canvasY: 0, createdAt: 1 },
      { id: "right", name: "Right Screen", canvasX: 336, canvasY: 0, createdAt: 2 },
    ];

    const input = baseInput({
      grid,
      subScreens,
      outputCanvasW: 1920,
      outputCanvasH: 1080,
      inputMode: "whole",
      wholeCanvasInputId: 4,
      canvasInputs: [],
    });

    const result = await buildNovaStarExport(input);
    expect(result.ok).toBe(true);
    expect(result.summary.inputAssignments).toEqual([{ name: "Whole Canvas", inputLabel: "HDMI 1" }]);

    const bytes = new Uint8Array(await result.blob!.arrayBuffer());
    const parsed = await parseUprj(bytes);
    expect(parsed.inputLayers).toHaveLength(1);
    expect(parsed.inputLayers[0]).toMatchObject({ interfacePk: 4, x: 0, y: 0, width: 1920, height: 1080 });
  });

  it("warns when the whole canvas has no input assigned, without requiring per-entry assignments", () => {
    const grid = patchSequentially(makeGridPanels(2, 2, "MG9"), 1);
    const input = baseInput({ grid, inputMode: "whole", wholeCanvasInputId: null, canvasInputs: [] });
    const { errors, warnings } = buildExportSummaryAndCabinets(input);
    expect(errors).toEqual([]);
    expect(warnings.some((w) => w.includes("whole canvas has no input assigned"))).toBe(true);
  });

  it("blocks the export when the whole-canvas input is not available on the selected processor", () => {
    const grid = patchSequentially(makeGridPanels(2, 2, "MG9"), 1);
    const input = baseInput({ grid, inputMode: "whole", wholeCanvasInputId: 3 /* DP - not on VX1000 Pro */, canvasInputs: [] });
    const { errors } = buildExportSummaryAndCabinets(input);
    expect(errors.some((e) => e.includes("not available on"))).toBe(true);
  });
});

describe("buildNovaStarExport - single processor only", () => {
  it("never bundles more than one nodeList entry / device zip, unlike piev3.uprj", async () => {
    const grid = patchSequentially(makeGridPanels(2, 2, "MG9"), 1);
    const input = baseInput({ grid, outputCanvasW: 336, outputCanvasH: 336 });
    const result = await buildNovaStarExport(input);
    expect(result.ok).toBe(true);

    const bytes = new Uint8Array(await result.blob!.arrayBuffer());
    const parsed = await parseUprj(bytes);
    expect(parsed.deviceCount).toBe(1);
  });
});

describe("buildExportSummaryAndCabinets - validation", () => {
  it("blocks the export when total pixels exceed the selected processor's capacity", () => {
    // A wall far bigger than VX1000 Pro's 6.5M px cap, all on one port.
    const grid = patchSequentially(makeGridPanels(50, 50, "MG9"), 1);
    const input = baseInput({ grid, outputCanvasW: 8400, outputCanvasH: 8400 });
    const { errors } = buildExportSummaryAndCabinets(input);
    expect(errors.some((e) => e.includes("exceeds") && e.includes("capacity"))).toBe(true);
  });

  it("blocks the export when a single port exceeds its pixel limit", () => {
    const grid = patchSequentially(makeGridPanels(20, 20, "MG9"), 1); // 400 panels * 168*168 > 650,000
    const input = baseInput({ grid, outputCanvasW: 3360, outputCanvasH: 3360 });
    const { errors } = buildExportSummaryAndCabinets(input);
    expect(errors.some((e) => e.includes("px/port limit"))).toBe(true);
  });

  it("blocks the export when the output canvas exceeds the processor's max resolution", () => {
    const grid = patchSequentially(makeGridPanels(2, 2, "MG9"), 1);
    const input = baseInput({ grid, outputCanvasW: 20000, outputCanvasH: 20000 });
    const { errors } = buildExportSummaryAndCabinets(input);
    expect(errors.some((e) => e.includes("maximum"))).toBe(true);
  });

  it("blocks the export when a signal port exceeds the processor's output count", () => {
    const grid = patchSequentially(makeGridPanels(2, 2, "MG9"), 11); // VX1000 Pro only has 10 outputs
    const input = baseInput({ grid });
    const { errors } = buildExportSummaryAndCabinets(input);
    expect(errors.some((e) => e.includes("available Ethernet outputs"))).toBe(true);
  });

  it("blocks the export when an assigned input is not available on the selected processor", () => {
    const grid = patchSequentially(makeGridPanels(2, 2, "MG9"), 1);
    const input = baseInput({ grid, canvasInputs: [{ key: WHOLE_LAYOUT_KEY, name: "Whole Layout", interfacePk: 3 /* DP - VX1000 has none */ }] });
    const { errors } = buildExportSummaryAndCabinets(input);
    expect(errors.some((e) => e.includes("not available on"))).toBe(true);
  });

  it("blocks the export when nothing is patched to a signal port", () => {
    const grid = makeGridPanels(2, 2, "MG9"); // no assignedPort/sequence set
    const input = baseInput({ grid });
    const { errors } = buildExportSummaryAndCabinets(input);
    expect(errors.some((e) => e.includes("nothing to export"))).toBe(true);
  });

  it("warns (but does not block) when panels have no assigned port", () => {
    const grid = makeGridPanels(2, 2, "MG9");
    grid[0] = { ...grid[0], assignedPort: 1, sequence: 1 };
    const input = baseInput({ grid });
    const { errors, warnings } = buildExportSummaryAndCabinets(input);
    expect(errors).toEqual([]);
    expect(warnings.some((w) => w.includes("no assigned signal port"))).toBe(true);
  });

  it("confirms VX1000 Pro and VX2000 Pro capacities used in validation match the public spec sheets", () => {
    expect(PROCESSOR_SPECS.VX1000_PRO.ethernetOutputCount).toBe(10);
    expect(PROCESSOR_SPECS.VX1000_PRO.maxPixelsPerPort).toBe(650_000);
    expect(PROCESSOR_SPECS.VX1000_PRO.maxTotalPixels).toBe(6_500_000);
    expect(PROCESSOR_SPECS.VX2000_PRO.ethernetOutputCount).toBe(20);
    expect(PROCESSOR_SPECS.VX2000_PRO.maxTotalPixels).toBe(13_100_000);
  });
});
