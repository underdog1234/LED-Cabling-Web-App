import { describe, it, expect } from "vitest";
import { makeGridPanels, type Cell } from "../App";
import { computeTestPatternLayout } from "./drawTestPattern";

// Regression coverage for a real bug: panels were positioned by tightly
// packing each row band left-to-right in array order (summing pixel widths),
// which silently closes up any gap left by a missing/removed panel and
// shifts every panel after the gap out of its true position. The fix
// positions each panel from its own real mm offset instead.
describe("computeTestPatternLayout with a gap in the middle of a row", () => {
  it("leaves a gap-sized hole instead of shifting later panels left", () => {
    // 3x1 row of MG9 (168px/500mm each); remove the middle panel.
    const grid = makeGridPanels(3, 1, "MG9");
    const withGap: Cell[] = grid.map((cell) => (cell.x === 500 ? { ...cell, isRemoved: true } : cell));

    const layout = computeTestPatternLayout({ projectName: "Test", panelType: "MG9", panels: withGap });

    expect(layout.totalPanels).toBe(2);
    // Full 3-panel-wide bounding box (504mm gap included), not squeezed to 2.
    expect(layout.W).toBe(504);

    const first = grid.find((c) => c.x === 0)!;
    const third = grid.find((c) => c.x === 1000)!;
    const firstRect = layout.panelPixelRects.get(first.id)!;
    const thirdRect = layout.panelPixelRects.get(third.id)!;

    expect(firstRect).toMatchObject({ x: 0, w: 168, h: 168 });
    // Must stay at its true offset (2 module-widths in = 336px), not slide
    // left to 168px just because the middle panel is missing.
    expect(thirdRect).toMatchObject({ x: 336, w: 168, h: 168 });
  });

  it("preserves true positions with a gap-riddled wall matching the reported repro shape", () => {
    // Mirrors the user's broken-MG9-5x5 repro: a 5x5 grid with several
    // panels missing from the middle of various rows/columns.
    const grid = makeGridPanels(5, 5, "MG9");
    const removedXY = new Set(["500,0", "1500,0", "500,500", "1500,500", "500,1000", "1500,1000", "500,1500", "1500,1500", "1500,2000"]);
    const withGaps: Cell[] = grid.map((cell) => (removedXY.has(`${cell.x},${cell.y}`) ? { ...cell, isRemoved: true } : cell));

    const layout = computeTestPatternLayout({ projectName: "Test", panelType: "MG9", panels: withGaps });

    // Full 5x5 bounding box preserved (2500mm x 2500mm -> 840x840px), even
    // though many interior panels are missing.
    expect(layout.W).toBe(840);
    expect(layout.H).toBe(840);

    // Every remaining panel keeps its true grid-relative pixel position.
    for (const cell of withGaps.filter((c) => !c.isRemoved)) {
      const rect = layout.panelPixelRects.get(cell.id)!;
      expect(rect.x).toBe((cell.x / 500) * 168);
      expect(rect.y).toBe((cell.y / 500) * 168);
    }
  });

  it("still tightly packs a uniform gap-free wall exactly as before", () => {
    const grid = makeGridPanels(4, 2, "MG9");
    const layout = computeTestPatternLayout({ projectName: "Test", panelType: "MG9", panels: grid });
    expect(layout.W).toBe(4 * 168);
    expect(layout.H).toBe(2 * 168);
    for (const cell of grid) {
      const rect = layout.panelPixelRects.get(cell.id)!;
      expect(rect.x).toBe((cell.x / 500) * 168);
      expect(rect.y).toBe((cell.y / 500) * 168);
    }
  });
});

// Regression coverage for a real bug: column numbers were computed from the
// panel's raw back-view x position, but the pattern always renders mirrored
// (front view) - so the printed number didn't match the column the audience
// actually sees it in. The panel that appears in the top-left corner of the
// rendered image must read row 1, column 1.
describe("computeTestPatternLayout column/row numbering reads from the front", () => {
  it("labels the rendered top-left panel as row 1, column 1", () => {
    const grid = makeGridPanels(4, 3, "MG9");
    const layout = computeTestPatternLayout({ projectName: "Test", panelType: "MG9", panels: grid });

    // Rendered top-left = smallest y (top row) and, after the horizontal
    // mirror, the panel with the LARGEST back-view x (rightmost when
    // standing behind the wall becomes leftmost from the front).
    const topLeft = grid.reduce((best, c) => (c.y === 0 && c.x > best.x ? c : best), grid.find((c) => c.y === 0)!);
    expect(layout.rowLabel(topLeft)).toBe(1);
    expect(layout.colLabel(topLeft)).toBe("1");

    // Rendered top-right (back-view leftmost, x=0) must read column 4 (the
    // wall is 4 columns wide).
    const topRight = grid.find((c) => c.y === 0 && c.x === 0)!;
    expect(layout.colLabel(topRight)).toBe("4");

    // Bottom row (largest y) must still read the highest row number.
    const bottomLeft = grid.reduce((best, c) => (c.y > best.y ? c : c.y === best.y && c.x > best.x ? c : best), grid[0]);
    expect(layout.rowLabel(bottomLeft)).toBe(3);
  });
});
