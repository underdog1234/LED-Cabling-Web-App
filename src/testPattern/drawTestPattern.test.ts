import { describe, it, expect, vi } from "vitest";
import { makeGridPanels, type Cell } from "../App";
import {
  computeTestPatternLayout,
  getContentPixelHeight,
  drawBouncingLogo,
  drawAlignmentOverlay,
  drawSurfaceBoundaries,
  WHOLE_WALL_SURFACE_ID,
  UNASSIGNED_SURFACE_ID,
  type TestPatternLayout,
} from "./drawTestPattern";

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
    expect(layout.rowLabel(topLeft)).toBe("1");
    expect(layout.colLabel(topLeft)).toBe("1");

    // Rendered top-right (back-view leftmost, x=0) must read column 4 (the
    // wall is 4 columns wide).
    const topRight = grid.find((c) => c.y === 0 && c.x === 0)!;
    expect(layout.colLabel(topRight)).toBe("4");

    // Bottom row (largest y) must still read the highest row number.
    const bottomLeft = grid.reduce((best, c) => (c.y > best.y ? c : c.y === best.y && c.x > best.x ? c : best), grid[0]);
    expect(layout.rowLabel(bottomLeft)).toBe("3");
  });

  // Regression coverage for a real bug: column numbers were computed by
  // dividing raw x-position by the fixed 500mm MODULE_MM, which silently
  // counts each 1000mm-wide MT panel as 2 columns instead of 1 (three real
  // MT panel-columns came out labeled 5, 3, 1 instead of 3, 2, 1).
  it("numbers MT columns 1-per-panel, not 2-per-panel (MT panels are 1000mm wide, not 500mm)", () => {
    const grid = makeGridPanels(3, 1, "MT");
    const layout = computeTestPatternLayout({ projectName: "Test", panelType: "MT", panels: grid });
    const byX = (x: number) => grid.find((c) => c.x === x)!;
    // Front-view mirrored, same convention as the MG9 case above: back-view
    // leftmost (x=0) reads the highest column number.
    expect(layout.colLabel(byX(0))).toBe("3");
    expect(layout.colLabel(byX(1000))).toBe("2");
    expect(layout.colLabel(byX(2000))).toBe("1");
    expect(layout.activeColsCount).toBe(3);
  });
});

// Regression coverage for the MT "Recommended Content Resolution" rework:
// this used to be a display-only label (contentPixelH computed but never
// used to size anything) - these cases pin down the actual computed value,
// including the "no blended content resolution for a mixed wall" rule.
describe("computeTestPatternLayout content resolution (MT vs. MG9)", () => {
  it("keeps contentPixelH/contentPixelW equal to H/W for an MG9-only wall", () => {
    const grid = makeGridPanels(4, 2, "MG9");
    const layout = computeTestPatternLayout({ projectName: "Test", panelType: "MG9", panels: grid });
    expect(layout.contentPixelW).toBe(layout.W);
    expect(layout.contentPixelH).toBe(layout.H);
  });

  it("doubles contentPixelH (not contentPixelW) for an MT-only wall", () => {
    const grid = makeGridPanels(2, 1, "MT");
    const layout = computeTestPatternLayout({ projectName: "Test", panelType: "MT", panels: grid });
    expect(layout.W).toBe(512); // 2 panels x 256px
    expect(layout.H).toBe(64);
    expect(layout.contentPixelW).toBe(layout.W);
    expect(layout.contentPixelH).toBe(layout.H * 2);
  });

  it("does not apply the MT doubling to a mixed MG9+MT wall", () => {
    const mg9 = makeGridPanels(1, 1, "MG9");
    const mt = makeGridPanels(1, 1, "MT").map((cell) => ({ ...cell, x: 500 }));
    const layout = computeTestPatternLayout({ projectName: "Test", panelType: "MG9", panels: [...mg9, ...mt] });
    expect(layout.contentPixelH).toBe(layout.H);
  });

  it("getContentPixelHeight is a pure no-op on an empty panel list", () => {
    expect(getContentPixelHeight([], 100)).toBe(100);
  });
});

// Regression coverage for a real bug: the bouncing logo (live-preview-only
// decoration, never exported/recorded, so never later "squeezed" back down
// by real MT receiving hardware the way the rest of the pattern is) is drawn
// in native W x H space and inherits whatever vertical scale the live view's
// caller has applied for MT's content-resolution doubling. Without
// compensating for that here, it would render twice as tall as its real
// shape on an MT wall, with nothing downstream to cancel it out.
describe("drawBouncingLogo aspect ratio", () => {
  const makeCtx = () => ({ save: vi.fn(), restore: vi.fn(), drawImage: vi.fn(), imageSmoothingEnabled: false });
  const makeImage = (naturalWidth: number, naturalHeight: number) => ({ naturalWidth, naturalHeight, complete: true }) as unknown as HTMLImageElement;

  it("preserves the image's natural aspect ratio for a non-MT wall (no ambient stretch)", () => {
    const ctx = makeCtx();
    const layout = { W: 672, H: 336, contentPixelH: 336, tileWidthPx: 168 } as unknown as TestPatternLayout;
    drawBouncingLogo(ctx as unknown as CanvasRenderingContext2D, layout, 0, makeImage(200, 100));
    const [, , , logoW, logoH] = ctx.drawImage.mock.calls[0];
    expect(logoH / logoW).toBeCloseTo(100 / 200, 10);
  });

  it("pre-compensates the native-space height for an MT wall, so the ambient vertical stretch restores the correct on-screen aspect ratio", () => {
    const ctx = makeCtx();
    const layout = { W: 512, H: 64, contentPixelH: 128, tileWidthPx: 256 } as unknown as TestPatternLayout;
    drawBouncingLogo(ctx as unknown as CanvasRenderingContext2D, layout, 0, makeImage(200, 100));
    const [, , , logoW, logoH] = ctx.drawImage.mock.calls[0];
    const contentScaleY = layout.contentPixelH / layout.H; // 2 for this MT wall
    // The live view applies ctx.scale(1, contentScaleY) before drawing - the
    // EFFECTIVE on-screen size is (logoW, logoH * contentScaleY), which must
    // match the image's real aspect ratio.
    expect((logoH * contentScaleY) / logoW).toBeCloseTo(100 / 200, 10);
  });
});

// Regression coverage for a real bug: the alignment overlay's centre circle
// is meant as a TRUE circle reference (for spotting warp/stretch by eye) -
// drawn as ctx.arc with equal radii in native W x H space, it rendered as a
// squashed ellipse once the ambient content-resolution stretch reached an
// MT wall, since nothing downstream corrects it back (unlike the RGB/panel
// geometry, this shape's whole purpose is to look right in the actual
// generated image, not to represent physical panel positions).
describe("drawAlignmentOverlay circle", () => {
  const makeCtx = () => ({
    save: vi.fn(),
    restore: vi.fn(),
    beginPath: vi.fn(),
    moveTo: vi.fn(),
    lineTo: vi.fn(),
    stroke: vi.fn(),
    ellipse: vi.fn(),
    strokeStyle: "",
    lineWidth: 0,
  });

  it("draws a true circle (equal radii) for a non-MT wall", () => {
    const ctx = makeCtx();
    const layout = { W: 672, H: 336, contentPixelH: 336 } as unknown as TestPatternLayout;
    drawAlignmentOverlay(ctx as unknown as CanvasRenderingContext2D, layout);
    const [, , radiusX, radiusY] = ctx.ellipse.mock.calls[0];
    expect(radiusY).toBe(radiusX);
  });

  it("pre-compensates the vertical radius for an MT wall, so it renders as a true circle at the content resolution", () => {
    const ctx = makeCtx();
    const layout = { W: 512, H: 64, contentPixelH: 128 } as unknown as TestPatternLayout;
    drawAlignmentOverlay(ctx as unknown as CanvasRenderingContext2D, layout);
    const [, , radiusX, radiusY] = ctx.ellipse.mock.calls[0];
    const contentScaleY = layout.contentPixelH / layout.H; // 2
    // After the caller's ctx.scale(1, contentScaleY), the effective on-screen radii must be equal.
    expect(radiusY * contentScaleY).toBeCloseTo(radiusX, 10);
  });
});

// Per-sub-screen surfaces: each sub-screen gets its own independently
// animating pattern (see TestPatternSurface / drawTestPatternFrame), so the
// pattern lives inside that screen's panels only rather than every screen
// showing one slice of a single wall-wide animation.
describe("computeTestPatternLayout surfaces", () => {
  const twoScreenWall = () => {
    // 4x1 MG9 row; left two panels are "Left", right two are "Right".
    const grid = makeGridPanels(4, 1, "MG9");
    return grid.map((cell) => ({ ...cell, subScreenId: cell.x < 1000 ? "left" : "right" }));
  };

  it("collapses to one whole-wall surface when the project has no sub-screens", () => {
    const panels = makeGridPanels(3, 2, "MG9");
    const layout = computeTestPatternLayout({ projectName: "Test", panelType: "MG9", panels });

    expect(layout.surfaces).toHaveLength(1);
    expect(layout.surfaces[0].id).toBe(WHOLE_WALL_SURFACE_ID);
    expect(layout.surfaces[0].color).toBeNull();
    expect(layout.surfaces[0].cells).toHaveLength(6);
    expect(layout.surfaces[0].bbox).toMatchObject({ x: 0, y: 0, w: layout.W, h: layout.H });
  });

  it("gives each sub-screen its own surface, colour and bounding box", () => {
    const layout = computeTestPatternLayout({
      projectName: "Test",
      panelType: "MG9",
      panels: twoScreenWall(),
      subScreens: [
        { id: "left", name: "Left Screen", color: "#38bdf8" },
        { id: "right", name: "Right Screen", color: "#fb923c" },
      ],
    });

    expect(layout.surfaces.map((s) => s.id)).toEqual(["left", "right"]);
    expect(layout.surfaces.map((s) => s.color)).toEqual(["#38bdf8", "#fb923c"]);
    layout.surfaces.forEach((surface) => expect(surface.cells).toHaveLength(2));
    // Each bbox covers exactly two panels wide, and the two are side by side
    // in the mirrored display space with no overlap.
    const [left, right] = layout.surfaces;
    expect(left.bbox.w).toBe(336);
    expect(right.bbox.w).toBe(336);
    expect(left.bbox.x + left.bbox.w === right.bbox.x || right.bbox.x + right.bbox.w === left.bbox.x).toBe(true);
  });

  it("names a single sub-screen rendered on its own", () => {
    // Rendering one screen is where its name matters most - the shape alone
    // no longer says which screen it is - and the boundary/name pass used to
    // bail out whenever there was only one surface.
    const panels = twoScreenWall().filter((cell) => cell.subScreenId === "left");
    const layout = computeTestPatternLayout({
      projectName: "Test",
      panelType: "MG9",
      panels,
      subScreens: [{ id: "left", name: "Left Screen", color: "#38bdf8" }],
    });
    expect(layout.surfaces).toHaveLength(1);
    expect(layout.surfaces[0]).toMatchObject({ id: "left", name: "Left Screen", isSubScreen: true });

    const drawn: string[] = [];
    const ctx = {
      save: () => {}, restore: () => {},
      strokeRect: () => {}, fillRect: () => {},
      measureText: () => ({ width: 50 }),
      fillText: (text: string) => drawn.push(text),
      set font(_v: string) {}, set strokeStyle(_v: string) {}, set fillStyle(_v: string) {},
      set lineWidth(_v: number) {}, set textAlign(_v: string) {}, set textBaseline(_v: string) {},
      set lineJoin(_v: string) {},
    };
    drawSurfaceBoundaries(ctx as unknown as CanvasRenderingContext2D, layout);
    expect(drawn).toEqual(["LEFT SCREEN"]);
  });

  it("leaves a plain wall with no sub-screens unmarked", () => {
    // Nothing is drawn over a wall that has no sub-screens at all, exactly as
    // before - the single-surface case is only named when it is a real screen.
    const layout = computeTestPatternLayout({ projectName: "Test", panelType: "MG9", panels: makeGridPanels(3, 2, "MG9") });
    expect(layout.surfaces[0].isSubScreen).toBe(false);
    const drawn: string[] = [];
    const ctx = {
      save: () => { drawn.push("save"); }, restore: () => {},
      strokeRect: () => { drawn.push("rect"); }, fillRect: () => {},
      measureText: () => ({ width: 50 }),
      fillText: (text: string) => drawn.push(text),
      set font(_v: string) {}, set strokeStyle(_v: string) {}, set fillStyle(_v: string) {},
      set lineWidth(_v: number) {}, set textAlign(_v: string) {}, set textBaseline(_v: string) {},
      set lineJoin(_v: string) {},
    };
    drawSurfaceBoundaries(ctx as unknown as CanvasRenderingContext2D, layout);
    expect(drawn).toEqual([]);
  });

  it("puts panels outside any sub-screen in their own Unassigned surface", () => {
    const panels = twoScreenWall().map((cell) => (cell.x === 1500 ? { ...cell, subScreenId: null } : cell));
    const layout = computeTestPatternLayout({
      projectName: "Test",
      panelType: "MG9",
      panels,
      subScreens: [
        { id: "left", name: "Left Screen", color: "#38bdf8" },
        { id: "right", name: "Right Screen", color: "#fb923c" },
      ],
    });

    expect(layout.surfaces.map((s) => s.id)).toEqual(["left", "right", UNASSIGNED_SURFACE_ID]);
    expect(layout.surfaces[2].cells).toHaveLength(1);
    expect(layout.surfaces[2].color).toBeNull();
  });

  it("assigns every active panel to exactly one surface", () => {
    const layout = computeTestPatternLayout({
      projectName: "Test",
      panelType: "MG9",
      panels: twoScreenWall(),
      subScreens: [
        { id: "left", name: "Left Screen", color: "#38bdf8" },
        { id: "right", name: "Right Screen", color: "#fb923c" },
      ],
    });

    const ids = layout.surfaces.flatMap((surface) => surface.cells.map((cell) => cell.id));
    expect(new Set(ids).size).toBe(ids.length);
    expect(ids.length).toBe(layout.activePanels.length);
  });

  it("ignores a declared sub-screen that has no panels", () => {
    const layout = computeTestPatternLayout({
      projectName: "Test",
      panelType: "MG9",
      panels: makeGridPanels(2, 1, "MG9"),
      subScreens: [{ id: "empty", name: "Empty", color: "#38bdf8" }],
    });

    expect(layout.surfaces.map((s) => s.id)).toEqual([WHOLE_WALL_SURFACE_ID]);
  });
});
