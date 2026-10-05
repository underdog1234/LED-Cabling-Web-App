import { describe, it, expect } from "vitest";
import { makeGridPanels, getPanelSymbol, PANEL_TYPES, type Cell } from "./App";
import { resolutionOf, subScreenResolutionOf, wallFootprintResolutionOf } from "./canvasView/canvasModel";

// Two different pixel spaces live in this app and they must not be confused:
//
//  - resolutionOf packs each row's panels shoulder to shoulder. That is the
//    NovaStar cabinet-topology space, where there is no such thing as an empty
//    gap pixel, and the .uprj export is built in it.
//  - wallFootprintResolutionOf measures the rectangle the wall physically
//    stands in. That is what the wall's quoted Resolution, aspect ratio and
//    content size mean, because content has to span the whole rectangle.
//
// They agree exactly on a full rectangle. They part company the moment the
// layout is stepped, and the footprint is the one that is right there.

const MG9 = PANEL_TYPES.MG9;

describe("wallFootprintResolutionOf", () => {
  it("matches the packed resolution on a plain rectangular wall", () => {
    for (const [cols, rows] of [[6, 3], [1, 1], [24, 8]] as Array<[number, number]>) {
      const panels = makeGridPanels(cols, rows, "MG9");
      expect(wallFootprintResolutionOf(panels)).toEqual({ w: cols * MG9.pixW, h: rows * MG9.pixH });
      expect(wallFootprintResolutionOf(panels)).toEqual(resolutionOf(panels));
    }
  });

  it("spans every column a stepped wall stands in, not just its longest row", () => {
    // Four columns wide, but no single row holds more than three panels - the
    // shape of the 18.5m wall that reported 5,880px instead of 6,216px.
    const panels = makeGridPanels(4, 2, "MG9").filter((cell) => {
      const col = Math.round(cell.x / (MG9.w * 1000));
      const row = Math.round(cell.y / (MG9.h * 1000));
      return !(row === 0 && col === 3) && !(row === 1 && col === 0);
    });
    expect(panels).toHaveLength(6);
    // The packed space stops at the longest row...
    expect(resolutionOf(panels)).toEqual({ w: 3 * MG9.pixW, h: 2 * MG9.pixH });
    // ...the wall itself still stands across all four columns.
    expect(wallFootprintResolutionOf(panels)).toEqual({ w: 4 * MG9.pixW, h: 2 * MG9.pixH });
  });

  it("ignores removed panels and copes with an empty wall", () => {
    const panels = makeGridPanels(3, 1, "MG9");
    panels[2].isRemoved = true;
    expect(wallFootprintResolutionOf(panels)).toEqual({ w: 2 * MG9.pixW, h: MG9.pixH });
    expect(wallFootprintResolutionOf([])).toEqual({ w: 0, h: 0 });
  });

  it("uses each panel type's own pitch on an MT wall", () => {
    const panels = makeGridPanels(3, 2, "MT");
    expect(wallFootprintResolutionOf(panels)).toEqual({ w: 3 * PANEL_TYPES.MT.pixW, h: 2 * PANEL_TYPES.MT.pixH });
  });
});

describe("subScreenResolutionOf", () => {
  // A sub-screen is a wall in its own right, so its quoted resolution is the
  // rectangle it stands in - the same figure the test pattern renders at.
  // Reading it off the packed space instead lost a module column per gap: a
  // 22-wide sub-screen with a hole in each row read 3,528px where the test
  // pattern, the PDF's own layout image and the real content are 3,696px.
  const screened = (cols: number, rows: number, holeOf: (row: number) => number | null): Cell[] =>
    makeGridPanels(cols, rows, "MG9")
      .filter((cell) => {
        const col = Math.round(cell.x / (MG9.w * 1000));
        const row = Math.round(cell.y / (MG9.h * 1000));
        return holeOf(row) !== col;
      })
      .map((cell) => ({ ...cell, subScreenId: "screen-1" }));

  it("spans every column the sub-screen stands in, gaps included", () => {
    const panels = screened(22, 8, (row) => row + 1);
    expect(subScreenResolutionOf(panels, "screen-1")).toEqual({ w: 22 * MG9.pixW, h: 8 * MG9.pixH });
    // The exact numbers off the report: 3696 x 1344, not 3528 x 1344.
    expect(subScreenResolutionOf(panels, "screen-1")).toEqual({ w: 3696, h: 1344 });
    // The packed space is what it used to read, and still what the NovaStar
    // cabinet topology is built in - both correct, for different jobs.
    expect(resolutionOf(panels)).toEqual({ w: 21 * MG9.pixW, h: 8 * MG9.pixH });
  });

  it("agrees with the packed space on a solid rectangular sub-screen", () => {
    const panels = screened(6, 3, () => null);
    expect(subScreenResolutionOf(panels, "screen-1")).toEqual({ w: 6 * MG9.pixW, h: 3 * MG9.pixH });
  });

  it("measures only its own panels, not the wall around them", () => {
    const mine = screened(4, 2, () => null);
    const theirs = makeGridPanels(4, 2, "MG9").map((cell) => ({
      ...cell,
      id: `other-${cell.id}`,
      x: cell.x + 10_000,
      subScreenId: "screen-2",
    }));
    expect(subScreenResolutionOf([...mine, ...theirs], "screen-1")).toEqual({ w: 4 * MG9.pixW, h: 2 * MG9.pixH });
  });
});

describe("getPanelSymbol", () => {
  const shaped = (variant: Cell["panelVariant"], rotation: number): Cell => ({
    ...makeGridPanels(1, 1, "MG9")[0],
    panelVariant: variant,
    rotation,
  });

  it("names which physical part a shaped panel is, the way the shelf names it", () => {
    // The same LU / LD / RU / RD buckets Stock Calculations counts against the
    // shelf, read from the FRONT of the wall - so the drawing names the part
    // you pick. Left/Right is the corner the right angle is drawn in; Up/Down
    // is the shelf's own vertical sense, which runs the opposite way (a
    // rotation-0 triangle draws its corner at the bottom and is the "Up"
    // part). See TRIANGLE_ORIENTATION - these are deliberate.
    expect(getPanelSymbol(shaped("TRIANGLE", 0))).toBe("△ LU");
    expect(getPanelSymbol(shaped("TRIANGLE", 90))).toBe("△ LD");
    expect(getPanelSymbol(shaped("TRIANGLE", 180))).toBe("△ RD");
    expect(getPanelSymbol(shaped("TRIANGLE", 270))).toBe("△ RU");

    expect(getPanelSymbol(shaped("CURVED", 0))).toBe("◜ RU");
    expect(getPanelSymbol(shaped("CURVED", 90))).toBe("◜ LU");
    expect(getPanelSymbol(shaped("CURVED", 180))).toBe("◜ LD");
    expect(getPanelSymbol(shaped("CURVED", 270))).toBe("◜ RD");
  });

  it("keeps each rotation on its own part, with the four never colliding", () => {
    // Whatever the naming, the one thing that must hold is that the four
    // rotations map onto four DIFFERENT one-way parts - two rotations sharing
    // a bucket would order half the wall wrong.
    for (const variant of ["TRIANGLE", "CURVED"] as const) {
      const names = [0, 90, 180, 270].map((rotation) => getPanelSymbol(shaped(variant, rotation)));
      expect(new Set(names).size).toBe(4);
    }
  });

  it("leaves a plain panel alone, and still flags one that is merely rotated", () => {
    expect(getPanelSymbol(shaped("STANDARD", 0))).toBe("");
    expect(getPanelSymbol(shaped("STANDARD", 45))).toBe("🔄");
  });
});
