import { describe, it, expect } from "vitest";
import {
  gridRefLabel,
  mirrorGridRef,
  panelFramePoint,
  panelGridRefs,
  panelShapeLabelPoint,
  panelShapeNeedsInsetLabel,
  type GridRef,
  type PanelRecord,
  type RectMm,
} from "./model/panels";

// Row and column references are what a crew calls a panel over the radio, so
// an offset panel must not shift every number after it. The wall's own module
// grid fixes the numbering; a panel that straddles two cells stands in both.

let seq = 0;
const at = (x: number, y: number, w = 500, h = 500): PanelRecord => ({
  id: `p${(seq += 1)}`,
  panelType: "MG9",
  panelVariant: "STANDARD",
  x,
  y,
  rotation: 0,
  isRemoved: false,
  assignedPort: null,
  sequence: null,
  assignedPowerPort: null,
  powerSequence: null,
  powerManual: false,
  // Carried on the record only so the test's rectOf can read the size back.
  ...({ testW: w, testH: h } as unknown as object),
});
const rectOf = (p: PanelRecord): RectMm => {
  const size = p as unknown as { testW: number; testH: number };
  return { x: p.x, y: p.y, w: size.testW, h: size.testH };
};
const label = (refs: ReturnType<typeof panelGridRefs>, p: PanelRecord, axis: "rows" | "cols") =>
  gridRefLabel(refs.refs.get(p.id)?.[axis]);

describe("panelGridRefs", () => {
  it("numbers a plain grid 1-per-panel from the top-left", () => {
    const panels = [at(0, 0), at(500, 0), at(1000, 0), at(0, 500), at(500, 500), at(1000, 500)];
    const refs = panelGridRefs(panels, rectOf);
    expect({ rows: refs.rows, cols: refs.cols }).toEqual({ rows: 2, cols: 3 });
    expect(panels.map((p) => `R${label(refs, p, "rows")} C${label(refs, p, "cols")}`)).toEqual([
      "R1 C1", "R1 C2", "R1 C3",
      "R2 C1", "R2 C2", "R2 C3",
    ]);
  });

  it("says a half-module offset panel is in BOTH rows, and leaves its neighbours alone", () => {
    // The bug this replaces: a panel dropped half a module down was clustered
    // into a row of its own, so a 3-high wall counted 5 rows and every number
    // after the offset panel was wrong.
    const top = at(0, 0);
    const mid = at(500, 250); // half a module down
    const bottom = at(0, 500);
    const refs = panelGridRefs([top, mid, bottom], rectOf);
    expect(refs.rows).toBe(2);
    expect(label(refs, top, "rows")).toBe("1");
    expect(label(refs, mid, "rows")).toBe("1 & 2");
    expect(label(refs, bottom, "rows")).toBe("2");
  });

  it("counts a panel that spans three cells as a range", () => {
    const tall = at(0, 0, 500, 1500);
    const refs = panelGridRefs([tall, at(500, 0), at(500, 500), at(500, 1000)], rectOf);
    expect(label(refs, tall, "rows")).toBe("1-3");
  });

  it("keeps a wall of 1m-wide MT panels at one column each", () => {
    // The module grid follows what the wall is mostly built from, so MT panels
    // number 1, 2, 3 - not 1&2, 3&4, 5&6 as a fixed 500mm grid would say.
    const panels = [at(0, 0, 1000, 500), at(1000, 0, 1000, 500), at(2000, 0, 1000, 500)];
    const refs = panelGridRefs(panels, rectOf);
    expect(refs.cols).toBe(3);
    expect(panels.map((p) => label(refs, p, "cols"))).toEqual(["1", "2", "3"]);
  });

  it("measures from the wall's own top-left, wherever it sits in the workspace", () => {
    const panels = [at(7000, 3000), at(7500, 3000)];
    const refs = panelGridRefs(panels, rectOf);
    expect(panels.map((p) => `R${label(refs, p, "rows")} C${label(refs, p, "cols")}`)).toEqual(["R1 C1", "R1 C2"]);
  });

  it("ignores removed panels, and copes with an empty wall", () => {
    const gone = at(1000, 0);
    gone.isRemoved = true;
    const refs = panelGridRefs([at(0, 0), at(500, 0), gone], rectOf);
    expect(refs.cols).toBe(2);
    expect(refs.refs.has(gone.id)).toBe(false);
    expect(panelGridRefs([], rectOf)).toMatchObject({ rows: 0, cols: 0 });
  });
});

describe("gridRefLabel", () => {
  const ref = (from: number, to: number): GridRef => ({ from, to });
  it("reads a single cell as a number, two as '&', more as a range", () => {
    expect(gridRefLabel(ref(4, 4))).toBe("4");
    expect(gridRefLabel(ref(5, 6))).toBe("5 & 6");
    expect(gridRefLabel(ref(2, 5))).toBe("2-5");
  });
  it("has something to print for a panel with no reference", () => {
    expect(gridRefLabel(undefined)).toBe("-");
  });
});

describe("mirrorGridRef", () => {
  it("flips a reference for the front view, ends included", () => {
    expect(mirrorGridRef({ from: 1, to: 1 }, 4)).toEqual({ from: 4, to: 4 });
    expect(mirrorGridRef({ from: 4, to: 4 }, 4)).toEqual({ from: 1, to: 1 });
    // A straddling panel still reads low-to-high after the flip.
    expect(gridRefLabel(mirrorGridRef({ from: 2, to: 3 }, 4))).toBe("2 & 3");
    expect(mirrorGridRef(undefined, 4)).toBeUndefined();
  });

  // Every reference number the app prints is read from the FRONT of the wall:
  // the panel labels, the PDF's chain tables, the test pattern and the hover
  // readout. These pin the one relationship they all depend on - column N
  // counted from the front sits (total - N) modules along the layout's own
  // left-to-right geometry - so the number on a panel and the position quoted
  // for it can never drift apart.
  it("puts front column 1 at the far end of the layout's own geometry", () => {
    const panels = [at(0, 0), at(500, 0), at(1000, 0), at(1500, 0)];
    const { refs, cols } = panelGridRefs(panels, (p) => ({ x: p.x, y: p.y, w: 500, h: 500 }));
    expect(cols).toBe(4);
    const frontCol = (p: PanelRecord) => gridRefLabel(mirrorGridRef(refs.get(p.id)?.cols, cols));
    // Left-to-right in the layout's own space counts DOWN from the front.
    expect(panels.map(frontCol)).toEqual(["4", "3", "2", "1"]);
    // ...and the front-referenced x offset of each panel agrees with it: the
    // panel called "1" is at front x 0, which is the far end in this space.
    const frontX = (p: PanelRecord) => 4 * 500 - p.x - 500;
    expect(frontX(panels[3])).toBe(0);
    expect(frontX(panels[0])).toBe(1500);
    panels.forEach((p) => {
      expect({ col: frontCol(p), x: frontX(p) }).toEqual({ col: frontCol(p), x: (Number(frontCol(p)) - 1) * 500 });
    });
  });

  it("keeps rows alone - flipping a wall left to right moves nothing up or down", () => {
    const panels = [at(0, 0), at(0, 500), at(0, 1000)];
    const { refs } = panelGridRefs(panels, (p) => ({ x: p.x, y: p.y, w: 500, h: 500 }));
    expect(panels.map((p) => gridRefLabel(refs.get(p.id)?.rows))).toEqual(["1", "2", "3"]);
  });
});

// ---------------------------------------------------------------------------
// Where a shaped panel's label sits. A triangle or quarter circle leaves one
// corner of its rect empty, and a label placed there is cut away with the
// shape - which is exactly what the PNG exports were doing.
// ---------------------------------------------------------------------------

type Pt = { x: number; y: number };

const ROTATIONS = [0, 90, 180, 270];

// The panel's real silhouette in its own space, traced exactly as
// tracePanelShapePath in App.tsx does: a triangle with its right angle bottom
// left, and a quarter circle with its right angle bottom right, whose curve is
// the quadratic Bezier (w,0) -> control (0,0) -> (0,h), sampled finely.
const SHAPE_OUTLINE: Record<"triangle" | "curve", (w: number, h: number) => Pt[]> = {
  triangle: (w, h) => [{ x: 0, y: 0 }, { x: w, y: h }, { x: 0, y: h }],
  curve: (w, h) => {
    const pts: Pt[] = [{ x: w, y: 0 }, { x: w, y: h }, { x: 0, y: h }];
    const steps = 64;
    for (let i = 1; i < steps; i += 1) {
      const t = i / steps;
      const u = 1 - t;
      // (1-t)^2 * (0,h) + 2t(1-t) * (0,0) + t^2 * (w,0)
      pts.push({ x: t * t * w, y: u * u * h });
    }
    return pts;
  },
};

// Ray casting, so it holds for the sampled curve as well as the triangle.
const insidePolygon = (p: Pt, poly: Pt[]): boolean => {
  let inside = false;
  for (let i = 0, j = poly.length - 1; i < poly.length; j = i, i += 1) {
    const a = poly[i];
    const b = poly[j];
    if (a.y > p.y !== b.y > p.y && p.x < a.x + ((p.y - a.y) / (b.y - a.y)) * (b.x - a.x)) inside = !inside;
  }
  return inside;
};

describe("panelShapeLabelPoint / panelFramePoint", () => {
  it("keeps the whole label block inside the shape at every rotation, front and back", () => {
    const size = 168; // one MG9 panel, native pixels
    const rect: RectMm = { x: 400, y: 900, w: size, h: size };
    // The block the PNG exporters draw: two lines of ~0.08 of the panel, and
    // at most the 0.46 width budget they shrink a long reference to.
    const fontPx = Math.floor(size * 0.08);
    const lineH = Math.round(fontPx * 1.15);
    const halfW = (size * 0.46) / 2;

    for (const shape of ["triangle", "curve"] as const) {
      expect(panelShapeNeedsInsetLabel(shape)).toBe(true);
      const local = panelShapeLabelPoint(shape, rect.w, rect.h);
      for (const rotation of ROTATIONS) {
        for (const mirrorX of [false, true]) {
          // The silhouette's corners, carried through the same frame as the
          // label - so this asserts the label sits inside the shape as drawn.
          const hull = SHAPE_OUTLINE[shape](rect.w, rect.h).map((v) => panelFramePoint(rect, rotation, mirrorX, v.x, v.y));
          const anchor = panelFramePoint(rect, rotation, mirrorX, local.x, local.y);
          const corners: Pt[] = [
            { x: anchor.x - halfW, y: anchor.y - lineH },
            { x: anchor.x + halfW, y: anchor.y - lineH },
            { x: anchor.x - halfW, y: anchor.y + lineH },
            { x: anchor.x + halfW, y: anchor.y + lineH },
          ];
          for (const corner of corners) {
            expect({ shape, rotation, mirrorX, corner, inside: insidePolygon(corner, hull) })
              .toMatchObject({ inside: true });
          }
        }
      }
    }
  });

  it("leaves a plain or corner panel centred, and asks for no inset", () => {
    for (const shape of ["rect", "corner"] as const) {
      expect(panelShapeNeedsInsetLabel(shape)).toBe(false);
      expect(panelShapeLabelPoint(shape, 200, 100)).toEqual({ x: 100, y: 50 });
    }
  });

  it("maps a panel-local point through rotation and then the mirror", () => {
    const rect: RectMm = { x: 0, y: 0, w: 100, h: 100 };
    // Unrotated, unmirrored: the point is where it says it is.
    expect(panelFramePoint(rect, 0, false, 30, 70)).toEqual({ x: 30, y: 70 });
    // Mirrored: reflected about the panel's own centre line.
    expect(panelFramePoint(rect, 0, true, 30, 70)).toEqual({ x: 70, y: 70 });
    // A quarter turn clockwise about the panel centre: (-20, +20) from the
    // centre becomes (-20, -20).
    const turned = panelFramePoint(rect, 90, false, 30, 70);
    expect(turned.x).toBeCloseTo(30, 6);
    expect(turned.y).toBeCloseTo(30, 6);
    const turned180 = panelFramePoint(rect, 180, false, 30, 70);
    expect(turned180.x).toBeCloseTo(70, 6);
    expect(turned180.y).toBeCloseTo(30, 6);
  });
});
