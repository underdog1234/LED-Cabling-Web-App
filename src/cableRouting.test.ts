import { describe, it, expect } from "vitest";
import {
  CABLE_LANE,
  CABLE_STROKE,
  cableChevronPoints,
  cableColorFor,
  panelLabelFontPx,
  routeCablePx,
  PORT_COLORS,
  type CableKind,
} from "./App";

// Cable runs are painted OVER the panels in both the workspace and the PDF, so
// the route is the only thing keeping a panel's text readable. Every panel
// prints its labels centred and stacked up from the bottom edge, which leaves
// the strip above the topmost label and the margin either side of the text as
// the only places a run may cross a panel. These tests pin that down, plus the
// styling rules the two renderers share: signal is a plain line, power carries
// exactly one outline ">" per panel it enters.

const panel = (x: number, y: number, size = 78) => ({ x, y, w: size, h: size });

/**
 * Widest label block a panel can print, as a box inside the panel's rect: four
 * centred lines of the widest label there is, sitting 4px off the bottom edge.
 * Deliberately the worst case in both directions - no real panel prints every
 * line at that width - so a route that clears this clears anything.
 */
const labelBox = (r: { x: number; y: number; w: number; h: number }, fontPx: number) => {
  const textW = 5.9 * fontPx;
  const stackH = 4.9 * fontPx;
  return { x: r.x + (r.w - textW) / 2, y: r.y + r.h - 4 - stackH, w: textW, h: stackH };
};

/** Does a stroked segment of `width` touch `box`? */
const segmentHitsBox = (
  a: { x: number; y: number },
  b: { x: number; y: number },
  width: number,
  box: { x: number; y: number; w: number; h: number },
) => {
  const half = width / 2;
  const steps = Math.max(2, Math.ceil(Math.hypot(b.x - a.x, b.y - a.y) * 4));
  for (let i = 0; i <= steps; i += 1) {
    const t = i / steps;
    const px = a.x + (b.x - a.x) * t;
    const py = a.y + (b.y - a.y) * t;
    const dx = Math.max(box.x - px, 0, px - (box.x + box.w));
    const dy = Math.max(box.y - py, 0, py - (box.y + box.h));
    if (Math.hypot(dx, dy) < half - 0.001) return true;
  }
  return false;
};

const routeHitsLabels = (
  a: { x: number; y: number; w: number; h: number },
  b: { x: number; y: number; w: number; h: number },
  kind: CableKind,
  fontPx: number,
) => {
  const { pts, entry } = routeCablePx(a, b, kind);
  const strokes: Array<[{ x: number; y: number }, { x: number; y: number }, number]> = [];
  for (let i = 1; i < pts.length; i += 1) strokes.push([pts[i - 1], pts[i], CABLE_STROKE.casing]);
  if (kind === "power" && entry) {
    const chevron = cableChevronPoints(entry, CABLE_STROKE.chevron);
    for (let i = 1; i < chevron.length; i += 1) strokes.push([chevron[i - 1], chevron[i], CABLE_STROKE.casing - 1]);
  }
  return [a, b].some((rect) =>
    strokes.some(([from, to, width]) => segmentHitsBox(from, to, width, labelBox(rect, fontPx))),
  );
};

describe("routeCablePx", () => {
  it("runs a straight line along the lane between panels side by side", () => {
    const { pts } = routeCablePx(panel(0, 0), panel(78, 0), "signal");
    expect(pts).toEqual([
      { x: 39, y: 78 * CABLE_LANE.signal.across },
      { x: 117, y: 78 * CABLE_LANE.signal.across },
    ]);
  });

  it("is visible between touching panels rather than collapsing onto their shared edge", () => {
    // The old edge-to-edge router produced a zero-length hop for a contiguous
    // wall, which is exactly the case every LED wall is made of.
    for (const kind of ["signal", "power"] as CableKind[]) {
      const { pts } = routeCablePx(panel(0, 0), panel(78, 0), kind);
      const length = pts.slice(1).reduce((sum, p, i) => sum + Math.hypot(p.x - pts[i].x, p.y - pts[i].y), 0);
      expect(length).toBeGreaterThan(70);
    }
  });

  it("steps out to the side lane before running up or down", () => {
    const { pts } = routeCablePx(panel(0, 0), panel(0, 78), "signal");
    const sideX = 78 * CABLE_LANE.signal.beside;
    expect(pts).toEqual([
      { x: 39, y: 78 * CABLE_LANE.signal.across },
      { x: sideX, y: 78 * CABLE_LANE.signal.across },
      { x: sideX, y: 78 + 78 * CABLE_LANE.signal.across },
      { x: 39, y: 78 + 78 * CABLE_LANE.signal.across },
    ]);
  });

  it("keeps signal and power apart on every hop direction", () => {
    const pairs: Array<[ReturnType<typeof panel>, ReturnType<typeof panel>]> = [
      [panel(0, 0), panel(78, 0)],
      [panel(78, 0), panel(0, 0)],
      [panel(0, 0), panel(0, 78)],
      [panel(0, 78), panel(0, 0)],
    ];
    for (const [a, b] of pairs) {
      const signal = routeCablePx(a, b, "signal").pts;
      const power = routeCablePx(a, b, "power").pts;
      for (const s of signal) {
        for (const p of power) {
          // Same point would mean the two runs meet; they share end x/y only
          // because both start and finish at a panel's centre line.
          expect(Math.hypot(s.x - p.x, s.y - p.y)).toBeGreaterThan(CABLE_STROKE.casing / 2);
        }
      }
    }
  });

  it("never lays a run or its entry mark over a panel's labels", () => {
    const offsets: Array<[number, number]> = [
      [78, 0], [-78, 0], [0, 78], [0, -78], [156, 0], [0, 156], [156, 78],
    ];
    for (const kind of ["signal", "power"] as CableKind[]) {
      for (const [dx, dy] of offsets) {
        const a = panel(200, 200);
        const b = panel(200 + dx, 200 + dy);
        const fontPx = panelLabelFontPx(a.w, a.h, 10, 4.9, 4);
        expect({ kind, dx, dy, hit: routeHitsLabels(a, b, kind, fontPx) }).toEqual({ kind, dx, dy, hit: false });
      }
    }
  });

  it("marks where a power run enters the panel it feeds, pointing the way it travels", () => {
    const right = routeCablePx(panel(0, 0), panel(78, 0), "power").entry;
    expect(right).toEqual({ x: 78, y: 78 * CABLE_LANE.power.across, angle: 0 });

    const left = routeCablePx(panel(78, 0), panel(0, 0), "power").entry;
    expect(left).toEqual({ x: 78, y: 78 * CABLE_LANE.power.across, angle: Math.PI });

    const down = routeCablePx(panel(0, 0), panel(0, 78), "power").entry;
    expect(down?.y).toBe(78);
    expect(down?.angle).toBeCloseTo(Math.PI / 2);

    const up = routeCablePx(panel(0, 78), panel(0, 0), "power").entry;
    expect(up?.y).toBe(78);
    expect(up?.angle).toBeCloseTo(-Math.PI / 2);
  });

  it("gives a single-panel chain nothing to draw", () => {
    const { pts, entry } = routeCablePx(panel(0, 0), panel(0, 0), "power");
    expect(pts).toHaveLength(1);
    expect(entry).toBeNull();
  });
});

describe("cableChevronPoints", () => {
  it("is an open '>' pointing the way the run travels, just inside the panel it enters", () => {
    const [back1, apex, back2] = cableChevronPoints({ x: 100, y: 50, angle: 0 }, 10);
    // Apex ahead of the crossing point, legs back at the panel's own edge.
    expect(apex.x).toBeGreaterThan(100);
    expect(back1.x).toBeLessThan(apex.x);
    expect(back2.x).toBeLessThan(apex.x);
    expect(Math.min(back1.x, back2.x)).toBeGreaterThan(99);
    // One leg either side of the run.
    expect(Math.sign(back1.y - apex.y)).toBe(-Math.sign(back2.y - apex.y));
    // Three points, so it can only ever be stroked open - never filled into a
    // solid arrowhead.
    expect(cableChevronPoints({ x: 0, y: 0, angle: Math.PI / 2 }, 10)).toHaveLength(3);
  });
});

describe("cableColorFor", () => {
  it("darkens a signal run so it reads over panels filled with its own port colour", () => {
    expect(cableColorFor("signal", "#48d7d2")).toBe("#287674");
    expect(cableColorFor("signal", PORT_COLORS[0])).not.toBe(PORT_COLORS[0]);
  });

  it("leaves power alone - its colour never matches a panel fill", () => {
    expect(cableColorFor("power", "#f97316")).toBe("#f97316");
  });
});

describe("panelLabelFontPx", () => {
  it("keeps the sizes both renderers have always used on a full panel", () => {
    expect(panelLabelFontPx(78, 78, 9, 5, 9.5)).toBe(9); // workspace at 100% zoom
    expect(panelLabelFontPx(78, 78, 10, 4.9, 4)).toBe(10); // PDF
  });

  it("never grows past the renderer's own size on a big panel", () => {
    expect(panelLabelFontPx(234, 234, 9, 5, 9.5)).toBe(9);
  });

  it("steps down on a panel too narrow for the text, like a poster section", () => {
    const poster = panelLabelFontPx(49.9, 74.9, 10, 4.9, 4);
    expect(poster).toBeGreaterThan(0);
    expect(poster).toBeLessThan(10);
  });

  it("gives up rather than spill text over a panel that is too small for any of it", () => {
    expect(panelLabelFontPx(39, 39, 9, 5, 9.5)).toBe(0);
  });
});
