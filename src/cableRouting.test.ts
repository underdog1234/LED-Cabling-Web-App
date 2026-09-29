import { describe, it, expect } from "vitest";
import {
  CABLE_LANE,
  PANEL_LABEL_BOTTOM_PX,
  CABLE_SEPARATION,
  CABLE_STROKE,
  CABLE_CASING_COLOR,
  POWER_COLOR,
  SIGNAL_CABLE_COLOR,
  cableChevronPoints,
  cableRunsClash,
  cableStrokes,
  spreadCableRun,
  panelLabelFontPx,
  routeCablePx,
  type CableKind,
} from "./App";

// Cable runs are painted OVER the panels in both the workspace and the PDF, so
// the route is the only thing keeping a panel's text readable. Every panel
// prints its labels centred and stacked up from the bottom edge, which leaves
// the strip above the topmost label and the margin either side of the text as
// the only places a run may cross a panel. These tests pin that down, plus the
// styling rules the two renderers share: one colour per service, one run (not
// two) where signal and power travel together, and one outline ">" per panel
// entered.

const panel = (x: number, y: number, size = 78) => ({ x, y, w: size, h: size });

/**
 * Widest label block a panel can print, as a box inside the panel's rect: four
 * centred lines of the widest label there is, at its own distance off the
 * bottom edge.
 * Deliberately the worst case in both directions - no real panel prints every
 * line at that width - so a route that clears this clears anything.
 */
const labelBox = (r: { x: number; y: number; w: number; h: number }, fontPx: number) => {
  const textW = 5.9 * fontPx;
  const stackH = 4.9 * fontPx;
  return { x: r.x + (r.w - textW) / 2, y: r.y + r.h - PANEL_LABEL_BOTTOM_PX - stackH, w: textW, h: stackH };
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
  const widest = Math.max(...cableStrokes(kind).map((s) => s.width));
  const strokes: Array<[{ x: number; y: number }, { x: number; y: number }, number]> = [];
  for (let i = 1; i < pts.length; i += 1) strokes.push([pts[i - 1], pts[i], widest]);
  if (entry) {
    const chevron = cableChevronPoints(entry, CABLE_STROKE.chevron);
    for (let i = 1; i < chevron.length; i += 1) strokes.push([chevron[i - 1], chevron[i], widest]);
  }
  return [a, b].some((rect) =>
    strokes.some(([from, to, width]) => segmentHitsBox(from, to, width, labelBox(rect, fontPx))),
  );
};

const KINDS: CableKind[] = ["signal", "power", "both"];

describe("routeCablePx", () => {
  it("runs a straight line along the lane between panels side by side", () => {
    const { pts } = routeCablePx(panel(0, 0), panel(78, 0), "both");
    const laneY = 78 * CABLE_LANE.across;
    const laneX = 78 * CABLE_LANE.beside;
    expect(pts).toEqual([
      { x: laneX, y: laneY },
      { x: 78 + laneX, y: laneY },
    ]);
  });

  it("is visible between touching panels rather than collapsing onto their shared edge", () => {
    // The old edge-to-edge router produced a zero-length hop for a contiguous
    // wall, which is exactly the case every LED wall is made of.
    for (const kind of KINDS) {
      for (const [dx, dy] of [[78, 0], [0, 78]] as Array<[number, number]>) {
        const { pts } = routeCablePx(panel(0, 0), panel(dx, dy), kind);
        const length = pts.slice(1).reduce((sum, p, i) => sum + Math.hypot(p.x - pts[i].x, p.y - pts[i].y), 0);
        expect(length).toBeGreaterThan(70);
      }
    }
  });

  it("meets itself exactly where a chain turns a corner", () => {
    // One anchor per panel, shifted by the same amount in x and y, so the run
    // arriving at a panel ends on the very point the next run leaves from -
    // no step, whichever way either of them travels.
    for (const kind of KINDS) {
      const corner = panel(78, 0);
      const arriving = routeCablePx(panel(0, 0), corner, kind).pts;
      const leaving = routeCablePx(corner, panel(78, 78), kind).pts;
      const end = arriving[arriving.length - 1];
      expect(Math.hypot(end.x - leaving[0].x, end.y - leaving[0].y)).toBeLessThan(0.01);
    }
  });

  it("keeps signal and power apart wherever they do not share the hop", () => {
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
          expect(Math.hypot(s.x - p.x, s.y - p.y)).toBeGreaterThanOrEqual(CABLE_STROKE.casing - 0.01);
        }
      }
    }
  });

  it("never lays a run or its entry mark over a panel's labels", () => {
    const offsets: Array<[number, number]> = [
      [78, 0], [-78, 0], [0, 78], [0, -78], [156, 0], [0, 156], [156, 78], [-156, -78],
    ];
    for (const kind of KINDS) {
      for (const [dx, dy] of offsets) {
        const a = panel(200, 200);
        const b = panel(200 + dx, 200 + dy);
        const fontPx = panelLabelFontPx(a.w, a.h, 10, 4.9, PANEL_LABEL_BOTTOM_PX);
        expect({ kind, dx, dy, hit: routeHitsLabels(a, b, kind, fontPx) }).toEqual({ kind, dx, dy, hit: false });
      }
    }
  });

  it("marks where a run enters the panel it feeds, pointing the way it travels", () => {
    const laneY = 78 * CABLE_LANE.across;
    const right = routeCablePx(panel(0, 0), panel(78, 0), "both").entry;
    expect(right).toEqual({ x: 78, y: laneY, angle: 0 });

    const left = routeCablePx(panel(78, 0), panel(0, 0), "both").entry;
    expect(left).toEqual({ x: 78, y: laneY, angle: Math.PI });

    const down = routeCablePx(panel(0, 0), panel(0, 78), "both").entry;
    expect(down?.y).toBe(78);
    expect(down?.angle).toBeCloseTo(Math.PI / 2);

    const up = routeCablePx(panel(0, 78), panel(0, 0), "both").entry;
    expect(up?.y).toBe(78);
    expect(up?.angle).toBeCloseTo(-Math.PI / 2);
  });

  it("gives a single-panel chain nothing to draw", () => {
    const { pts, entry } = routeCablePx(panel(0, 0), panel(0, 0), "power");
    expect(pts).toHaveLength(1);
    expect(entry).toBeNull();
  });
});

describe("runs that land on the same lane", () => {
  // A chain that doubles back on itself, or a return leg retracing one that
  // went out earlier, puts two cables on one lane. Drawn as they come out of
  // the router that reads as a SINGLE line with two arrows piled on it, which
  // is exactly the case that has to show as two.
  const routeOf = (a: ReturnType<typeof panel>, b: ReturnType<typeof panel>) => routeCablePx(a, b, "both");
  const segmentsOf = (pts: Array<{ x: number; y: number }>) => {
    const segs: Array<{ horiz: boolean; fixed: number; lo: number; hi: number }> = [];
    for (let i = 1; i < pts.length; i += 1) {
      const from = pts[i - 1];
      const to = pts[i];
      const horiz = Math.abs(from.y - to.y) < 0.01;
      if (!horiz && Math.abs(from.x - to.x) >= 0.01) continue;
      segs.push(horiz
        ? { horiz, fixed: from.y, lo: Math.min(from.x, to.x), hi: Math.max(from.x, to.x) }
        : { horiz, fixed: from.x, lo: Math.min(from.y, to.y), hi: Math.max(from.y, to.y) });
    }
    return segs;
  };

  it("spots two runs drawn on top of each other, and leaves unrelated ones alone", () => {
    const there = routeOf(panel(0, 0), panel(78, 0));
    const back = routeOf(panel(78, 0), panel(0, 0));
    expect(cableRunsClash(segmentsOf(there.pts), segmentsOf(back.pts))).toBe(true);

    const elsewhere = routeOf(panel(0, 78), panel(78, 78));
    expect(cableRunsClash(segmentsOf(there.pts), segmentsOf(elsewhere.pts))).toBe(false);
  });

  it("moves the second run clear, keeping its own end points and its own arrow", () => {
    const dest = panel(0, 0);
    const back = routeOf(panel(78, 0), dest);
    const moved = spreadCableRun(back, dest, -CABLE_STROKE.casing);

    // Same start and finish, so it still meets the hops either side of it.
    expect(moved.pts[0]).toEqual(back.pts[0]);
    expect(moved.pts[moved.pts.length - 1]).toEqual(back.pts[back.pts.length - 1]);
    // ...but no longer sharing a lane with the run that was already there.
    const there = routeOf(panel(0, 0), panel(78, 0));
    expect(cableRunsClash(segmentsOf(there.pts), segmentsOf(moved.pts))).toBe(false);
    // Its direction mark comes with it, so both runs are arrowed.
    expect(moved.entry).not.toBeNull();
    expect(moved.entry!.y).toBeLessThan(back.entry!.y);
  });

  it("always steps a run AWAY from the labels it would otherwise land on", () => {
    // The nudge is negative - up for a run crossing a panel, left for one
    // running down it - so it can never eat the clearance the label block was
    // sized against.
    const dest = panel(200, 200);
    const route = routeCablePx(panel(200, 122), dest, "both");
    const moved = spreadCableRun(route, dest, -CABLE_STROKE.casing);
    const fontPx = panelLabelFontPx(dest.w, dest.h, 10, 4.9, PANEL_LABEL_BOTTOM_PX);
    const widest = Math.max(...cableStrokes("both").map((s) => s.width));
    const hit = moved.pts.slice(1).some((p, i) => segmentHitsBox(moved.pts[i], p, widest, labelBox(dest, fontPx)));
    expect(hit).toBe(false);
  });

  it("leaves a run untouched when there is nothing to step around", () => {
    const route = routeOf(panel(0, 0), panel(78, 0));
    expect(spreadCableRun(route, panel(78, 0), 0)).toBe(route);
  });
});

describe("cableStrokes", () => {
  it("draws signal blue and power orange, whatever port they belong to", () => {
    expect(cableStrokes("signal").map((s) => s.color)).toEqual([CABLE_CASING_COLOR, SIGNAL_CABLE_COLOR]);
    expect(cableStrokes("power").map((s) => s.color)).toEqual([CABLE_CASING_COLOR, POWER_COLOR]);
  });

  it("draws a shared hop as one cable carrying both, not two side by side", () => {
    const both = cableStrokes("both");
    expect(both.map((s) => s.color)).toEqual([CABLE_CASING_COLOR, SIGNAL_CABLE_COLOR, POWER_COLOR]);
    // The power pass is dashed, so the single run reads as blue-and-orange
    // rather than hiding the signal underneath it.
    expect(both[1].dash).toBeUndefined();
    expect(both[2].dash?.length).toBe(2);
    // ...and it is no wider than a run of one service on its own.
    expect(Math.max(...both.map((s) => s.width))).toBe(Math.max(...cableStrokes("signal").map((s) => s.width)));
  });

  it("scales every pass together", () => {
    const half = cableStrokes("both", 0.5);
    expect(half.map((s) => s.width)).toEqual(cableStrokes("both", 1).map((s) => s.width / 2));
    expect(half[2].dash).toEqual([3, 3]);
  });
});

describe("cableChevronPoints", () => {
  it("is an open '>' pointing the way the run travels, straddling the edge it crosses", () => {
    const [back1, apex, back2] = cableChevronPoints({ x: 100, y: 50, angle: 0 }, 10);
    // Apex just past the crossing point, legs just short of it, so the mark
    // sits ON the edge rather than wholly inside either panel.
    expect(apex.x).toBeGreaterThan(100);
    expect(back1.x).toBeLessThan(100);
    expect(back2.x).toBeLessThan(100);
    expect(Math.min(back1.x, back2.x)).toBeGreaterThan(100 - 10 / 2);
    // One leg either side of the run.
    expect(Math.sign(back1.y - apex.y)).toBe(-Math.sign(back2.y - apex.y));
    // Three points, so it can only ever be stroked open - never filled into a
    // solid arrowhead.
    expect(cableChevronPoints({ x: 0, y: 0, angle: Math.PI / 2 }, 10)).toHaveLength(3);
  });
});

describe("panelLabelFontPx", () => {
  it("keeps the sizes both renderers have always used on a full panel", () => {
    expect(panelLabelFontPx(78, 78, 9, 5, PANEL_LABEL_BOTTOM_PX + 5.5)).toBe(9); // workspace at 100% zoom
    expect(panelLabelFontPx(78, 78, 10, 4.9, PANEL_LABEL_BOTTOM_PX)).toBe(10); // PDF
  });

  it("never grows past the renderer's own size on a big panel", () => {
    expect(panelLabelFontPx(234, 234, 9, 5, PANEL_LABEL_BOTTOM_PX + 5.5)).toBe(9);
  });

  it("steps down on a panel too narrow for the text, like a poster section", () => {
    const poster = panelLabelFontPx(49.9, 74.9, 10, 4.9, PANEL_LABEL_BOTTOM_PX);
    expect(poster).toBeGreaterThan(0);
    expect(poster).toBeLessThan(10);
  });

  it("gives up rather than spill text over a panel that is too small for any of it", () => {
    expect(panelLabelFontPx(39, 39, 9, 5, PANEL_LABEL_BOTTOM_PX + 5.5)).toBe(0);
  });

  it("leaves room for the separation between an unshared signal and power run", () => {
    expect(CABLE_SEPARATION * 2).toBeGreaterThanOrEqual(CABLE_STROKE.casing);
  });
});
