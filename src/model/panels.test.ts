import { describe, it, expect } from "vitest";
import {
  panelWorldAnchors,
  computeAnchorSnapDelta,
  panelsAnchorJoined,
  sharedEdgeOrientation,
  QUARTER_MODULE_MM,
  type PanelAnchorSpec,
} from "./panels";

// Regression coverage for a real bug: panelWorldAnchors rounded a panel's
// rotation to the nearest 90deg before computing its connector anchor
// positions, so a panel spun to a custom angle (e.g. 45deg) snapped/joined
// as if it were unrotated (or at the nearest cardinal angle) instead of
// along its own true rotated edges. Harmless while rotation was always a
// multiple of 90, but a real bug once custom-angle rotation shipped.
describe("panelWorldAnchors respects the panel's full rotation", () => {
  it("rotates anchors by the exact angle, not the nearest 90 degrees", () => {
    const g: PanelAnchorSpec = { cx: 0, cy: 0, halfW: 250, halfH: 250, rotation: 45, shape: "rect" };
    const anchors = panelWorldAnchors(g);
    const expectedRightMid = { x: 250 * Math.cos(Math.PI / 4), y: 250 * Math.sin(Math.PI / 4) };
    const rotated = anchors.some((a) => Math.abs(a.x - expectedRightMid.x) < 0.5 && Math.abs(a.y - expectedRightMid.y) < 0.5);
    expect(rotated).toBe(true);
    // The old (buggy) nearest-90 rounding would have left this anchor at (250, 0).
    const stillCardinal = anchors.some((a) => Math.abs(a.x - 250) < 0.5 && Math.abs(a.y) < 0.5);
    expect(stillCardinal).toBe(false);
  });
});

describe("computeAnchorSnapDelta snaps along a rotated panel's own axes", () => {
  it("snaps a 45deg-rotated moving panel toward a matching 45deg-rotated stationary panel's anchor, diagonally", () => {
    const stationary: PanelAnchorSpec = { cx: 0, cy: 0, halfW: 250, halfH: 250, rotation: 45, shape: "rect" };
    const targetX = 250 * Math.cos(Math.PI / 4);
    const targetY = 250 * Math.sin(Math.PI / 4);
    // Moving panel's own "left-mid" anchor (-250, 0), rotated 45deg, sits at
    // (mcx - targetX, mcy - targetY) - placed 5mm off in both x and y from
    // the stationary panel's "right-mid" anchor (within SNAP_DISTANCE_MM).
    const offset = 5;
    const moving: PanelAnchorSpec = { cx: 2 * targetX + offset, cy: 2 * targetY + offset, halfW: 250, halfH: 250, rotation: 45, shape: "rect" };
    const result = computeAnchorSnapDelta([moving], [stationary], true, { x: moving.cx - 250, y: moving.cy - 250, w: 500, h: 500 });
    expect(result.snappedTo).toBe("panel");
    // The correcting delta must itself be diagonal (~equal x/y), matching the
    // true 45deg geometry - the old bug computed anchors as if both panels
    // were unrotated squares, giving a materially different (wrong) delta.
    expect(result.dx).toBeCloseTo(-offset, 1);
    expect(result.dy).toBeCloseTo(-offset, 1);
  });
});

// ---------------------------------------------------------------------------
// Which offsets two panels can actually be joined on.
//
// The anchors used to sit at fractions of the half extent (+/-0.8, +/-0.4, 0),
// which on a 500mm edge put them every 100mm. That let panels snap to 100, 200
// and 300 - none of them a half or a quarter of a module - while 125 and 250
// had no anchor pair at all. Dragging a panel exactly half a module down fell
// out of anchor range, dropped through to the 250mm grid snap, and landed
// LOOKING right while counting as not joined: no connector ordered for it, and
// the two panels moved independently afterwards. That is the bug these cover.
// ---------------------------------------------------------------------------

/** A 500x500 panel whose top-left corner is at (x, y). */
const panelAt = (x: number, y: number, shape: PanelAnchorSpec["shape"] = "rect", rotation = 0): PanelAnchorSpec =>
  ({ cx: x + 250, cy: y + 250, halfW: 250, halfH: 250, rotation, shape });

const snapTo = (wantedY: number) => {
  const moving = panelAt(500, wantedY);
  const { dy } = computeAnchorSnapDelta([moving], [panelAt(0, 0)], true, { x: 500, y: wantedY, w: 500, h: 500 });
  return wantedY + dy;
};

describe("side-by-side joins land on halves and quarters", () => {
  it("snaps a half-module offset to exactly half, and counts it as joined", () => {
    // The reported case: dragged to 250mm down, it used to land on 250 only
    // via the grid fallback and read as NOT joined.
    expect(snapTo(250)).toBe(250);
    expect(panelsAnchorJoined(panelAt(0, 0), panelAt(500, 250))).toBe(true);
    expect(sharedEdgeOrientation(panelAt(0, 0), panelAt(500, 250))).toBe("vertical");
  });

  it("only ever lands on a quarter of a module", () => {
    for (let wanted = 0; wanted <= 500; wanted += 10) {
      const landed = snapTo(wanted);
      expect({ wanted, onLattice: landed % QUARTER_MODULE_MM === 0 }).toEqual({ wanted, onLattice: true });
    }
  });

  it("joins on every quarter, and on nothing in between", () => {
    for (const offset of [0, 125, 250, 375]) {
      expect({ offset, joined: panelsAnchorJoined(panelAt(0, 0), panelAt(500, offset)) }).toEqual({ offset, joined: true });
    }
    for (const offset of [100, 200, 300, 400]) {
      expect({ offset, joined: panelsAnchorJoined(panelAt(0, 0), panelAt(500, offset)) }).toEqual({ offset, joined: false });
    }
  });

  it("sees a brick-bond stack, which the old lattice could not", () => {
    // Half a module across and one row down - panels really do share half an
    // edge there, and the connectors for it were never being ordered.
    expect(panelsAnchorJoined(panelAt(0, 0), panelAt(250, 500))).toBe(true);
    expect(sharedEdgeOrientation(panelAt(0, 0), panelAt(250, 500))).toBe("horizontal");
  });

  it("still refuses two panels that only touch at a corner", () => {
    // Anchors now include the edge ends, and a corner belongs to two edges -
    // counted twice it would read as the two coincident anchors a join needs.
    expect(panelsAnchorJoined(panelAt(0, 0), panelAt(500, 500))).toBe(false);
    expect(panelsAnchorJoined(panelAt(0, 0), panelAt(-500, 500))).toBe(false);
    // A full module offset side by side is the same corner-only touch.
    expect(panelsAnchorJoined(panelAt(0, 0), panelAt(500, 500))).toBe(false);
  });

  it("keeps flush joins, stacked and side by side", () => {
    expect(panelsAnchorJoined(panelAt(0, 0), panelAt(500, 0))).toBe(true);
    expect(panelsAnchorJoined(panelAt(0, 0), panelAt(0, 500))).toBe(true);
  });

  it("gives a shaped panel the same lattice on its straight legs", () => {
    // A triangle's legs are its left and bottom edges; a rect sitting on its
    // left edge, half a module down, joins exactly as two rects would.
    expect(panelsAnchorJoined(panelAt(500, 0, "triangle"), panelAt(0, 250))).toBe(true);
    // Its hypotenuse carries no anchors, so nothing joins along it.
    expect(panelsAnchorJoined(panelAt(0, 0, "triangle"), panelAt(500, 250))).toBe(false);
  });

  it("puts an MT panel on the same lattice as an MG9 one", () => {
    // MT is 1000mm wide, so a fraction-of-the-edge lattice would have spaced
    // its anchors twice as far apart and the two could not meet on a quarter.
    const mt: PanelAnchorSpec = { cx: 500, cy: 250, halfW: 500, halfH: 250, rotation: 0, shape: "rect" };
    expect(panelsAnchorJoined(mt, panelAt(1000, 125))).toBe(true);
    expect(panelsAnchorJoined(mt, panelAt(1000, 250))).toBe(true);
    expect(panelsAnchorJoined(mt, panelAt(1000, 100))).toBe(false);
  });
});
