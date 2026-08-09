import { describe, it, expect } from "vitest";
import { panelWorldAnchors, computeAnchorSnapDelta, type PanelAnchorSpec } from "./panels";

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
