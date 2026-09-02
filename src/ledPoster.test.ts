import { describe, it, expect } from "vitest";
import { PANEL_TYPES, POSTER_SECTIONS, makePosterAt, makePosterPanels, normalizePanels, spareBucketOfCell, cellRect } from "./App";

// A complete LED poster is 640 x 1920mm / 344 x 1032px, but it lives in the
// grid as four stacked 640 x 480mm / 344 x 258px sections sharing a
// posterGroupId, so layout/patching/pixel maths all work in one panel unit.
describe("LED poster geometry", () => {
  it("stores a section, not the whole poster", () => {
    const p = PANEL_TYPES.POSTER;
    expect(p.w).toBe(0.64);
    expect(p.h).toBe(0.48);
    expect(p.pixW).toBe(344);
    expect(p.pixH).toBe(258);
  });

  it("four sections reconstruct the complete poster exactly", () => {
    const p = PANEL_TYPES.POSTER;
    expect(p.w * 1000).toBe(640);
    expect(p.h * 1000 * POSTER_SECTIONS).toBe(1920);
    expect(p.pixW).toBe(344);
    expect(p.pixH * POSTER_SECTIONS).toBe(1032);
  });

  it("makePosterAt stacks four sections with no gap or overlap", () => {
    const cells = makePosterAt(0, 0);
    expect(cells).toHaveLength(POSTER_SECTIONS);
    const rects = cells.map(cellRect).sort((a, b) => a.y - b.y);
    rects.forEach((r, i) => {
      expect(r.x).toBe(0);
      expect(r.w).toBe(640);
      expect(r.h).toBe(480);
      expect(r.y).toBe(i * 480);
      if (i > 0) expect(r.y).toBe(rects[i - 1].y + rects[i - 1].h);
    });
    const top = rects[0], bottom = rects[rects.length - 1];
    expect(bottom.y + bottom.h - top.y).toBe(1920);
  });

  it("all four sections of one poster share a group id, and posters differ", () => {
    const a = makePosterAt(0, 0);
    const b = makePosterAt(640, 0);
    expect(new Set(a.map((c) => c.posterGroupId)).size).toBe(1);
    expect(new Set(b.map((c) => c.posterGroupId)).size).toBe(1);
    expect(a[0].posterGroupId).not.toBe(b[0].posterGroupId);
    expect(a[0].posterGroupId).toBeTruthy();
  });

  it("makePosterPanels lays complete posters side by side", () => {
    const cells = makePosterPanels(3);
    expect(cells).toHaveLength(3 * POSTER_SECTIONS);
    expect(new Set(cells.map((c) => c.posterGroupId)).size).toBe(3);
    expect(new Set(cells.map((c) => c.x)).size).toBe(3);
    // One poster high: every section sits within the poster's 1920mm height.
    cells.forEach((c) => expect(c.y).toBeLessThan(1920));
  });

  it("posters get their own spare bucket rather than being lumped in with MT", () => {
    expect(spareBucketOfCell(makePosterAt(0, 0)[0])).toBe("POSTER");
  });

  it("survives a save/load round trip with its grouping intact", () => {
    const cells = makePosterPanels(2);
    const restored = normalizePanels(JSON.parse(JSON.stringify(cells)));
    expect(restored).toHaveLength(8);
    expect(new Set(restored.map((c) => c.posterGroupId)).size).toBe(2);
    restored.forEach((c) => expect(c.panelType).toBe("POSTER"));
  });

  it("leaves non-poster panels ungrouped", () => {
    const restored = normalizePanels([{ x: 0, y: 0, panelType: "MG9" }]);
    expect(restored[0].posterGroupId).toBeNull();
  });
});
