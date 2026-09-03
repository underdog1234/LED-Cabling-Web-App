import { describe, it, expect } from "vitest";
import {
  PANEL_TYPES,
  POSTER_COLS,
  POSTER_ROWS,
  POSTER_SECTIONS,
  POSTER_WATTS_PER_UNIT,
  makePosterAt,
  makePosterPanels,
  normalizePanels,
  spareBucketOfCell,
  cellRect,
  getSelectedIds,
  makePosterUnits,
  POSTER_WIDTH_MM,
} from "./App";

// A complete LED poster is 640 x 1920mm / 344 x 1032px, but it lives in the
// grid as a 2-wide x 4-high block of 320 x 480mm / 172 x 258px sections sharing
// a posterGroupId, so layout/patching/pixel maths all work in one panel unit.
describe("LED poster geometry", () => {
  it("stores a section, not the whole poster", () => {
    const p = PANEL_TYPES.POSTER;
    expect(p.w).toBe(0.32);
    expect(p.h).toBe(0.48);
    expect(p.pixW).toBe(172);
    expect(p.pixH).toBe(258);
  });

  it("splits two wide and four high", () => {
    expect(POSTER_COLS).toBe(2);
    expect(POSTER_ROWS).toBe(4);
    expect(POSTER_SECTIONS).toBe(8);
  });

  it("the section block reconstructs the complete poster exactly", () => {
    const p = PANEL_TYPES.POSTER;
    expect(p.w * 1000 * POSTER_COLS).toBe(640);
    expect(p.h * 1000 * POSTER_ROWS).toBe(1920);
    expect(p.pixW * POSTER_COLS).toBe(344);
    expect(p.pixH * POSTER_ROWS).toBe(1032);
    expect(POSTER_WIDTH_MM).toBe(640);
  });

  it("makePosterAt tiles a 2x4 block with no gap or overlap", () => {
    const cells = makePosterAt(0, 0);
    expect(cells).toHaveLength(POSTER_SECTIONS);
    const rects = cells.map(cellRect);
    rects.forEach((r) => {
      expect(r.w).toBe(320);
      expect(r.h).toBe(480);
    });
    // Every one of the 8 grid positions is occupied exactly once.
    const seen = rects.map((r) => `${r.x},${r.y}`).sort();
    const want: string[] = [];
    for (let row = 0; row < POSTER_ROWS; row += 1) {
      for (let col = 0; col < POSTER_COLS; col += 1) want.push(`${col * 320},${row * 480}`);
    }
    expect(seen).toEqual([...want].sort());
    // ...and together they cover exactly the poster's real footprint.
    const right = Math.max(...rects.map((r) => r.x + r.w));
    const bottom = Math.max(...rects.map((r) => r.y + r.h));
    expect(right - Math.min(...rects.map((r) => r.x))).toBe(640);
    expect(bottom - Math.min(...rects.map((r) => r.y))).toBe(1920);
    expect(rects.reduce((sum, r) => sum + r.w * r.h, 0)).toBe(640 * 1920);
  });

  it("places the block at the requested origin", () => {
    const rects = makePosterAt(1000, 500).map(cellRect);
    expect(Math.min(...rects.map((r) => r.x))).toBe(1000);
    expect(Math.min(...rects.map((r) => r.y))).toBe(500);
  });

  it("all sections of one poster share a group id, and posters differ", () => {
    const a = makePosterAt(0, 0);
    const b = makePosterAt(POSTER_WIDTH_MM, 0);
    expect(new Set(a.map((c) => c.posterGroupId)).size).toBe(1);
    expect(new Set(b.map((c) => c.posterGroupId)).size).toBe(1);
    expect(a[0].posterGroupId).not.toBe(b[0].posterGroupId);
    expect(a[0].posterGroupId).toBeTruthy();
  });

  it("makePosterPanels lays complete posters side by side without overlapping", () => {
    const cells = makePosterPanels(3);
    expect(cells).toHaveLength(3 * POSTER_SECTIONS);
    expect(new Set(cells.map((c) => c.posterGroupId)).size).toBe(3);
    // 3 posters x 2 section columns each = 6 distinct x positions.
    expect(new Set(cells.map((c) => c.x)).size).toBe(3 * POSTER_COLS);
    // One poster high: every section sits within the poster's 1920mm height.
    cells.forEach((c) => expect(c.y).toBeLessThan(1920));
    // Neighbouring posters butt up against each other, never overlap.
    const rects = cells.map(cellRect);
    expect(Math.max(...rects.map((r) => r.x + r.w))).toBe(3 * POSTER_WIDTH_MM);
  });

  it("posters get their own spare bucket rather than being lumped in with MT", () => {
    expect(spareBucketOfCell(makePosterAt(0, 0)[0])).toBe("POSTER");
  });

  it("survives a save/load round trip with its grouping intact", () => {
    const cells = makePosterPanels(2);
    const restored = normalizePanels(JSON.parse(JSON.stringify(cells)));
    expect(restored).toHaveLength(2 * POSTER_SECTIONS);
    expect(new Set(restored.map((c) => c.posterGroupId)).size).toBe(2);
    restored.forEach((c) => expect(c.panelType).toBe("POSTER"));
  });

  it("leaves non-poster panels ungrouped", () => {
    const restored = normalizePanels([{ x: 0, y: 0, panelType: "MG9" }]);
    expect(restored[0].posterGroupId).toBeNull();
  });
});

// Power is specified per COMPLETE poster (575.00W); the catalog stores one
// section, so the per-section figure is that divided by eight. Amps use the
// same 230V basis as every other panel in the catalog.
describe("LED poster power", () => {
  const VOLTAGE = 230;

  it("all sections add up to exactly 575W per poster", () => {
    expect(POSTER_WATTS_PER_UNIT).toBe(575);
    expect(PANEL_TYPES.POSTER.power.maxW * POSTER_SECTIONS).toBe(575);
    expect(PANEL_TYPES.POSTER.power.maxW).toBe(71.875);
  });

  it("derives amps from watts at 230V, consistently with MG9 and MT", () => {
    const p = PANEL_TYPES.POSTER.power;
    expect(p.maxA * POSTER_SECTIONS).toBeCloseTo(575 / VOLTAGE, 6);
    expect(p.maxA).toBeCloseTo(p.maxW / VOLTAGE, 6);
    // The existing entries are stored rounded to 2dp, so they only imply
    // ~227-229V rather than exactly 230 - check they share the same basis
    // within that rounding, not that they match to the decimal.
    for (const spec of [PANEL_TYPES.MG9, PANEL_TYPES.MT, PANEL_TYPES.POSTER]) {
      expect(spec.power.maxW / spec.power.maxA).toBeGreaterThan(225);
      expect(spec.power.maxW / spec.power.maxA).toBeLessThanOrEqual(230);
    }
  });

  it("uses peak for average too, since only one figure was supplied", () => {
    const p = PANEL_TYPES.POSTER.power;
    expect(p.avgW).toBe(p.maxW);
    expect(p.avgA).toBe(p.maxA);
  });

  it("keeps a whole poster's draw inside one 16A outlet allowance", () => {
    const perOutletW = PANEL_TYPES.POSTER.defaults.powerPanelsPerOutlet * PANEL_TYPES.POSTER.power.maxW;
    expect(perOutletW).toBeLessThanOrEqual(16 * VOLTAGE);
    // Whole posters only - never a part-poster on an outlet.
    expect(PANEL_TYPES.POSTER.defaults.powerPanelsPerOutlet % POSTER_SECTIONS).toBe(0);
  });

  it("carries weight of 0 by choice, so posters add nothing to weight totals", () => {
    expect(PANEL_TYPES.POSTER.weight).toBe(0);
  });
});

// Regression: signalPanelsPerPort was 28 sections, computed by conflating
// posters with sections - 28 x 88,752px was 382% of the 650,000px-per-port
// ceiling the UI checks against. Every type's default must fit under it.
describe("panels per signal port stay inside the 650,000px ceiling", () => {
  const MAX_PIXELS_PER_PORT = 650000;

  it.each(["MG9", "MT", "POSTER"] as const)("%s default fits on one port", (key) => {
    const spec = PANEL_TYPES[key];
    expect(spec.defaults.signalPanelsPerPort * spec.pixW * spec.pixH).toBeLessThanOrEqual(MAX_PIXELS_PER_PORT);
  });

  it("gives a poster port whole posters, not part of one", () => {
    expect(PANEL_TYPES.POSTER.defaults.signalPanelsPerPort % POSTER_SECTIONS).toBe(0);
  });
});

// Regression: a catch-all `else` in topRowBars swept POSTER sections into the
// MG9 tally, which handed them MG9's fly-bar/sling weights and ordered MG9
// hanging bars for them. Posters carry their own (zero) rigging hardware.
describe("LED poster rigging hardware", () => {
  it("has its own fly bar and sling weights, not MG9's", () => {
    expect(PANEL_TYPES.POSTER.defaults.flyBarWeight).toBe(0);
    expect(PANEL_TYPES.POSTER.defaults.slingWeight).toBe(0);
    expect(PANEL_TYPES.MG9.defaults.flyBarWeight).not.toBe(PANEL_TYPES.POSTER.defaults.flyBarWeight);
  });
});

// The whole point of posterGroupId: the eight sections are one physical
// fixture, so touching any one of them acts on all eight - and only those.
describe("LED poster grouping survives the 2x4 split", () => {
  it("selecting one section expands to all eight of that poster", () => {
    const cells = makePosterPanels(2);
    const [first] = cells;
    const expanded = getSelectedIds(new Set([first.id]), null, cells);
    expect(expanded.size).toBe(POSTER_SECTIONS);
    const group = cells.filter((c) => c.posterGroupId === first.posterGroupId);
    group.forEach((c) => expect(expanded.has(c.id)).toBe(true));
  });

  it("does not drag the neighbouring poster in with it", () => {
    const cells = makePosterPanels(2);
    const first = cells[0];
    const other = cells.find((c) => c.posterGroupId !== first.posterGroupId)!;
    const expanded = getSelectedIds(new Set([first.id]), null, cells);
    expect(expanded.has(other.id)).toBe(false);
  });

  it("expands a bare selectedId too, not just a multi-selection", () => {
    const cells = makePosterPanels(1);
    expect(getSelectedIds(new Set(), cells[POSTER_SECTIONS - 1].id, cells).size).toBe(POSTER_SECTIONS);
  });
});

// Every poster is its own sub-screen, however it was created - each is a
// separate physical fixture with its own content feed.
describe("posters arrive as their own sub-screens", () => {
  it("gives each poster one sub-screen, and every section of it that id", () => {
    const { cells, subScreens } = makePosterUnits(3);
    expect(subScreens).toHaveLength(3);
    expect(cells).toHaveLength(3 * POSTER_SECTIONS);
    subScreens.forEach((screen) => {
      const members = cells.filter((c) => c.subScreenId === screen.id);
      expect(members).toHaveLength(POSTER_SECTIONS);
      // ...and those sections are exactly one poster group, not a mix.
      expect(new Set(members.map((c) => c.posterGroupId)).size).toBe(1);
    });
    expect(cells.every((c) => c.subScreenId)).toBe(true);
  });

  it("names them Poster 1..n and gives each its own colour", () => {
    const { subScreens } = makePosterUnits(3);
    expect(subScreens.map((s) => s.name)).toEqual(["Poster 1", "Poster 2", "Poster 3"]);
    expect(new Set(subScreens.map((s) => s.color)).size).toBe(3);
  });

  it("continues past existing sub-screens instead of reusing a name", () => {
    const first = makePosterUnits(2);
    const second = makePosterUnits(2, first.subScreens);
    expect(second.subScreens.map((s) => s.name)).toEqual(["Poster 3", "Poster 4"]);
    const ids = [...first.subScreens, ...second.subScreens].map((s) => s.id);
    expect(new Set(ids).size).toBe(4);
  });

  it("skips a name that is already taken by an unrelated sub-screen", () => {
    const existing = makePosterUnits(1).subScreens;
    const { subScreens } = makePosterUnits(1, [...existing, { ...existing[0], id: "x", name: "Poster 2" }]);
    expect(subScreens[0].name).toBe("Poster 3");
  });

  it("lays the posters out side by side from the given origin", () => {
    const { cells } = makePosterUnits(2, [], 1000, 250);
    const rects = cells.map(cellRect);
    expect(Math.min(...rects.map((r) => r.x))).toBe(1000);
    expect(Math.min(...rects.map((r) => r.y))).toBe(250);
    expect(Math.max(...rects.map((r) => r.x + r.w))).toBe(1000 + 2 * POSTER_WIDTH_MM);
  });

  it("builds nothing for a count of zero", () => {
    expect(makePosterUnits(0)).toEqual({ cells: [], subScreens: [] });
  });
});
