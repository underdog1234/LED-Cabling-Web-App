import { describe, expect, it } from "vitest";
import {
  RESOLUTION_PRESETS,
  alignRects,
  applyPlaylistStep,
  arrangeRects,
  defaultConfig,
  distributeRects,
  distributeWithGap,
  exportFileName,
  layoutWarnings,
  makeScreen,
  normalizeConfig,
  playlistStepAt,
  reorderLayers,
  resizeRect,
  resizeWithAspect,
  snapMove,
  snapTargetsFor,
  uniqueNames,
  type Rect,
} from "./model";

const isInt = (n: number) => Number.isInteger(n);

describe("resolution presets", () => {
  it("has the common sizes and a portrait version of each", () => {
    const sizes = RESOLUTION_PRESETS.map((p) => `${p.w}x${p.h}`);
    ["1280x720", "1920x1080", "1920x1200", "2560x1440", "3840x2160", "4096x2160"].forEach((s) => {
      expect(sizes).toContain(s);
      const [w, h] = s.split("x");
      expect(sizes).toContain(`${h}x${w}`);
    });
  });

  it("keeps a locked aspect ratio to the nearest whole pixel", () => {
    expect(resizeWithAspect({ w: 1920, h: 1080 }, "w", 1280, true)).toEqual({ w: 1280, h: 720 });
    expect(resizeWithAspect({ w: 1920, h: 1080 }, "h", 1000, true)).toEqual({ w: 1778, h: 1000 });
    expect(resizeWithAspect({ w: 1920, h: 1080 }, "w", 1000, false)).toEqual({ w: 1000, h: 1080 });
  });
});

describe("alignment", () => {
  const rects: Rect[] = [
    { x: 13, y: 7, w: 101, h: 51 },
    { x: 400, y: 90, w: 333, h: 17 },
  ];
  const canvas = { x: 0, y: 0, w: 1921, h: 1081 };

  it("aligns to the canvas in whole pixels", () => {
    expect(alignRects(rects, "left", canvas).map((p) => p.x)).toEqual([0, 0]);
    expect(alignRects(rects, "right", canvas).map((p) => p.x)).toEqual([1820, 1588]);
    const centre = alignRects(rects, "centre", canvas);
    centre.forEach((p) => expect(isInt(p.x)).toBe(true));
    expect(centre.map((p) => p.x)).toEqual([910, 794]);
    expect(alignRects(rects, "bottom", canvas).map((p) => p.y)).toEqual([1030, 1064]);
    expect(alignRects(rects, "middle", canvas).map((p) => p.y)).toEqual([515, 532]);
  });

  it("aligns to a reference screen", () => {
    const ref = { x: 500, y: 300, w: 200, h: 100 };
    expect(alignRects(rects, "top", ref).map((p) => p.y)).toEqual([300, 300]);
    expect(alignRects(rects, "right", ref).map((p) => p.x)).toEqual([599, 367]);
  });
});

describe("distribution", () => {
  it("spaces screens evenly with whole-pixel gaps that differ by at most one", () => {
    const rects: Rect[] = [
      { x: 0, y: 0, w: 100, h: 10 },
      { x: 120, y: 0, w: 50, h: 10 },
      { x: 130, y: 0, w: 70, h: 10 },
      { x: 1000, y: 0, w: 100, h: 10 },
    ];
    const out = distributeRects(rects, "horizontal");
    // Outermost stay put.
    expect(out[0].x).toBe(0);
    expect(out[3].x).toBe(1000);
    const placed = rects.map((r, i) => ({ ...r, x: out[i].x })).sort((a, b) => a.x - b.x);
    const gaps = placed.slice(1).map((r, i) => r.x - (placed[i].x + placed[i].w));
    gaps.forEach((g) => expect(isInt(g)).toBe(true));
    expect(Math.max(...gaps) - Math.min(...gaps)).toBeLessThanOrEqual(1);
    expect(gaps.reduce((a, b) => a + b, 0)).toBe(1100 - 320);
  });

  it("packs with a fixed gap", () => {
    const out = distributeWithGap(
      [
        { x: 50, y: 0, w: 10, h: 10 },
        { x: 0, y: 5, w: 20, h: 10 },
      ],
      "horizontal",
      7,
    );
    expect(out).toEqual([{ x: 27, y: 0 }, { x: 0, y: 5 }]);
  });
});

describe("arrange", () => {
  const rects: Rect[] = [
    { x: 0, y: 0, w: 100, h: 50 },
    { x: 200, y: 0, w: 80, h: 70 },
    { x: 0, y: 200, w: 120, h: 30 },
  ];
  it("lays out a grid with each column as wide as its widest screen", () => {
    const out = arrangeRects(rects, "grid", { columns: 2, gapX: 10, gapY: 5, origin: { x: 3, y: 4 } });
    expect(out).toEqual([
      { x: 3, y: 4 },
      { x: 3 + 120 + 10, y: 4 },
      { x: 3, y: 4 + 70 + 5 },
    ]);
  });
  it("lays out a row and a column", () => {
    expect(arrangeRects(rects, "row", { columns: 1, gapX: 2, gapY: 0, origin: { x: 0, y: 0 } }).map((p) => p.x)).toEqual([0, 102, 184]);
    expect(arrangeRects(rects, "column", { columns: 1, gapX: 0, gapY: 1, origin: { x: 0, y: 0 } }).map((p) => p.y)).toEqual([0, 51, 122]);
  });
});

describe("snapping and resizing", () => {
  const snap = defaultConfig().snap;
  it("snaps an edge to the canvas and reports the guide", () => {
    const t = snapTargetsFor({ w: 1920, h: 1080 }, [], snap);
    const r = snapMove({ x: 4.6, y: 300.2, w: 100, h: 100 }, t, 8, null);
    expect(r.x).toBe(0);
    expect(r.guides).toContainEqual(expect.objectContaining({ axis: "x", at: 0 }));
    expect(isInt(r.y)).toBe(true);
  });
  it("snaps centre to centre", () => {
    const t = snapTargetsFor({ w: 1920, h: 1080 }, [], snap);
    expect(snapMove({ x: 857, y: 0, w: 200, h: 10 }, t, 8, null).x).toBe(860);
  });
  it("snaps to another screen's edge", () => {
    const t = snapTargetsFor({ w: 1920, h: 1080 }, [{ x: 500, y: 500, w: 300, h: 300 }], { ...snap, canvasEdges: false, canvasCentre: false });
    expect(snapMove({ x: 803, y: 503, w: 100, h: 100 }, t, 8, null)).toMatchObject({ x: 800, y: 500 });
  });
  it("resizes from any handle in whole pixels, keeping the opposite edge", () => {
    const start = { x: 100, y: 100, w: 400, h: 200 };
    expect(resizeRect(start, "w", -50.4, 0, false)).toEqual({ x: 50, y: 100, w: 450, h: 200 });
    expect(resizeRect(start, "se", 100, 0, true)).toEqual({ x: 100, y: 100, w: 500, h: 250 });
    expect(resizeRect(start, "n", 0, 500, false)).toEqual({ x: 100, y: 299, w: 400, h: 1 });
  });
});

describe("layers", () => {
  const items = ["a", "b", "c", "d"].map((id) => ({ id }));
  const order = (list: { id: string }[]) => list.map((x) => x.id).join("");
  it("moves selected layers", () => {
    expect(order(reorderLayers(items, new Set(["b"]), "front"))).toBe("acdb");
    expect(order(reorderLayers(items, new Set(["c"]), "back"))).toBe("cabd");
    expect(order(reorderLayers(items, new Set(["b"]), "forward"))).toBe("acbd");
    expect(order(reorderLayers(items, new Set(["c"]), "backward"))).toBe("acbd");
  });
});

describe("file names", () => {
  it("names a file after its screen, pattern and resolution", () => {
    expect(exportFileName("Stage Left / IMAG", "SMPTE Colour Bars", 1080, 1920, "png")).toBe("Stage-Left-IMAG_SMPTE-Colour-Bars_1080x1920.png");
  });
  it("keeps names unique inside one ZIP", () => {
    expect(uniqueNames(["a.png", "a.png", "b.png", "a.png"])).toEqual(["a.png", "a-2.png", "b.png", "a-3.png"]);
  });
});

describe("saved configurations", () => {
  it("round-trips every screen's position, resolution, pattern, settings and layer order", () => {
    const config = defaultConfig();
    config.screens[2] = { ...config.screens[2], x: -7, y: 13, w: 333, h: 777, locked: true, settings: { grid: { spacing: 32 } }, pattern: "grid" };
    config.screens.reverse();
    const back = normalizeConfig(JSON.parse(JSON.stringify(config)));
    expect(back.screens.map((s) => [s.id, s.x, s.y, s.w, s.h, s.pattern, s.locked])).toEqual(config.screens.map((s) => [s.id, s.x, s.y, s.w, s.h, s.pattern, s.locked]));
    expect(back.screens.find((s) => s.w === 333)?.settings.grid.spacing).toBe(32);
  });
  it("repairs a damaged file instead of failing", () => {
    const back = normalizeConfig({ canvas: { w: "abc", h: -5 }, screens: [{ w: 0.4, x: 1.6 }, null] });
    expect(back.canvas.w).toBe(3840);
    expect(back.canvas.h).toBe(1);
    expect(back.screens[0]).toMatchObject({ w: 1, x: 2 });
  });
});

describe("warnings", () => {
  it("flags off-canvas and overlapping screens", () => {
    const c = defaultConfig();
    c.canvas = { ...c.canvas, w: 1000, h: 1000 };
    c.screens = [makeScreen(0, { name: "A", x: -1, y: 0, w: 500, h: 500 }), makeScreen(1, { name: "B", x: 400, y: 400, w: 700, h: 100 })];
    const w = layoutWarnings(c).join("\n");
    expect(w).toContain("A: negative");
    expect(w).toContain("B: extends beyond");
    expect(w).toContain("A and B overlap");
  });
});

describe("playlists", () => {
  it("steps through by duration and loops", () => {
    const c = defaultConfig();
    const id = c.screens[0].id;
    c.playlist = [
      { id: "1", name: "one", duration: 5, assignments: { [id]: { pattern: "solid", settings: { color: "#ff0000" } } } },
      { id: "2", name: "two", duration: 10, assignments: { [id]: { pattern: "grid", settings: {} } } },
    ];
    expect(playlistStepAt(c.playlist, 4.9, true)?.index).toBe(0);
    expect(playlistStepAt(c.playlist, 5, true)?.index).toBe(1);
    expect(playlistStepAt(c.playlist, 15.5, true)?.index).toBe(0);
    expect(playlistStepAt(c.playlist, 99, false)?.index).toBe(1);
    const applied = applyPlaylistStep(c, c.playlist[0]);
    expect(applied.screens[0].pattern).toBe("solid");
    expect(applied.screens[0].settings.solid.color).toBe("#ff0000");
    expect(applied.screens[1]).toBe(c.screens[1]);
  });
});
