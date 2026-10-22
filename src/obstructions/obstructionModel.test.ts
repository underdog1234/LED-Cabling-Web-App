import { describe, it, expect } from "vitest";
import {
  MIN_OBSTRUCTION_MM,
  makeObstruction,
  metresLabel,
  normalizeObstructions,
  obstructionIntersectsSurface,
  obstructionMetresLabel,
  obstructionPixelRect,
} from "./obstructionModel";

// An obstruction is measured on site in metres but has to land on an exact
// pixel of the content, so the mm -> px conversion is the part that matters.
describe("obstruction geometry", () => {
  // A 4m x 2m wall at 1344 x 672 px - the MG9 8x4 grid used throughout.
  const wall = { x: 0, y: 0, w: 4000, h: 2000 };
  const W = 1344;
  const H = 672;

  it("converts metres to the surface's own pixels", () => {
    const obstruction = makeObstruction("PA stack", 1500, 750, 1000, 500);
    expect(obstructionPixelRect(obstruction, wall, W, H)).toEqual({ x: 504, y: 252, w: 336, h: 168 });
  });

  it("measures from the surface's own origin, not the workspace's", () => {
    // The same obstruction against a sub-screen whose bbox starts at 2000mm:
    // it is 500mm into that screen, so a quarter of the way across it.
    const subScreen = { x: 2000, y: 0, w: 2000, h: 2000 };
    const obstruction = makeObstruction("Pillar", 2500, 0, 500, 500);
    const rect = obstructionPixelRect(obstruction, subScreen, 672, 672);
    expect(rect.x).toBe(168);
    expect(rect.w).toBe(168);
  });

  it("never collapses a small obstruction to nothing", () => {
    const tiny = makeObstruction("Mic stand", 0, 0, MIN_OBSTRUCTION_MM, MIN_OBSTRUCTION_MM);
    const rect = obstructionPixelRect(tiny, { x: 0, y: 0, w: 100000, h: 100000 }, 100, 100);
    expect(rect.w).toBeGreaterThanOrEqual(1);
    expect(rect.h).toBeGreaterThanOrEqual(1);
  });

  it("survives a zero-sized surface rather than dividing by zero", () => {
    const rect = obstructionPixelRect(makeObstruction("X", 0, 0), { x: 0, y: 0, w: 0, h: 0 }, 100, 100);
    expect(Number.isFinite(rect.x)).toBe(true);
    expect(Number.isFinite(rect.y)).toBe(true);
  });

  it("knows which obstructions actually fall on a surface", () => {
    expect(obstructionIntersectsSurface({ x: 10, y: 10, w: 50, h: 50 }, W, H)).toBe(true);
    // Entirely past the right-hand edge, e.g. an obstruction in front of a
    // different tower - must not be drawn on this one.
    expect(obstructionIntersectsSurface({ x: W + 5, y: 10, w: 50, h: 50 }, W, H)).toBe(false);
    expect(obstructionIntersectsSurface({ x: -100, y: 10, w: 50, h: 50 }, W, H)).toBe(false);
    // Straddling the edge still counts - part of it is over the LED.
    expect(obstructionIntersectsSurface({ x: W - 10, y: 10, w: 50, h: 50 }, W, H)).toBe(true);
  });
});

describe("obstruction labels", () => {
  it("reads in metres to 2dp", () => {
    expect(metresLabel(1500)).toBe("1.50");
    expect(metresLabel(800)).toBe("0.80");
    expect(obstructionMetresLabel(makeObstruction("Stack", 0, 0, 2400, 1200))).toBe("2.40 x 1.20 m");
  });
});

describe("obstructions survive a save/load round trip", () => {
  it("restores a stored list unchanged", () => {
    const original = [makeObstruction("PA left", 100, 200, 1000, 2000), makeObstruction("PA right", 5000, 200, 1000, 2000)];
    const restored = normalizeObstructions(JSON.parse(JSON.stringify(original)));
    expect(restored).toEqual(original);
  });

  it("treats a project with no obstructions as an empty list", () => {
    expect(normalizeObstructions(undefined)).toEqual([]);
    expect(normalizeObstructions(null)).toEqual([]);
    expect(normalizeObstructions({})).toEqual([]);
  });

  it("repairs a damaged entry instead of dropping it", () => {
    const restored = normalizeObstructions([{ name: "  ", xMm: "oops", wMm: -5 }]);
    expect(restored).toHaveLength(1);
    expect(restored[0].name).toBe("Obstruction 1");
    expect(restored[0].xMm).toBe(0);
    expect(restored[0].wMm).toBeGreaterThanOrEqual(MIN_OBSTRUCTION_MM);
    expect(restored[0].id).toBeTruthy();
  });

  it("accepts numeric strings, which is what a hand-edited file holds", () => {
    const restored = normalizeObstructions([{ id: "a", name: "Pillar", xMm: "1500", yMm: "0", wMm: "500", hMm: "2000" }]);
    expect(restored[0]).toMatchObject({ xMm: 1500, wMm: 500, hMm: 2000 });
  });
});
