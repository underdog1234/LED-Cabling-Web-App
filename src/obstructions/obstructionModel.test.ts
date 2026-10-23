import { describe, it, expect } from "vitest";
import {
  MIN_OBSTRUCTION_MM,
  makeObstruction,
  metresLabel,
  migrateAbsoluteObstructions,
  normalizeObstructions,
  obstructionIntersectsSurface,
  obstructionMetresLabel,
  obstructionOffsetFor,
  obstructionPixelRect,
  obstructionRect,
} from "./obstructionModel";

// An obstruction is measured on site in metres but has to land on an exact
// pixel of the content, so the mm -> px conversion is the part that matters.
describe("obstruction geometry", () => {
  // A 4m x 2m wall at 1344 x 672 px - the MG9 8x4 grid used throughout.
  const wall = { x: 0, y: 0, w: 4000, h: 2000 };
  const W = 1344;
  const H = 672;

  it("0, 0 stands it centred on the wall, on the wall's bottom edge", () => {
    const obstruction = makeObstruction("PA stack", 0, 0, 1000, 500);
    expect(obstructionRect(obstruction, wall)).toEqual({ x: 1500, y: 1500, w: 1000, h: 500 });
  });

  it("offsets move it right and down from there", () => {
    const obstruction = makeObstruction("PA stack", 500, -250, 1000, 500);
    expect(obstructionRect(obstruction, wall)).toEqual({ x: 2000, y: 1250, w: 1000, h: 500 });
  });

  it("stays put when the wall grows, because the offset is measured from it", () => {
    const obstruction = makeObstruction("Centre pillar", 0, 0, 1000, 500);
    // A wall twice as wide, with its origin moved: still dead centre, still
    // standing on the bottom. An absolute coordinate would have been stranded.
    const wider = { x: -4000, y: 0, w: 8000, h: 2000 };
    const rect = obstructionRect(obstruction, wider);
    expect(rect.x + rect.w / 2).toBe(wider.x + wider.w / 2);
    expect(rect.y + rect.h).toBe(wider.y + wider.h);
  });

  it("converts metres to the surface's own pixels", () => {
    const obstruction = makeObstruction("PA stack", 0, -750, 1000, 500);
    // Centred: 1500mm from the left of a 4m wall -> 504px of 1344.
    // 750mm up from the bottom -> top edge at 750mm -> 252px of 672.
    expect(obstructionPixelRect(obstruction, wall, W, H)).toEqual({ x: 504, y: 252, w: 336, h: 168 });
  });

  it("a sub-screen renders it where it stands on the WHOLE wall", () => {
    // Sub-screen covering the right half of the 4m wall, at its own 672px.
    const subScreen = { x: 2000, y: 0, w: 2000, h: 2000 };
    const obstruction = makeObstruction("Pillar", 500, 0, 500, 500);
    // Offset 500 from the whole wall's centre (2000) -> absolute x 2250,
    // which is 250mm into the sub-screen = an eighth of its 672px width.
    const rect = obstructionPixelRect(obstruction, subScreen, 672, 672, wall);
    expect(rect.x).toBe(84);
    expect(rect.w).toBe(168);
  });

  it("round-trips an absolute position back to the offset that produces it", () => {
    const size = { wMm: 800, hMm: 1200 };
    const offset = obstructionOffsetFor({ x: 2600, y: 400 }, size, wall);
    const rect = obstructionRect({ ...makeObstruction("X", 0, 0, size.wMm, size.hMm), ...offset }, wall);
    expect(rect.x).toBe(2600);
    expect(rect.y).toBe(400);
  });

  it("never collapses a small obstruction to nothing", () => {
    const tiny = makeObstruction("Mic stand", 0, 0, MIN_OBSTRUCTION_MM, MIN_OBSTRUCTION_MM);
    const rect = obstructionPixelRect(tiny, { x: 0, y: 0, w: 100000, h: 100000 }, 100, 100);
    expect(rect.w).toBeGreaterThanOrEqual(1);
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

// Obstructions were stored as absolute workspace mm up to formatVersion 7.
// The same numbers mean something else now, so an old file has to be converted
// on the way in or every obstruction in it jumps.
describe("migrating pre-v8 obstructions", () => {
  const wall = { x: 0, y: 0, w: 4000, h: 2000 };

  it("keeps an old obstruction exactly where it was", () => {
    const stored = makeObstruction("PA stack", 1500, 750, 1000, 500); // absolute, v7
    const [migrated] = migrateAbsoluteObstructions([stored], wall);
    expect(obstructionRect(migrated, wall)).toEqual({ x: 1500, y: 750, w: 1000, h: 500 });
  });

  it("an old obstruction that was already centred on the floor becomes 0, 0", () => {
    const stored = makeObstruction("Centre stack", 1500, 1500, 1000, 500);
    const [migrated] = migrateAbsoluteObstructions([stored], wall);
    expect(migrated.xMm).toBe(0);
    expect(migrated.yMm).toBe(0);
  });

  it("leaves name, size and id alone", () => {
    const stored = makeObstruction("Truss leg", 100, 200, 800, 1200);
    const [migrated] = migrateAbsoluteObstructions([stored], wall);
    expect(migrated).toMatchObject({ id: stored.id, name: "Truss leg", wMm: 800, hMm: 1200 });
  });
});
