import { describe, it, expect } from "vitest";
import {
  MIN_CUSTOM_RESOLUTION_PX,
  SUB_SCREEN_ASPECT_PRESETS,
  normalizeCustomResolution,
  normalizeSubScreens,
} from "./App";

// A sub-screen's content region can be typed in by hand instead of being
// derived from its panels - a 16:9 feed landing on a 2:1 wall, say.
describe("custom sub-screen regions", () => {
  it("accepts a usable region", () => {
    expect(normalizeCustomResolution({ w: 1920, h: 1080 })).toEqual({ w: 1920, h: 1080 });
  });

  it("rounds to whole pixels - there is no half a pixel on a wall", () => {
    expect(normalizeCustomResolution({ w: 1920.4, h: 1080.6 })).toEqual({ w: 1920, h: 1081 });
  });

  it("rejects anything that is not a region, falling back to the panels", () => {
    expect(normalizeCustomResolution(null)).toBeNull();
    expect(normalizeCustomResolution(undefined)).toBeNull();
    expect(normalizeCustomResolution({ w: 0, h: 0 })).toBeNull();
    expect(normalizeCustomResolution({ w: -1920, h: 1080 })).toBeNull();
    expect(normalizeCustomResolution({ w: "wide", h: 1080 })).toBeNull();
    expect(normalizeCustomResolution({ w: MIN_CUSTOM_RESOLUTION_PX - 1, h: 1080 })).toBeNull();
  });

  it("offers the aspect ratios content actually arrives in", () => {
    expect(SUB_SCREEN_ASPECT_PRESETS.map((p) => p.label)).toEqual(["16:9", "16:10", "4:3", "1:1"]);
  });

  it("a preset keeps the width and solves for height", () => {
    const widthPx = 672;
    const expected = SUB_SCREEN_ASPECT_PRESETS.map((p) => Math.round((widthPx * p.h) / p.w));
    expect(expected).toEqual([378, 420, 504, 672]);
  });
});

describe("custom regions survive a save/load round trip", () => {
  const base = { id: "a", name: "Right Tower", canvasX: 100, canvasY: 50, createdAt: 2, color: "#fb923c" };

  it("restores a stored region", () => {
    const [restored] = normalizeSubScreens([{ ...base, customResolution: { w: 1920, h: 1080 } }]);
    expect(restored.customResolution).toEqual({ w: 1920, h: 1080 });
  });

  it("older projects have no region at all, which means 'use the panels'", () => {
    const [restored] = normalizeSubScreens([base]);
    expect(restored.customResolution).toBeNull();
  });

  it("a corrupt region falls back to the panels rather than breaking the screen", () => {
    const [restored] = normalizeSubScreens([{ ...base, customResolution: { w: "oops", h: 1080 } }]);
    expect(restored.customResolution).toBeNull();
    expect(restored.name).toBe("Right Tower");
  });
});
