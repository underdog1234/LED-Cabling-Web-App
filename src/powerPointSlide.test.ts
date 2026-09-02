import { describe, it, expect } from "vitest";
import { powerPointSlideSize, POWERPOINT_MAX_SLIDE_CM } from "./App";

// PowerPoint sizes slides physically, so what has to match the wall is the
// aspect ratio; 100cm on the longest edge keeps the numbers memorable and well
// inside PowerPoint's own 142.24cm ceiling.
describe("powerPointSlideSize", () => {
  it("matches the worked example for a 2016 x 1176 wall", () => {
    const s = powerPointSlideSize(2016, 1176)!;
    expect(s.widthCm).toBe(100);
    expect(s.heightCm.toFixed(3)).toBe("58.333");
    expect(`${s.ratioW}:${s.ratioH}`).toBe("12:7");
    expect(s.exceedsLimit).toBe(false);
  });

  it("puts the 100cm on the tall edge for a portrait wall", () => {
    const s = powerPointSlideSize(344, 1032)!;
    expect(s.heightCm).toBe(100);
    expect(s.widthCm.toFixed(3)).toBe("33.333");
    expect(`${s.ratioW}:${s.ratioH}`).toBe("1:3");
  });

  it("keeps a square wall square", () => {
    const s = powerPointSlideSize(1008, 1008)!;
    expect(s.widthCm).toBe(100);
    expect(s.heightCm).toBe(100);
    expect(`${s.ratioW}:${s.ratioH}`).toBe("1:1");
  });

  it("never exceeds PowerPoint's limit with the 100cm method", () => {
    for (const [w, h] of [[1920, 1080], [3840, 1080], [1008, 504], [344, 1032], [7680, 2160]]) {
      const s = powerPointSlideSize(w, h)!;
      expect(Math.max(s.widthCm, s.heightCm)).toBeLessThanOrEqual(POWERPOINT_MAX_SLIDE_CM);
      expect(s.exceedsLimit).toBe(false);
      // Slide proportions must match the wall's exactly.
      expect(s.widthCm / s.heightCm).toBeCloseTo(w / h, 10);
    }
  });

  it("returns null for an empty wall rather than dividing by zero", () => {
    expect(powerPointSlideSize(0, 0)).toBeNull();
    expect(powerPointSlideSize(1920, 0)).toBeNull();
  });
});
