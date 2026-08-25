import { describe, it, expect } from "vitest";
import { computePixelMappingStatus } from "./pixelMapping";

describe("computePixelMappingStatus", () => {
  it("reports 1:1 when the rendered CSS box, scaled by dpr, exactly matches the canvas backing size", () => {
    expect(computePixelMappingStatus(1680, 840, 1680, 840, 1)).toBe("1:1");
    expect(computePixelMappingStatus(3360, 1680, 1680, 840, 2)).toBe("1:1");
  });

  it("reports scaled when the rendered CSS box doesn't match the backing size", () => {
    expect(computePixelMappingStatus(1680, 840, 800, 400, 1)).toBe("scaled");
  });

  it("reports scaled when the pattern is larger than the viewport and clipped smaller, not just resized", () => {
    // A canvas that overflows its container still renders at its true CSS
    // size (getBoundingClientRect reports the full box even if part of it is
    // clipped) - so this case is really just "doesn't match", same as above.
    expect(computePixelMappingStatus(4032, 1344, 1920, 640, 1)).toBe("scaled");
  });

  it("tolerates a fraction-of-a-physical-pixel rounding difference", () => {
    expect(computePixelMappingStatus(1680, 840, 1679.7, 840, 1)).toBe("1:1");
  });

  it("does not tolerate a full pixel of mismatch", () => {
    expect(computePixelMappingStatus(1680, 840, 1678, 840, 1)).toBe("scaled");
  });
});
