import { describe, it, expect } from "vitest";
import { makeGridPanels, spareBucketOfCell, spareForBucket, PANEL_TYPES, type Cell } from "./App";

// Spare-panel bucketing/rounding rules (see App.tsx spareBucketOfCell /
// spareForBucket): MG9's four buckets get 7% spare; required + spare is then
// taken up to a whole number of equipment boxes, because whole boxes are what
// physically leaves the warehouse. The "spare rounded to full boxes" figure is
// whatever spare falls out of that (total - required), NOT the raw spare
// rounded on its own - the required panels already part-fill a box, so the
// spare only tops up what's left of it. Standard and Corner box in 10s; MT
// boxes in 6s but intentionally gets a 0% spare ratio (a deliberate catalog
// choice, not an oversight); shaped panels (Triangle/Curved) are one-way
// pieces bought individually and don't box-round at all.

const cellOfVariant = (variant: Cell["panelVariant"]): Cell => ({
  ...makeGridPanels(1, 1, "MG9")[0],
  panelVariant: variant,
});

describe("spareBucketOfCell", () => {
  it("buckets a standard MG9 panel", () => {
    expect(spareBucketOfCell(cellOfVariant("STANDARD"))).toBe("MG9_STANDARD");
  });

  it("buckets each MG9 variant separately", () => {
    expect(spareBucketOfCell(cellOfVariant("TRIANGLE"))).toBe("MG9_TRIANGLE");
    expect(spareBucketOfCell(cellOfVariant("CURVED"))).toBe("MG9_CURVED");
    expect(spareBucketOfCell(cellOfVariant("CORNER"))).toBe("MG9_CORNER");
  });

  it("buckets any MT panel as MT regardless of variant field", () => {
    const cell = { ...makeGridPanels(1, 1, "MT")[0], panelVariant: "STANDARD" as const };
    expect(spareBucketOfCell(cell)).toBe("MT");
  });
});

describe("spareForBucket", () => {
  it("computes 7% spare, ceiled, for every MG9 bucket", () => {
    // 50 * 0.07 = 3.5 -> ceil 4
    expect(spareForBucket(50, "MG9_STANDARD").spare).toBe(4);
    expect(spareForBucket(50, "MG9_CORNER").spare).toBe(4);
    expect(spareForBucket(50, "MG9_TRIANGLE").spare).toBe(4);
    expect(spareForBucket(50, "MG9_CURVED").spare).toBe(4);
  });

  it("MT always gets 0 spare", () => {
    expect(PANEL_TYPES.MT.defaults.spareRatio).toBe(0);
    expect(spareForBucket(50, "MT").spare).toBe(0);
  });

  it("takes required + spare up to a full box, and reports the spare that falls out", () => {
    expect(PANEL_TYPES.MG9.defaults.panelsPerBox).toBe(10);
    // The worked example: 45 required needs 4 spare; 45 + 4 = 49, which rounds
    // up to 50 - a clean 5 boxes - so 5 spares go out, not 4 and not a whole
    // extra box of 10.
    expect(spareForBucket(45, "MG9_STANDARD")).toEqual({ spare: 4, spareRounded: 5, total: 50 });
    expect(spareForBucket(45, "MG9_CORNER")).toEqual({ spare: 4, spareRounded: 5, total: 50 });
  });

  it("only tops up the part-filled box, never adds a whole spare box on top", () => {
    // 50 + 4 = 54 -> 60, so 10 spare (not 50 + a full box of 10 = 60 by way of
    // rounding 4 up to 10 first, which happens to agree here) ...
    expect(spareForBucket(50, "MG9_STANDARD")).toEqual({ spare: 4, spareRounded: 10, total: 60 });
    // ... whereas here the two approaches genuinely differ: 18 + 2 = 20 is
    // already an exact box multiple, so no extra rounding at all.
    expect(spareForBucket(18, "MG9_STANDARD")).toEqual({ spare: 2, spareRounded: 2, total: 20 });
  });

  it("rounds a larger wall the same way", () => {
    // 120 * 0.07 = 8.4 -> 9 spare; 129 -> 130, so 10 spare.
    expect(spareForBucket(120, "MG9_STANDARD")).toEqual({ spare: 9, spareRounded: 10, total: 130 });
    // 150 * 0.07 = 10.5 -> 11 spare; 161 -> 170, so 20 spare.
    expect(spareForBucket(150, "MG9_STANDARD")).toEqual({ spare: 11, spareRounded: 20, total: 170 });
  });

  it("never reports fewer spares than the raw ratio asked for", () => {
    for (let used = 0; used <= 200; used += 1) {
      const { spare, spareRounded, total } = spareForBucket(used, "MG9_STANDARD");
      expect(spareRounded).toBeGreaterThanOrEqual(spare);
      expect(total).toBe(used + spareRounded);
      expect(total % PANEL_TYPES.MG9.defaults.panelsPerBox).toBe(0);
    }
  });

  it("MT boxes in 6s off its own zero spare", () => {
    expect(PANEL_TYPES.MT.defaults.panelsPerBox).toBe(6);
    // 50 + 0 = 50 -> next box of 6 is 54, so 4 spare come along with the box.
    expect(spareForBucket(50, "MT")).toEqual({ spare: 0, spareRounded: 4, total: 54 });
    // 48 is already an exact box multiple -> nothing extra.
    expect(spareForBucket(48, "MT")).toEqual({ spare: 0, spareRounded: 0, total: 48 });
  });

  it("does not box-round shaped panels - the raw spare is used as-is", () => {
    expect(spareForBucket(50, "MG9_TRIANGLE")).toEqual({ spare: 4, spareRounded: 4, total: 54 });
    expect(spareForBucket(50, "MG9_CURVED")).toEqual({ spare: 4, spareRounded: 4, total: 54 });
  });

  it("returns all zeros when nothing is used", () => {
    (["MG9_STANDARD", "MG9_TRIANGLE", "MG9_CURVED", "MG9_CORNER", "MT"] as const).forEach((bucket) => {
      expect(spareForBucket(0, bucket)).toEqual({ spare: 0, spareRounded: 0, total: 0 });
    });
  });
});
