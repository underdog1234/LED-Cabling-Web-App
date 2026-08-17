import { describe, it, expect } from "vitest";
import { makeGridPanels, spareBucketOfCell, spareForBucket, PANEL_TYPES, type Cell } from "./App";

// Spare-panel bucketing/rounding rules (see App.tsx spareBucketOfCell /
// spareForBucket): MG9's four buckets get 7% spare, then rounded up to a
// full box - Standard and Corner box in 10s, shaped panels (Triangle/Curved)
// are bought individually so they don't box-round at all. MT intentionally
// gets 0% spare (a deliberate catalog choice, not an oversight), but still
// boxes in 6s when its (always-zero) spare is added on top of what's used.

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

  it("rounds MG9 Standard and Corner up to a full box of 10", () => {
    expect(PANEL_TYPES.MG9.defaults.panelsPerBox).toBe(10);
    // 50 used + 4 spare = 54 -> next box of 10 is 60.
    expect(spareForBucket(50, "MG9_STANDARD").rounded).toBe(60);
    expect(spareForBucket(50, "MG9_CORNER").rounded).toBe(60);
  });

  it("rounds MT up to a full box of 6, with no spare added on top", () => {
    expect(PANEL_TYPES.MT.defaults.panelsPerBox).toBe(6);
    // 50 used + 0 spare = 50 -> next box of 6 is 54.
    expect(spareForBucket(50, "MT").rounded).toBe(54);
    // 48 used is already an exact box multiple -> stays 48.
    expect(spareForBucket(48, "MT").rounded).toBe(48);
  });

  it("does not box-round shaped panels - just used + spare as-is", () => {
    // 50 used + 4 spare = 54, with no box multiple applied.
    expect(spareForBucket(50, "MG9_TRIANGLE").rounded).toBe(54);
    expect(spareForBucket(50, "MG9_CURVED").rounded).toBe(54);
  });

  it("returns zero spare/rounded when nothing is used", () => {
    (["MG9_STANDARD", "MG9_TRIANGLE", "MG9_CURVED", "MG9_CORNER", "MT"] as const).forEach((bucket) => {
      expect(spareForBucket(0, bucket)).toEqual({ spare: 0, rounded: 0 });
    });
  });
});
