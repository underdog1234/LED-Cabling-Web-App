import { describe, it, expect } from "vitest";
import type { StockRow } from "./App";
import {
  applyStockEdits,
  calculatedTotalOf,
  normalizeStockEdits,
  removedStockRows,
  stockEditCount,
  withStockQty,
  withStockRemoved,
  withStockRowReset,
  type StockEdits,
} from "./stock/stockEdits";

// A manual change to the pull list has to behave like a note pinned to one
// row: it changes the number that leaves the warehouse, it never quietly
// rewrites what the tool worked out, and taking it off puts the row back
// exactly as it was.

const row = (code: string, required: number, spareRounded = 0, stock = 100): StockRow => ({
  code,
  name: `Item ${code}`,
  required,
  stock,
  net: stock - (required + spareRounded),
  method: "test",
  spare: spareRounded,
  spareRounded,
  rounded: required + spareRounded,
});

const rows = [row("12254", 10, 2), row("12263", 6), row("12398-LU", 1)];

describe("applyStockEdits", () => {
  it("changes nothing at all with no edits", () => {
    expect(applyStockEdits(rows, {})).toEqual(rows);
  });

  it("replaces the order quantity and what is left against stock, keeping the calculated figure", () => {
    const [first] = applyStockEdits(rows, { "12254": { qty: 4 } });
    expect(first).toMatchObject({ code: "12254", rounded: 4, net: 96, calculated: 12, edited: true });
    // The derivation is untouched - the row still says where 12 came from.
    expect(first).toMatchObject({ required: 10, spare: 2, spareRounded: 2 });
  });

  it("takes a removed row off the list entirely, and can list what went", () => {
    const edits: StockEdits = { "12263": { removed: true } };
    expect(applyStockEdits(rows, edits).map((r) => r.code)).toEqual(["12254", "12398-LU"]);
    expect(removedStockRows(rows, edits).map((r) => r.code)).toEqual(["12263"]);
  });

  it("does not mark a row edited when the typed quantity is the calculated one", () => {
    const [first] = applyStockEdits(rows, { "12254": { qty: 12 } });
    expect(first.edited).toBeUndefined();
    expect(first.rounded).toBe(12);
  });

  it("keeps an edit of zero - it is a real answer, not an empty box", () => {
    const [first] = applyStockEdits(rows, { "12254": { qty: 0 } });
    expect(first).toMatchObject({ rounded: 0, net: 100, edited: true, calculated: 12 });
  });

  it("leaves a note for a row that is not on the list alone", () => {
    expect(applyStockEdits(rows, { "99999": { qty: 5 } })).toEqual(rows);
  });
});

describe("stockEditCount", () => {
  it("counts each edited or removed row once, and ignores notes that change nothing", () => {
    expect(stockEditCount(rows, {})).toBe(0);
    expect(stockEditCount(rows, { "12254": { qty: 4 }, "12263": { removed: true } })).toBe(2);
    expect(stockEditCount(rows, { "12254": { qty: 12 } })).toBe(0);
    expect(stockEditCount(rows, { "99999": { qty: 4 } })).toBe(0);
  });
});

describe("editing one row", () => {
  it("drops the note when the quantity is typed back to the calculated one", () => {
    const edited = withStockQty({}, "12254", 4, calculatedTotalOf(rows[0]));
    expect(edited).toEqual({ "12254": { qty: 4 } });
    expect(withStockQty(edited, "12254", 12, 12)).toEqual({});
    // An emptied box is the same as no change.
    expect(withStockQty(edited, "12254", null, 12)).toEqual({});
  });

  it("rounds and floors what is typed", () => {
    expect(withStockQty({}, "12254", 3.6, 12)).toEqual({ "12254": { qty: 4 } });
    expect(withStockQty({}, "12254", -5, 12)).toEqual({ "12254": { qty: 0 } });
  });

  it("keeps a row's quantity when it is put back on the list", () => {
    const removed = withStockRemoved({ "12254": { qty: 4 } }, "12254", true);
    expect(removed).toEqual({ "12254": { qty: 4, removed: true } });
    expect(withStockRemoved(removed, "12254", false)).toEqual({ "12254": { qty: 4 } });
  });

  it("clears a row completely on reset, and leaves the others", () => {
    const edits: StockEdits = { "12254": { qty: 4 }, "12263": { removed: true } };
    expect(withStockRowReset(edits, "12254")).toEqual({ "12263": { removed: true } });
    expect(withStockRowReset(edits, "99999")).toBe(edits);
  });

  it("keeps a quantity edit while the row is off the list, so putting it back restores it", () => {
    let edits = withStockQty({}, "12254", 4, 12);
    edits = withStockRemoved(edits, "12254", true);
    expect(applyStockEdits(rows, edits)).toHaveLength(2);
    edits = withStockRemoved(edits, "12254", false);
    expect(applyStockEdits(rows, edits)[0]).toMatchObject({ rounded: 4, edited: true });
  });
});

describe("normalizeStockEdits", () => {
  it("accepts what a project file should hold and refuses the rest", () => {
    expect(normalizeStockEdits({ "12254": { qty: 4 }, "12263": { removed: true } }))
      .toEqual({ "12254": { qty: 4 }, "12263": { removed: true } });
    expect(normalizeStockEdits({ a: { qty: "4" }, b: { removed: "yes" }, c: null, d: { qty: NaN } })).toEqual({});
    expect(normalizeStockEdits(undefined)).toEqual({});
    expect(normalizeStockEdits("nonsense")).toEqual({});
  });
});
