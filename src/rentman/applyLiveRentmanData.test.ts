import { describe, it, expect } from "vitest";
import type { StockRow } from "../App";
import { applyLiveRentmanData, baseCodeOf, isMappableStockRow } from "./applyLiveRentmanData";

const row = (overrides: Partial<StockRow> = {}): StockRow => ({
  code: "12224",
  name: "MG9 LED Panel",
  required: 50,
  stock: 319,
  net: 259,
  method: "test",
  spare: 4,
  rounded: 60,
  ...overrides,
});

describe("baseCodeOf", () => {
  it("strips a shaped-panel orientation suffix", () => {
    expect(baseCodeOf("12398-LU")).toBe("12398");
    expect(baseCodeOf("12398-LD")).toBe("12398");
    expect(baseCodeOf("12399-RU")).toBe("12399");
  });

  it("leaves a plain code unchanged", () => {
    expect(baseCodeOf("12224")).toBe("12224");
  });
});

describe("isMappableStockRow", () => {
  it("excludes the synthetic box-count rows", () => {
    expect(isMappableStockRow({ code: "BOX-MG9" })).toBe(false);
    expect(isMappableStockRow({ code: "BOX-MT" })).toBe(false);
  });

  it("includes real equipment codes", () => {
    expect(isMappableStockRow({ code: "12224" })).toBe(true);
  });
});

describe("applyLiveRentmanData", () => {
  it("overrides stock/net/available for a mapped and fetched row", () => {
    const rows = [row()];
    const liveStock = { "12224": { name: "MG9 LED Panel (Rentman)", currentQuantity: 500 } };
    const liveAvailable = { "12224": 200 };
    const result = applyLiveRentmanData(rows, liveStock, liveAvailable);
    expect(result[0].stock).toBe(500);
    expect(result[0].net).toBe(500 - 60); // rounded (60) is unchanged
    expect(result[0].available).toBe(200);
  });

  it("passes an unmapped row through completely unchanged", () => {
    const rows = [row({ code: "99999" })];
    const result = applyLiveRentmanData(rows, {}, {});
    expect(result[0]).toEqual(rows[0]);
  });

  it("is a full no-op with empty live-data maps (zero behaviour change when unconfigured)", () => {
    const rows = [row(), row({ code: "12223", name: "MT Mesh Panel" })];
    const result = applyLiveRentmanData(rows, {}, {});
    expect(result).toEqual(rows);
  });

  it("falls back cleanly when a row is mapped but missing from the fetched live data", () => {
    // e.g. a stale/deleted Rentman equipment id, or one failed lookup in a batch -
    // liveStock/liveAvailable simply don't have an entry for this code.
    const rows = [row()];
    const result = applyLiveRentmanData(rows, { "99999": { name: "unrelated", currentQuantity: 10 } }, {});
    expect(result[0]).toEqual(rows[0]);
    expect(result[0].stock).toBe(319); // original hardcoded value, not NaN/undefined
    expect(result[0].available).toBeUndefined();
  });

  it("strips the orientation suffix before looking up shaped-panel rows", () => {
    const rows = [row({ code: "12398-LU", name: "Triangle Panel LU" }), row({ code: "12398-LD", name: "Triangle Panel LD" })];
    const liveStock = { "12398": { name: "Triangle Panel", currentQuantity: 15 } };
    const result = applyLiveRentmanData(rows, liveStock, {});
    expect(result[0].stock).toBe(15);
    expect(result[1].stock).toBe(15);
  });
});
