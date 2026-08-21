import { describe, expect, it } from "vitest";
import type { StockRow } from "../App";
import { applyStockOverrides, baseCodeOf, buildStockComparison } from "./stockOverrides";

const row = (overrides: Partial<StockRow> = {}): StockRow => ({
  code: "12245",
  name: "32A Distro",
  required: 4,
  stock: 4,
  net: 0,
  method: "test",
  ...overrides,
});

describe("baseCodeOf", () => {
  it("strips shaped-panel orientation suffixes", () => {
    expect(baseCodeOf("12398-LU")).toBe("12398");
    expect(baseCodeOf("12398-RD")).toBe("12398");
  });

  it("leaves plain codes unchanged", () => {
    expect(baseCodeOf("12245")).toBe("12245");
  });
});

describe("applyStockOverrides", () => {
  it("overrides stock and net for a matching code", () => {
    const [result] = applyStockOverrides([row({ rounded: 5 })], { "12245": 10 });
    expect(result.stock).toBe(10);
    expect(result.net).toBe(5);
  });

  it("falls back to required when rounded is absent", () => {
    const [result] = applyStockOverrides([row({ required: 6, rounded: undefined })], { "12245": 10 });
    expect(result.net).toBe(4);
  });

  it("passes through a row with no override unchanged", () => {
    const input = row();
    const [result] = applyStockOverrides([input], {});
    expect(result).toEqual(input);
  });

  it("is a full no-op with an empty overrides map", () => {
    const rows = [row(), row({ code: "12280", name: "Joiner" })];
    expect(applyStockOverrides(rows, {})).toEqual(rows);
  });

  it("matches shaped-panel rows via their base code", () => {
    const [result] = applyStockOverrides([row({ code: "12398-LU", rounded: 2 })], { "12398": 8 });
    expect(result.stock).toBe(8);
    expect(result.net).toBe(6);
  });
});

describe("buildStockComparison", () => {
  it("pairs current stock with the fetched Rentman quantity and name", () => {
    const [result] = buildStockComparison(
      [{ code: "12245", name: "32A Distro", stock: 4 }],
      { "12245": { name: "YES TECH 32A 3-phase Distro", currentQuantity: 4 } },
    );
    expect(result).toEqual({ code: "12245", localName: "32A Distro", rentmanName: "YES TECH 32A 3-phase Distro", oldQuantity: 4, newQuantity: 4 });
  });

  it("reports null newQuantity when a code did not resolve in Rentman", () => {
    const [result] = buildStockComparison([{ code: "99999", name: "Ghost Item", stock: 1 }], { "99999": null });
    expect(result.rentmanName).toBeNull();
    expect(result.newQuantity).toBeNull();
  });
});
