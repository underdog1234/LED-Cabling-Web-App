import { describe, it, expect } from "vitest";
import { PANEL_TYPES, STOCK_CATALOG, stockCatalogLookup } from "./App";

// The catalogue is the one place a Rentman code is written down, and getting
// one wrong orders the wrong part - which has happened: 12398/12399 came in
// swapped, and 12250/12251/12252 were on the floor parts while Rentman had
// those codes against the MT corner bracket, its bolt and the patch cable.
// These assertions are the "do not tidy these back" note in executable form.

describe("STOCK_CATALOG", () => {
  it("never puts two items on one code", () => {
    // Two rows on one code read as a duplicate requirement and break the
    // Rentman stock comparison, so a collision has to fail here, not on a
    // pull sheet. (stockItemByCode would also silently resolve to whichever
    // of the two came first.)
    const byCode = new Map<string, string[]>();
    Object.values(STOCK_CATALOG).forEach((item) => {
      byCode.set(item.code, [...(byCode.get(item.code) ?? []), item.name]);
    });
    const collisions = [...byCode.entries()].filter(([, names]) => names.length > 1);
    expect(collisions).toEqual([]);
  });

  it("has a positive stock figure or an honest zero for every item", () => {
    Object.values(STOCK_CATALOG).forEach((item) => {
      expect({ code: item.code, ok: Number.isInteger(item.stock) && item.stock >= 0 }).toEqual({ code: item.code, ok: true });
    });
  });

  it("keeps the confirmed codes that read backwards", () => {
    // Triangle is the HIGHER number and the quarter circle the lower, the
    // opposite way round to what the names suggest.
    expect(STOCK_CATALOG.mg12Triangle.code).toBe("12399");
    expect(STOCK_CATALOG.mg13Curved.code).toBe("12398");
    // The floor parts and the three items that held their old codes.
    expect(STOCK_CATALOG.temperedGlass.code).toBe("12250");
    expect(STOCK_CATALOG.floorReinforcementBar.code).toBe("12251");
    expect(STOCK_CATALOG.floorTaperPin.code).toBe("12252");
    expect(STOCK_CATALOG.patchSignalCable.code).toBe("12272");
    expect(STOCK_CATALOG.mtCornerBracket.code).toBe("12274");
    expect(STOCK_CATALOG.mtCornerBracketBolt.code).toBe("12275");
  });

  it("carries both processors on their own codes", () => {
    expect(STOCK_CATALOG.vx1000Pro).toMatchObject({ code: "12247", name: "NovaStar VX1000 Pro LED Processor" });
    expect(STOCK_CATALOG.vx2000Pro).toMatchObject({ code: "12353", name: "NovaStar VX2000 Pro LED Processor Rack" });
    // The VX2000 is the racked unit, so the two are genuinely different line
    // items - a project pulls one or the other, never a shared code.
    expect(STOCK_CATALOG.vx1000Pro.code).not.toBe(STOCK_CATALOG.vx2000Pro.code);
  });

  it("reads each processor's shelf quantity from the MG9 catalogue", () => {
    // One number to keep current: the figures live in PANEL_TYPES.MG9.stock,
    // which is where a stock check updates them.
    expect(STOCK_CATALOG.vx1000Pro.stock).toBe(PANEL_TYPES.MG9.stock.vx1000);
    expect(STOCK_CATALOG.vx2000Pro.stock).toBe(PANEL_TYPES.MG9.stock.vx2000);
  });

  it("finds every catalogue item by its code", () => {
    // stockCatalogLookup is what a by-hand added row reads its name and shelf
    // quantity from; an item it cannot find lands on the list unnamed.
    Object.values(STOCK_CATALOG).forEach((item) => {
      expect(stockCatalogLookup(item.code)).toMatchObject({ name: item.name, stock: item.stock });
    });
    expect(stockCatalogLookup("00000")).toBeNull();
  });
});
