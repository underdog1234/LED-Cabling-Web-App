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

  it("carries the levelling packers as the one assorted set Rentman stocks", () => {
    // Checked live against Rentman: code 12374, current_quantity 1. It is a
    // set, so the row is never a per-panel count.
    expect(STOCK_CATALOG.levellingPackers).toEqual({
      code: "12374",
      name: "Assorted Levelling Packers & Shims",
      stock: 1,
    });
  });

  it("carries the MT mesh module, the part that actually gets swapped", () => {
    // Checked live against Rentman: code 12626, current_quantity 20. It only
    // ever appears as a spare - see MT_MESH_MODULE_SPARE_RATIO - because the
    // modules go out fitted in the panels.
    expect(STOCK_CATALOG.mtMeshModule).toEqual({
      code: "12626",
      name: "YES TECH MT P3.9-7.8 LED Mesh Module",
      stock: 20,
    });
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

// "Add an item" offers STOCK_CATALOG and nothing else, so a line that the
// layout calculates but the catalogue does not hold can never be added by
// hand. That is how 12257 (MG9 Floor / Hanging Bar) went missing: it was
// written inline in the row builder and existed nowhere else, so a job
// needing a couple of spare bars had no way to put them on the list - and
// neither did a job wanting a panel type other than the one the wall is
// built in.
describe("every calculated stock line is in the catalogue", () => {
  // The codes stockRows can emit that are not driven by a catalogue entry
  // already - panels, cables, bars, distros, reinforcement.
  const CALCULATED_CODES = [
    "12224", // MG9 LED Panel
    "12223", // MT Mesh Panel
    "12257", // MG9 Floor / Hanging Bar
    "12262", // MT Floor / Hanging Bar
    "12245", // 32A 3-phase Power Distro
    "12246", // 63A 3-phase Power Distro
    "12254", // 15m PowerCON Cable
    "12263", // 15m Signal Cable
    "12264", // MG9 Reinforcement Plate
    "12265", // MG9 Reinforcement Screw
  ];

  it.each(CALCULATED_CODES)("%s can be added by hand", (code) => {
    expect(stockCatalogLookup(code)).toBeTruthy();
  });

  it("carries the stock figure the calculated row uses", () => {
    expect(stockCatalogLookup("12224")?.stock).toBe(PANEL_TYPES.MG9.stock.panels);
    expect(stockCatalogLookup("12223")?.stock).toBe(PANEL_TYPES.MT.stock.panels);
    expect(stockCatalogLookup("12257")?.stock).toBe(PANEL_TYPES.MG9.stock.hangingBar);
    expect(stockCatalogLookup("12262")?.stock).toBe(PANEL_TYPES.MT.stock.hangingBar);
    expect(stockCatalogLookup("12254")?.stock).toBe(PANEL_TYPES.MG9.stock.powerCable15m);
    expect(stockCatalogLookup("12263")?.stock).toBe(PANEL_TYPES.MG9.stock.signalCable15m);
  });

  it("offers every panel type, so a wall can be topped up with another kind", () => {
    const names = Object.values(STOCK_CATALOG).map((item) => item.name);
    expect(names).toContain("MG9 LED Panel");
    expect(names).toContain("MT Mesh Panel");
    expect(names).toContain("Tentec P1.86 LED Poster");
    // ...and the shaped parts, which were already there.
    expect(names).toContain("Triangle Panel");
    expect(names).toContain("1/4 Curved Panel");
  });
});
