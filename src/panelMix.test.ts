import { describe, it, expect } from "vitest";
import { makeGridPanels, makePosterAt, panelTypeMix, type Cell } from "./App";

// What the selection toolbar reports. The split is the one Stock Calculations
// orders against, so "24 selected" can be read as the parts it would pull -
// 20 standard panels and 4 corners is two different pulls, and a bare count
// never said so.

const typed = (cells: Cell[], panelType: Cell["panelType"], variant?: Cell["panelVariant"]): Cell[] =>
  cells.map((cell) => ({ ...cell, panelType, panelVariant: variant ?? "STANDARD" }));

describe("panelTypeMix", () => {
  it("splits a mixed selection by part, biggest group first", () => {
    const mix = panelTypeMix([
      ...typed(makeGridPanels(6, 1, "MG9"), "MG9"),
      ...typed(makeGridPanels(2, 1, "MG9"), "MG9", "CORNER"),
      ...typed(makeGridPanels(3, 1, "MT"), "MT"),
    ]);
    expect(mix.map((entry) => [entry.label, entry.count])).toEqual([
      ["MG9", 6],
      ["MT", 3],
      ["MG9 Corner", 2],
    ]);
  });

  it("names each shape as its own part", () => {
    const mix = panelTypeMix([
      ...typed(makeGridPanels(1, 1, "MG9"), "MG9", "TRIANGLE"),
      ...typed(makeGridPanels(1, 1, "MG9"), "MG9", "CURVED"),
      ...typed(makeGridPanels(1, 1, "MG9"), "MG9", "CORNER_FLAT"),
      ...typed(makeGridPanels(1, 1, "MT"), "MT", "CORNER"),
    ]);
    expect(mix.map((entry) => entry.label).sort()).toEqual([
      "MG9 Corner (flat)",
      "MG9 Curved",
      "MG9 Triangle",
      "MT Corner",
    ]);
    // The catalogue's own name rides along for the tooltip, with MT named as
    // itself rather than as the MG9 the label was written for.
    expect(mix.find((entry) => entry.label === "MT Corner")?.detail).toBe("MT LED Corner Panel");
  });

  it("counts whole posters, not the eight sections selecting one pulls in", () => {
    const mix = panelTypeMix([...makePosterAt(0, 0), ...makePosterAt(1000, 0)]);
    expect(mix).toHaveLength(1);
    expect(mix[0]).toMatchObject({ label: "LED Poster", count: 2 });
  });

  it("says how many of a group are inactive, so a Delete is not a surprise", () => {
    const panels = typed(makeGridPanels(4, 1, "MG9"), "MG9");
    panels[0].isRemoved = true;
    panels[1].isRemoved = true;
    expect(panelTypeMix(panels)[0]).toMatchObject({ label: "MG9", count: 4, removed: 2 });
  });

  it("counts a poster as inactive only when every section of it is", () => {
    const whole = makePosterAt(0, 0).map((cell) => ({ ...cell, isRemoved: true }));
    const partial = makePosterAt(1000, 0);
    partial[0].isRemoved = true;
    expect(panelTypeMix([...whole, ...partial])[0]).toMatchObject({ count: 2, removed: 1 });
  });

  it("has nothing to say about an empty selection", () => {
    expect(panelTypeMix([])).toEqual([]);
  });
});
