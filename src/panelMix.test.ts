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

  it("names each shape as its own part, down to the orientation", () => {
    const mix = panelTypeMix([
      ...typed(makeGridPanels(1, 1, "MG9"), "MG9", "TRIANGLE"),
      ...typed(makeGridPanels(1, 1, "MG9"), "MG9", "CURVED"),
      ...typed(makeGridPanels(1, 1, "MG9"), "MG9", "CORNER_FLAT"),
      ...typed(makeGridPanels(1, 1, "MT"), "MT", "CORNER"),
    ]);
    expect(mix.map((entry) => entry.label).sort()).toEqual([
      "MG9 Corner (flat)",
      "MG9 Curved RU",
      "MG9 Triangle LU",
      "MT Corner",
    ]);
    // The catalogue's own name rides along for the tooltip, with MT named as
    // itself rather than as the MG9 the label was written for.
    expect(mix.find((entry) => entry.label === "MT Corner")?.detail).toBe("MT LED Corner Panel");
    expect(mix.find((entry) => entry.label === "MG9 Triangle LU")?.detail).toBe("MG12 Triangle Panel \u2196 Left Up");
  });

  it("keeps each triangle orientation on its own line", () => {
    // Four triangles facing four ways are four different one-way parts off
    // four different shelves - exactly how Stock Calculations counts them.
    const mix = panelTypeMix(
      [0, 90, 180, 270].map((rotation) => ({
        ...makeGridPanels(1, 1, "MG9")[0],
        id: `t-${rotation}`,
        panelType: "MG9" as const,
        panelVariant: "TRIANGLE" as const,
        rotation,
      })),
    );
    expect(mix).toHaveLength(4);
    expect(mix.map((entry) => entry.label).sort()).toEqual([
      "MG9 Triangle LD",
      "MG9 Triangle LU",
      "MG9 Triangle RD",
      "MG9 Triangle RU",
    ]);
    mix.forEach((entry) => expect(entry.count).toBe(1));
  });

  it("snaps a shaped panel spun off the quarter turns to the nearest part", () => {
    // There is no half-way triangle on the shelf, so 45 degrees reads as the
    // quarter turn it is nearest (see normalizeRotation) rather than being
    // dropped off the list - the same bucket Stock Calculations will order.
    const mix = panelTypeMix([
      { ...makeGridPanels(1, 1, "MG9")[0], panelType: "MG9", panelVariant: "TRIANGLE", rotation: 45 },
    ]);
    expect(mix).toEqual([expect.objectContaining({ label: "MG9 Triangle LD", count: 1 })]);
  });

  it("leaves a plain or corner panel with no orientation on its name", () => {
    const mix = panelTypeMix([
      ...typed(makeGridPanels(2, 1, "MG9"), "MG9"),
      ...typed(makeGridPanels(1, 1, "MG9"), "MG9", "CORNER"),
    ]);
    expect(mix.map((entry) => entry.label).sort()).toEqual(["MG9", "MG9 Corner"]);
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
