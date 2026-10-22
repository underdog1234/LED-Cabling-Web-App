import { describe, it, expect } from "vitest";
import { makeGridPanels, snakeOrderSegments, type Cell } from "./App";

// Auto Snake's reading order. One function answers both "what would this
// patch" (the preview overlay) and "patch it" (the button), so these cover the
// walk once for both.

/** A row of `cols` MG9 panels starting at (x, y), tagged to a sub-screen. */
const row = (cols: number, x: number, y: number, subScreenId: string | null): Cell[] =>
  makeGridPanels(cols, 1, "MG9").map((cell, i) => ({
    ...cell,
    id: `${subScreenId ?? "none"}-${x}-${y}-${i}`,
    x: x + i * 500,
    y,
    subScreenId,
  }));

const ids = (segments: Cell[][]) => segments.map((seg) => seg.map((c) => c.id));
const opts = (over: Partial<Parameters<typeof snakeOrderSegments>[1]> = {}) => ({
  direction: "LR",
  alternates: true,
  perSubScreen: false,
  subScreenIds: [] as string[],
  ...over,
});

describe("snakeOrderSegments over one wall", () => {
  it("reads left to right, alternating on the way back", () => {
    const panels = [...row(3, 0, 0, null), ...row(3, 0, 500, null)];
    const [segment] = snakeOrderSegments(panels, opts());
    expect(segment.map((c) => `${c.x},${c.y}`)).toEqual([
      "0,0", "500,0", "1000,0",      // top row, left to right
      "1000,500", "500,500", "0,500", // and back along the second
    ]);
  });

  it("stops alternating when the option is off", () => {
    const panels = [...row(3, 0, 0, null), ...row(3, 0, 500, null)];
    const [segment] = snakeOrderSegments(panels, opts({ alternates: false }));
    expect(segment.map((c) => `${c.x},${c.y}`)).toEqual([
      "0,0", "500,0", "1000,0",
      "0,500", "500,500", "1000,500",
    ]);
  });
});

describe("snakeOrderSegments per sub-screen", () => {
  // Two screens side by side at the SAME height, touching - the case where a
  // row band swallows both and the snake runs straight from one into the
  // other.
  const twoScreensTouching = () => [
    ...row(2, 0, 0, "left"),
    ...row(2, 0, 500, "left"),
    ...row(2, 1000, 0, "right"),
    ...row(2, 1000, 500, "right"),
  ];

  it("runs one path through both screens when the option is off", () => {
    const segments = snakeOrderSegments(twoScreensTouching(), opts({ subScreenIds: ["left", "right"] }));
    expect(segments).toHaveLength(1);
    // The top band holds all four panels, so the walk crosses the boundary.
    expect(segments[0].slice(0, 4).map((c) => c.subScreenId)).toEqual(["left", "left", "right", "right"]);
  });

  it("keeps each screen's path inside its own screen when the option is on", () => {
    const segments = snakeOrderSegments(twoScreensTouching(), opts({ perSubScreen: true, subScreenIds: ["left", "right"] }));
    expect(segments).toHaveLength(2);
    segments.forEach((segment) => expect(new Set(segment.map((c) => c.subScreenId)).size).toBe(1));
    expect(segments.map((s) => s[0].subScreenId)).toEqual(["left", "right"]);
    // Each screen snakes on its own: left to right, then back.
    expect(segments[0].map((c) => `${c.x},${c.y}`)).toEqual(["0,0", "500,0", "500,500", "0,500"]);
    expect(segments[1].map((c) => `${c.x},${c.y}`)).toEqual(["1000,0", "1500,0", "1500,500", "1000,500"]);
  });

  it("follows the order the screens are listed in, not where they sit", () => {
    const segments = snakeOrderSegments(twoScreensTouching(), opts({ perSubScreen: true, subScreenIds: ["right", "left"] }));
    expect(segments.map((s) => s[0].subScreenId)).toEqual(["right", "left"]);
  });

  it("walks panels in no sub-screen last, as a group of their own", () => {
    const panels = [...row(2, 0, 0, "left"), ...row(2, 2000, 0, null)];
    const segments = snakeOrderSegments(panels, opts({ perSubScreen: true, subScreenIds: ["left"] }));
    expect(segments).toHaveLength(2);
    expect(segments[0].every((c) => c.subScreenId === "left")).toBe(true);
    expect(segments[1].every((c) => c.subScreenId === null)).toBe(true);
  });

  it("loses no panel, whichever way it is walked", () => {
    const panels = [...twoScreensTouching(), ...row(2, 3000, 0, null)];
    for (const perSubScreen of [false, true]) {
      for (const direction of ["LR", "RL", "LRB", "RLB", "TB", "BT", "LOOP_TOGETHER"]) {
        const segments = snakeOrderSegments(panels, opts({ perSubScreen, direction, subScreenIds: ["left", "right"] }));
        const walked = segments.flat().map((c) => c.id).sort();
        expect({ perSubScreen, direction, count: walked.length }).toEqual({ perSubScreen, direction, count: panels.length });
        expect(new Set(walked).size).toBe(panels.length);
      }
    }
  });

  it("is a plain walk again when there are no sub-screens to split on", () => {
    const panels = [...row(3, 0, 0, null), ...row(3, 0, 500, null)];
    expect(ids(snakeOrderSegments(panels, opts({ perSubScreen: true })))).toEqual(
      ids(snakeOrderSegments(panels, opts())),
    );
  });
});
