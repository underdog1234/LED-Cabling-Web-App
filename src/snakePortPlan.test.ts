import { describe, it, expect } from "vitest";
import { makeGridPanels, planSnakePorts, snakeOrderSegments, type Cell } from "./App";

// The preview and the Auto Snake button share this one function, because they
// did not used to: the preview drew a single unbroken line through every panel
// in reading order while the button cut that line into one run per port. On a
// 72-panel wall at 23 panels a port the drawing said one run and the button
// made four.
const plan = (segments: Cell[][], over: Partial<Parameters<typeof planSnakePorts>[1]> = {}) =>
  planSnakePorts(segments, {
    letterMode: false,
    maxPorts: 20,
    maxPerPort: 23,
    freshPortPerSegment: false,
    ...over,
  });

const portsUsed = (assignment: Map<string, number>) => {
  const counts = new Map<number, number>();
  assignment.forEach((port) => counts.set(port, (counts.get(port) ?? 0) + 1));
  return [...counts.entries()].sort((a, b) => a[0] - b[0]);
};

describe("planSnakePorts fills ports in order", () => {
  it("splits a long walk at the per-port ceiling", () => {
    const wall = makeGridPanels(12, 6, "MG9"); // 72 panels
    const segments = snakeOrderSegments(wall, { direction: "LR", alternates: true, perSubScreen: false, subScreenIds: [] });
    expect(portsUsed(plan(segments))).toEqual([[1, 23], [2, 23], [3, 23], [4, 3]]);
  });

  it("keeps a walk that fits on one port on one port", () => {
    const wall = makeGridPanels(6, 3, "MG9"); // 18
    const segments = snakeOrderSegments(wall, { direction: "LR", alternates: true, perSubScreen: false, subScreenIds: [] });
    expect(portsUsed(plan(segments))).toEqual([[1, 18]]);
  });

  it("assigns every panel exactly once, in walk order", () => {
    const wall = makeGridPanels(12, 6, "MG9");
    const segments = snakeOrderSegments(wall, { direction: "LR", alternates: true, perSubScreen: false, subScreenIds: [] });
    const assignment = plan(segments);
    const walked = segments.flat();
    expect(assignment.size).toBe(walked.length);
    // Ports never go backwards along the walk.
    let last = 0;
    walked.forEach((cell) => {
      const port = assignment.get(cell.id)!;
      expect(port).toBeGreaterThanOrEqual(last);
      last = port;
    });
  });

  it("works around ports another sub-screen already holds", () => {
    const wall = makeGridPanels(5, 1, "MG9"); // 5 panels
    const segments = snakeOrderSegments(wall, { direction: "LR", alternates: true, perSubScreen: false, subScreenIds: [] });
    // Port 1 is full, port 2 has room for three.
    const assignment = plan(segments, { startingCounts: new Map([[1, 23], [2, 20]]) });
    expect(portsUsed(assignment)).toEqual([[2, 3], [3, 2]]);
  });

  it("leaves panels unassigned rather than inventing hardware", () => {
    const wall = makeGridPanels(10, 1, "MG9");
    const segments = snakeOrderSegments(wall, { direction: "LR", alternates: true, perSubScreen: false, subScreenIds: [] });
    const assignment = plan(segments, { maxPorts: 1, maxPerPort: 4 });
    expect(assignment.size).toBe(4);
  });

  it("starts each run on a fresh port when asked", () => {
    const a = makeGridPanels(3, 1, "MG9").map((c, i) => ({ ...c, id: `a${i}` }));
    const b = makeGridPanels(3, 1, "MG9").map((c, i) => ({ ...c, id: `b${i}` }));
    expect(portsUsed(plan([a, b], { freshPortPerSegment: true }))).toEqual([[1, 3], [2, 3]]);
    // ...and packs them together when not.
    expect(portsUsed(plan([a, b], { freshPortPerSegment: false }))).toEqual([[1, 6]]);
  });
});

describe("planSnakePorts also fills plugs on amps", () => {
  // A plug has a panel ceiling AND an amp ceiling; whichever is hit first
  // moves the walk on.
  it("moves to the next plug when the watts run out before the count does", () => {
    const wall = makeGridPanels(10, 1, "MG9");
    const segments = snakeOrderSegments(wall, { direction: "LR", alternates: true, perSubScreen: false, subScreenIds: [] });
    const assignment = plan(segments, {
      maxPerPort: 100, // never the limiting factor
      weightOf: () => 175, // MG9 max watts
      maxWeightPerPort: 700, // room for exactly four
    });
    expect(portsUsed(assignment)).toEqual([[1, 4], [2, 4], [3, 2]]);
  });

  it("counts the load a plug already carries", () => {
    const wall = makeGridPanels(4, 1, "MG9");
    const segments = snakeOrderSegments(wall, { direction: "LR", alternates: true, perSubScreen: false, subScreenIds: [] });
    const assignment = plan(segments, {
      maxPerPort: 100,
      startingWeights: new Map([[1, 600]]),
      weightOf: () => 175,
      maxWeightPerPort: 700,
    });
    // Only 100W spare on plug 1, so nothing fits there.
    expect(portsUsed(assignment)).toEqual([[2, 4]]);
  });
});

describe("letter mode keeps a letter whole where it fits", () => {
  const letter = (id: string, n: number) => makeGridPanels(n, 1, "MG9").map((c, i) => ({ ...c, id: `${id}${i}` }));

  it("moves a letter to a fresh port rather than splitting it", () => {
    // Port 1 already holds 20 of 23. The next letter is 5 panels: it fits a
    // whole port but not the remaining 3, so it starts a new one.
    const assignment = plan([letter("a", 5)], { letterMode: true, startingCounts: new Map([[1, 20]]) });
    expect(portsUsed(assignment)).toEqual([[2, 5]]);
  });

  it("splits a letter that is bigger than a whole port - it has nowhere else to go", () => {
    const assignment = plan([letter("a", 30)], { letterMode: true });
    expect(portsUsed(assignment)).toEqual([[1, 23], [2, 7]]);
  });
});
