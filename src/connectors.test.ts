import { describe, it, expect } from "vitest";
import { CONNECTOR_NAMES, connectorForEdge, type ConnectorPanelClass, type SharedEdgeOrientation } from "./model/connectors";
import { sharedEdgeOrientation, type PanelAnchorSpec } from "./model/panels";

// The connector rules decide what comes off the shelf, so they are pinned
// here rule by rule: which part, how many, and that the answer never depends
// on which panel of the pair you name first.

const CLASSES: ConnectorPanelClass[] = ["mg9", "corner", "cornerFlat", "shape"];
const EDGES: SharedEdgeOrientation[] = ["horizontal", "vertical"];

describe("connectorForEdge", () => {
  it("gives a horizontal flat corner-to-corner join 3 x 150 Connector", () => {
    const need = connectorForEdge("cornerFlat", "cornerFlat", "horizontal");
    expect(need).toMatchObject({ connector: "connector150", qty: 3 });
    expect(CONNECTOR_NAMES[need!.connector]).toBe("150 Connector");
  });

  it("gives a horizontal corner-to-corner join 3 x MG9 Corner Connector", () => {
    const need = connectorForEdge("corner", "corner", "horizontal");
    expect(need).toMatchObject({ connector: "cornerConnector", qty: 3 });
    expect(CONNECTOR_NAMES[need!.connector]).toBe("MG9 Corner Connector");
  });

  it("treats a pair where only one panel is flat as a corner join", () => {
    // One panel still folded round the corner means the corner part is what
    // physically fits.
    expect(connectorForEdge("corner", "cornerFlat", "horizontal")).toMatchObject({ connector: "cornerConnector", qty: 3 });
    expect(connectorForEdge("cornerFlat", "corner", "horizontal")).toMatchObject({ connector: "cornerConnector", qty: 3 });
  });

  it("follows the shape-panel rules on both edge orientations", () => {
    expect(connectorForEdge("shape", "mg9", "horizontal")).toMatchObject({ connector: "connector150", qty: 3 });
    expect(connectorForEdge("shape", "shape", "horizontal")).toMatchObject({ connector: "connector180", qty: 3 });
    expect(connectorForEdge("shape", "mg9", "vertical")).toMatchObject({ connector: "horizontalConnector", qty: 2 });
    expect(connectorForEdge("shape", "shape", "vertical")).toMatchObject({ connector: "connector180", qty: 2 });
  });

  it("gives a corner panel meeting a shape its own rule, folded or flat", () => {
    // Not the shape-to-MG9 rule, which is what a corner panel used to fall
    // back to: 3 x 180 stacked, 2 x 150 side by side, either way up.
    for (const corner of ["corner", "cornerFlat"] as ConnectorPanelClass[]) {
      expect(connectorForEdge(corner, "shape", "horizontal")).toMatchObject({ connector: "connector180", qty: 3 });
      expect(connectorForEdge("shape", corner, "horizontal")).toMatchObject({ connector: "connector180", qty: 3 });
      expect(connectorForEdge(corner, "shape", "vertical")).toMatchObject({ connector: "connector150", qty: 2 });
      expect(connectorForEdge("shape", corner, "vertical")).toMatchObject({ connector: "connector150", qty: 2 });
    }
  });

  it("keeps a PLAIN MG9 against a shape on the shape-to-MG9 rule", () => {
    expect(connectorForEdge("mg9", "shape", "horizontal")).toMatchObject({ connector: "connector150", qty: 3 });
    expect(connectorForEdge("mg9", "shape", "vertical")).toMatchObject({ connector: "horizontalConnector", qty: 2 });
  });

  it("keeps the corner-to-plain-MG9 rule that was already there", () => {
    for (const edge of EDGES) {
      expect(connectorForEdge("corner", "mg9", edge)).toMatchObject({ connector: "connector150", qty: 3 });
      expect(connectorForEdge("cornerFlat", "mg9", edge)).toMatchObject({ connector: "connector150", qty: 3 });
    }
  });

  it("asks for nothing between two plain MG9 panels", () => {
    for (const edge of EDGES) expect(connectorForEdge("mg9", "mg9", edge)).toBeNull();
  });

  it("gives the same answer whichever panel is named first", () => {
    for (const a of CLASSES) {
      for (const b of CLASSES) {
        for (const edge of EDGES) {
          expect({ a, b, edge, need: connectorForEdge(a, b, edge) })
            .toEqual({ a, b, edge, need: connectorForEdge(b, a, edge) });
        }
      }
    }
  });
});

// A shared edge is found from the panels' real world-space anchors, so turning
// a panel turns the edges it joins along - and the connectors move with them.
const panelAt = (cx: number, cy: number, rotation = 0): PanelAnchorSpec => ({
  cx,
  cy,
  halfW: 250,
  halfH: 250,
  rotation,
  shape: "rect",
});

describe("sharedEdgeOrientation", () => {
  it("reads stacked panels as a horizontal edge and side-by-side as a vertical one", () => {
    expect(sharedEdgeOrientation(panelAt(0, 0), panelAt(0, 500))).toBe("horizontal");
    expect(sharedEdgeOrientation(panelAt(0, 0), panelAt(500, 0))).toBe("vertical");
  });

  it("finds no edge between panels that do not touch", () => {
    expect(sharedEdgeOrientation(panelAt(0, 0), panelAt(2000, 0))).toBeNull();
    expect(sharedEdgeOrientation(panelAt(0, 0), panelAt(500, 500))).toBeNull();
  });

  it("only joins shaped panels along their straight legs", () => {
    // A triangle carries connectors on its two right-angle legs only - at
    // rotation 0 the left and bottom edges. Two of them stacked the same way
    // up meet hypotenuse-to-leg and do not join at all, so they ask for no
    // connector; turn one of them round and the legs meet on a horizontal
    // edge, which is the shape-to-shape 180 Connector case.
    const tri = (cx: number, cy: number, rotation: number): PanelAnchorSpec => ({ cx, cy, halfW: 250, halfH: 250, rotation, shape: "triangle" });
    expect(sharedEdgeOrientation(tri(0, 0, 0), tri(0, 500, 0))).toBeNull();
    expect(sharedEdgeOrientation(tri(0, 0, 0), tri(0, 500, 180))).toBe("horizontal");
    expect(connectorForEdge("shape", "shape", sharedEdgeOrientation(tri(0, 0, 0), tri(0, 500, 180))!))
      .toMatchObject({ connector: "connector180", qty: 3 });
    // Legs meeting side by side is the vertical-edge case - the same 180
    // Connector as stacked, only two of them rather than three.
    expect(sharedEdgeOrientation(tri(0, 0, 0), tri(-500, 0, 180))).toBe("vertical");
    expect(connectorForEdge("shape", "shape", sharedEdgeOrientation(tri(0, 0, 0), tri(-500, 0, 180))!))
      .toMatchObject({ connector: "connector180", qty: 2 });
  });

  it("turns the edge with the panel", () => {
    // An MT-shaped panel (1m x 0.5m) on its side meets a neighbour along a
    // different edge than it would lying flat - the connectors have to follow.
    const wide = (cx: number, cy: number, rotation: number): PanelAnchorSpec => ({ cx, cy, halfW: 500, halfH: 250, rotation, shape: "rect" });
    // Upright (rotated a quarter turn) its long edges run vertically, so a
    // neighbour beside it shares a vertical edge...
    expect(sharedEdgeOrientation(wide(0, 0, 90), wide(500, 0, 90))).toBe("vertical");
    // ...and lying flat, a neighbour above it shares a horizontal one.
    expect(sharedEdgeOrientation(wide(0, 0, 0), wide(0, 500, 0))).toBe("horizontal");
  });
});

// The whole decision table in one place, as a person would read it off a
// pull sheet. Every pair of panel kinds, both edge orientations: 20 rows, and
// nothing outside them. If a rule changes, this is the row that has to be
// edited to say so - and the table in the README has to be edited with it.
describe("the connector table, whole", () => {
  const label: Record<ConnectorPanelClass, string> = {
    mg9: "Plain MG9",
    corner: "MG9 Corner (folded)",
    cornerFlat: "MG9 Corner (laid flat)",
    shape: "MG12 triangle / MG13 curve",
  };

  it("decides these 20 joins and no others", () => {
    const rows: string[] = [];
    for (let i = 0; i < CLASSES.length; i += 1) {
      for (let j = i; j < CLASSES.length; j += 1) {
        for (const edge of EDGES) {
          const need = connectorForEdge(CLASSES[i], CLASSES[j], edge);
          rows.push(
            `${label[CLASSES[i]]} + ${label[CLASSES[j]]} | ${edge} | ${need ? `${CONNECTOR_NAMES[need.connector]} x${need.qty}` : "none"}`,
          );
        }
      }
    }
    expect(rows).toEqual([
      "Plain MG9 + Plain MG9 | horizontal | none",
      "Plain MG9 + Plain MG9 | vertical | none",
      "Plain MG9 + MG9 Corner (folded) | horizontal | 150 Connector x3",
      "Plain MG9 + MG9 Corner (folded) | vertical | 150 Connector x3",
      "Plain MG9 + MG9 Corner (laid flat) | horizontal | 150 Connector x3",
      "Plain MG9 + MG9 Corner (laid flat) | vertical | 150 Connector x3",
      "Plain MG9 + MG12 triangle / MG13 curve | horizontal | 150 Connector x3",
      "Plain MG9 + MG12 triangle / MG13 curve | vertical | Horizontal Connector x2",
      "MG9 Corner (folded) + MG9 Corner (folded) | horizontal | MG9 Corner Connector x3",
      "MG9 Corner (folded) + MG9 Corner (folded) | vertical | MG9 Corner Connector x3",
      "MG9 Corner (folded) + MG9 Corner (laid flat) | horizontal | MG9 Corner Connector x3",
      "MG9 Corner (folded) + MG9 Corner (laid flat) | vertical | MG9 Corner Connector x3",
      "MG9 Corner (folded) + MG12 triangle / MG13 curve | horizontal | 180 Connector x3",
      "MG9 Corner (folded) + MG12 triangle / MG13 curve | vertical | 150 Connector x2",
      "MG9 Corner (laid flat) + MG9 Corner (laid flat) | horizontal | 150 Connector x3",
      "MG9 Corner (laid flat) + MG9 Corner (laid flat) | vertical | 150 Connector x3",
      "MG9 Corner (laid flat) + MG12 triangle / MG13 curve | horizontal | 180 Connector x3",
      "MG9 Corner (laid flat) + MG12 triangle / MG13 curve | vertical | 150 Connector x2",
      "MG12 triangle / MG13 curve + MG12 triangle / MG13 curve | horizontal | 180 Connector x3",
      "MG12 triangle / MG13 curve + MG12 triangle / MG13 curve | vertical | 180 Connector x2",
    ]);
  });
});

// What the edge finder does NOT see. These are not bugs being pinned in
// place - they are the boundaries of the model, written down so a change to
// it shows up here rather than in somebody's connector count on site.
describe("joins that produce no connector at all", () => {
  const at = (cx: number, cy: number, rotation = 0, shape: PanelAnchorSpec["shape"] = "rect", halfW = 250, halfH = 250): PanelAnchorSpec =>
    ({ cx, cy, halfW, halfH, rotation, shape });

  it("sees a flush join, and an offset of a whole anchor spacing", () => {
    expect(sharedEdgeOrientation(at(0, 0), at(500, 0))).toBe("vertical");
    expect(sharedEdgeOrientation(at(0, 0), at(0, 500))).toBe("horizontal");
    // Anchors sit every 100mm along a 500mm edge, so these still line up.
    expect(sharedEdgeOrientation(at(0, 0), at(100, 500))).toBe("horizontal");
    expect(sharedEdgeOrientation(at(0, 0), at(200, 500))).toBe("horizontal");
  });

  it("does NOT see a brick bond - a neighbour offset half a panel", () => {
    // 250mm falls between the anchors on both edges, so nothing coincides and
    // the pair reads as unjoined. A brick-bonded wall therefore counts no
    // connectors along its staggered joins.
    expect(sharedEdgeOrientation(at(0, 0), at(250, 500))).toBeNull();
    expect(sharedEdgeOrientation(at(0, 0), at(500, 250))).toBeNull();
  });

  it("does NOT see an MT panel meeting an MG9 one", () => {
    // Their anchor spacings differ (MT is 1m wide), so even before the MG9-only
    // filter in App.tsx, no anchor of one lands on an anchor of the other.
    expect(sharedEdgeOrientation(at(0, 0, 0, "rect", 500, 250), at(-250, 500))).toBeNull();
  });

  it("does NOT see panels turned to an odd angle", () => {
    expect(sharedEdgeOrientation(at(0, 0, 45), at(707, 0, 45))).toBeNull();
  });

  it("does NOT join a triangle along its hypotenuse", () => {
    expect(sharedEdgeOrientation(at(0, 0, 0, "triangle"), at(500, 0, 0, "triangle"))).toBeNull();
  });
});
