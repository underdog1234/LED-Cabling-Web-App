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
    expect(connectorForEdge("shape", "shape", "vertical")).toMatchObject({ connector: "connector150", qty: 2 });
  });

  it("lets a shape panel's rule win over the corner rules", () => {
    // A corner panel counts as MG9 when the other side of the edge is a shape.
    expect(connectorForEdge("shape", "corner", "vertical")).toMatchObject({ connector: "horizontalConnector", qty: 2 });
    expect(connectorForEdge("cornerFlat", "shape", "horizontal")).toMatchObject({ connector: "connector150", qty: 3 });
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
    // Legs meeting side by side is the vertical-edge case.
    expect(sharedEdgeOrientation(tri(0, 0, 0), tri(-500, 0, 180))).toBe("vertical");
    expect(connectorForEdge("shape", "shape", sharedEdgeOrientation(tri(0, 0, 0), tri(-500, 0, 180))!))
      .toMatchObject({ connector: "connector150", qty: 2 });
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
