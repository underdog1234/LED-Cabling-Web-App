// ---------------------------------------------------------------------------
// Which connector a joined pair of panels needs, and how many of it.
//
// Kept as a pure table away from App.tsx so the rules can be read, tested and
// corrected on their own - a wrong connector here is a wrong pull sheet, and
// nobody finds out until the truck is on site.
//
// TWO THINGS DECIDE THE ANSWER:
//
//  1. What the two panels are. A "shape panel" is an MG12 triangle or an MG13
//     quarter circle. Everything else on an MG9 wall counts as MG9 - including
//     a corner panel, when it is joined to a shape panel.
//
//  2. Which way the edge they share runs. A HORIZONTAL edge is a horizontal
//     line, so the panels are stacked one above the other; a VERTICAL edge is
//     a vertical line, so they sit side by side. That is measured off the
//     panels' real world-space anchors, so a rotated panel's edges rotate with
//     it and its connectors move to match (see sharedEdgeOrientation in
//     App.tsx).
//
// Horizontal edges, 3 per shared edge:
//   MG9 Corner <-> MG9 Corner, both used flat  -> 150 Connector
//   MG9 Corner <-> MG9 Corner, otherwise       -> MG9 Corner Connector
//   shape      <-> MG9                         -> 150 Connector
//   shape      <-> shape                       -> 180 Connector
//
// Vertical edges, 2 per shared edge:
//   shape      <-> MG9                         -> Horizontal Connector
//   shape      <-> shape                       -> 180 Connector
//
// SHAPE TO SHAPE IS ALWAYS THE 180 CONNECTOR, whichever way the edge runs.
// The vertical case took the 150 Connector until v0.53.0; two shaped panels
// meeting leg to leg need the 180 either way round.
//
// Anything else keeps the rule it had before these were added:
//   MG9 Corner <-> plain MG9, either way round -> 3 x 150 Connector
//   MG9 Corner <-> MG9 Corner on a VERTICAL edge is not in the table above, so
//     it keeps the 3-per-join rule too, split flat/corner the same way as the
//     horizontal case - it is the same pair of parts either way round.
//   plain MG9  <-> plain MG9                   -> nothing
// ---------------------------------------------------------------------------

/** What a panel counts as when working out the connector for a join. */
export type ConnectorPanelClass = "mg9" | "corner" | "cornerFlat" | "shape";

/** Orientation of the LINE two panels meet along, in world space. */
export type SharedEdgeOrientation = "horizontal" | "vertical";

export type ConnectorKey = "connector150" | "cornerConnector" | "connector180" | "horizontalConnector";

/**
 * The connector names the brief specifies, kept separate from whatever the
 * stock catalogue calls the item, so a pull sheet can still say which rule
 * asked for it.
 */
export const CONNECTOR_NAMES: Record<ConnectorKey, string> = {
  connector150: "150 Connector",
  cornerConnector: "MG9 Corner Connector",
  connector180: "180 Connector",
  horizontalConnector: "Horizontal Connector",
};

export type ConnectorNeed = {
  connector: ConnectorKey;
  /** Per shared edge - counted ONCE for the edge, never once per panel. */
  qty: number;
  /** Plain-language rule, for the stock row's method text. */
  rule: string;
};

const isShape = (panel: ConnectorPanelClass) => panel === "shape";
const isCorner = (panel: ConnectorPanelClass) => panel === "corner" || panel === "cornerFlat";

/**
 * The connector one shared edge needs, or null when that pair needs none.
 * Symmetric: the order of `a` and `b` never changes the answer.
 */
export const connectorForEdge = (
  a: ConnectorPanelClass,
  b: ConnectorPanelClass,
  edge: SharedEdgeOrientation,
): ConnectorNeed | null => {
  // A shape panel in the pair decides the rule, whatever the other panel is -
  // the shape rows are the specific case, the corner rows below the general one.
  if (isShape(a) || isShape(b)) {
    const bothShape = isShape(a) && isShape(b);
    if (edge === "horizontal") {
      return bothShape
        ? { connector: "connector180", qty: 3, rule: "shape-to-shape horizontal edge" }
        : { connector: "connector150", qty: 3, rule: "shape-to-MG9 horizontal edge" };
    }
    return bothShape
      ? { connector: "connector180", qty: 2, rule: "shape-to-shape vertical edge" }
      : { connector: "horizontalConnector", qty: 2, rule: "shape-to-MG9 vertical edge" };
  }

  if (isCorner(a) && isCorner(b)) {
    // Flat only when BOTH are set flat: one panel still folded round the corner
    // means the corner part is what physically fits, and pulling the corner
    // part when it turns out not to be needed is the safe way to be wrong.
    const bothFlat = a === "cornerFlat" && b === "cornerFlat";
    return bothFlat
      ? { connector: "connector150", qty: 3, rule: "flat corner-to-corner join" }
      : { connector: "cornerConnector", qty: 3, rule: "corner-to-corner join" };
  }

  if (isCorner(a) || isCorner(b)) {
    return { connector: "connector150", qty: 3, rule: "corner-to-flat join" };
  }

  return null;
};
