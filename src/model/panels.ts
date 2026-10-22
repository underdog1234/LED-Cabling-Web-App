// ---------------------------------------------------------------------------
// Free-placement panel model (Stage 2 of the non-uniform overhaul).
//
// A project is a flat list of Panel records positioned in workspace
// millimetres (x/y = TOP-LEFT corner). The old rows×cols grid becomes a
// generator that emits panels on a 500 mm pitch; after generation panels are
// freely movable. MT panels are plain 1000×500 mm records — the old
// head/tail module pairing is gone (kept only in legacy-file migration).
// ---------------------------------------------------------------------------

export const MODULE_MM = 500; // base 0.5 m module
export const HALF_MODULE_MM = 250; // fine snap grid
// The finest interval two panels may be joined on. Halves fall out of this for
// free (two quarters), and nothing between a quarter and the next is reachable
// - see edgeAnchorOffsets.
export const QUARTER_MODULE_MM = 125;
export const SNAP_DISTANCE_MM = 32; // edge-anchor snap radius (matches layout tool)
export const JOIN_GAP_MM = 2; // max gap for two edges to count as joined
export const JOIN_MIN_SHARED_MM = 100; // min shared edge length for a join

export type PanelSizeSpec = { wMm: number; hMm: number };

export type PanelRecord = {
  id: string;
  panelType: string; // "MG9" | "MT"
  panelVariant: string; // STANDARD | TRIANGLE | CURVED | CORNER
  x: number; // top-left, workspace mm
  y: number; // top-left, workspace mm
  rotation: number; // 0/90/180/270 (clockwise)
  isRemoved: boolean;
  assignedPort: number | null;
  sequence: number | null;
  assignedPowerPort: number | null;
  powerSequence: number | null;
  powerManual: boolean;
  /** Unknown fields from imported files, preserved for round-tripping. */
  _extra?: Record<string, unknown>;
};

let idCounter = 0;
export const newPanelId = () => {
  try {
    return crypto.randomUUID();
  } catch {
    idCounter += 1;
    return `p-${Date.now().toString(36)}-${idCounter}`;
  }
};

export const makePanel = (
  panelType: string,
  x: number,
  y: number,
  overrides: Partial<PanelRecord> = {},
): PanelRecord => ({
  id: newPanelId(),
  panelType,
  panelVariant: "STANDARD",
  x,
  y,
  rotation: 0,
  isRemoved: false,
  assignedPort: null,
  sequence: null,
  assignedPowerPort: null,
  powerSequence: null,
  powerManual: false,
  ...overrides,
});

/** Footprint in mm for a panel, honouring rotation (90/270 swaps w/h). */
export const panelSizeMm = (panel: PanelRecord, baseSize: PanelSizeSpec): PanelSizeSpec => {
  const rot = ((Math.round(panel.rotation / 90) * 90) % 360 + 360) % 360;
  if (rot === 90 || rot === 270) return { wMm: baseSize.hMm, hMm: baseSize.wMm };
  return baseSize;
};

export type RectMm = { x: number; y: number; w: number; h: number };

export const panelRect = (panel: PanelRecord, baseSize: PanelSizeSpec): RectMm => {
  const { wMm, hMm } = panelSizeMm(panel, baseSize);
  return { x: panel.x, y: panel.y, w: wMm, h: hMm };
};

export const rectsOverlap = (a: RectMm, b: RectMm, epsilon = 1) =>
  a.x + epsilon < b.x + b.w && b.x + epsilon < a.x + a.w && a.y + epsilon < b.y + b.h && b.y + epsilon < a.y + a.h;

export const activeBBox = (rects: RectMm[]): RectMm => {
  if (!rects.length) return { x: 0, y: 0, w: 0, h: 0 };
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  rects.forEach((r) => {
    minX = Math.min(minX, r.x);
    minY = Math.min(minY, r.y);
    maxX = Math.max(maxX, r.x + r.w);
    maxY = Math.max(maxY, r.y + r.h);
  });
  return { x: minX, y: minY, w: maxX - minX, h: maxY - minY };
};

export const snapToIncrement = (value: number, step: number) => Math.round(value / step) * step;

/**
 * Edge-snap: given the moving panels' rects and the stationary rects, find the
 * smallest translation (within SNAP_DISTANCE_MM) that makes a moving edge
 * coincide with a stationary edge on one axis while overlapping on the other
 * (so panels join flush). Falls back to the half-module grid.
 */
export const computeSnapDelta = (
  moving: RectMm[],
  others: RectMm[],
  snapEnabled: boolean,
): { dx: number; dy: number; snappedTo: "panel" | "grid" | null } => {
  if (!snapEnabled) return { dx: 0, dy: 0, snappedTo: null };
  let best: { dx: number; dy: number; dist: number } | null = null;
  for (const m of moving) {
    for (const o of others) {
      // Candidate x-deltas that make vertical edges touch, and y-deltas for horizontal edges.
      const xCandidates = [o.x - (m.x + m.w), o.x + o.w - m.x, o.x - m.x, o.x + o.w - (m.x + m.w)];
      const yCandidates = [o.y - (m.y + m.h), o.y + o.h - m.y, o.y - m.y, o.y + o.h - (m.y + m.h)];
      for (const dx of xCandidates) {
        if (Math.abs(dx) > SNAP_DISTANCE_MM) continue;
        // Require some vertical overlap so the snap is a real edge join, then
        // also try to align vertically to the neighbour's top edge if close.
        const vOverlap = Math.min(m.y + m.h, o.y + o.h) - Math.max(m.y, o.y);
        if (vOverlap < -SNAP_DISTANCE_MM) continue;
        let dy = 0;
        for (const cand of yCandidates) {
          if (Math.abs(cand) <= SNAP_DISTANCE_MM && (dy === 0 || Math.abs(cand) < Math.abs(dy))) dy = cand;
        }
        const dist = Math.hypot(dx, dy);
        if (!best || dist < best.dist) best = { dx, dy, dist };
      }
      for (const dy of yCandidates) {
        if (Math.abs(dy) > SNAP_DISTANCE_MM) continue;
        const hOverlap = Math.min(m.x + m.w, o.x + o.w) - Math.max(m.x, o.x);
        if (hOverlap < -SNAP_DISTANCE_MM) continue;
        let dx = 0;
        for (const cand of xCandidates) {
          if (Math.abs(cand) <= SNAP_DISTANCE_MM && (dx === 0 || Math.abs(cand) < Math.abs(dx))) dx = cand;
        }
        const dist = Math.hypot(dx, dy);
        if (!best || dist < best.dist) best = { dx, dy, dist };
      }
    }
  }
  if (best) return { dx: best.dx, dy: best.dy, snappedTo: "panel" };
  // Grid fallback: snap the first moving rect's corner to the half-module grid.
  const first = moving[0];
  if (!first) return { dx: 0, dy: 0, snappedTo: null };
  return {
    dx: snapToIncrement(first.x, HALF_MODULE_MM) - first.x,
    dy: snapToIncrement(first.y, HALF_MODULE_MM) - first.y,
    snappedTo: "grid",
  };
};

// ---------------------------------------------------------------------------
// Connector-anchor snap/join model, ported from the YES TECH layout tool.
//
// Every panel exposes connector anchor points along its *joinable* edges. Two
// panels snap/join when anchors coincide (within SNAP_DISTANCE_MM). Shaped
// panels only expose anchors on their straight legs - a triangle has none on
// its hypotenuse, a quarter-circle none on its curve - so those edges can never
// join to another panel. Anchors rotate with the panel.
// ---------------------------------------------------------------------------

export type PanelShape = "rect" | "triangle" | "curve" | "corner";
export type PanelAnchorSpec = {
  /** Panel centre in workspace mm. */
  cx: number;
  cy: number;
  /** Half extents of the UNROTATED base footprint. */
  halfW: number;
  halfH: number;
  rotation: number;
  shape: PanelShape;
};

// ---------------------------------------------------------------------------
// Where a panel's own label goes, so that it lands on the lit part of the
// panel and not in the empty corner beside it.
//
// A rectangle can take its label anywhere. A TRIANGLE or QUARTER CIRCLE
// cannot: at rotation 0 the triangle's right angle is bottom-left and the
// quarter circle's is bottom-right, so the opposite corner is empty and a
// label placed there is cut away with the shape. These two helpers give a
// point well inside the solid area, mapped through the panel's own rotation
// and mirror, so the label follows the shape whichever way round it is.
// ---------------------------------------------------------------------------

/**
 * A point deep inside a shape's solid area, in the panel's own unrotated
 * 0..w / 0..h space. Chosen for room around it rather than exact centre of
 * area, so a two-line label centred there still fits.
 */
export const panelShapeLabelPoint = (shape: PanelShape, w: number, h: number): { x: number; y: number } => {
  if (shape === "triangle") return { x: w * 0.3, y: h * 0.7 };
  if (shape === "curve") return { x: w * 0.58, y: h * 0.58 };
  return { x: w / 2, y: h / 2 };
};

/** True when a shape's label has to be moved off the panel's corner. */
export const panelShapeNeedsInsetLabel = (shape: PanelShape) => shape === "triangle" || shape === "curve";

/**
 * A point in a panel's own 0..w / 0..h space, mapped to where it lands on the
 * canvas once the panel's rotation and any front-view mirror are applied -
 * the same frame the panel's shape itself is drawn in, so a point inside the
 * shape stays inside it.
 */
export const panelFramePoint = (
  r: RectMm,
  rotation: number,
  mirrorX: boolean,
  lx: number,
  ly: number,
): { x: number; y: number } => {
  const rad = (rotation * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  const px = lx - r.w / 2;
  const py = ly - r.h / 2;
  // Rotate first, then mirror - the order the panel's own draw frame uses.
  const rx = px * cos - py * sin;
  const ry = px * sin + py * cos;
  return { x: r.x + r.w / 2 + (mirrorX ? -rx : rx), y: r.y + r.h / 2 + ry };
};

/**
 * Is a point inside the panel's own silhouette? Local 0..w / 0..h space, the
 * same shapes tracePanelShapePath draws: the triangle's hypotenuse runs from
 * the top-left corner to the bottom-right one, and the quarter circle's curve
 * is the quadratic Bezier (w,0) - control (0,0) - (0,h), whose region is
 * exactly sqrt(x/w) + sqrt(y/h) >= 1.
 */
export const pointInPanelShape = (shape: PanelShape, w: number, h: number, x: number, y: number): boolean => {
  if (w <= 0 || h <= 0) return false;
  if (x < 0 || y < 0 || x > w || y > h) return false;
  if (shape === "triangle") return y * w >= x * h;
  if (shape === "curve") return Math.sqrt(x / w) + Math.sqrt(y / h) >= 1;
  return true;
};

/** panelFramePoint the other way round: canvas coordinates back to the panel's own space. */
export const panelFrameInverse = (
  r: RectMm,
  rotation: number,
  mirrorX: boolean,
  x: number,
  y: number,
): { x: number; y: number } => {
  const rad = (rotation * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  let dx = x - (r.x + r.w / 2);
  const dy = y - (r.y + r.h / 2);
  if (mirrorX) dx = -dx;
  // Undo the rotation (transpose of the rotation matrix).
  const px = dx * cos + dy * sin;
  const py = -dx * sin + dy * cos;
  return { x: px + r.w / 2, y: py + r.h / 2 };
};

/** Where a panel's label block is centred, in canvas coordinates. */
export const panelLabelAnchor = (r: RectMm, shape: PanelShape, rotation: number, mirrorX: boolean) => {
  const local = panelShapeLabelPoint(shape, r.w, r.h);
  return panelFramePoint(r, rotation, mirrorX, local.x, local.y);
};

/**
 * Does an upright block of text centred at (cx, cy) sit wholly on the lit part
 * of the panel? Checked corner by corner against the real silhouette, so it
 * holds at any rotation and for a mirrored (front-view) panel alike.
 */
export const panelLabelBlockFitsAt = (
  r: RectMm,
  shape: PanelShape,
  rotation: number,
  mirrorX: boolean,
  cx: number,
  cy: number,
  halfW: number,
  halfH: number,
): boolean => {
  for (const [sx, sy] of [[-1, -1], [1, -1], [-1, 1], [1, 1]] as Array<[number, number]>) {
    const local = panelFrameInverse(r, rotation, mirrorX, cx + sx * halfW, cy + sy * halfH);
    if (!pointInPanelShape(shape, r.w, r.h, local.x, local.y)) return false;
  }
  return true;
};

/** The same test, for a block centred on the shape's own label point. */
export const panelLabelBlockFits = (
  r: RectMm,
  shape: PanelShape,
  rotation: number,
  mirrorX: boolean,
  halfW: number,
  halfH: number,
): boolean => {
  const anchor = panelLabelAnchor(r, shape, rotation, mirrorX);
  return panelLabelBlockFitsAt(r, shape, rotation, mirrorX, anchor.x, anchor.y, halfW, halfH);
};

/**
 * Where the connector anchors sit along an edge, as offsets in mm from that
 * edge's midpoint: every QUARTER_MODULE_MM out to the edge's ends, plus the
 * two end corners themselves.
 *
 * These used to be fractions of the half extent - [-0.8, -0.4, 0, 0.4, 0.8] -
 * which on a 500mm edge put them every 100mm. That is not a lattice any panel
 * can be joined on: the offsets it allowed (100, 200, 300) are not halves or
 * quarters of a module, and the ones that are (125, 250, 375) had no anchor
 * pair at all. Dragging a panel to sit exactly half a module down therefore
 * fell out of anchor range entirely, dropped through to the 250mm grid snap,
 * and landed looking right while counting as NOT JOINED - no connector
 * ordered for it, and the two panels moved independently afterwards.
 *
 * Measured in absolute mm rather than as a fraction so the lattice is the same
 * on every panel size: an MT panel's 1000mm edge gets anchors every 125mm just
 * like an MG9's 500mm one, so the two can join each other on a quarter.
 */
const edgeAnchorOffsets = (halfExtent: number): number[] => {
  const offsets = new Set<number>([-halfExtent, 0, halfExtent]);
  for (let d = QUARTER_MODULE_MM; d < halfExtent; d += QUARTER_MODULE_MM) {
    offsets.add(-d);
    offsets.add(d);
  }
  return [...offsets].sort((a, b) => a - b);
};
export const ANCHOR_JOIN_TOL = 4; // mm - anchors this close count as coincident

// Local (centre-relative, unrotated) anchor points for a panel shape. Base
// orientations match the layout tool: triangle right-angle bottom-left (legs =
// left + bottom edges); quarter-circle right-angle bottom-right (legs = bottom
// + right edges); rect/corner = all four edges. Shaped panels expose NO anchors
// on their hypotenuse / curve, so those edges can never join.
const localAnchors = (shape: PanelShape, halfW: number, halfH: number): Array<{ x: number; y: number }> => {
  const pts: Array<{ x: number; y: number }> = [];
  const alongH = edgeAnchorOffsets(halfH);
  const alongW = edgeAnchorOffsets(halfW);
  const left = () => alongH.forEach((t) => pts.push({ x: -halfW, y: t }));
  const right = () => alongH.forEach((t) => pts.push({ x: halfW, y: t }));
  const top = () => alongW.forEach((t) => pts.push({ x: t, y: -halfH }));
  const bottom = () => alongW.forEach((t) => pts.push({ x: t, y: halfH }));
  if (shape === "triangle") {
    left();
    bottom();
  } else if (shape === "curve") {
    bottom();
    right();
  } else {
    top();
    bottom();
    left();
    right();
  }
  // A corner belongs to two edges, so it was pushed twice. Left as duplicates
  // it would let two panels that merely TOUCH AT A CORNER report the two
  // coincident anchors panelsAnchorJoined reads as a join - one point counted
  // twice - and order connectors for an edge they do not share.
  const seen = new Set<string>();
  return pts.filter((p) => {
    const key = `${p.x}|${p.y}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
};

/**
 * World-space connector anchors for a panel, rotated by its FULL rotation
 * (any angle, not just multiples of 90) - so a panel spun to e.g. 45deg
 * snaps and joins along its own true rotated edges, not the nearest cardinal
 * axis. (Rounding to the nearest 90 here used to be harmless when rotation
 * was always cardinal; once custom angles became possible it silently
 * snapped every rotated panel as if it were unrotated or at the nearest
 * cardinal angle instead.)
 */
export const panelWorldAnchors = (g: PanelAnchorSpec): Array<{ x: number; y: number }> => {
  const rad = (g.rotation * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  return localAnchors(g.shape, g.halfW, g.halfH).map((p) => ({
    x: g.cx + (p.x * cos - p.y * sin),
    y: g.cy + (p.x * sin + p.y * cos),
  }));
};

/**
 * Anchor snap: find the smallest translation that makes a moving panel anchor
 * coincide with a stationary panel anchor (within SNAP_DISTANCE_MM), snapping
 * panels together only along compatible edges. Falls back to the half-module
 * grid when nothing is in range. `firstRect` is the moving group's reference
 * rect for the grid fallback.
 */
export const computeAnchorSnapDelta = (
  moving: PanelAnchorSpec[],
  others: PanelAnchorSpec[],
  snapEnabled: boolean,
  firstRect: RectMm | null,
): { dx: number; dy: number; snappedTo: "panel" | "grid" | null } => {
  if (!snapEnabled) return { dx: 0, dy: 0, snappedTo: null };
  const movingAnchors = moving.flatMap(panelWorldAnchors);
  const otherAnchors = others.flatMap(panelWorldAnchors);
  let best: { dx: number; dy: number; dist: number } | null = null;
  for (const am of movingAnchors) {
    for (const ao of otherAnchors) {
      const dx = ao.x - am.x;
      const dy = ao.y - am.y;
      const dist = Math.hypot(dx, dy);
      if (dist <= SNAP_DISTANCE_MM && (!best || dist < best.dist)) best = { dx, dy, dist };
    }
  }
  if (best) return { dx: best.dx, dy: best.dy, snappedTo: "panel" };
  if (!firstRect) return { dx: 0, dy: 0, snappedTo: null };
  return {
    dx: snapToIncrement(firstRect.x, HALF_MODULE_MM) - firstRect.x,
    dy: snapToIncrement(firstRect.y, HALF_MODULE_MM) - firstRect.y,
    snappedTo: "grid",
  };
};

/** Two panels are joined when they share >= 2 coincident connector anchors. */
export const panelsAnchorJoined = (a: PanelAnchorSpec, b: PanelAnchorSpec): boolean => {
  const aw = panelWorldAnchors(a);
  const bw = panelWorldAnchors(b);
  let shared = 0;
  for (const pa of aw) {
    for (const pb of bw) {
      if (Math.abs(pa.x - pb.x) <= ANCHOR_JOIN_TOL && Math.abs(pa.y - pb.y) <= ANCHOR_JOIN_TOL) {
        shared += 1;
        break;
      }
    }
    if (shared >= 2) return true;
  }
  return false;
};

/**
 * Orientation of the LINE two joined panels meet along, in world space, or
 * null when they do not meet at all.
 *
 * "horizontal" means the shared edge is a horizontal line, so one panel sits
 * on top of the other; "vertical" means a vertical line, so they sit side by
 * side. Taken from the panels' rotated world anchors rather than their
 * unrotated footprints, so turning a panel turns the edges it joins along -
 * and therefore the connectors that edge needs (see model/connectors.ts).
 *
 * A panel spun to something other than a quarter turn meets its neighbour on a
 * slope; that is classified by whichever way the edge leans furthest, which is
 * the connector a builder would reach for.
 */
export const sharedEdgeOrientation = (a: PanelAnchorSpec, b: PanelAnchorSpec): "horizontal" | "vertical" | null => {
  const bw = panelWorldAnchors(b);
  const shared = panelWorldAnchors(a).filter((pa) =>
    bw.some((pb) => Math.abs(pa.x - pb.x) <= ANCHOR_JOIN_TOL && Math.abs(pa.y - pb.y) <= ANCHOR_JOIN_TOL),
  );
  if (shared.length < 2) return null;
  // Every shared anchor lies on the one edge, so the two furthest apart span it.
  let spanX = 0;
  let spanY = 0;
  let span = 0;
  for (let i = 0; i < shared.length; i += 1) {
    for (let j = i + 1; j < shared.length; j += 1) {
      const dx = Math.abs(shared[i].x - shared[j].x);
      const dy = Math.abs(shared[i].y - shared[j].y);
      const d = Math.hypot(dx, dy);
      if (d > span) {
        span = d;
        spanX = dx;
        spanY = dy;
      }
    }
  }
  if (span <= ANCHOR_JOIN_TOL) return null;
  return spanX >= spanY ? "horizontal" : "vertical";
};

/** Connected components over the anchor-join relation; id -> group index. */
export const connectedGroupsByGeom = (
  panels: PanelRecord[],
  geomOf: (p: PanelRecord) => PanelAnchorSpec,
): Map<string, number> => {
  const active = panels.filter((p) => !p.isRemoved);
  const geoms = active.map(geomOf);
  const parent = active.map((_, i) => i);
  const find = (i: number): number => (parent[i] === i ? i : (parent[i] = find(parent[i])));
  const union = (a: number, b: number) => {
    const ra = find(a);
    const rb = find(b);
    if (ra !== rb) parent[rb] = ra;
  };
  for (let i = 0; i < active.length; i += 1) {
    for (let j = i + 1; j < active.length; j += 1) {
      if (panelsAnchorJoined(geoms[i], geoms[j])) union(i, j);
    }
  }
  const groups = new Map<string, number>();
  active.forEach((p, i) => groups.set(p.id, find(i)));
  return groups;
};

/** Ids joined (directly or transitively) to any seed id, by anchor connectivity. */
export const joinedGroupIdsByGeom = (
  panels: PanelRecord[],
  geomOf: (p: PanelRecord) => PanelAnchorSpec,
  seedIds: Set<string>,
): Set<string> => {
  const groups = connectedGroupsByGeom(panels, geomOf);
  const seedGroups = new Set<number>();
  seedIds.forEach((id) => {
    const g = groups.get(id);
    if (g !== undefined) seedGroups.add(g);
  });
  const out = new Set<string>();
  groups.forEach((g, id) => {
    if (seedGroups.has(g)) out.add(id);
  });
  return out;
};

/** Shared-edge join test between two rects. */
export const rectsJoined = (a: RectMm, b: RectMm): boolean => {
  const gapX = Math.max(a.x, b.x) - Math.min(a.x + a.w, b.x + b.w);
  const gapY = Math.max(a.y, b.y) - Math.min(a.y + a.h, b.y + b.h);
  const sharedY = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y);
  const sharedX = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x);
  // Vertical edges touching (side by side)
  if (Math.abs(gapX) <= JOIN_GAP_MM && sharedY >= JOIN_MIN_SHARED_MM) return true;
  // Horizontal edges touching (stacked)
  if (Math.abs(gapY) <= JOIN_GAP_MM && sharedX >= JOIN_MIN_SHARED_MM) return true;
  return false;
};

/** Connected components over the join relation; returns id → group index. */
export const connectedGroups = (
  panels: PanelRecord[],
  rectOf: (p: PanelRecord) => RectMm,
): Map<string, number> => {
  const active = panels.filter((p) => !p.isRemoved);
  const rects = active.map(rectOf);
  const parent = active.map((_, i) => i);
  const find = (i: number): number => (parent[i] === i ? i : (parent[i] = find(parent[i])));
  const union = (a: number, b: number) => {
    const ra = find(a);
    const rb = find(b);
    if (ra !== rb) parent[rb] = ra;
  };
  for (let i = 0; i < active.length; i += 1) {
    for (let j = i + 1; j < active.length; j += 1) {
      if (rectsJoined(rects[i], rects[j])) union(i, j);
    }
  }
  const groups = new Map<string, number>();
  active.forEach((p, i) => groups.set(p.id, find(i)));
  return groups;
};

/** Ids in the same joined group as any of the seed ids. */
export const joinedGroupIds = (
  panels: PanelRecord[],
  rectOf: (p: PanelRecord) => RectMm,
  seedIds: Set<string>,
): Set<string> => {
  const groups = connectedGroups(panels, rectOf);
  const seedGroups = new Set<number>();
  seedIds.forEach((id) => {
    const g = groups.get(id);
    if (g !== undefined) seedGroups.add(g);
  });
  const out = new Set<string>();
  groups.forEach((g, id) => {
    if (seedGroups.has(g)) out.add(id);
  });
  return out;
};

/** Overlapping active panel id pairs. */
export const findOverlaps = (
  panels: PanelRecord[],
  rectOf: (p: PanelRecord) => RectMm,
): Array<[string, string]> => {
  const active = panels.filter((p) => !p.isRemoved);
  const rects = active.map(rectOf);
  const out: Array<[string, string]> = [];
  for (let i = 0; i < active.length; i += 1) {
    for (let j = i + 1; j < active.length; j += 1) {
      if (rectsOverlap(rects[i], rects[j])) out.push([active[i].id, active[j].id]);
    }
  }
  return out;
};

/**
 * Row-banding for non-uniform layouts: group active panels into visual rows by
 * their vertical centre (tolerance = half module). Bands are ordered top→bottom
 * and panels within a band left→right. Used by snake ordering, pixel maths,
 * and the PNG test pattern.
 */
/**
 * Row / column reference for a panel: which cells of the wall's own module
 * grid it stands in, 1-based and inclusive.
 *
 * Replaces clustering panels into rows by proximity, which cannot describe a
 * wall whose panels are not all on the same lines: a panel dropped half a
 * module down sits in neither of the rows around it, and the old clustering
 * gave it a row of its own - so a three-high brick-bonded wall counted five
 * rows and every number after the offset panel was wrong.
 *
 * Here the grid is fixed by the wall itself, so a panel that straddles two
 * rows is simply IN both of them and says so ("3 & 4"), and every other panel
 * keeps the number it always had.
 */
export type GridRef = { from: number; to: number };

/**
 * Cell size for the grid: the panel extent that occurs most often on that
 * axis. That keeps one panel to one cell for whatever the wall is mostly
 * built from - a wall of 1m-wide MT panels numbers them 1, 2, 3, not 1&2,
 * 3&4 - while a panel of a different size on the same wall honestly spans
 * the cells it covers.
 */
const commonExtent = (extents: number[]): number => {
  const counts = new Map<number, number>();
  extents.forEach((value) => {
    const key = Math.round(value);
    if (key > 0) counts.set(key, (counts.get(key) ?? 0) + 1);
  });
  let best = 0;
  let bestCount = 0;
  counts.forEach((count, key) => {
    if (count > bestCount || (count === bestCount && key < best)) {
      best = key;
      bestCount = count;
    }
  });
  return best || MODULE_MM;
};

// Fraction of a cell that a panel may poke into the next one before it counts
// as being in it. Absorbs rounding, not a real offset.
const GRID_REF_TOL = 0.12;

const gridSpan = (offset: number, extent: number, step: number): GridRef => {
  const from = Math.floor(offset / step + GRID_REF_TOL);
  const to = Math.max(from, Math.ceil((offset + extent) / step - GRID_REF_TOL) - 1);
  return { from: from + 1, to: to + 1 };
};

export type PanelGridRefs = {
  /** id -> the rows and columns that panel stands in. */
  refs: Map<string, { rows: GridRef; cols: GridRef }>;
  /** How many grid cells the wall spans on each axis. */
  rows: number;
  cols: number;
};

export const panelGridRefs = (panels: PanelRecord[], rectOf: (p: PanelRecord) => RectMm): PanelGridRefs => {
  const active = panels.filter((p) => !p.isRemoved);
  const rects = active.map(rectOf);
  const refs = new Map<string, { rows: GridRef; cols: GridRef }>();
  if (!rects.length) return { refs, rows: 0, cols: 0 };
  const bbox = activeBBox(rects);
  const stepX = commonExtent(rects.map((r) => r.w));
  const stepY = commonExtent(rects.map((r) => r.h));
  let rows = 0;
  let cols = 0;
  active.forEach((panel, i) => {
    const r = rects[i];
    const ref = {
      cols: gridSpan(r.x - bbox.x, r.w, stepX),
      rows: gridSpan(r.y - bbox.y, r.h, stepY),
    };
    refs.set(panel.id, ref);
    rows = Math.max(rows, ref.rows.to);
    cols = Math.max(cols, ref.cols.to);
  });
  return { refs, rows, cols };
};

/** "3", or "3 & 4" for a panel straddling two, or "3-6" for a longer run. */
export const gridRefLabel = (ref: GridRef | undefined): string => {
  if (!ref) return "-";
  if (ref.to <= ref.from) return String(ref.from);
  if (ref.to === ref.from + 1) return `${ref.from} & ${ref.to}`;
  return `${ref.from}-${ref.to}`;
};

/** The same reference read from the other side of the wall. */
export const mirrorGridRef = (ref: GridRef | undefined, total: number): GridRef | undefined =>
  ref ? { from: total + 1 - ref.to, to: total + 1 - ref.from } : undefined;

export const bandPanels = (
  panels: PanelRecord[],
  rectOf: (p: PanelRecord) => RectMm,
): PanelRecord[][] => {
  const active = panels.filter((p) => !p.isRemoved);
  const entries = active
    .map((p) => ({ p, r: rectOf(p) }))
    .sort((a, b) => a.r.y + a.r.h / 2 - (b.r.y + b.r.h / 2));
  const bands: { centerY: number; items: { p: PanelRecord; r: RectMm }[] }[] = [];
  entries.forEach((e) => {
    const cy = e.r.y + e.r.h / 2;
    const band = bands.find((b) => Math.abs(b.centerY - cy) < HALF_MODULE_MM);
    if (band) {
      band.items.push(e);
      band.centerY = band.items.reduce((s, i) => s + i.r.y + i.r.h / 2, 0) / band.items.length;
    } else {
      bands.push({ centerY: cy, items: [e] });
    }
  });
  return bands.map((b) => b.items.sort((a, c) => a.r.x - c.r.x).map((i) => i.p));
};

/**
 * Column-banding for non-uniform layouts: group active panels into visual
 * columns by their horizontal centre (tolerance = half module), mirroring
 * bandPanels' row logic exactly. Bands are ordered left→right and panels
 * within a band top→bottom.
 *
 * This exists because panel WIDTH varies by panel type - MG9 is a 500mm
 * square module, but MT is 1000x500mm (twice as wide) - so a column index
 * computed by dividing raw x-position by the fixed 500mm module (as this
 * codebase's column-numbering used to do, in both App.tsx and
 * drawTestPattern.ts) silently counts every MT panel as 2 columns instead of
 * 1. Banding by actual panel adjacency, like rows already do, works
 * correctly for any panel width without that assumption.
 */
export const bandPanelsByColumn = (
  panels: PanelRecord[],
  rectOf: (p: PanelRecord) => RectMm,
): PanelRecord[][] => {
  const active = panels.filter((p) => !p.isRemoved);
  const entries = active
    .map((p) => ({ p, r: rectOf(p) }))
    .sort((a, b) => a.r.x + a.r.w / 2 - (b.r.x + b.r.w / 2));
  const bands: { centerX: number; items: { p: PanelRecord; r: RectMm }[] }[] = [];
  entries.forEach((e) => {
    const cx = e.r.x + e.r.w / 2;
    const band = bands.find((b) => Math.abs(b.centerX - cx) < HALF_MODULE_MM);
    if (band) {
      band.items.push(e);
      band.centerX = band.items.reduce((s, i) => s + i.r.x + i.r.w / 2, 0) / band.items.length;
    } else {
      bands.push({ centerX: cx, items: [e] });
    }
  });
  return bands.map((b) => b.items.sort((a, c) => a.r.y - c.r.y).map((i) => i.p));
};
