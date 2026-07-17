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

// Anchor offsets along an edge, as fractions of the half extent. Every straight
// edge uses the same pattern so flush edges (rect-rect, rect-to-shaped-leg) share
// exact anchor positions and join cleanly.
const EDGE_ANCHORS = [-0.8, -0.4, 0, 0.4, 0.8];
const ANCHOR_JOIN_TOL = 4; // mm - anchors this close count as coincident

// Local (centre-relative, unrotated) anchor points for a panel shape. Base
// orientations match the layout tool: triangle right-angle bottom-left (legs =
// left + bottom edges); quarter-circle right-angle bottom-right (legs = bottom
// + right edges); rect/corner = all four edges. Shaped panels expose NO anchors
// on their hypotenuse / curve, so those edges can never join.
const localAnchors = (shape: PanelShape, halfW: number, halfH: number): Array<{ x: number; y: number }> => {
  const pts: Array<{ x: number; y: number }> = [];
  const left = () => EDGE_ANCHORS.forEach((t) => pts.push({ x: -halfW, y: t * halfH }));
  const right = () => EDGE_ANCHORS.forEach((t) => pts.push({ x: halfW, y: t * halfH }));
  const top = () => EDGE_ANCHORS.forEach((t) => pts.push({ x: t * halfW, y: -halfH }));
  const bottom = () => EDGE_ANCHORS.forEach((t) => pts.push({ x: t * halfW, y: halfH }));
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
  return pts;
};

/** World-space connector anchors for a panel, rotated by its rotation. */
export const panelWorldAnchors = (g: PanelAnchorSpec): Array<{ x: number; y: number }> => {
  const rad = (((Math.round(g.rotation / 90) * 90) % 360) * Math.PI) / 180;
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
