// ---------------------------------------------------------------------------
// Test Pattern Generator - data model and pure geometry.
//
// Everything in here is framework-free and DOM-free so it can be unit-tested
// in node: the config shape that is saved and reopened, the resolution
// presets, and every piece of arithmetic that moves a sub-screen (align,
// distribute, arrange, snap, nudge). All positions and sizes leave these
// functions as whole pixels - a test pattern that lands half a pixel off is
// the one thing it must never do.
// ---------------------------------------------------------------------------

export type Rect = { x: number; y: number; w: number; h: number };

export type PatternSettings = Record<string, number | string | boolean>;

/** Drawn inside the sub-screen's own pixels, so they travel with it into a single-screen export. */
export type OverlaySettings = {
  border: boolean;
  label: boolean;
  resolution: boolean;
  crosshair: boolean;
};

/** Resolution derived from an LED panel grid - kept on the screen so the LED pattern can draw the real panel grid. */
export type LedGrid = {
  cols: number;
  rows: number;
  panelPxW: number;
  panelPxH: number;
  /** Name of the panel preset used, purely for display. */
  panelName?: string;
};

/**
 * Panels imported from an LED Cabling Planner project. Kept as the raw saved
 * panel records (normalised by the planner's own code when drawn), so the
 * screen renders the planner's exact Moving Test Pattern: shaped, rotated
 * and gapped layouts included.
 */
export type LedLayoutSource = {
  projectName: string;
  surfaceName: string;
  panelType: string;
  panels: unknown[];
  subScreen: { id: string; name: string; color: string } | null;
};

export type SubScreen = {
  id: string;
  name: string;
  /** Top-left on the main canvas, whole pixels. */
  x: number;
  y: number;
  /** Native resolution, whole pixels. */
  w: number;
  h: number;
  pattern: string;
  /** Per-pattern settings, keyed by pattern id, so switching back restores what was set. */
  settings: Record<string, PatternSettings>;
  overlays: OverlaySettings;
  color: string;
  locked: boolean;
  visible: boolean;
  aspectLock: boolean;
  /** Seconds added to the animation clock, so two screens running the same pattern need not move in step. */
  phase: number;
  led: LedGrid | null;
  ledLayout: LedLayoutSource | null;
  /** NovaStar processor input (interface pk), when a processor is selected. */
  input: number | null;
};

export type ResolutionPreset = { name: string; w: number; h: number };

export type PlaylistStep = {
  id: string;
  name: string;
  /** Seconds this step is shown for. */
  duration: number;
  /** Pattern and settings per screen id. Screens not listed keep their own. */
  assignments: Record<string, { pattern: string; settings: PatternSettings }>;
};

export type SnapSettings = {
  enabled: boolean;
  canvasEdges: boolean;
  canvasCentre: boolean;
  screens: boolean;
  grid: boolean;
  gridSize: number;
  /** Snap distance in on-screen (preview) pixels. */
  threshold: number;
};

export type ProcessorSettings = {
  model: "" | "VX1000_PRO" | "VX2000_PRO";
  inputMode: "whole" | "perEntry";
  wholeInput: number | null;
};

export type GeneratorConfig = {
  formatVersion: 1;
  app: "test-pattern-generator";
  name: string;
  canvas: { w: number; h: number; background: string; aspectLock: boolean };
  /** Layer order: index 0 is drawn first (bottom). */
  screens: SubScreen[];
  customPresets: ResolutionPreset[];
  snap: SnapSettings;
  nudge: { step: number; shiftStep: number };
  /** One loop of every animated pattern, in seconds. Videos are exactly this long. */
  loopSeconds: number;
  playlist: PlaylistStep[];
  processor: ProcessorSettings;
};

export const CONFIG_FORMAT_VERSION = 1;
export const MAX_DIMENSION = 16384;

export const SCREEN_COLORS = ["#38bdf8", "#fb923c", "#a3e635", "#f472b6", "#c084fc", "#facc15", "#2dd4bf", "#fb7185"];

const LANDSCAPE: Array<[number, number]> = [
  [1280, 720],
  [1920, 1080],
  [1920, 1200],
  [2560, 1440],
  [3840, 2160],
  [4096, 2160],
];

/** Common presets, landscape then portrait. The LED planner's own wide canvases are kept after them. */
export const RESOLUTION_PRESETS: ResolutionPreset[] = [
  ...LANDSCAPE.map(([w, h]) => ({ name: `${w} × ${h}`, w, h })),
  ...LANDSCAPE.map(([w, h]) => ({ name: `${h} × ${w} (portrait)`, w: h, h: w })),
  { name: "3840 × 1080 (wide)", w: 3840, h: 1080 },
  { name: "7680 × 2160 (wide)", w: 7680, h: 2160 },
  { name: "7680 × 4320 (8K)", w: 7680, h: 4320 },
];

/** Pixel size of common LED panels, for working out a wall's resolution from rows and columns. */
export const LED_PANEL_PRESETS: Array<{ name: string; pxW: number; pxH: number }> = [
  { name: "MG9 (500 × 500mm, P2.9)", pxW: 168, pxH: 168 },
  { name: "MT (1000 × 500mm transparent)", pxW: 256, pxH: 64 },
  { name: "LED Poster section", pxW: 172, pxH: 258 },
  { name: "500 × 500mm P1.9", pxW: 256, pxH: 256 },
  { name: "500 × 500mm P2.6", pxW: 192, pxH: 192 },
  { name: "500 × 500mm P3.9", pxW: 128, pxH: 128 },
  { name: "500 × 1000mm P2.6", pxW: 192, pxH: 384 },
  { name: "500 × 1000mm P3.9", pxW: 128, pxH: 256 },
];

let idCounter = 0;
export const newId = (): string => {
  try {
    return crypto.randomUUID();
  } catch {
    idCounter += 1;
    return `id-${Date.now().toString(36)}-${idCounter}`;
  }
};

/** A whole number of pixels, at least `min`, at most MAX_DIMENSION. */
export const clampDimension = (value: number, min = 1): number => {
  if (!Number.isFinite(value)) return min;
  return Math.min(MAX_DIMENSION, Math.max(min, Math.round(value)));
};

export const toInt = (value: number): number => (Number.isFinite(value) ? Math.round(value) : 0);

export const gcd = (a: number, b: number): number => {
  let x = Math.abs(Math.round(a));
  let y = Math.abs(Math.round(b));
  while (y) [x, y] = [y, x % y];
  return x || 1;
};

/** "16:9", "9:16", "32:9"... reduced to whole numbers, with the decimal ratio for odd sizes. */
export const aspectLabel = (w: number, h: number): string => {
  if (w <= 0 || h <= 0) return "-";
  const d = gcd(w, h);
  const a = w / d;
  const b = h / d;
  if (a <= 64 && b <= 64) return `${a}:${b}`;
  return `${(w / h).toFixed(3)}:1`;
};

/**
 * New size after one dimension is edited with the aspect ratio locked. The
 * other side follows, rounded to a whole pixel - so a locked ratio is kept as
 * closely as whole pixels allow, never exactly at the cost of a fraction.
 */
export const resizeWithAspect = (
  current: { w: number; h: number },
  edited: "w" | "h",
  value: number,
  lock: boolean,
): { w: number; h: number } => {
  const v = clampDimension(value);
  if (!lock || current.w <= 0 || current.h <= 0) return edited === "w" ? { w: v, h: current.h } : { w: current.w, h: v };
  const ratio = current.w / current.h;
  return edited === "w" ? { w: v, h: clampDimension(v / ratio) } : { w: clampDimension(v * ratio), h: v };
};

export const defaultOverlays = (): OverlaySettings => ({ border: true, label: true, resolution: true, crosshair: false });

export const makeScreen = (index: number, partial: Partial<SubScreen> = {}): SubScreen => ({
  id: newId(),
  name: `Screen ${index + 1}`,
  x: 0,
  y: 0,
  w: 1920,
  h: 1080,
  pattern: "info",
  settings: {},
  overlays: defaultOverlays(),
  color: SCREEN_COLORS[index % SCREEN_COLORS.length],
  locked: false,
  visible: true,
  aspectLock: false,
  phase: 0,
  led: null,
  ledLayout: null,
  input: null,
  ...partial,
});

export const defaultSnap = (): SnapSettings => ({
  enabled: true,
  canvasEdges: true,
  canvasCentre: true,
  screens: true,
  grid: false,
  gridSize: 60,
  threshold: 8,
});

export const defaultConfig = (): GeneratorConfig => ({
  formatVersion: CONFIG_FORMAT_VERSION,
  app: "test-pattern-generator",
  name: "Untitled output",
  canvas: { w: 3840, h: 2160, background: "#000000", aspectLock: false },
  screens: [
    makeScreen(0, { name: "Screen 1", x: 0, y: 0, w: 1920, h: 1080, pattern: "smpte" }),
    makeScreen(1, { name: "Screen 2", x: 1920, y: 0, w: 1920, h: 1080, pattern: "grid" }),
    makeScreen(2, { name: "Screen 3", x: 0, y: 1080, w: 1920, h: 1080, pattern: "led-moving", led: { cols: 12, rows: 6, panelPxW: 160, panelPxH: 180, panelName: "Custom" } }),
    makeScreen(3, { name: "Screen 4", x: 1920, y: 1080, w: 1920, h: 1080, pattern: "faces" }),
  ],
  customPresets: [],
  snap: defaultSnap(),
  nudge: { step: 1, shiftStep: 10 },
  loopSeconds: 10,
  playlist: [],
  processor: { model: "", inputMode: "perEntry", wholeInput: null },
});

export const screenRect = (s: { x: number; y: number; w: number; h: number }): Rect => ({ x: s.x, y: s.y, w: s.w, h: s.h });

export const boundsOf = (rects: Rect[]): Rect => {
  if (!rects.length) return { x: 0, y: 0, w: 0, h: 0 };
  const x0 = Math.min(...rects.map((r) => r.x));
  const y0 = Math.min(...rects.map((r) => r.y));
  const x1 = Math.max(...rects.map((r) => r.x + r.w));
  const y1 = Math.max(...rects.map((r) => r.y + r.h));
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
};

export const rectsOverlap = (a: Rect, b: Rect): boolean => a.x < b.x + b.w && b.x < a.x + a.w && a.y < b.y + b.h && b.y < a.y + a.h;

// ---------------------------------------------------------------------------
// Alignment, distribution and arrangement. Each takes the rects to move and
// returns their new top-left corners in the same order, as whole pixels.
// ---------------------------------------------------------------------------

export type AlignMode = "left" | "centre" | "right" | "top" | "middle" | "bottom";

export const alignRects = (rects: Rect[], mode: AlignMode, ref: Rect): Array<{ x: number; y: number }> =>
  rects.map((r) => {
    switch (mode) {
      case "left":
        return { x: toInt(ref.x), y: r.y };
      case "right":
        return { x: toInt(ref.x + ref.w - r.w), y: r.y };
      case "centre":
        return { x: Math.round(ref.x + (ref.w - r.w) / 2), y: r.y };
      case "top":
        return { x: r.x, y: toInt(ref.y) };
      case "bottom":
        return { x: r.x, y: toInt(ref.y + ref.h - r.h) };
      case "middle":
        return { x: r.x, y: Math.round(ref.y + (ref.h - r.h) / 2) };
      default:
        return { x: r.x, y: r.y };
    }
  });

/**
 * Even spacing between the outermost rects, which stay where they are. The
 * free space rarely divides exactly, so the gaps are whole pixels that differ
 * by at most one - the spare pixels go to the first gaps - rather than
 * fractional positions.
 */
export const distributeRects = (rects: Rect[], axis: "horizontal" | "vertical"): Array<{ x: number; y: number }> => {
  const out = rects.map((r) => ({ x: r.x, y: r.y }));
  if (rects.length < 3) return out;
  const pos = (r: Rect) => (axis === "horizontal" ? r.x : r.y);
  const size = (r: Rect) => (axis === "horizontal" ? r.w : r.h);
  const order = rects.map((r, i) => i).sort((a, b) => pos(rects[a]) - pos(rects[b]) || a - b);
  const first = rects[order[0]];
  const last = rects[order[order.length - 1]];
  const span = pos(last) + size(last) - pos(first);
  const used = order.reduce((sum, i) => sum + size(rects[i]), 0);
  const gaps = order.length - 1;
  const free = span - used;
  const base = Math.floor(free / gaps);
  let spare = free - base * gaps;
  let cursor = pos(first);
  order.forEach((i, n) => {
    if (n === order.length - 1) return; // the last one stays put
    if (n > 0) {
      if (axis === "horizontal") out[i].x = cursor;
      else out[i].y = cursor;
    }
    cursor += size(rects[i]) + base + (spare > 0 ? 1 : 0);
    if (spare > 0) spare -= 1;
  });
  return out;
};

/** Distribute with a fixed gap: packs the rects in their current order starting from the first one. */
export const distributeWithGap = (rects: Rect[], axis: "horizontal" | "vertical", gap: number): Array<{ x: number; y: number }> => {
  const out = rects.map((r) => ({ x: r.x, y: r.y }));
  if (rects.length < 2) return out;
  const g = toInt(gap);
  const order = rects.map((r, i) => i).sort((a, b) => (axis === "horizontal" ? rects[a].x - rects[b].x : rects[a].y - rects[b].y) || a - b);
  let cursor = axis === "horizontal" ? rects[order[0]].x : rects[order[0]].y;
  order.forEach((i) => {
    if (axis === "horizontal") out[i].x = cursor;
    else out[i].y = cursor;
    cursor += (axis === "horizontal" ? rects[i].w : rects[i].h) + g;
  });
  return out;
};

export type ArrangeMode = "row" | "column" | "grid";

/**
 * Lays the rects out in reading order (top to bottom, then left to right, as
 * they currently sit) in a row, a column or a grid of `columns` columns.
 * Each grid column is as wide as its widest member and each row as tall as
 * its tallest, so mixed resolutions never overlap.
 */
export const arrangeRects = (
  rects: Rect[],
  mode: ArrangeMode,
  options: { columns: number; gapX: number; gapY: number; origin: { x: number; y: number } },
): Array<{ x: number; y: number }> => {
  const out = rects.map((r) => ({ x: r.x, y: r.y }));
  if (!rects.length) return out;
  const order = rects.map((r, i) => i).sort((a, b) => rects[a].y - rects[b].y || rects[a].x - rects[b].x || a - b);
  const cols = mode === "row" ? order.length : mode === "column" ? 1 : Math.max(1, Math.round(options.columns));
  const rowsCount = Math.ceil(order.length / cols);
  const colW = new Array(cols).fill(0);
  const rowH = new Array(rowsCount).fill(0);
  order.forEach((i, n) => {
    colW[n % cols] = Math.max(colW[n % cols], rects[i].w);
    rowH[Math.floor(n / cols)] = Math.max(rowH[Math.floor(n / cols)], rects[i].h);
  });
  const gapX = toInt(options.gapX);
  const gapY = toInt(options.gapY);
  const colX: number[] = [];
  colW.reduce((x, w, c) => {
    colX[c] = x;
    return x + w + gapX;
  }, toInt(options.origin.x));
  const rowY: number[] = [];
  rowH.reduce((y, h, r) => {
    rowY[r] = y;
    return y + h + gapY;
  }, toInt(options.origin.y));
  order.forEach((i, n) => {
    out[i] = { x: colX[n % cols], y: rowY[Math.floor(n / cols)] };
  });
  return out;
};

// ---------------------------------------------------------------------------
// Snapping. Targets are lines on the canvas (x positions for vertical lines,
// y positions for horizontal ones). A moving rect offers its left, centre and
// right edges (top, middle, bottom vertically); the closest pair inside the
// threshold wins, and its line is reported as a guide to draw.
// ---------------------------------------------------------------------------

export type SnapGuide = { axis: "x" | "y"; at: number; from: number; to: number };

export type SnapTargets = { xs: number[]; ys: number[] };

export const snapTargetsFor = (
  canvas: { w: number; h: number },
  others: Rect[],
  snap: SnapSettings,
): SnapTargets => {
  const xs: number[] = [];
  const ys: number[] = [];
  if (snap.canvasEdges) {
    xs.push(0, canvas.w);
    ys.push(0, canvas.h);
  }
  if (snap.canvasCentre) {
    xs.push(Math.round(canvas.w / 2));
    ys.push(Math.round(canvas.h / 2));
  }
  if (snap.screens) {
    others.forEach((r) => {
      xs.push(r.x, r.x + r.w, Math.round(r.x + r.w / 2));
      ys.push(r.y, r.y + r.h, Math.round(r.y + r.h / 2));
    });
  }
  return { xs, ys };
};

const nearestGrid = (value: number, size: number) => Math.round(value / size) * size;

/**
 * Snap a moving rect. `threshold` is in canvas pixels. Grid snapping applies
 * to the rect's top-left corner only (grid cells are where things start),
 * and only when no edge/centre target is closer.
 */
export const snapMove = (
  rect: Rect,
  targets: SnapTargets,
  threshold: number,
  grid: number | null,
): { x: number; y: number; guides: SnapGuide[] } => {
  const pick = (offers: number[], lines: number[]) => {
    let best: { delta: number; line: number } | null = null;
    offers.forEach((offer) => {
      lines.forEach((line) => {
        const delta = line - offer;
        if (Math.abs(delta) <= threshold && (!best || Math.abs(delta) < Math.abs(best.delta))) best = { delta, line };
      });
    });
    return best as { delta: number; line: number } | null;
  };
  const xHit = pick([rect.x, rect.x + rect.w / 2, rect.x + rect.w], targets.xs);
  const yHit = pick([rect.y, rect.y + rect.h / 2, rect.y + rect.h], targets.ys);
  let x = rect.x;
  let y = rect.y;
  const guides: SnapGuide[] = [];
  if (xHit) {
    x = rect.x + xHit.delta;
    guides.push({ axis: "x", at: xHit.line, from: 0, to: 0 });
  } else if (grid && grid > 0) {
    const g = nearestGrid(rect.x, grid);
    if (Math.abs(g - rect.x) <= threshold) x = g;
  }
  if (yHit) {
    y = rect.y + yHit.delta;
    guides.push({ axis: "y", at: yHit.line, from: 0, to: 0 });
  } else if (grid && grid > 0) {
    const g = nearestGrid(rect.y, grid);
    if (Math.abs(g - rect.y) <= threshold) y = g;
  }
  return { x: Math.round(x), y: Math.round(y), guides };
};

/** Snap a single edge being dragged during a resize. */
export const snapEdge = (value: number, lines: number[], threshold: number, grid: number | null): { value: number; guide: number | null } => {
  let best: number | null = null;
  lines.forEach((line) => {
    if (Math.abs(line - value) <= threshold && (best === null || Math.abs(line - value) < Math.abs(best - value))) best = line;
  });
  if (best !== null) return { value: Math.round(best), guide: best };
  if (grid && grid > 0) {
    const g = nearestGrid(value, grid);
    if (Math.abs(g - value) <= threshold) return { value: g, guide: null };
  }
  return { value: Math.round(value), guide: null };
};

export type ResizeHandle = "n" | "s" | "e" | "w" | "ne" | "nw" | "se" | "sw";

/**
 * A rect resized by dragging one handle by (dx, dy) canvas pixels, with an
 * optional locked aspect ratio. Always returns whole pixels and at least 1x1;
 * the opposite edge stays fixed.
 */
export const resizeRect = (start: Rect, handle: ResizeHandle, dx: number, dy: number, aspectLock: boolean): Rect => {
  let { x, y, w, h } = start;
  const right = start.x + start.w;
  const bottom = start.y + start.h;
  if (handle.includes("e")) w = start.w + dx;
  if (handle.includes("w")) w = start.w - dx;
  if (handle.includes("s")) h = start.h + dy;
  if (handle.includes("n")) h = start.h - dy;
  w = Math.max(1, Math.round(w));
  h = Math.max(1, Math.round(h));
  if (aspectLock && start.w > 0 && start.h > 0) {
    const ratio = start.w / start.h;
    if (handle === "n" || handle === "s") w = Math.max(1, Math.round(h * ratio));
    else if (handle === "e" || handle === "w") h = Math.max(1, Math.round(w / ratio));
    else if (Math.abs(w - start.w) / start.w >= Math.abs(h - start.h) / start.h) h = Math.max(1, Math.round(w / ratio));
    else w = Math.max(1, Math.round(h * ratio));
  }
  w = Math.min(MAX_DIMENSION, w);
  h = Math.min(MAX_DIMENSION, h);
  if (handle.includes("w")) x = right - w;
  if (handle.includes("n")) y = bottom - h;
  return { x: Math.round(x), y: Math.round(y), w, h };
};

// ---------------------------------------------------------------------------
// Layer order.
// ---------------------------------------------------------------------------

export type LayerMove = "forward" | "backward" | "front" | "back";

/** Moves the screens with the given ids, keeping their relative order. Index 0 is the bottom layer. */
export const reorderLayers = <T extends { id: string }>(items: T[], ids: Set<string>, move: LayerMove): T[] => {
  const moving = items.filter((s) => ids.has(s.id));
  const rest = items.filter((s) => !ids.has(s.id));
  if (!moving.length) return items;
  if (move === "front") return [...rest, ...moving];
  if (move === "back") return [...moving, ...rest];
  const out = [...items];
  if (move === "forward") {
    for (let i = out.length - 2; i >= 0; i -= 1) {
      if (ids.has(out[i].id) && !ids.has(out[i + 1].id)) [out[i], out[i + 1]] = [out[i + 1], out[i]];
    }
  } else {
    for (let i = 1; i < out.length; i += 1) {
      if (ids.has(out[i].id) && !ids.has(out[i - 1].id)) [out[i], out[i - 1]] = [out[i - 1], out[i]];
    }
  }
  return out;
};

// ---------------------------------------------------------------------------
// File names.
// ---------------------------------------------------------------------------

export const fileSafe = (text: string): string =>
  (text || "untitled")
    .trim()
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, "-")
    .replace(/\s+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "") || "untitled";

/** `<screen>_<pattern>_<w>x<h>.<ext>` - every exported file says what is on it and how big it is. */
export const exportFileName = (screenName: string, patternName: string, w: number, h: number, ext: string): string =>
  `${fileSafe(screenName)}_${fileSafe(patternName)}_${w}x${h}.${ext}`;

/** Names stay unique inside one ZIP even when two screens share a name. */
export const uniqueNames = (names: string[]): string[] => {
  const seen = new Map<string, number>();
  return names.map((name) => {
    const count = seen.get(name) ?? 0;
    seen.set(name, count + 1);
    if (!count) return name;
    const dot = name.lastIndexOf(".");
    return dot > 0 ? `${name.slice(0, dot)}-${count + 1}${name.slice(dot)}` : `${name}-${count + 1}`;
  });
};

// ---------------------------------------------------------------------------
// Saving and reopening. Anything missing or malformed in a file falls back to
// a default rather than failing the whole load.
// ---------------------------------------------------------------------------

const HEX = /^#[0-9a-fA-F]{6}$/;
const num = (v: unknown, fallback: number) => (typeof v === "number" && Number.isFinite(v) ? v : Number.isFinite(Number(v)) && v !== "" && v !== null && v !== undefined ? Number(v) : fallback);
const bool = (v: unknown, fallback: boolean) => (typeof v === "boolean" ? v : fallback);
const str = (v: unknown, fallback: string) => (typeof v === "string" ? v : fallback);
const color = (v: unknown, fallback: string) => (typeof v === "string" && HEX.test(v) ? v.toLowerCase() : fallback);

const normalizeSettings = (raw: unknown): PatternSettings => {
  if (!raw || typeof raw !== "object") return {};
  const out: PatternSettings = {};
  Object.entries(raw as Record<string, unknown>).forEach(([k, v]) => {
    if (typeof v === "number" || typeof v === "string" || typeof v === "boolean") out[k] = v;
  });
  return out;
};

const normalizeLed = (raw: unknown): LedGrid | null => {
  if (!raw || typeof raw !== "object") return null;
  const r = raw as Record<string, unknown>;
  return {
    cols: clampDimension(num(r.cols, 1)),
    rows: clampDimension(num(r.rows, 1)),
    panelPxW: clampDimension(num(r.panelPxW, 168)),
    panelPxH: clampDimension(num(r.panelPxH, 168)),
    panelName: typeof r.panelName === "string" ? r.panelName : undefined,
  };
};

const normalizeLedLayout = (raw: unknown): LedLayoutSource | null => {
  if (!raw || typeof raw !== "object") return null;
  const r = raw as Record<string, unknown>;
  if (!Array.isArray(r.panels) || !r.panels.length) return null;
  const sub = r.subScreen as Record<string, unknown> | null | undefined;
  return {
    projectName: str(r.projectName, ""),
    surfaceName: str(r.surfaceName, ""),
    panelType: str(r.panelType, "MG9"),
    panels: r.panels,
    subScreen: sub && typeof sub.id === "string" ? { id: sub.id, name: str(sub.name, ""), color: color(sub.color, "#ffffff") } : null,
  };
};

export const normalizeScreen = (raw: unknown, index: number): SubScreen => {
  const r = (raw && typeof raw === "object" ? raw : {}) as Record<string, unknown>;
  const settings: Record<string, PatternSettings> = {};
  if (r.settings && typeof r.settings === "object") {
    Object.entries(r.settings as Record<string, unknown>).forEach(([k, v]) => {
      settings[k] = normalizeSettings(v);
    });
  }
  const ov = (r.overlays && typeof r.overlays === "object" ? r.overlays : {}) as Record<string, unknown>;
  const d = defaultOverlays();
  return {
    id: typeof r.id === "string" && r.id ? r.id : newId(),
    name: str(r.name, `Screen ${index + 1}`),
    x: toInt(num(r.x, 0)),
    y: toInt(num(r.y, 0)),
    w: clampDimension(num(r.w, 1920)),
    h: clampDimension(num(r.h, 1080)),
    pattern: str(r.pattern, "info"),
    settings,
    overlays: {
      border: bool(ov.border, d.border),
      label: bool(ov.label, d.label),
      resolution: bool(ov.resolution, d.resolution),
      crosshair: bool(ov.crosshair, d.crosshair),
    },
    color: color(r.color, SCREEN_COLORS[index % SCREEN_COLORS.length]),
    locked: bool(r.locked, false),
    visible: bool(r.visible, true),
    aspectLock: bool(r.aspectLock, false),
    phase: num(r.phase, 0),
    led: normalizeLed(r.led),
    ledLayout: normalizeLedLayout(r.ledLayout),
    input: typeof r.input === "number" ? r.input : null,
  };
};

export const normalizePresets = (raw: unknown): ResolutionPreset[] =>
  Array.isArray(raw)
    ? raw.flatMap((p) => {
        const r = p as Record<string, unknown>;
        if (!r || typeof r.name !== "string") return [];
        return [{ name: r.name, w: clampDimension(num(r.w, 1920)), h: clampDimension(num(r.h, 1080)) }];
      })
    : [];

export const normalizeConfig = (raw: unknown): GeneratorConfig => {
  const base = defaultConfig();
  if (!raw || typeof raw !== "object") return base;
  const r = raw as Record<string, unknown>;
  const canvas = (r.canvas && typeof r.canvas === "object" ? r.canvas : {}) as Record<string, unknown>;
  const snap = (r.snap && typeof r.snap === "object" ? r.snap : {}) as Record<string, unknown>;
  const nudge = (r.nudge && typeof r.nudge === "object" ? r.nudge : {}) as Record<string, unknown>;
  const proc = (r.processor && typeof r.processor === "object" ? r.processor : {}) as Record<string, unknown>;
  const screens = Array.isArray(r.screens) ? r.screens.map(normalizeScreen) : [];
  const screenIds = new Set(screens.map((s) => s.id));
  const playlist: PlaylistStep[] = Array.isArray(r.playlist)
    ? r.playlist.flatMap((p, i) => {
        const step = p as Record<string, unknown>;
        if (!step || typeof step !== "object") return [];
        const assignments: PlaylistStep["assignments"] = {};
        if (step.assignments && typeof step.assignments === "object") {
          Object.entries(step.assignments as Record<string, unknown>).forEach(([id, a]) => {
            const entry = a as Record<string, unknown>;
            if (!screenIds.has(id) || !entry || typeof entry.pattern !== "string") return;
            assignments[id] = { pattern: entry.pattern, settings: normalizeSettings(entry.settings) };
          });
        }
        return [{ id: str(step.id, newId()) || newId(), name: str(step.name, `Step ${i + 1}`), duration: Math.max(1, num(step.duration, 10)), assignments }];
      })
    : [];
  const ds = defaultSnap();
  return {
    formatVersion: CONFIG_FORMAT_VERSION,
    app: "test-pattern-generator",
    name: str(r.name, base.name),
    canvas: {
      w: clampDimension(num(canvas.w, base.canvas.w)),
      h: clampDimension(num(canvas.h, base.canvas.h)),
      background: color(canvas.background, "#000000"),
      aspectLock: bool(canvas.aspectLock, false),
    },
    screens,
    customPresets: normalizePresets(r.customPresets),
    snap: {
      enabled: bool(snap.enabled, ds.enabled),
      canvasEdges: bool(snap.canvasEdges, ds.canvasEdges),
      canvasCentre: bool(snap.canvasCentre, ds.canvasCentre),
      screens: bool(snap.screens, ds.screens),
      grid: bool(snap.grid, ds.grid),
      gridSize: clampDimension(num(snap.gridSize, ds.gridSize)),
      threshold: clampDimension(num(snap.threshold, ds.threshold)),
    },
    nudge: { step: clampDimension(num(nudge.step, 1)), shiftStep: clampDimension(num(nudge.shiftStep, 10)) },
    loopSeconds: Math.min(600, Math.max(1, num(r.loopSeconds, base.loopSeconds))),
    playlist,
    processor: {
      model: proc.model === "VX1000_PRO" || proc.model === "VX2000_PRO" ? proc.model : "",
      inputMode: proc.inputMode === "whole" ? "whole" : "perEntry",
      wholeInput: typeof proc.wholeInput === "number" ? proc.wholeInput : null,
    },
  };
};

// ---------------------------------------------------------------------------
// Warnings - the same checks the LED planner's Output Canvas makes.
// ---------------------------------------------------------------------------

export const layoutWarnings = (config: GeneratorConfig): string[] => {
  const { w: W, h: H } = config.canvas;
  const visible = config.screens.filter((s) => s.visible);
  const warnings: string[] = [];
  visible.forEach((s) => {
    if (s.x < 0 || s.y < 0) warnings.push(`${s.name}: negative canvas position (X ${s.x}, Y ${s.y}).`);
    if (s.x + s.w > W || s.y + s.h > H) warnings.push(`${s.name}: extends beyond the ${W}×${H} canvas.`);
  });
  for (let i = 0; i < visible.length; i += 1) {
    for (let j = i + 1; j < visible.length; j += 1) {
      if (rectsOverlap(visible[i], visible[j])) warnings.push(`${visible[i].name} and ${visible[j].name} overlap (${visible[j].name} is on top).`);
    }
  }
  if (visible.length > 1) {
    const b = boundsOf(visible.map(screenRect));
    if (b.w > W || b.h > H) warnings.push("The complete layout exceeds the configured canvas resolution.");
  }
  return warnings;
};

// ---------------------------------------------------------------------------
// Playlists. Given the time since a playlist started, which step is showing,
// and the config as that step makes it. Pure, so the editor and the output
// window derive exactly the same picture from the same clock.
// ---------------------------------------------------------------------------

export const playlistStepAt = (playlist: PlaylistStep[], elapsedSeconds: number, loop: boolean): { index: number; stepElapsed: number } | null => {
  if (!playlist.length) return null;
  const total = playlist.reduce((sum, s) => sum + Math.max(1, s.duration), 0);
  let t = elapsedSeconds;
  if (loop) t = ((t % total) + total) % total;
  else if (t >= total) return { index: playlist.length - 1, stepElapsed: Math.max(1, playlist[playlist.length - 1].duration) };
  for (let i = 0; i < playlist.length; i += 1) {
    const d = Math.max(1, playlist[i].duration);
    if (t < d) return { index: i, stepElapsed: t };
    t -= d;
  }
  return { index: playlist.length - 1, stepElapsed: 0 };
};

export const applyPlaylistStep = (config: GeneratorConfig, step: PlaylistStep | null): GeneratorConfig => {
  if (!step) return config;
  return {
    ...config,
    screens: config.screens.map((s) => {
      const a = step.assignments[s.id];
      if (!a) return s;
      return { ...s, pattern: a.pattern, settings: { ...s.settings, [a.pattern]: { ...(s.settings[a.pattern] ?? {}), ...a.settings } } };
    }),
  };
};
