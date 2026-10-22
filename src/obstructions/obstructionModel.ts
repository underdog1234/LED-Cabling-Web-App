// Pure, framework-free helpers for Obstructions: physical things standing in
// front of the LED (a speaker stack, a truss leg, a pillar) that mask part of
// the picture. They are NOT panels - they never affect panel counts, power,
// weight, patching or resolution. They exist only to show, on every test
// pattern and on the Basic LED Map, which part of the content will be hidden.
//
// Stored in the SAME workspace-mm coordinate space as Cell.x/y (see
// App.tsx's Cell comments), so an obstruction's position is meaningful
// against the panels regardless of zoom, front/back view or output canvas.
// Entered and shown in metres, because that is how a site is measured.

export type Obstruction = {
  id: string;
  name: string;
  /** Workspace mm, top-left origin - the same space as Cell.x / Cell.y. */
  xMm: number;
  yMm: number;
  wMm: number;
  hMm: number;
};

let obstructionIdCounter = 0;
export const newObstructionId = (): string => {
  try {
    return crypto.randomUUID();
  } catch {
    obstructionIdCounter += 1;
    return `obs-${Date.now().toString(36)}-${obstructionIdCounter}`;
  }
};

/** Smallest obstruction we allow, in mm - below this it cannot be grabbed or read. */
export const MIN_OBSTRUCTION_MM = 50;

export const makeObstruction = (
  name: string,
  xMm: number,
  yMm: number,
  wMm = 1000,
  hMm = 1000,
): Obstruction => ({
  id: newObstructionId(),
  name,
  xMm,
  yMm,
  wMm: Math.max(MIN_OBSTRUCTION_MM, wMm),
  hMm: Math.max(MIN_OBSTRUCTION_MM, hMm),
});

const finiteOr = (raw: unknown, fallback: number): number => {
  const n = typeof raw === "string" ? Number.parseFloat(raw) : raw;
  return typeof n === "number" && Number.isFinite(n) ? n : fallback;
};

/**
 * Accepts whatever a saved project file holds and returns usable
 * obstructions - old files have no obstructions at all, and a hand-edited
 * one could hold anything. A bad entry is repaired, never thrown away,
 * so a typo in a file can't silently drop an obstruction someone relies on.
 */
export const normalizeObstructions = (raw: unknown): Obstruction[] => {
  if (!Array.isArray(raw)) return [];
  return raw.map((entry, index) => {
    const row = (entry ?? {}) as Record<string, unknown>;
    return {
      id: typeof row.id === "string" && row.id ? row.id : newObstructionId(),
      name: typeof row.name === "string" && row.name.trim() ? row.name.trim() : `Obstruction ${index + 1}`,
      xMm: finiteOr(row.xMm, 0),
      yMm: finiteOr(row.yMm, 0),
      wMm: Math.max(MIN_OBSTRUCTION_MM, finiteOr(row.wMm, 1000)),
      hMm: Math.max(MIN_OBSTRUCTION_MM, finiteOr(row.hMm, 1000)),
    };
  });
};

/** mm -> metres for display, always 2dp: site measurements, not pixel maths. */
export const metresLabel = (mm: number): string => (mm / 1000).toFixed(2);

/** "2.40 x 1.20 m" - the physical size, as it would be measured on site. */
export const obstructionMetresLabel = (obstruction: Obstruction): string =>
  `${metresLabel(obstruction.wMm)} x ${metresLabel(obstruction.hMm)} m`;

/**
 * The obstruction's footprint in an LED surface's own pixel space.
 *
 * `wallBBoxMm` is that surface's physical bounding box and `W`/`H` its native
 * pixel resolution, so the ratio is taken from the surface actually being
 * drawn - the whole wall, or one sub-screen, which have different origins and
 * different pixel sizes. Returns back-view (unmirrored) pixels; callers that
 * render front view mirror it themselves, exactly as they do for panels.
 */
export const obstructionPixelRect = (
  obstruction: Obstruction,
  wallBBoxMm: { x: number; y: number; w: number; h: number },
  W: number,
  H: number,
): { x: number; y: number; w: number; h: number } => {
  const pxPerMmX = wallBBoxMm.w > 0 ? W / wallBBoxMm.w : 0;
  const pxPerMmY = wallBBoxMm.h > 0 ? H / wallBBoxMm.h : 0;
  return {
    x: Math.round((obstruction.xMm - wallBBoxMm.x) * pxPerMmX),
    y: Math.round((obstruction.yMm - wallBBoxMm.y) * pxPerMmY),
    w: Math.max(1, Math.round(obstruction.wMm * pxPerMmX)),
    h: Math.max(1, Math.round(obstruction.hMm * pxPerMmY)),
  };
};

/** True when any part of the obstruction falls over the surface's pixel area. */
export const obstructionIntersectsSurface = (
  pixelRect: { x: number; y: number; w: number; h: number },
  W: number,
  H: number,
): boolean => pixelRect.x < W && pixelRect.y < H && pixelRect.x + pixelRect.w > 0 && pixelRect.y + pixelRect.h > 0;
