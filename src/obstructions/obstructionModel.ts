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
  /**
   * Position in mm, measured as an OFFSET FROM THE LAYOUT, not as an absolute
   * workspace coordinate: 0, 0 puts the obstruction horizontally centred on
   * the wall with its bottom edge level with the wall's bottom edge - a
   * speaker stack on the floor in the middle of the screen, which is where
   * one usually starts.
   *
   * It is relative because the wall moves. A panel added on the left shifts
   * every absolute coordinate in the layout; an obstruction standing in front
   * of the middle of the screen should still be in front of the middle of the
   * screen afterwards. Converted to absolute mm at the point of use by
   * obstructionRect (see below).
   *
   * xMm runs right-positive, yMm runs down-positive, matching Cell.x / Cell.y
   * and every other measurement in the app.
   */
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

/**
 * Where an obstruction actually sits, in absolute workspace mm.
 *
 * The anchor is the BOTTOM CENTRE of the layout: x 0 centres it on the wall,
 * y 0 stands it on the wall's bottom edge. Everything that draws or hit-tests
 * an obstruction goes through here, so the stored offset only ever has one
 * meaning.
 */
export const obstructionRect = (
  obstruction: Obstruction,
  layoutBBoxMm: { x: number; y: number; w: number; h: number },
): { x: number; y: number; w: number; h: number } => ({
  x: layoutBBoxMm.x + layoutBBoxMm.w / 2 - obstruction.wMm / 2 + obstruction.xMm,
  y: layoutBBoxMm.y + layoutBBoxMm.h - obstruction.hMm + obstruction.yMm,
  w: obstruction.wMm,
  h: obstruction.hMm,
});

/**
 * The offset that would put an obstruction at a given absolute position -
 * the inverse of obstructionRect, used when a drag finishes.
 */
export const obstructionOffsetFor = (
  absolute: { x: number; y: number },
  size: { wMm: number; hMm: number },
  layoutBBoxMm: { x: number; y: number; w: number; h: number },
): { xMm: number; yMm: number } => ({
  xMm: absolute.x - (layoutBBoxMm.x + layoutBBoxMm.w / 2 - size.wMm / 2),
  yMm: absolute.y - (layoutBBoxMm.y + layoutBBoxMm.h - size.hMm),
});

/**
 * Converts obstructions saved BEFORE the offsets became relative (project
 * files up to formatVersion 7, where x/y were absolute workspace mm) into the
 * relative form. Without this an old file's obstructions would jump, because
 * the same numbers now mean something else.
 */
export const migrateAbsoluteObstructions = (
  obstructions: Obstruction[],
  layoutBBoxMm: { x: number; y: number; w: number; h: number },
): Obstruction[] =>
  obstructions.map((obstruction) => ({
    ...obstruction,
    ...obstructionOffsetFor({ x: obstruction.xMm, y: obstruction.yMm }, obstruction, layoutBBoxMm),
  }));

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
  /**
   * The layout the obstruction's offset is measured against - the WHOLE wall.
   * It is a separate argument from wallBBoxMm because a sub-screen renders at
   * its own bbox and resolution, but an obstruction is still positioned
   * against the whole layout: it does not move because you opened one screen.
   * Defaults to the surface itself, which is the whole wall in the common
   * single-surface case.
   */
  layoutBBoxMm: { x: number; y: number; w: number; h: number } = wallBBoxMm,
): { x: number; y: number; w: number; h: number } => {
  const pxPerMmX = wallBBoxMm.w > 0 ? W / wallBBoxMm.w : 0;
  const pxPerMmY = wallBBoxMm.h > 0 ? H / wallBBoxMm.h : 0;
  const absolute = obstructionRect(obstruction, layoutBBoxMm);
  return {
    x: Math.round((absolute.x - wallBBoxMm.x) * pxPerMmX),
    y: Math.round((absolute.y - wallBBoxMm.y) * pxPerMmY),
    w: Math.max(1, Math.round(absolute.w * pxPerMmX)),
    h: Math.max(1, Math.round(absolute.h * pxPerMmY)),
  };
};

/** True when any part of the obstruction falls over the surface's pixel area. */
export const obstructionIntersectsSurface = (
  pixelRect: { x: number; y: number; w: number; h: number },
  W: number,
  H: number,
): boolean => pixelRect.x < W && pixelRect.y < H && pixelRect.x + pixelRect.w > 0 && pixelRect.y + pixelRect.h > 0;
