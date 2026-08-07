// Pure helpers for the Output Canvas Positioning feature. Deliberately kept
// separate from the mm-based physical layout: everything here only READS
// panel mm positions to derive a non-persisted canvas-pixel offset, it never
// writes back into panel x/y. Physical layout and canvas position are, and
// must stay, two separate coordinate systems (see App.tsx SubScreen/Cell
// comments).
import { type RectMm, bandPanels } from "../model/panels";
import { type Cell, PANEL_TYPES, cellPanelType, cellRect } from "../App";

export type CanvasResolution = { w: number; h: number };

/**
 * Pixel resolution of an already-filtered panel list, using the same
 * per-band pixel-summation algorithm as the whole-layout wallPixels
 * calculation in App.tsx (width = widest band's pixels, height = sum of each
 * band's tallest panel) - reused rather than reinvented so the two stay
 * consistent for mixed MG9/MT layouts.
 */
export const resolutionOf = (panels: Cell[]): CanvasResolution => {
  const active = panels.filter((cell) => !cell.isRemoved);
  const bands = bandPanels(active, cellRect) as Cell[][];
  let pixelW = 0;
  let pixelH = 0;
  bands.forEach((band) => {
    let rowPixelW = 0;
    let rowPixelH = 0;
    band.forEach((cell) => {
      const p = PANEL_TYPES[cellPanelType(cell)];
      rowPixelW += p.pixW;
      rowPixelH = Math.max(rowPixelH, p.pixH);
    });
    pixelW = Math.max(pixelW, rowPixelW);
    pixelH += rowPixelH;
  });
  return { w: pixelW, h: pixelH };
};

/** A sub-screen's own pixel resolution, derived from its member panels. */
export const subScreenResolutionOf = (panels: Cell[], subScreenId: string): CanvasResolution =>
  resolutionOf(panels.filter((cell) => cell.subScreenId === subScreenId));

/**
 * Each panel's tightly-packed pixel position within the wall's own native
 * resolution (see resolutionOf above - same per-band packing algorithm,
 * returning per-panel positions instead of just the aggregate W/H).
 *
 * This is the coordinate space a NovaStar processor's own LED-screen canvas
 * uses for cabinet topology - confirmed by comparing generated exports
 * against real NovaStar-software-generated project files: the processor's
 * pixel-addressing canvas is sized and positioned purely from real (patched)
 * panel pixels, with no such thing as an "empty gap pixel" for a
 * missing/unpatched panel, and is entirely independent of the output-canvas/
 * sub-screen placement used for on-screen positioning elsewhere in this app.
 * Deliberately NOT the same function as the test pattern's own panel
 * placement (drawTestPattern.ts), which intentionally preserves physical
 * gaps as empty space for a technician's visual reference - both are
 * correct, for two different purposes.
 */
export const wallNativePositionsOf = (panels: Cell[]): Map<string, RectMm> => {
  const active = panels.filter((cell) => !cell.isRemoved);
  const bands = bandPanels(active, cellRect) as Cell[][];
  const positions = new Map<string, RectMm>();
  let bandY = 0;
  bands.forEach((band) => {
    let x = 0;
    let bandH = 0;
    band.forEach((cell) => {
      const spec = PANEL_TYPES[cellPanelType(cell)];
      positions.set(cell.id, { x, y: bandY, w: spec.pixW, h: spec.pixH });
      x += spec.pixW;
      bandH = Math.max(bandH, spec.pixH);
    });
    bandY += bandH;
  });
  return positions;
};

/**
 * A panel's final position on the output canvas: the sub-screen's own canvas
 * X/Y plus the panel's position within the sub-screen, converted from mm to
 * canvas pixels using THAT PANEL'S OWN native mm->px ratio (MG9 and MT have
 * different pitches, so a single blended sub-screen-wide ratio would be
 * wrong) - the same ratio the front-view PNG export already uses.
 */
export const finalCanvasPositionOf = (
  cell: Cell,
  subScreenBBox: RectMm,
  canvasX: number,
  canvasY: number,
): { x: number; y: number } => {
  const rect = cellRect(cell);
  const spec = PANEL_TYPES[cellPanelType(cell)];
  const pxPerMm = spec.pixW / (spec.w * 1000);
  const localOffsetXMm = rect.x - subScreenBBox.x;
  const localOffsetYMm = rect.y - subScreenBBox.y;
  return {
    x: canvasX + Math.round(localOffsetXMm * pxPerMm),
    y: canvasY + Math.round(localOffsetYMm * pxPerMm),
  };
};
