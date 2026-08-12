// ---------------------------------------------------------------------------
// Animated LED wall test pattern - shared renderer.
//
// A single pure function (`drawTestPatternFrame`) draws one frame of the
// animation onto any canvas context. It is called identically by the live
// requestAnimationFrame loop in TestPatternView.tsx and by the WebM recorder's
// frame loop, so the live tab and the downloaded video are pixel-identical.
//
// Design: the RGB checkerboard and the greyscale gradient are defined as
// repeating patterns in WALL pixel-coordinate space (not per-panel), then
// revealed through each panel's own clip mask (true shape + rotation + the
// always-on front-view mirror). This is what makes uniform grids, freely
// placed/imported layouts, mixed MG9/MT and rotated/shaped panels all work
// through one code path with no special-casing.
// ---------------------------------------------------------------------------

import {
  type Cell,
  type PanelTypeKey,
  PANEL_TYPES,
  PANEL_VARIANTS,
  cellRect,
  cellPanelType,
  isPanelHead,
  applyPanelFrame,
  tracePanelShapePath,
} from "../App";
import { activeBBox, bandPanels, MODULE_MM, type RectMm } from "../model/panels";

export type TestPatternProject = {
  projectName: string;
  /** Name of this LED surface/sub-screen, if the project defines one. */
  surfaceName?: string;
  panelType: PanelTypeKey;
  panels: Cell[];
};

export type TestPatternLayout = {
  activePanels: Cell[];
  wallBBox: RectMm;
  W: number;
  H: number;
  /** Recommended Content Resolution height - equals H unless every active
   * panel is MT (a transparent panel missing every second LED row, so its
   * vertical pixel pitch is twice its horizontal pitch and H alone doesn't
   * represent the wall's true physical aspect ratio) - see drawInfoText. */
  contentPixelH: number;
  /** Each panel's TRUE native-resolution pixel rect (back view, unmirrored),
   * keyed by cell id - see computeTestPatternLayout. */
  panelPixelRects: Map<string, RectMm>;
  totalPanels: number;
  activeColsCount: number;
  activeRowsCount: number;
  wallWidthM: number;
  wallHeightM: number;
  projectName: string;
  surfaceName: string;
  /** RGB checkerboard cell size, in canvas px - one full panel footprint for
   * the project's panel type (see buildRgbTile). */
  tileWidthPx: number;
  tileHeightPx: number;
  rowLabel: (cell: Cell) => number;
  colLabel: (cell: Cell) => string;
};

// Pixel rect in canvas/output space (post front-view mirror). Snap both edges
// to the integer pixel grid - not just the width/height - so adjacent panels'
// rounded edges stay consistent with each other (no accumulating gaps) and a
// 1px stroke along that edge lands crisply on one pixel row/column instead of
// straddling two.
const snapRectPx = (r: RectMm): RectMm => {
  const x0 = Math.round(r.x);
  const y0 = Math.round(r.y);
  const x1 = Math.round(r.x + r.w);
  const y1 = Math.round(r.y + r.h);
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
};

// One full loop of the animation, in seconds. Both the RGB slide and the
// greyscale gradient return to their exact starting phase at this mark, so a
// video recorded for exactly this long loops seamlessly on repeat.
export const LOOP_SECONDS = 20;
// Cap the animation's internal draw rate - this is an alignment-check tool,
// not a smooth video, and a lower rate keeps hundreds of panels responsive.
export const DRAW_FPS = 24;

// RGB tile = one full panel footprint (of the project's own panel type) at
// the wall's native pixel pitch, so checkerboard cells land on real panel
// boundaries and each panel reads as a single solid colour - MG9 is a 500mm
// square module; MT is 1000x500mm (twice as wide), so its tile is twice as
// wide too (see computeTestPatternLayout's tileWidthPx/tileHeightPx).
const RGB_COLORS = ["#ff0000", "#00ff00", "#0000ff"];

let cachedRgbTile: { tileW: number; tileH: number; canvas: HTMLCanvasElement } | null = null;
let cachedGreyTile: { period: number; canvas: HTMLCanvasElement } | null = null;
let cachedPatternLayer: { w: number; h: number; canvas: HTMLCanvasElement } | null = null;

// Staggered diagonal R/G/B supercell: row R, col C -> colour (R+C) mod 3. This
// makes adjacent panels differ AND makes each row offset from the one above,
// and tiles seamlessly in both axes because the pattern's own period matches
// the tile size - no per-panel logic needed. Tile cells are `tileW x tileH`
// (one full panel footprint), not necessarily square - MT panels are twice
// as wide as they are tall.
const buildRgbTile = (tileW: number, tileH: number): HTMLCanvasElement => {
  const c = document.createElement("canvas");
  c.width = tileW * 3;
  c.height = tileH * 3;
  const ctx = c.getContext("2d")!;
  for (let row = 0; row < 3; row += 1) {
    for (let col = 0; col < 3; col += 1) {
      ctx.fillStyle = RGB_COLORS[(row + col) % 3];
      ctx.fillRect(col * tileW, row * tileH, tileW, tileH);
    }
  }
  return c;
};

const getRgbTile = (tileW: number, tileH: number): HTMLCanvasElement => {
  if (cachedRgbTile && cachedRgbTile.tileW === tileW && cachedRgbTile.tileH === tileH) return cachedRgbTile.canvas;
  const canvas = buildRgbTile(tileW, tileH);
  cachedRgbTile = { tileW, tileH, canvas };
  return canvas;
};

// Seamless diagonal greyscale sweep: a single band travels corner-to-corner
// across the WHOLE wall (period = the canvas diagonal, not a small repeating
// cell), so only one bright band is ever visible at a time, at wall scale.
// Brightness is a function of (x+y) mod `period` folded into a
// black->white->black triangle wave, which is exactly what a diagonal linear
// gradient from (0,0) to (period,period) with alternating floor/peak stops
// produces - periodic in (x+y), so it still repeats with no seam as it slides
// through a full loop.
const buildGreyTile = (period: number): HTMLCanvasElement => {
  const c = document.createElement("canvas");
  c.width = period;
  c.height = period;
  const ctx = c.getContext("2d")!;
  // Floor of #404040 (not black): with a 'multiply' blend, a black stop would
  // briefly extinguish a panel's colour entirely. A dim floor keeps the
  // brightness sweep clearly visible while every panel stays identifiable.
  const grad = ctx.createLinearGradient(0, 0, period, period);
  grad.addColorStop(0, "#404040");
  grad.addColorStop(0.25, "#ffffff");
  grad.addColorStop(0.5, "#404040");
  grad.addColorStop(0.75, "#ffffff");
  grad.addColorStop(1, "#404040");
  ctx.fillStyle = grad;
  ctx.fillRect(0, 0, period, period);
  return c;
};

const getGreyTile = (period: number): HTMLCanvasElement => {
  if (cachedGreyTile && cachedGreyTile.period === period) return cachedGreyTile.canvas;
  const canvas = buildGreyTile(period);
  cachedGreyTile = { period, canvas };
  return canvas;
};

const getPatternLayer = (w: number, h: number): HTMLCanvasElement => {
  if (cachedPatternLayer && cachedPatternLayer.w === w && cachedPatternLayer.h === h) return cachedPatternLayer.canvas;
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, w);
  canvas.height = Math.max(1, h);
  cachedPatternLayer = { w, h, canvas };
  return canvas;
};

// Compute the wall geometry once (on project load), not per animation frame.
export const computeTestPatternLayout = (project: TestPatternProject): TestPatternLayout => {
  const activePanels = project.panels.filter((cell) => isPanelHead(cell));
  const wallBBox = activeBBox(activePanels.map(cellRect));
  const panelBands = bandPanels(activePanels, cellRect) as Cell[][];
  const bandIndexById = new Map<string, number>();
  panelBands.forEach((band, index) => band.forEach((cell) => bandIndexById.set(cell.id, index)));
  const occupiedCols = new Set<number>();
  activePanels.forEach((cell) => {
    const r = cellRect(cell);
    const first = Math.floor((r.x - wallBBox.x) / MODULE_MM);
    const last = Math.ceil((r.x + r.w - wallBBox.x) / MODULE_MM) - 1;
    for (let i = first; i <= last; i += 1) occupiedCols.add(i);
  });

  // Each panel's TRUE native-resolution pixel rect (back view, unmirrored):
  // its own pixW x pixH from PANEL_TYPES (e.g. MT is 256x64, NOT a
  // physical-size scaling of MG9's pitch), placed at its own mm position
  // converted to pixels via that panel type's own px-per-mm ratio (per
  // axis, and swapped the same way cellRect's own mm footprint swaps on a
  // 90/270 rotation, so the ratio stays internally consistent).
  //
  // Deliberately NOT accumulated band-by-band (tightly packing each row's
  // panels left-to-right in array order): that silently closes up any gap
  // from a missing/removed panel in the middle of a row, shifting every
  // panel after the gap out of its true position - the wall's pixel
  // positions must reflect real relative mm spacing, gaps included, or the
  // pattern no longer matches which physical panel is which.
  const panelPixelRects = new Map<string, RectMm>();
  activePanels.forEach((cell) => {
    const rect = cellRect(cell);
    const spec = PANEL_TYPES[cellPanelType(cell)];
    const rot = ((Math.round((cell.rotation ?? 0) / 90) * 90) % 360 + 360) % 360;
    const swapped = rot === 90 || rot === 270;
    const nativeW = swapped ? spec.pixH : spec.pixW;
    const nativeH = swapped ? spec.pixW : spec.pixH;
    const physWMm = (swapped ? spec.h : spec.w) * 1000;
    const physHMm = (swapped ? spec.w : spec.h) * 1000;
    const x = Math.round(((rect.x - wallBBox.x) / physWMm) * nativeW);
    const y = Math.round(((rect.y - wallBBox.y) / physHMm) * nativeH);
    panelPixelRects.set(cell.id, { x, y, w: nativeW, h: nativeH });
  });
  let W = 0;
  let H = 0;
  panelPixelRects.forEach((r) => {
    W = Math.max(W, r.x + r.w);
    H = Math.max(H, r.y + r.h);
  });
  W = Math.max(1, W);
  H = Math.max(1, H);

  // RGB tile size = one full panel footprint at its OWN native resolution
  // (not a physical-size scaling), so each panel shows a single solid
  // colour. Derived from the panels actually on the wall (most common type
  // wins) rather than the project's panelType default, which can go stale
  // if panels were changed individually after the fact.
  const typeCounts = new Map<PanelTypeKey, number>();
  activePanels.forEach((cell) => {
    const type = cellPanelType(cell);
    typeCounts.set(type, (typeCounts.get(type) ?? 0) + 1);
  });
  let dominantType: PanelTypeKey = project.panelType;
  let dominantCount = -1;
  typeCounts.forEach((count, type) => {
    if (count > dominantCount) {
      dominantCount = count;
      dominantType = type;
    }
  });
  const tileSpec = PANEL_TYPES[dominantType] ?? PANEL_TYPES.MG9;
  const tileWidthPx = tileSpec.pixW;
  const tileHeightPx = tileSpec.pixH;

  // MT's pixel pitch is 2x taller than it is wide (see contentPixelH above) -
  // only apply the correction when EVERY active panel is MT, matching the
  // same conservative scope used elsewhere (a mixed MG9+MT wall doesn't get
  // a blended "content resolution").
  const allMt = activePanels.length > 0 && activePanels.every((cell) => cellPanelType(cell) === "MT");
  const contentPixelH = allMt ? H * 2 : H;

  return {
    activePanels,
    wallBBox,
    W,
    H,
    contentPixelH,
    panelPixelRects,
    totalPanels: activePanels.length,
    activeColsCount: occupiedCols.size,
    activeRowsCount: panelBands.length,
    wallWidthM: wallBBox.w / 1000,
    wallHeightM: wallBBox.h / 1000,
    projectName: (project.projectName || "").trim(),
    surfaceName: (project.surfaceName || "").trim(),
    tileWidthPx,
    tileHeightPx,
    rowLabel: (cell) => (bandIndexById.get(cell.id) ?? 0) + 1,
    // Column numbers must read from the FRONT (the pattern is always
    // rendered mirrored - see dispRectPx below), not the panel's raw
    // back-view x - otherwise the printed number doesn't match the
    // column the audience actually sees it in. Mirror the same way
    // dispRectPx does: a panel's front-view left edge is the wall's
    // width minus its back-view right edge.
    colLabel: (cell) => {
      const rect = cellRect(cell);
      const mirroredLeft = wallBBox.w - (rect.x - wallBBox.x) - rect.w;
      const col = mirroredLeft / MODULE_MM + 1;
      return Number.isInteger(col) ? String(col) : col.toFixed(1);
    },
  };
};

// Always front view (mirrored), independent of any on-screen toggle - matches
// the static PNG test-pattern exporter's convention. Mirrors the panel's
// native-pixel rect horizontally within the total wall pixel width.
const dispRectPx = (layout: TestPatternLayout, cell: Cell): RectMm => {
  const r = layout.panelPixelRects.get(cell.id);
  if (!r) return { x: 0, y: 0, w: 0, h: 0 };
  return { x: layout.W - r.x - r.w, y: r.y, w: r.w, h: r.h };
};

// Wall info text: centred in the middle of the wall - right where the
// alignment overlay's diagonal lines and circle cross - so a thin black
// stroke behind the white fill keeps it legible there too.
const drawInfoText = (ctx: CanvasRenderingContext2D, layout: TestPatternLayout) => {
  const { W, H } = layout;
  const fontPx = Math.max(16, Math.min(40, Math.round(W * 0.016)));
  const lineH = Math.round(fontPx * 1.4);
  const lines: string[] = [];
  if (layout.projectName) lines.push(layout.projectName);
  if (layout.surfaceName) lines.push(layout.surfaceName);
  const isContentResScaled = layout.contentPixelH !== layout.H;
  lines.push(`${isContentResScaled ? "LED Wall Resolution" : "Resolution"}: ${layout.W} x ${layout.H} px`);
  if (isContentResScaled) lines.push(`Recommended Content Resolution: ${layout.W} x ${layout.contentPixelH} px`);
  lines.push(
    `Physical Size: ${layout.wallWidthM.toFixed(1)} x ${layout.wallHeightM.toFixed(1)} m`,
    `Panels: ${layout.totalPanels}`,
    `Grid: ${layout.activeColsCount} columns x ${layout.activeRowsCount} rows`,
  );
  const cx = W / 2;
  const startY = H / 2 - (lineH * (lines.length - 1)) / 2;
  ctx.save();
  ctx.font = `bold ${fontPx}px Arial`;
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.lineJoin = "round";
  ctx.strokeStyle = "#000000";
  ctx.lineWidth = Math.max(2, Math.round(fontPx * 0.12));
  ctx.fillStyle = "#ffffff";
  lines.forEach((line, i) => {
    const ly = startY + i * lineH;
    ctx.strokeText(line, cx, ly);
    ctx.fillText(line, cx, ly);
  });
  ctx.restore();
};

// Single, thicker white outline around the TRUE outer extremity of the whole
// assembled LED surface (not per-panel): render every active panel's true
// (rotated/mirrored) shape as a solid fill into one silhouette - adjacent/
// touching panels naturally fuse into a single region with no internal seam
// lines - then peel off one band INSET from that silhouette's own edge
// (erosion, not dilation: growing the band OUTWARD would put it past the
// wall's own bounding canvas - since the wall's bbox has zero margin by
// definition, an outward band is entirely clipped off-canvas and invisible
// for the common rectangular-wall case). Erosion is implemented as
// invert -> dilate -> invert (dilation = stamping 8 one-pixel-offset copies,
// a standard binary-mask dilate), applied ERODE_PX times for a band that
// thick. Cached per layout (keyed by object identity) since the wall's
// geometry doesn't change frame to frame.
const outerOutlineCache = new WeakMap<TestPatternLayout, HTMLCanvasElement>();
const DILATE_OFFSETS: Array<[number, number]> = [
  [-1, -1], [0, -1], [1, -1],
  [-1, 0], [1, 0],
  [-1, 1], [0, 1], [1, 1],
];
const ERODE_PX = 3; // outline thickness, visibly heavier than the 1px panel outlines

// Erosion here is implemented as invert -> dilate -> invert, which only
// works if there's real "background" around the shape for the inverted
// region to dilate into. The wall's silhouette always touches its own
// canvas on every side by definition (the canvas IS the wall's tight
// bounding box), so with no margin the inverted region is empty right at
// the edges and erosion silently no-ops there - the outline would be
// invisible on every straight wall. Working in a padded canvas (the
// silhouette drawn inset by PAD on all sides) gives the algorithm real
// background to erode from; the result is cropped back to the true W x H
// canvas at the end by drawing it at a -PAD offset.
const OUTLINE_PAD = ERODE_PX + 3;

const buildOuterExtremityOutline = (layout: TestPatternLayout): HTMLCanvasElement => {
  const cached = outerOutlineCache.get(layout);
  if (cached) return cached;

  const { W, H } = layout;
  const PW = W + OUTLINE_PAD * 2;
  const PH = H + OUTLINE_PAD * 2;
  const makeCanvas = () => {
    const c = document.createElement("canvas");
    c.width = PW;
    c.height = PH;
    return c;
  };
  const dilateInto = (dstCtx: CanvasRenderingContext2D, src: HTMLCanvasElement) => {
    dstCtx.clearRect(0, 0, PW, PH);
    dstCtx.drawImage(src, 0, 0);
    DILATE_OFFSETS.forEach(([dx, dy]) => dstCtx.drawImage(src, dx, dy));
  };
  const subtractInto = (dstCtx: CanvasRenderingContext2D, minuend: HTMLCanvasElement, subtrahend: HTMLCanvasElement) => {
    dstCtx.clearRect(0, 0, PW, PH);
    dstCtx.drawImage(minuend, 0, 0);
    dstCtx.globalCompositeOperation = "destination-out";
    dstCtx.drawImage(subtrahend, 0, 0);
    dstCtx.globalCompositeOperation = "source-over";
  };

  // Silhouette: union of every active panel's true shape, solid white,
  // offset by PAD so it sits inset from the padded canvas's own edges.
  const silhouette = makeCanvas();
  const sCtx = silhouette.getContext("2d")!;
  sCtx.fillStyle = "#ffffff";
  layout.activePanels.forEach((cell) => {
    const r = snapRectPx(dispRectPx(layout, cell));
    const shape = PANEL_VARIANTS[cell.panelVariant ?? "STANDARD"].shape;
    sCtx.save();
    applyPanelFrame(sCtx, r.x + OUTLINE_PAD, r.y + OUTLINE_PAD, r.w, r.h, cell.rotation ?? 0, true);
    tracePanelShapePath(sCtx, r.w, r.h, shape);
    sCtx.fill();
    sCtx.restore();
  });

  const fullWhite = makeCanvas();
  const fCtx = fullWhite.getContext("2d")!;
  fCtx.fillStyle = "#ffffff";
  fCtx.fillRect(0, 0, PW, PH);

  const scratchA = makeCanvas();
  const scratchB = makeCanvas();
  const aCtx = scratchA.getContext("2d")!;
  const bCtx = scratchB.getContext("2d")!;
  // erode(src) -> dst = fullWhite - dilate(fullWhite - src), using A/B as scratch.
  const erodeInto = (dstCtx: CanvasRenderingContext2D, src: HTMLCanvasElement) => {
    subtractInto(aCtx, fullWhite, src); // A = invert(src)
    dilateInto(bCtx, scratchA); // B = dilate(A)
    subtractInto(dstCtx, fullWhite, scratchB); // dst = fullWhite - B
  };

  // Erode repeatedly to reach the full band thickness.
  const eroded = makeCanvas();
  erodeInto(eroded.getContext("2d")!, silhouette);
  let current = eroded;
  for (let i = 1; i < ERODE_PX; i += 1) {
    const next = makeCanvas();
    erodeInto(next.getContext("2d")!, current);
    current = next;
  }

  const bandPadded = makeCanvas();
  const bandCtx = bandPadded.getContext("2d")!;
  subtractInto(bandCtx, silhouette, current); // silhouette minus the fully-eroded core = the outline band

  // Crop back to the true wall resolution.
  const out = document.createElement("canvas");
  out.width = W;
  out.height = H;
  out.getContext("2d")!.drawImage(bandPadded, -OUTLINE_PAD, -OUTLINE_PAD);

  outerOutlineCache.set(layout, out);
  return out;
};

// Direction arrow icon (vector line + filled triangular head) for the
// per-panel row/column labels, replacing the old Unicode "↓"/"→" glyph
// characters. A font glyph's internal strokes are only ever as thick as the
// font renderer draws them - at small panel sizes (or when the output canvas
// is later scaled for display/video) that can render at sub-pixel width and
// disappear. Drawing the shaft with an explicit ctx.lineWidth (floored so it
// never rounds below a visible thickness) guarantees it stays visible at any
// output resolution, and the arrowhead is sized directly off that same line
// width so it always scales with it. (x, y) is the icon's top-left corner;
// `size` is its bounding box side length.
const drawDirectionArrow = (ctx: CanvasRenderingContext2D, x: number, y: number, size: number, direction: "down" | "right") => {
  const lineWidth = Math.min(3.5, Math.max(1.5, size * 0.16));
  const headLen = Math.max(lineWidth * 1.8, size * 0.42);
  const headWidth = headLen * 0.9;
  ctx.save();
  ctx.strokeStyle = "#ffffff";
  ctx.fillStyle = "#ffffff";
  ctx.lineWidth = lineWidth;
  ctx.lineCap = "round";
  ctx.lineJoin = "round";
  if (direction === "down") {
    const cx = x + size / 2;
    const yTop = y + lineWidth / 2;
    const yTip = y + size - lineWidth / 2;
    ctx.beginPath();
    ctx.moveTo(cx, yTop);
    ctx.lineTo(cx, yTip - headLen * 0.6);
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(cx, yTip);
    ctx.lineTo(cx - headWidth / 2, yTip - headLen);
    ctx.lineTo(cx + headWidth / 2, yTip - headLen);
    ctx.closePath();
    ctx.fill();
  } else {
    const cy = y + size / 2;
    const xLeft = x + lineWidth / 2;
    const xTip = x + size - lineWidth / 2;
    ctx.beginPath();
    ctx.moveTo(xLeft, cy);
    ctx.lineTo(xTip - headLen * 0.6, cy);
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(xTip, cy);
    ctx.lineTo(xTip - headLen, cy - headWidth / 2);
    ctx.lineTo(xTip - headLen, cy + headWidth / 2);
    ctx.closePath();
    ctx.fill();
  }
  ctx.restore();
};

// Corner-to-corner alignment cross + a centre circle as tall as the wall -
// classic geometry/alignment references for spotting warped, offset or
// stretched panels across the whole assembled surface.
const drawAlignmentOverlay = (ctx: CanvasRenderingContext2D, layout: TestPatternLayout) => {
  const { W, H } = layout;
  ctx.save();
  ctx.strokeStyle = "rgba(255, 255, 255, 0.3)";
  ctx.lineWidth = Math.max(2, Math.round(Math.min(W, H) * 0.0035));
  ctx.beginPath();
  ctx.moveTo(0, 0);
  ctx.lineTo(W, H);
  ctx.moveTo(W, 0);
  ctx.lineTo(0, H);
  ctx.stroke();
  ctx.beginPath();
  ctx.arc(W / 2, H / 2, H / 2, 0, Math.PI * 2);
  ctx.stroke();
  ctx.restore();
};

// 1D "DVD screensaver" ping-pong: reflects a constantly-increasing position
// back and forth across [0, range] with no stored state - the whole
// trajectory is a pure function of elapsed time, so it can be called from
// any frame without tracking velocity/position between calls.
const bounce1D = (t: number, speed: number, range: number): number => {
  if (range <= 0) return 0;
  const period = range * 2;
  let x = (t * speed) % period;
  if (x < 0) x += period;
  return x <= range ? x : period - x;
};

/**
 * Draws a small logo bouncing around the canvas, DVD-screensaver style.
 * Deliberately NOT called from drawTestPatternFrame (or anywhere in the
 * recorder/export paths) - this is decoration for the live browser tab only,
 * never the recorded video or the static PNG/PDF exports.
 */
export const drawBouncingLogo = (ctx: CanvasRenderingContext2D, layout: TestPatternLayout, timeSeconds: number, image: HTMLImageElement) => {
  const { W, H } = layout;
  if (W <= 0 || H <= 0 || !image.naturalWidth || !image.naturalHeight) return;
  // ~1/2 of a standard panel's native pixel width (2x the original 1/4), aspect ratio preserved.
  const logoW = Math.max(4, layout.tileWidthPx / 2);
  const logoH = logoW * (image.naturalHeight / image.naturalWidth);
  const rangeX = Math.max(0, W - logoW);
  const rangeY = Math.max(0, H - logoH);
  // Full one-way traverse takes several seconds on each axis (slow/subtle),
  // and the two periods share no common factor, so the path doesn't retrace
  // a short, obviously-repeating loop. /18 and /14 are double the original
  // /9 and /7 divisors - half the speed, same non-repeating ratio.
  const x = bounce1D(timeSeconds, rangeX / 18, rangeX);
  const y = bounce1D(timeSeconds, rangeY / 14, rangeY);
  ctx.save();
  ctx.imageSmoothingEnabled = true;
  ctx.drawImage(image, x, y, logoW, logoH);
  ctx.restore();
};

/**
 * Draw one animation frame. `timeSeconds` should be pre-wrapped or raw - the
 * function mods it by LOOP_SECONDS internally, so callers can pass either a
 * free-running elapsed time (live view) or a deterministic frame*interval
 * value (recorder), and both loop identically.
 */
export const drawTestPatternFrame = (ctx: CanvasRenderingContext2D, layout: TestPatternLayout, timeSeconds: number) => {
  const { W, H } = layout;
  if (W <= 0 || H <= 0) return;
  // The caller's transform when this function was entered - identity for the
  // WebM/MP4 recorder and PNG/PDF exports (canvas backing store == layout.W x
  // layout.H 1:1), but the live browser tab (TestPatternView.tsx) pre-applies
  // a devicePixelRatio/fit-to-window scale so the canvas stays crisp on
  // high-DPI displays. The per-panel loop below has to reset to THIS base
  // transform (not a hardcoded identity) before drawing the pre-rendered
  // world-space pattern layer, or that layer gets drawn at 1 source px = 1
  // physical px instead of 1 source px = 1 logical unit - which is exactly
  // what made the RGB tiles stop lining up with panel boundaries once the
  // live view started rendering at more than 1 physical pixel per unit.
  const baseTransform = ctx.getTransform();
  ctx.imageSmoothingEnabled = false;

  const rgbTile = getRgbTile(layout.tileWidthPx, layout.tileHeightPx);
  const rgbPeriodPx = layout.tileWidthPx * 3;
  // Grey period = the wall's own diagonal, so exactly one bright band sweeps
  // corner-to-corner across the whole wall at a time (not several small
  // repeats), scaled to each project's actual size.
  const greyPeriod = Math.max(1, Math.round(Math.hypot(W, H)));
  const greyTile = getGreyTile(greyPeriod);

  const phase = ((timeSeconds % LOOP_SECONDS) + LOOP_SECONDS) % LOOP_SECONDS;
  const rgbSlidePx = (phase / LOOP_SECONDS) * rgbPeriodPx;
  const greySlidePx = (phase / LOOP_SECONDS) * greyPeriod;

  // Build the world-space pattern layer once per frame (two full-wall draws,
  // O(1) regardless of panel count), then reveal it through each panel's clip.
  const layerCanvas = getPatternLayer(W, H);
  const layerCtx = layerCanvas.getContext("2d")!;
  layerCtx.clearRect(0, 0, W, H);
  // Nearest-neighbour sampling: at a fractional slide offset, a smoothed
  // pattern would blend adjacent red/green/blue tile pixels into a soft
  // rainbow seam. Each panel must show a SOLID colour with a hard boundary
  // as the slide reveals the next one, never a blended gradient.
  layerCtx.imageSmoothingEnabled = false;

  const rgbPattern = layerCtx.createPattern(rgbTile, "repeat")!;
  rgbPattern.setTransform(new DOMMatrix().translate(rgbSlidePx, 0));
  layerCtx.fillStyle = rgbPattern;
  layerCtx.fillRect(0, 0, W, H);

  layerCtx.save();
  // 'multiply' (not 'overlay'): overlay is a no-op on pure 0/255 channel
  // values, so it would never visibly affect these solid R/G/B panels.
  // Multiply scales the panel's existing channel toward the tile's floor and
  // leaves zero channels at zero, so brightness visibly sweeps across each
  // panel without ever introducing a new hue.
  layerCtx.globalCompositeOperation = "multiply";
  const greyPattern = layerCtx.createPattern(greyTile, "repeat")!;
  greyPattern.setTransform(new DOMMatrix().translate(greySlidePx, greySlidePx));
  layerCtx.fillStyle = greyPattern;
  layerCtx.fillRect(0, 0, W, H);
  layerCtx.restore();

  // Background + per-panel reveal.
  ctx.imageSmoothingEnabled = false;
  ctx.fillStyle = "#000000";
  ctx.fillRect(0, 0, W, H);

  layout.activePanels.forEach((cell) => {
    // Snap to the integer pixel grid so a 1px stroke lands crisply on one
    // pixel row/column (not straddling two, which anti-aliases into a soft
    // >1px-looking line), and so neighbouring panels' rounded edges agree
    // with each other rather than drifting apart.
    const r = snapRectPx(dispRectPx(layout, cell));
    const shape = PANEL_VARIANTS[cell.panelVariant ?? "STANDARD"].shape;
    const rotation = cell.rotation ?? 0;

    // Pattern fill: clip to the panel's true (rotated/mirrored) silhouette,
    // then draw the SAME world-space layer with the transform reset to
    // identity - so a rotated/shaped panel reveals the correct slice of the
    // global pattern instead of a rotated copy of it.
    ctx.save();
    applyPanelFrame(ctx, r.x, r.y, r.w, r.h, rotation, true);
    tracePanelShapePath(ctx, r.w, r.h, shape);
    ctx.clip();
    ctx.setTransform(baseTransform);
    ctx.drawImage(layerCanvas, 0, 0);
    ctx.restore();

    // 1px white outline, following the true shape. Rect/corner panels are
    // axis-aligned at every supported rotation (0/90/180/270 just swap the
    // already-baked w/h) and mirror-symmetric, so draw a plain strokeRect
    // offset by 0.5px - the standard crisp-1px-line trick - instead of
    // going through the rotate/mirror path, which can leave the line
    // sitting across a pixel boundary. Triangle/curve panels keep the true
    // traced path; their straight legs still land on this same snapped
    // rect, only the hypotenuse/arc is inherently anti-aliased.
    ctx.save();
    ctx.strokeStyle = "#ffffff";
    ctx.lineWidth = 1;
    if (shape === "rect" || shape === "corner") {
      ctx.strokeRect(r.x + 0.5, r.y + 0.5, Math.max(0, r.w - 1), Math.max(0, r.h - 1));
    } else {
      applyPanelFrame(ctx, r.x, r.y, r.w, r.h, rotation, true);
      tracePanelShapePath(ctx, r.w, r.h, shape);
      ctx.stroke();
    }
    ctx.restore();

    // Location label: top-left corner, two lines (row then column), always
    // axis-aligned/white/upright regardless of the panel's own rotation. Each
    // line is a vector direction-arrow icon followed by the row/column number.
    const pad = Math.max(3, Math.round(Math.min(r.w, r.h) * 0.06));
    const fontPx = Math.max(6, Math.floor(Math.min(r.w, r.h) * 0.08));
    const lineH = Math.round(fontPx * 1.15);
    const iconSize = fontPx;
    const textX = r.x + pad + iconSize + Math.max(2, Math.round(fontPx * 0.25));
    ctx.fillStyle = "#ffffff";
    ctx.textAlign = "left";
    ctx.textBaseline = "top";
    ctx.font = `bold ${fontPx}px Arial`;
    drawDirectionArrow(ctx, r.x + pad, r.y + pad, iconSize, "down");
    ctx.fillText(`${layout.rowLabel(cell)}`, textX, r.y + pad);
    drawDirectionArrow(ctx, r.x + pad, r.y + pad + lineH, iconSize, "right");
    ctx.fillText(`${layout.colLabel(cell)}`, textX, r.y + pad + lineH);
  });

  ctx.drawImage(buildOuterExtremityOutline(layout), 0, 0);
  drawAlignmentOverlay(ctx, layout);
  drawInfoText(ctx, layout);
};
