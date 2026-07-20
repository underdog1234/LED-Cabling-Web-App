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
  mirrorRectX,
  isPanelHead,
  applyPanelFrame,
  tracePanelShapePath,
} from "../App";
import { activeBBox, bandPanels, MODULE_MM, type RectMm } from "../model/panels";

export type TestPatternProject = {
  projectName: string;
  panelType: PanelTypeKey;
  panels: Cell[];
};

export type TestPatternLayout = {
  activePanels: Cell[];
  wallBBox: RectMm;
  W: number;
  H: number;
  pxPerMm: number;
  totalPanels: number;
  activeColsCount: number;
  activeRowsCount: number;
  wallWidthM: number;
  wallHeightM: number;
  projectName: string;
  rowLabel: (cell: Cell) => number;
  colLabel: (cell: Cell) => string;
};

// One full loop of the animation, in seconds. Both the RGB slide and the
// greyscale gradient return to their exact starting phase at this mark, so a
// video recorded for exactly this long loops seamlessly on repeat.
export const LOOP_SECONDS = 20;
// Cap the animation's internal draw rate - this is an alignment-check tool,
// not a smooth video, and a lower rate keeps hundreds of panels responsive.
export const DRAW_FPS = 24;

// RGB tile = one MG9 module footprint at native pixel pitch (matches the PNG
// test-pattern exporter's pxPerMm exactly, see below), so checkerboard cells
// land on real module boundaries for snap-placed layouts. Kept as a literal
// (not read from PANEL_TYPES.MG9.pixW) so this module has no import-time
// dependency on App.tsx - App.tsx also imports FROM this module for the
// in-app video recorder, and a live top-level read here would race against
// that circular import's module-initialisation order.
const TILE_PX = 168; // must match PANEL_TYPES.MG9.pixW in App.tsx
const RGB_PERIOD_PX = TILE_PX * 3; // 504 - one full stagger + slide period
const RGB_COLORS = ["#ff0000", "#00ff00", "#0000ff"];

let cachedRgbTile: HTMLCanvasElement | null = null;
let cachedGreyTile: { period: number; canvas: HTMLCanvasElement } | null = null;
let cachedPatternLayer: { w: number; h: number; canvas: HTMLCanvasElement } | null = null;

// Staggered diagonal R/G/B supercell: row R, col C -> colour (R+C) mod 3. This
// makes adjacent panels differ AND makes each row offset from the one above,
// and tiles seamlessly in both axes because the pattern's own period matches
// the tile size - no per-panel logic needed.
const buildRgbTile = (): HTMLCanvasElement => {
  const c = document.createElement("canvas");
  c.width = RGB_PERIOD_PX;
  c.height = RGB_PERIOD_PX;
  const ctx = c.getContext("2d")!;
  for (let row = 0; row < 3; row += 1) {
    for (let col = 0; col < 3; col += 1) {
      ctx.fillStyle = RGB_COLORS[(row + col) % 3];
      ctx.fillRect(col * TILE_PX, row * TILE_PX, TILE_PX, TILE_PX);
    }
  }
  return c;
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
  // Same formula as the static PNG test-pattern exporter, so resolutions match.
  const pxPerMm = PANEL_TYPES.MG9.pixW / (PANEL_TYPES.MG9.w * 1000);
  const W = Math.max(1, Math.round(wallBBox.w * pxPerMm));
  const H = Math.max(1, Math.round(wallBBox.h * pxPerMm));
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
  return {
    activePanels,
    wallBBox,
    W,
    H,
    pxPerMm,
    totalPanels: activePanels.length,
    activeColsCount: occupiedCols.size,
    activeRowsCount: panelBands.length,
    wallWidthM: wallBBox.w / 1000,
    wallHeightM: wallBBox.h / 1000,
    projectName: project.projectName,
    rowLabel: (cell) => (bandIndexById.get(cell.id) ?? 0) + 1,
    colLabel: (cell) => {
      const col = (cellRect(cell).x - wallBBox.x) / MODULE_MM + 1;
      return Number.isInteger(col) ? String(col) : col.toFixed(1);
    },
  };
};

// Always front view (mirrored), independent of any on-screen toggle - matches
// the static PNG test-pattern exporter's convention.
const dispRectPx = (layout: TestPatternLayout, cell: Cell): RectMm => {
  const d = mirrorRectX(cellRect(cell), layout.wallBBox);
  return {
    x: (d.x - layout.wallBBox.x) * layout.pxPerMm,
    y: (d.y - layout.wallBBox.y) * layout.pxPerMm,
    w: d.w * layout.pxPerMm,
    h: d.h * layout.pxPerMm,
  };
};

// Wall info text: centred in the middle of the wall, no background box. Each
// line is stroked in dark then filled in white so it stays readable over
// whatever colour/brightness happens to be underneath.
const drawInfoText = (ctx: CanvasRenderingContext2D, layout: TestPatternLayout) => {
  const { W, H } = layout;
  const fontPx = Math.max(14, Math.min(30, Math.round(W * 0.014)));
  const lineH = Math.round(fontPx * 1.4);
  const lines = [
    `Resolution: ${layout.W} x ${layout.H} px`,
    `Physical Size: ${layout.wallWidthM.toFixed(1)} x ${layout.wallHeightM.toFixed(1)} m`,
    `Panels: ${layout.totalPanels}`,
    `Grid: ${layout.activeColsCount} columns x ${layout.activeRowsCount} rows`,
  ];
  const cx = W / 2;
  const startY = H / 2 - (lineH * (lines.length - 1)) / 2;
  ctx.save();
  ctx.font = `bold ${fontPx}px Arial`;
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.lineJoin = "round";
  ctx.lineWidth = Math.max(2, Math.round(fontPx * 0.14));
  ctx.strokeStyle = "rgba(0, 0, 0, 0.85)";
  ctx.fillStyle = "#ffffff";
  lines.forEach((line, i) => {
    const y = startY + i * lineH;
    ctx.strokeText(line, cx, y);
    ctx.fillText(line, cx, y);
  });
  ctx.restore();
};

// Corner-to-corner alignment cross + a centre circle as tall as the wall -
// classic geometry/alignment references for spotting warped, offset or
// stretched panels across the whole assembled surface.
const drawAlignmentOverlay = (ctx: CanvasRenderingContext2D, layout: TestPatternLayout) => {
  const { W, H } = layout;
  ctx.save();
  ctx.strokeStyle = "rgba(255, 255, 255, 0.55)";
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

/**
 * Draw one animation frame. `timeSeconds` should be pre-wrapped or raw - the
 * function mods it by LOOP_SECONDS internally, so callers can pass either a
 * free-running elapsed time (live view) or a deterministic frame*interval
 * value (recorder), and both loop identically.
 */
export const drawTestPatternFrame = (ctx: CanvasRenderingContext2D, layout: TestPatternLayout, timeSeconds: number) => {
  const { W, H } = layout;
  if (W <= 0 || H <= 0) return;
  ctx.imageSmoothingEnabled = false;

  if (!cachedRgbTile) cachedRgbTile = buildRgbTile();
  // Grey period = the wall's own diagonal, so exactly one bright band sweeps
  // corner-to-corner across the whole wall at a time (not several small
  // repeats), scaled to each project's actual size.
  const greyPeriod = Math.max(1, Math.round(Math.hypot(W, H)));
  const greyTile = getGreyTile(greyPeriod);

  const phase = ((timeSeconds % LOOP_SECONDS) + LOOP_SECONDS) % LOOP_SECONDS;
  const rgbSlidePx = (phase / LOOP_SECONDS) * RGB_PERIOD_PX;
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

  const rgbPattern = layerCtx.createPattern(cachedRgbTile, "repeat")!;
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
    const r = dispRectPx(layout, cell);
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
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.drawImage(layerCanvas, 0, 0);
    ctx.restore();

    // 1px white outline, following the true shape.
    ctx.save();
    applyPanelFrame(ctx, r.x, r.y, r.w, r.h, rotation, true);
    tracePanelShapePath(ctx, r.w, r.h, shape);
    ctx.strokeStyle = "#ffffff";
    ctx.lineWidth = 1;
    ctx.stroke();
    ctx.restore();

    // Labels: always axis-aligned/white, matching the PNG exporter's convention.
    const cx = r.x + r.w / 2;
    ctx.fillStyle = "#ffffff";
    ctx.textAlign = "center";
    ctx.textBaseline = "alphabetic";
    ctx.font = `bold ${Math.max(11, Math.floor(r.h * 0.085))}px Arial`;
    ctx.fillText(`↓${layout.rowLabel(cell)} →${layout.colLabel(cell)}${cellPanelType(cell) === "MT" ? " (MT)" : ""}`, cx, r.y + r.h * 0.28);
    if (cell.assignedPort) ctx.fillText(`🔌 P${cell.assignedPort}`, cx, r.y + r.h * 0.52);
    if (cell.assignedPowerPort) ctx.fillText(`⚡ ${cell.assignedPowerPort}`, cx, r.y + r.h * 0.76);
  });

  drawAlignmentOverlay(ctx, layout);
  drawInfoText(ctx, layout);
};
