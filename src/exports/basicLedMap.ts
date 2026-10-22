// Basic LED Map - a plain picture of the finished surface.
//
// Not a test pattern and not a technical drawing: no port numbers, no panel
// numbering, no alignment overlay. Just the real shape of the LED (every
// separate piece of it), filled with either a flat colour or an uploaded
// image stretched across the whole surface, with the obstructions that will
// stand in front of it drawn on top and labelled.
//
// It is what you hand to whoever is making the content, so it answers only
// their question: what shape is my canvas, and what is going to cover it.
//
// Geometry comes from computeTestPatternLayout, deliberately rather than a
// private copy, so this map and the test patterns always agree about where a
// panel is - a second implementation is exactly how the two drifted apart
// before (see renderTestPatternCanvas's own comment).
import { type Cell, PANEL_VARIANTS, applyPanelFrame, cellRect, tracePanelShapePath } from "../App";
import { computeTestPatternLayout, type TestPatternProject } from "../testPattern/drawTestPattern";
import { drawObstructions } from "../obstructions/drawObstructions";
import { type Obstruction } from "../obstructions/obstructionModel";
import { connectedGroups } from "../model/panels";

export type BasicLedMapOptions = {
  projectName: string;
  surfaceName: string;
  panels: Cell[];
  panelType: TestPatternProject["panelType"];
  obstructions: Obstruction[];
  /** Flat fill behind the LED shape. Used on its own, or as the backdrop while an image loads. */
  backgroundColor: string;
  /** Optional artwork, stretched to cover the FULL surface (not per piece, not per panel). */
  backgroundImage: HTMLImageElement | null;
  /** Outline each separate piece of the layout and label it. */
  showPieceOutlines: boolean;
};

const PIECE_OUTLINE = "#ffffff";
const GAP_COLOR = "#000000";

/**
 * Renders the map and returns the canvas, at the surface's Recommended
 * Content Resolution - the same resolution the content itself should be
 * authored at, so the exported PNG can be dropped straight onto a timeline.
 */
export const renderBasicLedMap = (options: BasicLedMapOptions): HTMLCanvasElement => {
  const layout = computeTestPatternLayout({
    projectName: options.projectName,
    surfaceName: options.surfaceName,
    panelType: options.panelType,
    panels: options.panels,
  });
  const W = Math.max(1, layout.W);
  const H = Math.max(1, layout.H);

  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, layout.contentPixelW);
  canvas.height = Math.max(1, layout.contentPixelH);
  const ctx = canvas.getContext("2d");
  if (!ctx) throw new Error("Canvas context unavailable");
  // Same trick the test-pattern PNG uses: draw in native W x H space and let
  // one vertical scale carry an all-MT wall up to its content resolution.
  ctx.scale(1, layout.contentPixelH / H);

  // Anything that is not LED reads as nothing at all, so a gap in the wall is
  // obviously a gap and not a dark patch of content.
  ctx.fillStyle = GAP_COLOR;
  ctx.fillRect(0, 0, W, H);

  // Front view, like every other picture this app exports: mirror each
  // panel's pixel rect horizontally within the wall's total width.
  const dispRect = (cell: Cell) => {
    const r = layout.panelPixelRects.get(cell.id);
    if (!r) return null;
    return { x: W - r.x - r.w, y: r.y, w: r.w, h: r.h };
  };

  // 1 + 2. Paint the background through each panel's own true silhouette
  //        (real shape, rotation and front-view mirror), so it only ever
  //        appears where there is actually LED to show it on.
  //
  //        Clipped PER PANEL rather than accumulating one path for all of
  //        them: tracePanelShapePath starts with beginPath(), so building a
  //        combined path in a loop silently keeps only the last panel - which
  //        left the whole map black except for one module.
  //
  //        The background itself is drawn in WALL space every time (the clip
  //        survives the transform reset, since clips live in device space), so
  //        the image is stretched across the FULL surface exactly once rather
  //        than being repeated inside each panel - each panel reveals its own
  //        part of it, which is what the content will actually do.
  const baseTransform = ctx.getTransform();
  layout.activePanels.forEach((cell) => {
    const r = dispRect(cell);
    if (!r) return;
    const variant = PANEL_VARIANTS[cell.panelVariant ?? "STANDARD"];
    ctx.save();
    applyPanelFrame(ctx, r.x, r.y, r.w, r.h, cell.rotation ?? 0, true);
    tracePanelShapePath(ctx, r.w, r.h, variant.shape);
    ctx.clip();
    ctx.setTransform(baseTransform);
    ctx.fillStyle = options.backgroundColor;
    ctx.fillRect(0, 0, W, H);
    if (options.backgroundImage) {
      ctx.imageSmoothingEnabled = true;
      ctx.imageSmoothingQuality = "high";
      ctx.drawImage(options.backgroundImage, 0, 0, W, H);
    }
    ctx.restore();
  });

  // 3. Outline each separate piece. "Separate" means not physically joined -
  //    connectedGroups is the same joined-ness test the layout tool uses, so
  //    two towers read as two pieces and one wall reads as one.
  if (options.showPieceOutlines) {
    // connectedGroups returns cell id -> group index; invert it so each
    // physically separate piece can be outlined and labelled as one.
    const groupOf = connectedGroups(layout.activePanels, cellRect);
    const byGroup = new Map<number, Cell[]>();
    layout.activePanels.forEach((cell) => {
      const key = groupOf.get(cell.id);
      if (key === undefined) return;
      const list = byGroup.get(key);
      if (list) list.push(cell);
      else byGroup.set(key, [cell]);
    });
    const groups = [...byGroup.values()];
    const stroke = Math.max(2, Math.round(Math.min(W, H) * 0.003));
    const fontPx = Math.max(14, Math.round(Math.min(W, H) * 0.02));
    ctx.save();
    ctx.lineWidth = stroke;
    ctx.strokeStyle = PIECE_OUTLINE;
    ctx.textAlign = "left";
    ctx.textBaseline = "top";
    groups.forEach((group, index) => {
      const rects = group.map(dispRect).filter((r): r is NonNullable<typeof r> => r !== null);
      if (!rects.length) return;
      const x0 = Math.min(...rects.map((r) => r.x));
      const y0 = Math.min(...rects.map((r) => r.y));
      const x1 = Math.max(...rects.map((r) => r.x + r.w));
      const y1 = Math.max(...rects.map((r) => r.y + r.h));
      const inset = stroke / 2;
      ctx.strokeRect(x0 + inset, y0 + inset, x1 - x0 - stroke, y1 - y0 - stroke);
      // Only worth numbering when there is more than one piece to tell apart.
      if (groups.length < 2) return;
      const label = `PIECE ${index + 1} · ${x1 - x0} x ${y1 - y0} px`;
      ctx.font = `bold ${fontPx}px system-ui, -apple-system, Segoe UI, Roboto, sans-serif`;
      const padPx = Math.round(fontPx * 0.35);
      const boxW = ctx.measureText(label).width + padPx * 2;
      ctx.fillStyle = PIECE_OUTLINE;
      ctx.fillRect(x0 + stroke, y0 + stroke, boxW, fontPx + padPx * 2);
      ctx.fillStyle = "#020617";
      ctx.fillText(label, x0 + stroke + padPx, y0 + stroke + padPx);
    });
    ctx.restore();
  }

  // 4. Obstructions last and opaque - the whole point of the map is seeing
  //    what the content loses to them.
  drawObstructions(ctx, options.obstructions, {
    wallBBoxMm: layout.wallBBox,
    W,
    H,
    mirror: true,
  });

  return canvas;
};
