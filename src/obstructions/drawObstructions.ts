// Shared obstruction overlay renderer.
//
// Used by every surface that shows the LED picture: the live/animated test
// pattern, the static test-pattern PNGs, and the Basic LED Map. One function
// so an obstruction can never land in a different place on the video than on
// the PNG a technician is holding.
//
// Deliberately no dependency on App.tsx - this file imports only the pure
// obstruction model, so the test-pattern renderer can use it without pulling
// the monolith into a second import cycle.
import {
  type Obstruction,
  obstructionMetresLabel,
  obstructionPixelRect,
  obstructionIntersectsSurface,
} from "./obstructionModel";

export type ObstructionSurface = {
  /** Physical bounding box of the surface being drawn, workspace mm. */
  wallBBoxMm: { x: number; y: number; w: number; h: number };
  /** The surface's native pixel size. */
  W: number;
  H: number;
  /**
   * Front view mirrors horizontally, exactly as panels do (see dispRectPx in
   * drawTestPattern.ts). An obstruction is a real object in front of the
   * wall, so it has to mirror with the picture or it would be shown on the
   * wrong side of it.
   */
  mirror: boolean;
};

// Fully opaque: an obstruction hides the LED behind it, so showing the
// pattern through it would misrepresent what the audience actually sees.
const FILL = "#0f172a";
const BORDER = "#f43f5e";
const TEXT = "#ffffff";

/**
 * Draws each obstruction as an opaque block labelled with its name and its
 * size in both pixels and metres.
 *
 * Call it LAST, after the pattern and any panel outlines, so nothing is drawn
 * over the top of it. The caller's transform is inherited (so an all-MT
 * wall's content-resolution scale applies here too, same as everything else).
 */
export const drawObstructions = (
  ctx: CanvasRenderingContext2D,
  obstructions: Obstruction[],
  surface: ObstructionSurface,
): void => {
  if (!obstructions.length) return;
  const { W, H, wallBBoxMm, mirror } = surface;
  if (W <= 0 || H <= 0 || wallBBoxMm.w <= 0 || wallBBoxMm.h <= 0) return;

  // Scale the furniture to the surface, like drawInfoText does - a 3m-wide
  // sub-screen and a 20m wall must both end up readable.
  const fontPx = Math.max(12, Math.min(36, Math.round(W * 0.013)));
  const lineH = Math.round(fontPx * 1.25);
  const borderPx = Math.max(2, Math.round(W * 0.0012));

  obstructions.forEach((obstruction) => {
    const raw = obstructionPixelRect(obstruction, wallBBoxMm, W, H);
    // Off-surface obstructions are skipped rather than clamped: clamping
    // would park a block against the edge of a sub-screen it does not
    // actually cover, which reads as a real obstruction that isn't there.
    if (!obstructionIntersectsSurface(raw, W, H)) return;
    const x = mirror ? W - raw.x - raw.w : raw.x;
    const rect = { x, y: raw.y, w: raw.w, h: raw.h };

    ctx.save();
    ctx.fillStyle = FILL;
    ctx.fillRect(rect.x, rect.y, rect.w, rect.h);
    ctx.lineWidth = borderPx;
    ctx.strokeStyle = BORDER;
    ctx.strokeRect(rect.x + borderPx / 2, rect.y + borderPx / 2, Math.max(0, rect.w - borderPx), Math.max(0, rect.h - borderPx));

    // Hatching makes an obstruction unmistakable even in a monochrome print
    // or a photo of the wall, where the border colour alone may not survive.
    ctx.beginPath();
    ctx.rect(rect.x, rect.y, rect.w, rect.h);
    ctx.clip();
    ctx.globalAlpha = 0.25;
    ctx.strokeStyle = BORDER;
    ctx.lineWidth = Math.max(1, Math.round(borderPx / 2));
    const step = Math.max(12, Math.round(fontPx * 1.2));
    for (let d = -rect.h; d < rect.w + rect.h; d += step) {
      ctx.beginPath();
      ctx.moveTo(rect.x + d, rect.y);
      ctx.lineTo(rect.x + d + rect.h, rect.y + rect.h);
      ctx.stroke();
    }
    ctx.globalAlpha = 1;

    const lines = [obstruction.name, obstructionMetresLabel(obstruction), `${raw.w} x ${raw.h} px`];
    ctx.font = `bold ${fontPx}px system-ui, -apple-system, Segoe UI, Roboto, sans-serif`;
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.lineJoin = "round";
    const startY = rect.y + rect.h / 2 - (lineH * (lines.length - 1)) / 2;
    lines.forEach((line, index) => {
      const ly = startY + index * lineH;
      // Stroke-then-fill keeps the label readable against the hatching and
      // against whatever the obstruction happens to be sitting on.
      ctx.lineWidth = Math.max(2, Math.round(fontPx / 6));
      ctx.strokeStyle = "#000000";
      ctx.strokeText(line, rect.x + rect.w / 2, ly);
      ctx.fillStyle = TEXT;
      ctx.fillText(line, rect.x + rect.w / 2, ly);
    });
    ctx.restore();
  });
};
