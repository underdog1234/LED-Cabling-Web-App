// ---------------------------------------------------------------------------
// Rendering: one screen, or the whole canvas.
//
// A sub-screen is always drawn into its OWN canvas at its native resolution
// (pattern first, then its overlays), and the full canvas is built by placing
// those finished screens on the background in layer order. So the full-canvas
// output and a single-screen export are the same pixels by construction, and
// a single-screen export can never pick up a neighbour that overlaps it.
// ---------------------------------------------------------------------------

import { motionOffset, type GeneratorConfig, type LedLayoutSource, type SubScreen } from "./model";
import { LED_LAYOUT_PATTERN_ID, PATTERN_BY_ID, resolveSettings, type PatternInfo } from "./patterns";

// --- LED planner layouts (loaded on demand) ---------------------------------
// The planner's own Moving Test Pattern code is only pulled in when a screen
// actually carries an imported LED layout, so the generator stays light for
// everyone else.

type LedModule = {
  computeTestPatternLayout: typeof import("../testPattern/drawTestPattern").computeTestPatternLayout;
  drawTestPatternFrame: typeof import("../testPattern/drawTestPattern").drawTestPatternFrame;
  LOOP_SECONDS: number;
  normalizePanels: typeof import("../App").normalizePanels;
};

let ledModule: LedModule | null = null;
let ledLoading: Promise<LedModule> | null = null;

export const loadLedModule = (): Promise<LedModule> => {
  if (!ledLoading) {
    ledLoading = Promise.all([import("../testPattern/drawTestPattern"), import("../App")]).then(([d, a]) => {
      ledModule = {
        computeTestPatternLayout: d.computeTestPatternLayout,
        drawTestPatternFrame: d.drawTestPatternFrame,
        LOOP_SECONDS: d.LOOP_SECONDS,
        normalizePanels: a.normalizePanels,
      };
      return ledModule;
    });
  }
  return ledLoading;
};

export const ledModuleReady = () => ledModule !== null;

const ledLayoutCache = new WeakMap<LedLayoutSource, ReturnType<LedModule["computeTestPatternLayout"]>>();

export const ledLayoutFor = (source: LedLayoutSource) => {
  if (!ledModule) return null;
  const cached = ledLayoutCache.get(source);
  if (cached) return cached;
  const panels = ledModule.normalizePanels(source.panels);
  const panelType = source.panelType === "MT" ? "MT" : "MG9";
  const layout = ledModule.computeTestPatternLayout({
    projectName: source.projectName,
    surfaceName: source.surfaceName,
    panelType,
    panels,
    subScreens: source.subScreen ? [source.subScreen] : [],
  });
  ledLayoutCache.set(source, layout);
  return layout;
};

// --- one screen ---------------------------------------------------------------

export type RenderContext = { loopSeconds: number; canvas: { w: number; h: number } };

const mod = (a: number, b: number) => ((a % b) + b) % b;

export const patternInfoFor = (screen: SubScreen, time: number, rc: RenderContext): PatternInfo => {
  const loop = Math.max(0.001, rc.loopSeconds);
  const loopTime = mod(time + (screen.phase || 0), loop);
  return {
    progress: loopTime / loop,
    loopTime,
    loopSeconds: loop,
    screen: { name: screen.name, x: screen.x, y: screen.y, w: screen.w, h: screen.h, color: screen.color },
    led: screen.led,
    canvas: rc.canvas,
  };
};

/**
 * Analogue clock with a seconds hand, plus the time in digits under it, in
 * the top-right corner. It shows the time of day (this computer's clock), not
 * the pattern clock, so it is the one thing on screen that never loops.
 */
const drawClock = (ctx: CanvasRenderingContext2D, w: number, h: number, nowMs: number) => {
  const minD = Math.min(w, h);
  const r = Math.max(18, Math.round(minD * 0.11));
  const margin = Math.max(6, Math.round(minD * 0.03));
  const digitsPx = Math.max(10, Math.round(r * 0.36));
  const cx = w - margin - r;
  const cy = margin + r;
  const d = new Date(nowMs);
  const sec = d.getSeconds();
  const min = d.getMinutes() + sec / 60;
  const hr = (d.getHours() % 12) + min / 60;
  ctx.save();
  ctx.fillStyle = "rgba(2,6,23,0.8)";
  ctx.beginPath();
  ctx.arc(cx, cy, r, 0, Math.PI * 2);
  ctx.fill();
  ctx.strokeStyle = "#ffffff";
  ctx.lineWidth = Math.max(1.5, r * 0.04);
  ctx.stroke();
  ctx.lineCap = "round";
  for (let i = 0; i < 60; i += 1) {
    const a = (i / 60) * Math.PI * 2;
    const major = i % 5 === 0;
    const inner = r * (major ? 0.78 : 0.88);
    ctx.lineWidth = Math.max(1, r * (major ? 0.05 : 0.02));
    ctx.beginPath();
    ctx.moveTo(cx + Math.sin(a) * inner, cy - Math.cos(a) * inner);
    ctx.lineTo(cx + Math.sin(a) * r * 0.94, cy - Math.cos(a) * r * 0.94);
    ctx.stroke();
  }
  const hand = (turns: number, length: number, width: number, colour: string, tail = 0) => {
    const a = turns * Math.PI * 2;
    ctx.strokeStyle = colour;
    ctx.lineWidth = width;
    ctx.beginPath();
    ctx.moveTo(cx - Math.sin(a) * tail, cy + Math.cos(a) * tail);
    ctx.lineTo(cx + Math.sin(a) * length, cy - Math.cos(a) * length);
    ctx.stroke();
  };
  hand(hr / 12, r * 0.5, Math.max(2, r * 0.08), "#ffffff");
  hand(min / 60, r * 0.74, Math.max(1.5, r * 0.055), "#ffffff");
  hand(sec / 60, r * 0.86, Math.max(1, r * 0.025), "#ef4444", r * 0.18);
  ctx.fillStyle = "#ef4444";
  ctx.beginPath();
  ctx.arc(cx, cy, Math.max(2, r * 0.06), 0, Math.PI * 2);
  ctx.fill();
  const pad = (n: number) => String(n).padStart(2, "0");
  const text = `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(sec)}`;
  ctx.font = `bold ${digitsPx}px 'Courier New', monospace`;
  ctx.textAlign = "center";
  ctx.textBaseline = "top";
  const tw = ctx.measureText(text).width + digitsPx * 0.6;
  const ty = cy + r + Math.round(digitsPx * 0.25);
  ctx.fillStyle = "rgba(2,6,23,0.8)";
  ctx.fillRect(cx - tw / 2, ty, tw, digitsPx * 1.2);
  ctx.fillStyle = "#ffffff";
  ctx.fillText(text, cx, ty + digitsPx * 0.1);
  ctx.restore();
};

const drawOverlays = (ctx: CanvasRenderingContext2D, screen: SubScreen, nowMs: number) => {
  const { w, h } = screen;
  const o = screen.overlays;
  const stroke = Math.max(2, Math.round(Math.min(w, h) * 0.004));
  const fontPx = Math.max(12, Math.round(Math.min(w, h) * 0.03));
  ctx.save();
  if (o.crosshair) {
    ctx.fillStyle = screen.color;
    ctx.fillRect(Math.floor(w / 2), 0, 1, h);
    ctx.fillRect(0, Math.floor(h / 2), w, 1);
  }
  if (o.border) {
    ctx.fillStyle = screen.color;
    ctx.fillRect(0, 0, w, stroke);
    ctx.fillRect(0, h - stroke, w, stroke);
    ctx.fillRect(0, 0, stroke, h);
    ctx.fillRect(w - stroke, 0, stroke, h);
  }
  ctx.font = `bold ${fontPx}px Arial, Helvetica, sans-serif`;
  ctx.textBaseline = "top";
  ctx.textAlign = "left";
  const pad = Math.round(fontPx * 0.35);
  if (o.label && screen.name) {
    const label = screen.name.toUpperCase();
    const boxW = Math.min(w - stroke * 2, ctx.measureText(label).width + pad * 2);
    ctx.fillStyle = screen.color;
    ctx.fillRect(stroke, stroke, boxW, fontPx + pad * 2);
    ctx.fillStyle = "#020617";
    ctx.fillText(label, stroke + pad, stroke + pad, boxW - pad * 2);
  }
  if (o.resolution) {
    const text = `${w} × ${h}  ·  X ${screen.x}  Y ${screen.y}`;
    const boxW = Math.min(w - stroke * 2, ctx.measureText(text).width + pad * 2);
    const y = h - stroke - fontPx - pad * 2;
    ctx.fillStyle = "rgba(2,6,23,0.75)";
    ctx.fillRect(stroke, y, boxW, fontPx + pad * 2);
    ctx.fillStyle = "#ffffff";
    ctx.fillText(text, stroke + pad, y + pad, boxW - pad * 2);
  }
  if (o.clock) drawClock(ctx, w, h, nowMs);
  ctx.restore();
};

// One scratch canvas for scrolling screens. Every render is synchronous, so
// screens drawn one after another can share it.
let motionSurface: HTMLCanvasElement | null = null;

const drawPattern = (ctx: CanvasRenderingContext2D, screen: SubScreen, info: PatternInfo) => {
  const { w, h } = screen;
  if (screen.pattern === LED_LAYOUT_PATTERN_ID && screen.ledLayout) {
    const layout = ledLayoutFor(screen.ledLayout);
    if (layout && ledModule) {
      ctx.save();
      ctx.scale(w / layout.W, h / layout.H);
      // The planner's pattern loops every LOOP_SECONDS; map this generator's
      // loop onto it so a recorded file still repeats without a jump.
      ledModule.drawTestPatternFrame(ctx, layout, info.progress * ledModule.LOOP_SECONDS);
      ctx.restore();
    } else {
      void loadLedModule();
      ctx.fillStyle = "#94a3b8";
      ctx.font = `bold ${Math.max(12, Math.round(Math.min(w, h) * 0.05))}px Arial`;
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      ctx.fillText("Loading LED layout…", w / 2, h / 2);
    }
  } else {
    const def = PATTERN_BY_ID.get(screen.pattern) ?? PATTERN_BY_ID.get("info")!;
    ctx.save();
    def.draw(ctx, w, h, resolveSettings(def, screen.settings[def.id]), info);
    ctx.restore();
  }
};

/**
 * Paints one screen into a context whose origin is the screen's top-left and
 * whose drawable area is exactly screen.w x screen.h.
 */
export const renderScreen = (ctx: CanvasRenderingContext2D, screen: SubScreen, time: number, rc: RenderContext) => {
  const { w, h } = screen;
  ctx.save();
  ctx.beginPath();
  ctx.rect(0, 0, w, h);
  ctx.clip();
  ctx.fillStyle = "#000000";
  ctx.fillRect(0, 0, w, h);
  const info = patternInfoFor(screen, time, rc);
  const { dx, dy } = motionOffset(screen.motion, info.progress, w, h);
  if (!dx && !dy) {
    drawPattern(ctx, screen, info);
  } else {
    // Moving: the finished pattern scrolls across the screen and wraps round,
    // a whole pixel at a time, so it stays as sharp as the still version.
    if (!motionSurface) motionSurface = document.createElement("canvas");
    if (motionSurface.width !== w || motionSurface.height !== h) {
      motionSurface.width = w;
      motionSurface.height = h;
    }
    const mctx = motionSurface.getContext("2d")!;
    mctx.setTransform(1, 0, 0, 1, 0, 0);
    mctx.fillStyle = "#000000";
    mctx.fillRect(0, 0, w, h);
    drawPattern(mctx, screen, info);
    for (const ox of dx ? [dx - w, dx] : [0]) {
      for (const oy of dy ? [dy - h, dy] : [0]) ctx.drawImage(motionSurface, ox, oy);
    }
  }
  drawOverlays(ctx, screen, Date.now());
  ctx.restore();
};

/** A fresh canvas holding just this screen at its native resolution. */
export const renderScreenToCanvas = (screen: SubScreen, time: number, rc: RenderContext): HTMLCanvasElement => {
  const c = document.createElement("canvas");
  c.width = screen.w;
  c.height = screen.h;
  renderScreen(c.getContext("2d")!, screen, time, rc);
  return c;
};

// --- whole canvas ---------------------------------------------------------------

/**
 * Per-screen working canvases for the full-canvas render, reused frame to
 * frame. Each owner (the editor preview, the output window, a recorder)
 * passes its own cache so they never resize each other's canvases.
 */
export type SurfaceCache = Map<string, HTMLCanvasElement>;

/**
 * `scale` is for the editor's preview only: the screens are still drawn at
 * their native resolution and then placed scaled. Every output and export
 * uses 1.
 */
export const renderComposite = (ctx: CanvasRenderingContext2D, config: GeneratorConfig, time: number, cache: SurfaceCache, scale = 1) => {
  const { w, h, background } = config.canvas;
  const rc: RenderContext = { loopSeconds: config.loopSeconds, canvas: { w, h } };
  ctx.save();
  ctx.setTransform(scale, 0, 0, scale, 0, 0);
  ctx.imageSmoothingEnabled = scale < 1;
  ctx.fillStyle = background;
  ctx.fillRect(0, 0, w, h);
  const live = new Set<string>();
  config.screens.forEach((screen) => {
    if (!screen.visible || screen.w <= 0 || screen.h <= 0) return;
    // Nothing of this screen lands on the canvas - skip the work.
    if (screen.x >= w || screen.y >= h || screen.x + screen.w <= 0 || screen.y + screen.h <= 0) return;
    live.add(screen.id);
    let surface = cache.get(screen.id);
    if (!surface) {
      surface = document.createElement("canvas");
      cache.set(screen.id, surface);
    }
    if (surface.width !== screen.w || surface.height !== screen.h) {
      surface.width = screen.w;
      surface.height = screen.h;
    }
    const sctx = surface.getContext("2d")!;
    sctx.setTransform(1, 0, 0, 1, 0, 0);
    renderScreen(sctx, screen, time, rc);
    ctx.drawImage(surface, screen.x, screen.y);
  });
  // Drop canvases for screens that are gone, so a long session doesn't hold onto memory.
  [...cache.keys()].forEach((id) => {
    if (!live.has(id)) cache.delete(id);
  });
  ctx.restore();
};

export const renderCompositeToCanvas = (config: GeneratorConfig, time: number): HTMLCanvasElement => {
  const c = document.createElement("canvas");
  c.width = config.canvas.w;
  c.height = config.canvas.h;
  renderComposite(c.getContext("2d")!, config, time, new Map());
  return c;
};

/** Pattern name for a file holding the whole canvas: the one pattern if they all share it, otherwise "Mixed". */
export const canvasPatternLabel = (config: GeneratorConfig, nameOf: (id: string) => string): string => {
  const ids = [...new Set(config.screens.filter((s) => s.visible).map((s) => s.pattern))];
  if (ids.length === 1) return nameOf(ids[0]);
  if (!ids.length) return "Background";
  return "Mixed-Patterns";
};
