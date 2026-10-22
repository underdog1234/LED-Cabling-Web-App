// ---------------------------------------------------------------------------
// Test pattern library for the generator.
//
// Every pattern is a pure draw function: given a 2D context sized exactly to
// one screen's native resolution, its settings and where it is in the loop,
// it paints that whole screen. Nothing reads the surrounding canvas, so the
// same function produces a sub-screen inside the full canvas and the same
// sub-screen exported on its own - pixel for pixel.
//
// Animation is always a function of `progress` (0..1 through one loop), never
// of wall-clock time directly, so every animated pattern returns to its exact
// first frame at the end of the loop and a recorded video repeats seamlessly.
// ---------------------------------------------------------------------------

import type { LedGrid, PatternSettings } from "./model";
import { drawFaces } from "./faces";
import { drawPhotoFaces } from "./photoFaces";

export type ParamDef =
  | { key: string; label: string; type: "number"; min: number; max: number; step?: number; default: number }
  | { key: string; label: string; type: "color"; default: string }
  | { key: string; label: string; type: "boolean"; default: boolean }
  | { key: string; label: string; type: "select"; options: Array<{ value: string; label: string }>; default: string };

export type PatternInfo = {
  /** 0..1 through one loop. */
  progress: number;
  /** Seconds into the current loop. */
  loopTime: number;
  loopSeconds: number;
  screen: { name: string; x: number; y: number; w: number; h: number; color: string };
  led: LedGrid | null;
  canvas: { w: number; h: number };
};

/** A tone the pattern makes at `at` seconds into its loop. */
export type Beep = { at: number; freq: number; duration: number; gain?: number };

export type PatternDef = {
  id: string;
  name: string;
  category: "Broadcast" | "Geometry" | "Greyscale & colour" | "Motion" | "LED" | "Reference";
  animated: boolean;
  params: ParamDef[];
  draw: (ctx: CanvasRenderingContext2D, w: number, h: number, s: Required<PatternSettings>, info: PatternInfo) => void;
  /** Sync tones, for patterns that have them. */
  beeps?: (s: Required<PatternSettings>, loopSeconds: number) => Beep[];
};

// --- helpers ----------------------------------------------------------------

const fontFamily = "Arial, Helvetica, sans-serif";

const fitFont = (ctx: CanvasRenderingContext2D, text: string, maxW: number, maxPx: number, weight = "bold"): number => {
  let px = Math.max(6, Math.floor(maxPx));
  ctx.font = `${weight} ${px}px ${fontFamily}`;
  const width = ctx.measureText(text).width;
  if (width > maxW && width > 0) {
    px = Math.max(6, Math.floor((px * maxW) / width));
    ctx.font = `${weight} ${px}px ${fontFamily}`;
  }
  return px;
};

const edges = (n: number, total: number) => Array.from({ length: n + 1 }, (_, i) => Math.round((i * total) / n));

const rgb = (r: number, g: number, b: number) => `rgb(${r},${g},${b})`;

/** A crisp 1px-aligned rectangle outline drawn with fills, so it never anti-aliases across two pixels. */
const pixelFrame = (ctx: CanvasRenderingContext2D, x: number, y: number, w: number, h: number, t: number) => {
  ctx.fillRect(x, y, w, t);
  ctx.fillRect(x, y + h - t, w, t);
  ctx.fillRect(x, y, t, h);
  ctx.fillRect(x + w - t, y, t, h);
};

const lineWidthFor = (w: number, h: number) => Math.max(2, Math.round(Math.min(w, h) * 0.0035));

const timecode = (seconds: number, fps: number) => {
  const totalFrames = Math.floor(seconds * fps + 1e-6);
  const ff = totalFrames % fps;
  const s = Math.floor(totalFrames / fps);
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${pad(Math.floor(s / 3600))}:${pad(Math.floor(s / 60) % 60)}:${pad(s % 60)}:${pad(ff)}`;
};

// --- LED moving pattern -----------------------------------------------------
// The LED planner's Moving Test Pattern, for a rectangular panel grid that
// needs no panel layout: an R/G/B checkerboard one panel per cell sliding
// sideways, a greyscale band sweeping corner to corner over it, a 1px outline
// and row/column reference on every panel, a heavier outline round the edge,
// the alignment cross and circle, and the wall's details in the middle.

const RGB_COLORS = ["#ff0000", "#00ff00", "#0000ff"];
const rgbTileCache = new Map<string, HTMLCanvasElement>();
const rgbTile = (pw: number, ph: number) => {
  const key = `${pw}x${ph}`;
  const cached = rgbTileCache.get(key);
  if (cached) return cached;
  const c = document.createElement("canvas");
  c.width = pw * 3;
  c.height = ph * 3;
  const x = c.getContext("2d")!;
  for (let row = 0; row < 3; row += 1) {
    for (let col = 0; col < 3; col += 1) {
      x.fillStyle = RGB_COLORS[(row + col) % 3];
      x.fillRect(col * pw, row * ph, pw, ph);
    }
  }
  if (rgbTileCache.size > 16) rgbTileCache.clear();
  rgbTileCache.set(key, c);
  return c;
};

const drawArrow = (ctx: CanvasRenderingContext2D, x: number, y: number, size: number, dir: "down" | "right") => {
  const lw = Math.min(3.5, Math.max(1.5, size * 0.16));
  const head = Math.max(lw * 1.8, size * 0.42);
  const hw = head * 0.9;
  ctx.save();
  ctx.strokeStyle = "#ffffff";
  ctx.fillStyle = "#ffffff";
  ctx.lineWidth = lw;
  ctx.lineCap = "round";
  ctx.beginPath();
  if (dir === "down") {
    const cx = x + size / 2;
    const tip = y + size - lw / 2;
    ctx.moveTo(cx, y + lw / 2);
    ctx.lineTo(cx, tip - head * 0.6);
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(cx, tip);
    ctx.lineTo(cx - hw / 2, tip - head);
    ctx.lineTo(cx + hw / 2, tip - head);
  } else {
    const cy = y + size / 2;
    const tip = x + size - lw / 2;
    ctx.moveTo(x + lw / 2, cy);
    ctx.lineTo(tip - head * 0.6, cy);
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(tip, cy);
    ctx.lineTo(tip - head, cy - hw / 2);
    ctx.lineTo(tip - head, cy + hw / 2);
  }
  ctx.closePath();
  ctx.fill();
  ctx.restore();
};

const drawLedMoving: PatternDef["draw"] = (ctx, w, h, s, info) => {
  const pw = Math.max(4, Math.round(Number(info.led?.panelPxW ?? s.panelPxW)));
  const ph = Math.max(4, Math.round(Number(info.led?.panelPxH ?? s.panelPxH)));
  const cols = Math.ceil(w / pw);
  const rows = Math.ceil(h / ph);
  const p = s.motion ? info.progress : 0;
  ctx.save();
  ctx.imageSmoothingEnabled = false;
  const tile = rgbTile(pw, ph);
  const pattern = ctx.createPattern(tile, "repeat")!;
  pattern.setTransform(new DOMMatrix().translate(Math.round(p * pw * 3), 0));
  ctx.fillStyle = pattern;
  ctx.fillRect(0, 0, w, h);
  // Greyscale band: brightness is a triangle wave of (x + y), one band across
  // the whole screen, sliding one full period per loop so it repeats with no
  // seam. Multiply keeps every panel's hue and only dims it.
  const period = w + h;
  const slide = p * period;
  const start = slide - period;
  const k = Math.ceil((w + h - start) / period) + 1;
  const end = start + period * k;
  const grad = ctx.createLinearGradient(start / 2, start / 2, end / 2, end / 2);
  for (let j = 0; j < k; j += 1) {
    grad.addColorStop(j / k, "#404040");
    grad.addColorStop((j + 0.5) / k, "#ffffff");
  }
  grad.addColorStop(1, "#404040");
  ctx.globalCompositeOperation = "multiply";
  ctx.fillStyle = grad;
  ctx.fillRect(0, 0, w, h);
  ctx.globalCompositeOperation = "source-over";

  // Panel outlines and references.
  ctx.fillStyle = "#ffffff";
  for (let c = 0; c <= cols; c += 1) ctx.fillRect(Math.min(w - 1, c * pw), 0, 1, h);
  for (let r = 0; r <= rows; r += 1) ctx.fillRect(0, Math.min(h - 1, r * ph), w, 1);
  if (s.showLabels) {
    const fontPx = Math.max(6, Math.floor(Math.min(Math.max(Math.min(pw, ph) * 0.08, 18), ph / 2.6, pw / 4)));
    const lineH = Math.round(fontPx * 1.15);
    const gap = Math.max(2, Math.round(fontPx * 0.25));
    const pad = Math.max(3, Math.round(Math.min(pw, ph) * 0.06));
    ctx.font = `bold ${fontPx}px ${fontFamily}`;
    ctx.textBaseline = "top";
    ctx.textAlign = "left";
    for (let r = 0; r < rows; r += 1) {
      for (let c = 0; c < cols; c += 1) {
        const x = c * pw + pad;
        const y = r * ph + pad;
        drawArrow(ctx, x, y, fontPx, "down");
        ctx.fillStyle = "#ffffff";
        ctx.fillText(String(r + 1), x + fontPx + gap, y);
        drawArrow(ctx, x, y + lineH, fontPx, "right");
        ctx.fillText(String(c + 1), x + fontPx + gap, y + lineH);
      }
    }
  }
  ctx.fillStyle = "#ffffff";
  pixelFrame(ctx, 0, 0, w, h, 3);
  if (s.showAlignment) {
    ctx.strokeStyle = "rgba(255,255,255,0.3)";
    ctx.lineWidth = lineWidthFor(w, h);
    ctx.beginPath();
    ctx.moveTo(0, 0);
    ctx.lineTo(w, h);
    ctx.moveTo(w, 0);
    ctx.lineTo(0, h);
    ctx.stroke();
    ctx.beginPath();
    ctx.arc(w / 2, h / 2, Math.min(w, h) / 2, 0, Math.PI * 2);
    ctx.stroke();
  }
  if (s.showInfo) {
    const fontPx = Math.max(16, Math.min(40, Math.round(w * 0.016)));
    const lineH = Math.round(fontPx * 1.4);
    const lines = [
      info.screen.name,
      `Resolution: ${w} x ${h} px`,
      `Panel: ${pw} x ${ph} px${info.led?.panelName ? ` (${info.led.panelName})` : ""}`,
      `Grid: ${info.led?.cols ?? cols} columns x ${info.led?.rows ?? rows} rows`,
    ];
    ctx.font = `bold ${fontPx}px ${fontFamily}`;
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillStyle = "#ffffff";
    const y0 = h / 2 - (lineH * (lines.length - 1)) / 2;
    lines.forEach((line, i) => ctx.fillText(line, w / 2, y0 + i * lineH));
  }
  ctx.restore();
};

// --- broadcast bars ---------------------------------------------------------

const drawSmpte: PatternDef["draw"] = (ctx, w, h, s) => {
  const v = s.level === "100" ? 255 : 191;
  const top = [rgb(v, v, v), rgb(v, v, 0), rgb(0, v, v), rgb(0, v, 0), rgb(v, 0, v), rgb(v, 0, 0), rgb(0, 0, v)];
  const mid = [rgb(0, 0, v), rgb(19, 19, 19), rgb(v, 0, v), rgb(19, 19, 19), rgb(0, v, v), rgb(19, 19, 19), rgb(v, v, v)];
  const xs = edges(7, w);
  const y1 = Math.round(h * 0.67);
  const y2 = Math.round(h * 0.75);
  for (let i = 0; i < 7; i += 1) {
    ctx.fillStyle = top[i];
    ctx.fillRect(xs[i], 0, xs[i + 1] - xs[i], y1);
    ctx.fillStyle = mid[i];
    ctx.fillRect(xs[i], y1, xs[i + 1] - xs[i], y2 - y1);
  }
  // Bottom: -I, 100% white, +Q, black, then PLUGE (below black, black, above black), black.
  const bar = w / 7;
  const b = [
    { x1: 0, x2: bar * 1.25, c: rgb(0, 33, 76) },
    { x1: bar * 1.25, x2: bar * 2.5, c: rgb(255, 255, 255) },
    { x1: bar * 2.5, x2: bar * 3.75, c: rgb(50, 0, 106) },
    { x1: bar * 3.75, x2: bar * 5, c: rgb(19, 19, 19) },
    { x1: bar * 5, x2: bar * (5 + 1 / 3), c: rgb(9, 9, 9) },
    { x1: bar * (5 + 1 / 3), x2: bar * (5 + 2 / 3), c: rgb(19, 19, 19) },
    { x1: bar * (5 + 2 / 3), x2: bar * 6, c: rgb(29, 29, 29) },
    { x1: bar * 6, x2: w, c: rgb(19, 19, 19) },
  ];
  b.forEach((seg) => {
    ctx.fillStyle = seg.c;
    const x1 = Math.round(seg.x1);
    ctx.fillRect(x1, y2, Math.round(seg.x2) - x1, h - y2);
  });
};

const drawEbu: PatternDef["draw"] = (ctx, w, h, s) => {
  const v = s.level === "100" ? 255 : 191;
  const bars = [rgb(255, 255, 255), rgb(v, v, 0), rgb(0, v, v), rgb(0, v, 0), rgb(v, 0, v), rgb(v, 0, 0), rgb(0, 0, v), rgb(0, 0, 0)];
  const xs = edges(8, w);
  bars.forEach((c, i) => {
    ctx.fillStyle = c;
    ctx.fillRect(xs[i], 0, xs[i + 1] - xs[i], h);
  });
};

const drawColourChecker: PatternDef["draw"] = (ctx, w, h, s) => {
  const patches = [
    [115, 82, 68], [194, 150, 130], [98, 122, 157], [87, 108, 67], [133, 128, 177], [103, 189, 170],
    [214, 126, 44], [80, 91, 166], [193, 90, 99], [94, 60, 108], [157, 188, 64], [224, 163, 46],
    [56, 61, 150], [70, 148, 73], [175, 54, 60], [231, 199, 31], [187, 86, 149], [8, 133, 161],
    [243, 243, 242], [200, 200, 200], [160, 160, 160], [122, 122, 121], [85, 85, 85], [52, 52, 52],
  ];
  const names = [
    "Dark skin", "Light skin", "Blue sky", "Foliage", "Blue flower", "Bluish green",
    "Orange", "Purplish blue", "Moderate red", "Purple", "Yellow green", "Orange yellow",
    "Blue", "Green", "Red", "Yellow", "Magenta", "Cyan",
    "White", "Neutral 8", "Neutral 6.5", "Neutral 5", "Neutral 3.5", "Black",
  ];
  ctx.fillStyle = String(s.background);
  ctx.fillRect(0, 0, w, h);
  const xs = edges(6, w);
  const ys = edges(4, h);
  const gap = Math.max(2, Math.round(Math.min(w, h) * 0.012));
  patches.forEach(([r, g, b], i) => {
    const c = i % 6;
    const row = Math.floor(i / 6);
    const x = xs[c] + gap;
    const y = ys[row] + gap;
    const pw = xs[c + 1] - xs[c] - gap * 2;
    const ph = ys[row + 1] - ys[row] - gap * 2;
    ctx.fillStyle = rgb(r, g, b);
    ctx.fillRect(x, y, pw, ph);
    if (s.labels) {
      const px = fitFont(ctx, names[i], pw * 0.9, ph * 0.12, "bold");
      ctx.fillStyle = r + g + b > 380 ? "#000000" : "#ffffff";
      ctx.textAlign = "left";
      ctx.textBaseline = "bottom";
      ctx.fillText(names[i], x + px * 0.4, y + ph - px * 0.3);
    }
  });
};

// --- geometry ---------------------------------------------------------------

const drawGrid: PatternDef["draw"] = (ctx, w, h, s) => {
  const spacing = Math.max(2, Math.round(Number(s.spacing)));
  const lw = Math.max(1, Math.round(Number(s.lineWidth)));
  ctx.fillStyle = String(s.background);
  ctx.fillRect(0, 0, w, h);
  ctx.fillStyle = String(s.lineColor);
  const cx = Math.floor(w / 2);
  const cy = Math.floor(h / 2);
  const centred = s.origin === "centre";
  const x0 = centred ? cx % spacing : 0;
  const y0 = centred ? cy % spacing : 0;
  for (let x = x0; x < w; x += spacing) ctx.fillRect(x - (centred ? Math.floor(lw / 2) : 0), 0, lw, h);
  for (let y = y0; y < h; y += spacing) ctx.fillRect(0, y - (centred ? Math.floor(lw / 2) : 0), w, lw);
  // The last pixel column and row, so every edge of the output is lit.
  pixelFrame(ctx, 0, 0, w, h, lw);
  if (s.centreLines) {
    ctx.fillStyle = String(s.centreColor);
    const t = Math.max(lw, 2);
    ctx.fillRect(cx - Math.floor(t / 2), 0, t, h);
    ctx.fillRect(0, cy - Math.floor(t / 2), w, t);
  }
  if (s.circle) {
    ctx.strokeStyle = String(s.centreColor);
    ctx.lineWidth = Math.max(lw, 2);
    ctx.beginPath();
    ctx.arc(w / 2, h / 2, Math.min(w, h) / 2 - ctx.lineWidth, 0, Math.PI * 2);
    ctx.stroke();
  }
};

const drawGeometry: PatternDef["draw"] = (ctx, w, h, s) => {
  ctx.fillStyle = "#000000";
  ctx.fillRect(0, 0, w, h);
  const lw = lineWidthFor(w, h);
  const minD = Math.min(w, h);
  ctx.strokeStyle = "#ffffff";
  ctx.fillStyle = "#ffffff";
  ctx.lineWidth = lw;
  pixelFrame(ctx, 0, 0, w, h, 1);
  // Diagonals and centre cross.
  ctx.strokeStyle = "rgba(255,255,255,0.5)";
  ctx.beginPath();
  ctx.moveTo(0, 0);
  ctx.lineTo(w, h);
  ctx.moveTo(w, 0);
  ctx.lineTo(0, h);
  ctx.stroke();
  ctx.strokeStyle = "#ffffff";
  ctx.beginPath();
  ctx.moveTo(w / 2, 0);
  ctx.lineTo(w / 2, h);
  ctx.moveTo(0, h / 2);
  ctx.lineTo(w, h / 2);
  ctx.stroke();
  // Circles: one as big as the short side, one in each corner.
  ctx.beginPath();
  ctx.arc(w / 2, h / 2, minD / 2 - lw, 0, Math.PI * 2);
  ctx.stroke();
  ctx.beginPath();
  ctx.arc(w / 2, h / 2, minD / 4, 0, Math.PI * 2);
  ctx.stroke();
  const r = minD / 8;
  [[r, r], [w - r, r], [r, h - r], [w - r, h - r]].forEach(([x, y]) => {
    ctx.beginPath();
    ctx.arc(x, y, r - lw, 0, Math.PI * 2);
    ctx.stroke();
  });
  if (s.safeAreas) {
    ctx.setLineDash([lw * 4, lw * 3]);
    [[0.93, "#22d3ee", "Action safe 93%"], [0.9, "#facc15", "Title safe 90%"]].forEach(([f, c, label]) => {
      const fw = w * Number(f);
      const fh = h * Number(f);
      ctx.strokeStyle = String(c);
      ctx.strokeRect((w - fw) / 2, (h - fh) / 2, fw, fh);
      ctx.fillStyle = String(c);
      const px = Math.max(10, Math.round(minD * 0.018));
      ctx.font = `bold ${px}px ${fontFamily}`;
      ctx.textAlign = "left";
      ctx.textBaseline = "top";
      ctx.fillText(String(label), (w - fw) / 2 + lw * 2, (h - fh) / 2 + lw * 2 + (f === 0.9 ? px * 1.3 : 0));
    });
    ctx.setLineDash([]);
  }
  if (s.aspectFrames) {
    const frames: Array<[number, string, string]> = [[16 / 9, "16:9", "#a3e635"], [4 / 3, "4:3", "#f472b6"], [2.39, "2.39:1", "#c084fc"]];
    frames.forEach(([ratio, label, c]) => {
      let fw = w;
      let fh = Math.round(w / ratio);
      if (fh > h) {
        fh = h;
        fw = Math.round(h * ratio);
      }
      ctx.strokeStyle = c;
      ctx.lineWidth = Math.max(1, Math.round(lw / 2));
      ctx.strokeRect((w - fw) / 2 + 0.5, (h - fh) / 2 + 0.5, fw - 1, fh - 1);
      const px = Math.max(10, Math.round(minD * 0.016));
      ctx.font = `bold ${px}px ${fontFamily}`;
      ctx.fillStyle = c;
      ctx.textAlign = "right";
      ctx.textBaseline = "bottom";
      ctx.fillText(label, (w + fw) / 2 - px * 0.5, (h + fh) / 2 - px * 0.3);
    });
  }
};

// --- greyscale & colour -----------------------------------------------------

const drawGreySteps: PatternDef["draw"] = (ctx, w, h, s) => {
  const n = Math.max(2, Math.min(64, Math.round(Number(s.steps))));
  const vertical = s.direction === "vertical";
  const cuts = edges(n, vertical ? h : w);
  for (let i = 0; i < n; i += 1) {
    const v = Math.round((i * 255) / (n - 1));
    ctx.fillStyle = rgb(v, v, v);
    if (vertical) ctx.fillRect(0, cuts[i], w, cuts[i + 1] - cuts[i]);
    else ctx.fillRect(cuts[i], 0, cuts[i + 1] - cuts[i], h);
    if (s.labels) {
      const label = `${Math.round((i * 100) / (n - 1))}%`;
      const size = vertical ? cuts[i + 1] - cuts[i] : cuts[i + 1] - cuts[i];
      const px = fitFont(ctx, label, vertical ? w * 0.2 : size * 0.8, Math.min(size * 0.5, Math.min(w, h) * 0.05));
      ctx.fillStyle = v > 127 ? "#000000" : "#ffffff";
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      if (vertical) ctx.fillText(label, px * 2.5, (cuts[i] + cuts[i + 1]) / 2);
      else ctx.fillText(label, (cuts[i] + cuts[i + 1]) / 2, h / 2);
    }
  }
};

const drawGradient: PatternDef["draw"] = (ctx, w, h, s) => {
  const vertical = s.direction === "vertical";
  const channels: Record<string, string[]> = {
    grey: ["#ffffff"],
    red: ["#ff0000"],
    green: ["#00ff00"],
    blue: ["#0000ff"],
    rgbw: ["#ffffff", "#ff0000", "#00ff00", "#0000ff"],
  };
  const list = channels[String(s.channel)] ?? channels.grey;
  const bands = edges(list.length, vertical ? w : h);
  list.forEach((c, i) => {
    const g = vertical ? ctx.createLinearGradient(0, 0, 0, h) : ctx.createLinearGradient(0, 0, w, 0);
    g.addColorStop(0, "#000000");
    g.addColorStop(1, c);
    ctx.fillStyle = g;
    if (vertical) ctx.fillRect(bands[i], 0, bands[i + 1] - bands[i], h);
    else ctx.fillRect(0, bands[i], w, bands[i + 1] - bands[i]);
  });
};

const drawSolid: PatternDef["draw"] = (ctx, w, h, s) => {
  ctx.fillStyle = String(s.color);
  ctx.fillRect(0, 0, w, h);
};

const CYCLE = [
  ["#ffffff", "White"], ["#ff0000", "Red"], ["#00ff00", "Green"], ["#0000ff", "Blue"],
  ["#00ffff", "Cyan"], ["#ff00ff", "Magenta"], ["#ffff00", "Yellow"], ["#000000", "Black"],
];

const drawColourCycle: PatternDef["draw"] = (ctx, w, h, s, info) => {
  const list = s.includeSecondaries ? CYCLE : CYCLE.filter(([, n]) => ["White", "Red", "Green", "Blue", "Black"].includes(n));
  const idx = Math.min(list.length - 1, Math.floor(info.progress * list.length));
  ctx.fillStyle = list[idx][0];
  ctx.fillRect(0, 0, w, h);
  if (s.labels) {
    const px = fitFont(ctx, list[idx][1], w * 0.5, Math.min(w, h) * 0.06);
    ctx.fillStyle = list[idx][1] === "Black" || list[idx][1] === "Blue" ? "#808080" : "#000000";
    ctx.textAlign = "center";
    ctx.textBaseline = "bottom";
    ctx.fillText(list[idx][1], w / 2, h - px);
  }
};

// --- motion -----------------------------------------------------------------

const drawMovingBar: PatternDef["draw"] = (ctx, w, h, s, info) => {
  ctx.fillStyle = String(s.background);
  ctx.fillRect(0, 0, w, h);
  const vertical = s.direction === "vertical";
  const size = Math.max(1, Math.round(Number(s.barWidth)));
  const travel = (vertical ? h : w) + size;
  const cycles = Math.max(1, Math.round(Number(s.cycles)));
  const pos = Math.round(((info.progress * cycles) % 1) * travel) - size;
  ctx.fillStyle = String(s.color);
  if (vertical) ctx.fillRect(0, pos, w, size);
  else ctx.fillRect(pos, 0, size, h);
};

const drawTimecode: PatternDef["draw"] = (ctx, w, h, s, info) => {
  const fps = Math.max(1, Math.round(Number(s.fps)));
  ctx.fillStyle = "#101010";
  ctx.fillRect(0, 0, w, h);
  const minD = Math.min(w, h);
  // Sync flash on the first frame of every loop: all four corners and the
  // border, so a camera or a second output can be lined up against it.
  const frame = Math.floor(info.loopTime * fps + 1e-6);
  const flash = frame === 0;
  const box = Math.round(minD * 0.12);
  ctx.fillStyle = flash ? "#ffffff" : "#303030";
  [[0, 0], [w - box, 0], [0, h - box], [w - box, h - box]].forEach(([x, y]) => ctx.fillRect(x, y, box, box));
  // Sweep hand: one turn per loop.
  const r = minD * 0.32;
  ctx.strokeStyle = "#3f3f46";
  ctx.lineWidth = Math.max(2, minD * 0.01);
  ctx.beginPath();
  ctx.arc(w / 2, h / 2, r, 0, Math.PI * 2);
  ctx.stroke();
  const ticks = Math.max(2, Math.round(info.loopSeconds));
  ctx.strokeStyle = "#a1a1aa";
  for (let i = 0; i < ticks; i += 1) {
    const a = (i / ticks) * Math.PI * 2 - Math.PI / 2;
    ctx.beginPath();
    ctx.moveTo(w / 2 + Math.cos(a) * r * 0.9, h / 2 + Math.sin(a) * r * 0.9);
    ctx.lineTo(w / 2 + Math.cos(a) * r, h / 2 + Math.sin(a) * r);
    ctx.stroke();
  }
  ctx.fillStyle = "rgba(56,189,248,0.35)";
  ctx.beginPath();
  ctx.moveTo(w / 2, h / 2);
  ctx.arc(w / 2, h / 2, r, -Math.PI / 2, -Math.PI / 2 + info.progress * Math.PI * 2);
  ctx.closePath();
  ctx.fill();
  const tc = timecode(info.loopTime, fps);
  fitFont(ctx, "00:00:00:00", r * 1.6, minD * 0.1, "bold");
  ctx.font = ctx.font.replace(fontFamily, "'Courier New', monospace");
  ctx.fillStyle = "#ffffff";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText(tc, w / 2, h / 2);
  const small = Math.max(10, Math.round(minD * 0.035));
  ctx.font = `bold ${small}px ${fontFamily}`;
  ctx.fillStyle = "#a1a1aa";
  ctx.fillText(`Frame ${frame + 1} / ${Math.round(info.loopSeconds * fps)} @ ${fps} fps`, w / 2, h / 2 + minD * 0.09);
};

// --- AV sync ---------------------------------------------------------------------
// Flash and beep together, every interval, for lining up sound with picture
// (lip sync) through a camera, a processor or a whole show system. Between
// flashes a sweep counts round the circle one segment per frame, and the
// frame ruler underneath lights the frame offset from the flash: film the
// screen with sound, find the frame where the beep lands, and the lit cell
// reads the error in frames.

const avSyncTiming = (s: Required<PatternSettings>, loopTime: number) => {
  const fps = Math.max(1, Math.round(Number(s.fps)));
  const interval = Math.max(1, Math.round(Number(s.interval)));
  const perInterval = fps * interval;
  const frame = Math.floor(loopTime * fps + 1e-6);
  const inInterval = frame % perInterval;
  const flashFrames = Math.max(1, Math.min(perInterval, Math.round(Number(s.flashFrames))));
  // Offset from the nearest flash, in frames: negative before it, positive after.
  const half = Math.floor(perInterval / 2);
  const offset = ((inInterval + half) % perInterval) - half;
  return { fps, interval, perInterval, frame, inInterval, flashFrames, offset, flash: inInterval < flashFrames };
};

const drawAvSync: PatternDef["draw"] = (ctx, w, h, s, info) => {
  const t = avSyncTiming(s, info.loopTime);
  const minD = Math.min(w, h);
  const cx = w / 2;
  const cy = h * 0.42;
  const r = Math.min(minD * 0.3, w * 0.3, h * 0.3);
  const accent = String(s.color);
  ctx.fillStyle = "#000000";
  ctx.fillRect(0, 0, w, h);

  // Flash: the whole border and four corner blocks, plus the centre disc below.
  const box = Math.round(minD * 0.1);
  ctx.fillStyle = t.flash ? accent : "#1f1f1f";
  [[0, 0], [w - box, 0], [0, h - box], [w - box, h - box]].forEach(([x, y]) => ctx.fillRect(x, y, box, box));
  if (t.flash) pixelFrame(ctx, 0, 0, w, h, Math.max(4, Math.round(minD * 0.012)));

  // Countdown ring: one segment per frame of the interval.
  const segs = t.perInterval;
  const ringW = Math.max(4, r * 0.16);
  const gapA = segs > 60 ? 0 : (Math.PI * 2) / segs * 0.18;
  for (let i = 0; i < segs; i += 1) {
    const a0 = -Math.PI / 2 + (i / segs) * Math.PI * 2 + gapA / 2;
    const a1 = -Math.PI / 2 + ((i + 1) / segs) * Math.PI * 2 - gapA / 2;
    ctx.strokeStyle = i === t.inInterval ? "#ffffff" : i < t.inInterval ? "#52525b" : "#27272a";
    ctx.lineWidth = ringW;
    ctx.beginPath();
    ctx.arc(cx, cy, r - ringW / 2, a0, a1);
    ctx.stroke();
  }
  // Second ticks round the outside.
  ctx.fillStyle = "#a1a1aa";
  for (let i = 0; i < t.interval; i += 1) {
    const a = -Math.PI / 2 + (i / t.interval) * Math.PI * 2;
    ctx.beginPath();
    ctx.arc(cx + Math.cos(a) * (r + ringW * 0.6), cy + Math.sin(a) * (r + ringW * 0.6), Math.max(2, ringW * 0.22), 0, Math.PI * 2);
    ctx.fill();
  }
  // Centre disc: white on the flash, otherwise the frame count.
  const inner = r - ringW * 1.6;
  ctx.fillStyle = t.flash ? accent : "#0a0a0a";
  ctx.beginPath();
  ctx.arc(cx, cy, inner, 0, Math.PI * 2);
  ctx.fill();
  if (!t.flash) {
    const label = String(t.inInterval);
    fitFont(ctx, "00", inner * 1.3, inner * 0.9);
    ctx.fillStyle = "#ffffff";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText(label, cx, cy + inner * 0.04);
  }

  // Frame ruler: offset from the flash, -N .. +N frames.
  const n = Math.max(2, Math.min(12, Math.floor(t.perInterval / 2) - 1));
  const cells = n * 2 + 1;
  const rulerW = Math.min(w * 0.9, cells * minD * 0.08);
  const cw = rulerW / cells;
  const ch = Math.max(10, Math.min(cw * 1.1, h * 0.09));
  const rx = (w - rulerW) / 2;
  const ry = Math.min(h - ch - minD * 0.12, cy + r + minD * 0.06);
  const labelPx = Math.max(8, Math.floor(Math.min(cw * 0.42, ch * 0.5)));
  ctx.font = `bold ${labelPx}px ${fontFamily}`;
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  for (let i = 0; i < cells; i += 1) {
    const d = i - n;
    const lit = d === t.offset;
    const x = rx + i * cw;
    ctx.fillStyle = lit ? (d === 0 ? accent : "#f59e0b") : d === 0 ? "#3f3f46" : "#18181b";
    ctx.fillRect(Math.round(x + 1), Math.round(ry), Math.round(cw - 2), Math.round(ch));
    ctx.fillStyle = lit ? "#000000" : "#a1a1aa";
    ctx.fillText(d > 0 ? `+${d}` : String(d), x + cw / 2, ry + ch / 2);
  }
  const small = Math.max(10, Math.round(minD * 0.03));
  ctx.font = `bold ${small}px ${fontFamily}`;
  ctx.fillStyle = "#a1a1aa";
  ctx.fillText("frames from flash", w / 2, ry + ch + small * 0.9);

  // Timecode and settings, top centre.
  const tc = timecode(info.loopTime, t.fps);
  fitFont(ctx, "00:00:00:00", w * 0.4, minD * 0.07);
  ctx.font = ctx.font.replace(fontFamily, "'Courier New', monospace");
  ctx.fillStyle = "#ffffff";
  ctx.textBaseline = "top";
  ctx.fillText(tc, w / 2, Math.round(minD * 0.03));
  ctx.font = `bold ${small}px ${fontFamily}`;
  ctx.fillStyle = "#71717a";
  ctx.textBaseline = "bottom";
  ctx.fillText(`AV SYNC · ${t.fps} fps · flash + ${Number(s.beepFreq)} Hz beep every ${t.interval} s`, w / 2, h - Math.round(minD * 0.03));
};

const avSyncBeeps: NonNullable<PatternDef["beeps"]> = (s, loopSeconds) => {
  if (!s.beep) return [];
  const fps = Math.max(1, Math.round(Number(s.fps)));
  const interval = Math.max(1, Math.round(Number(s.interval)));
  const flashFrames = Math.max(1, Math.round(Number(s.flashFrames)));
  const out: Beep[] = [];
  for (let at = 0; at < loopSeconds - 1e-6; at += interval) out.push({ at, freq: Number(s.beepFreq), duration: flashFrames / fps });
  return out;
};

// --- pixel structure --------------------------------------------------------

const drawPixel: PatternDef["draw"] = (ctx, w, h, s) => {
  ctx.fillStyle = "#000000";
  ctx.fillRect(0, 0, w, h);
  const on = Math.max(1, Math.round(Number(s.pitch)));
  const period = on * 2;
  ctx.fillStyle = String(s.color);
  const mode = String(s.mode);
  if (mode === "vertical") {
    for (let x = 0; x < w; x += period) ctx.fillRect(x, 0, on, h);
  } else if (mode === "horizontal") {
    for (let y = 0; y < h; y += period) ctx.fillRect(0, y, w, on);
  } else {
    // Checker or dots: build one tile, repeat it.
    const tile = document.createElement("canvas");
    tile.width = period;
    tile.height = period;
    const t = tile.getContext("2d")!;
    t.fillStyle = String(s.color);
    t.fillRect(0, 0, on, on);
    if (mode === "checker") t.fillRect(on, on, on, on);
    ctx.fillStyle = ctx.createPattern(tile, "repeat")!;
    ctx.fillRect(0, 0, w, h);
  }
};

const zoneCache = new Map<string, HTMLCanvasElement>();
const drawZonePlate: PatternDef["draw"] = (ctx, w, h) => {
  const key = `${w}x${h}`;
  let c = zoneCache.get(key);
  if (!c) {
    c = document.createElement("canvas");
    c.width = w;
    c.height = h;
    const z = c.getContext("2d")!;
    const img = z.createImageData(w, h);
    const rMax = Math.hypot(w / 2, h / 2);
    // Frequency rises linearly with radius and reaches 0.5 cycles/pixel (the
    // pixel grid's own limit) at the corners.
    const k = Math.PI / (2 * rMax);
    for (let y = 0; y < h; y += 1) {
      const dy = y - h / 2 + 0.5;
      for (let x = 0; x < w; x += 1) {
        const dx = x - w / 2 + 0.5;
        const v = Math.round(127.5 + 127.5 * Math.cos(k * (dx * dx + dy * dy)));
        const i = (y * w + x) * 4;
        img.data[i] = v;
        img.data[i + 1] = v;
        img.data[i + 2] = v;
        img.data[i + 3] = 255;
      }
    }
    z.putImageData(img, 0, 0);
    if (zoneCache.size > 6) zoneCache.clear();
    zoneCache.set(key, c);
  }
  ctx.drawImage(c, 0, 0);
};

// --- reference / identification ----------------------------------------------

const drawInfo: PatternDef["draw"] = (ctx, w, h, s, info) => {
  const minD = Math.min(w, h);
  const accent = info.screen.color;
  ctx.fillStyle = String(s.background);
  ctx.fillRect(0, 0, w, h);
  // Fine grid so scaling or cropping shows up at a glance.
  const step = Math.max(8, Math.round(minD / 12));
  ctx.fillStyle = "rgba(255,255,255,0.12)";
  for (let x = 0; x < w; x += step) ctx.fillRect(x, 0, 1, h);
  for (let y = 0; y < h; y += step) ctx.fillRect(0, y, w, 1);
  // Corner markers: an L in each corner, exactly on the edge pixels.
  const arm = Math.round(minD * 0.12);
  const t = Math.max(2, Math.round(minD * 0.008));
  ctx.fillStyle = accent;
  [[0, 0, 1, 1], [w, 0, -1, 1], [0, h, 1, -1], [w, h, -1, -1]].forEach(([x, y, sx, sy]) => {
    ctx.fillRect(sx > 0 ? x : x - arm, sy > 0 ? y : y - t, arm, t);
    ctx.fillRect(sx > 0 ? x : x - t, sy > 0 ? y : y - arm, t, arm);
  });
  ctx.fillStyle = "#ffffff";
  pixelFrame(ctx, 0, 0, w, h, 1);
  // Text block.
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  const namePx = fitFont(ctx, info.screen.name || "Output", w * 0.85, minD * 0.16);
  ctx.fillStyle = accent;
  ctx.fillText(info.screen.name || "Output", w / 2, h / 2 - namePx * 0.9);
  const resText = `${w} × ${h}`;
  const resPx = fitFont(ctx, resText, w * 0.8, minD * 0.11);
  ctx.fillStyle = "#ffffff";
  ctx.fillText(resText, w / 2, h / 2 + resPx * 0.35);
  const ratio = (() => {
    let a = w;
    let b = h;
    while (b) [a, b] = [b, a % b];
    const x = w / a;
    const y = h / a;
    return x <= 64 && y <= 64 ? `${x}:${y}` : `${(w / h).toFixed(3)}:1`;
  })();
  const detail = `Aspect ${ratio}   ·   Canvas X ${info.screen.x}  Y ${info.screen.y}`;
  const detailPx = fitFont(ctx, detail, w * 0.85, minD * 0.045, "normal");
  ctx.fillStyle = "#cbd5e1";
  ctx.fillText(detail, w / 2, h / 2 + resPx * 0.35 + resPx * 0.6 + detailPx);
  if (s.running) {
    // A running block proves the output is live rather than a frozen frame.
    const barW = Math.round(w * 0.6);
    const barH = Math.max(4, Math.round(minD * 0.02));
    const bx = Math.round((w - barW) / 2);
    const by = Math.round(h * 0.82);
    ctx.fillStyle = "rgba(255,255,255,0.15)";
    ctx.fillRect(bx, by, barW, barH);
    ctx.fillStyle = accent;
    const blockW = Math.max(barH, Math.round(barW * 0.08));
    ctx.fillRect(bx + Math.round(info.progress * (barW - blockW)), by, blockW, barH);
  }
};

const drawBlank: PatternDef["draw"] = (ctx, w, h) => {
  ctx.fillStyle = "#000000";
  ctx.fillRect(0, 0, w, h);
};

const drawFacesPattern: PatternDef["draw"] = (ctx, w, h, s, info) => {
  drawFaces(ctx, w, h, {
    count: Math.max(1, Math.round(Number(s.count))),
    background: String(s.background),
    labels: Boolean(s.labels),
    animate: Boolean(s.animate),
    seed: Math.round(Number(s.seed)),
    toneStrip: Boolean(s.toneStrip),
    progress: info.progress,
  });
};

// --- registry -----------------------------------------------------------------

const levelOptions = [
  { value: "75", label: "75%" },
  { value: "100", label: "100%" },
];

const fpsOptions = ["24", "25", "30", "48", "50", "60"].map((v) => ({ value: v, label: `${v} fps` }));

export const LED_LAYOUT_PATTERN_ID = "led-layout";

export const PATTERNS: PatternDef[] = [
  { id: "info", name: "Screen ID & Resolution", category: "Reference", animated: true, params: [
    { key: "background", label: "Background", type: "color", default: "#0f172a" },
    { key: "running", label: "Running indicator", type: "boolean", default: true },
  ], draw: drawInfo },
  { id: "smpte", name: "SMPTE Colour Bars", category: "Broadcast", animated: false, params: [
    { key: "level", label: "Level", type: "select", options: levelOptions, default: "75" },
  ], draw: drawSmpte },
  { id: "ebu", name: "EBU Colour Bars", category: "Broadcast", animated: false, params: [
    { key: "level", label: "Colour level", type: "select", options: levelOptions, default: "75" },
  ], draw: drawEbu },
  { id: "colour-checker", name: "Colour Checker Patches", category: "Greyscale & colour", animated: false, params: [
    { key: "background", label: "Background", type: "color", default: "#202020" },
    { key: "labels", label: "Patch names", type: "boolean", default: true },
  ], draw: drawColourChecker },
  { id: "grid", name: "Grid / Crosshatch", category: "Geometry", animated: false, params: [
    { key: "spacing", label: "Spacing (px)", type: "number", min: 2, max: 4096, default: 64 },
    { key: "lineWidth", label: "Line width (px)", type: "number", min: 1, max: 64, default: 1 },
    { key: "origin", label: "Grid origin", type: "select", options: [{ value: "centre", label: "Centre" }, { value: "top-left", label: "Top-left" }], default: "centre" },
    { key: "lineColor", label: "Line colour", type: "color", default: "#ffffff" },
    { key: "background", label: "Background", type: "color", default: "#000000" },
    { key: "centreLines", label: "Centre lines", type: "boolean", default: true },
    { key: "circle", label: "Centre circle", type: "boolean", default: true },
    { key: "centreColor", label: "Centre colour", type: "color", default: "#ff3b30" },
  ], draw: drawGrid },
  { id: "geometry", name: "Circles, Safe Areas & Aspect", category: "Geometry", animated: false, params: [
    { key: "safeAreas", label: "Safe areas", type: "boolean", default: true },
    { key: "aspectFrames", label: "Aspect frames", type: "boolean", default: true },
  ], draw: drawGeometry },
  { id: "checker", name: "Checkerboard", category: "Geometry", animated: true, params: [
    { key: "cell", label: "Cell size (px)", type: "number", min: 1, max: 4096, default: 60 },
    { key: "colorA", label: "Colour A", type: "color", default: "#ffffff" },
    { key: "colorB", label: "Colour B", type: "color", default: "#000000" },
    { key: "invert", label: "Invert every half loop", type: "boolean", default: false },
  ], draw: (ctx, w, h, s, info) => {
    const cell = Math.max(1, Math.round(Number(s.cell)));
    const flip = s.invert && info.progress >= 0.5;
    const a = flip ? String(s.colorB) : String(s.colorA);
    const b = flip ? String(s.colorA) : String(s.colorB);
    ctx.fillStyle = b;
    ctx.fillRect(0, 0, w, h);
    ctx.fillStyle = a;
    for (let y = 0, r = 0; y < h; y += cell, r += 1) {
      for (let x = (r % 2) * cell; x < w; x += cell * 2) ctx.fillRect(x, y, cell, cell);
    }
  } },
  { id: "grey-steps", name: "Greyscale Steps", category: "Greyscale & colour", animated: false, params: [
    { key: "steps", label: "Steps", type: "number", min: 2, max: 64, default: 11 },
    { key: "direction", label: "Direction", type: "select", options: [{ value: "horizontal", label: "Left to right" }, { value: "vertical", label: "Top to bottom" }], default: "horizontal" },
    { key: "labels", label: "Labels", type: "boolean", default: true },
  ], draw: drawGreySteps },
  { id: "gradient", name: "Gradient Ramps", category: "Greyscale & colour", animated: false, params: [
    { key: "channel", label: "Channel", type: "select", options: [
      { value: "grey", label: "Grey" }, { value: "red", label: "Red" }, { value: "green", label: "Green" }, { value: "blue", label: "Blue" }, { value: "rgbw", label: "White, R, G, B stacked" },
    ], default: "rgbw" },
    { key: "direction", label: "Ramp", type: "select", options: [{ value: "horizontal", label: "Horizontal" }, { value: "vertical", label: "Vertical" }], default: "horizontal" },
  ], draw: drawGradient },
  { id: "solid", name: "Solid Colour Field", category: "Greyscale & colour", animated: false, params: [
    { key: "color", label: "Colour", type: "color", default: "#ffffff" },
  ], draw: drawSolid },
  { id: "colour-cycle", name: "Colour Field Cycle", category: "Greyscale & colour", animated: true, params: [
    { key: "includeSecondaries", label: "Include cyan, magenta, yellow", type: "boolean", default: true },
    { key: "labels", label: "Labels", type: "boolean", default: false },
  ], draw: drawColourCycle },
  { id: "moving-bar", name: "Moving Bar", category: "Motion", animated: true, params: [
    { key: "direction", label: "Direction", type: "select", options: [{ value: "horizontal", label: "Left to right" }, { value: "vertical", label: "Top to bottom" }], default: "horizontal" },
    { key: "barWidth", label: "Bar width (px)", type: "number", min: 1, max: 4096, default: 40 },
    { key: "cycles", label: "Passes per loop", type: "number", min: 1, max: 100, default: 2 },
    { key: "color", label: "Bar colour", type: "color", default: "#ffffff" },
    { key: "background", label: "Background", type: "color", default: "#000000" },
  ], draw: drawMovingBar },
  { id: "timecode", name: "Timecode & Sync Flash", category: "Motion", animated: true, params: [
    { key: "fps", label: "Frame rate", type: "select", options: fpsOptions, default: "25" },
    { key: "beep", label: "Beep with the flash", type: "boolean", default: true },
    { key: "beepFreq", label: "Beep pitch (Hz)", type: "number", min: 100, max: 8000, default: 1000 },
  ], draw: drawTimecode, beeps: (s) => (s.beep ? [{ at: 0, freq: Number(s.beepFreq), duration: 1 / Math.max(1, Math.round(Number(s.fps))) }] : []) },
  { id: "av-sync", name: "AV Sync Flash and Beep", category: "Motion", animated: true, params: [
    { key: "fps", label: "Frame rate", type: "select", options: fpsOptions, default: "25" },
    { key: "interval", label: "Flash every (s)", type: "number", min: 1, max: 60, default: 1 },
    { key: "flashFrames", label: "Flash length (frames)", type: "number", min: 1, max: 30, default: 2 },
    { key: "beep", label: "Beep", type: "boolean", default: true },
    { key: "beepFreq", label: "Beep pitch (Hz)", type: "number", min: 100, max: 8000, default: 1000 },
    { key: "color", label: "Flash colour", type: "color", default: "#ffffff" },
  ], draw: drawAvSync, beeps: avSyncBeeps },
  { id: "pixel", name: "Pixel Structure", category: "Geometry", animated: false, params: [
    { key: "mode", label: "Mode", type: "select", options: [
      { value: "vertical", label: "Vertical lines" }, { value: "horizontal", label: "Horizontal lines" }, { value: "checker", label: "Pixel checker" }, { value: "dots", label: "Dots" },
    ], default: "checker" },
    { key: "pitch", label: "Pixels on / off", type: "number", min: 1, max: 64, default: 1 },
    { key: "color", label: "Colour", type: "color", default: "#ffffff" },
  ], draw: drawPixel },
  { id: "zone-plate", name: "Zone Plate", category: "Geometry", animated: false, params: [], draw: drawZonePlate },
  { id: "faces", name: "Fictional Faces", category: "Reference", animated: true, params: [
    { key: "count", label: "Faces", type: "select", options: ["1", "2", "4", "6", "9", "12"].map((v) => ({ value: v, label: v })), default: "6" },
    { key: "seed", label: "Variation", type: "number", min: 1, max: 999, default: 1 },
    { key: "background", label: "Background", type: "color", default: "#7f8a99" },
    { key: "labels", label: "Labels", type: "boolean", default: true },
    { key: "toneStrip", label: "Skin-tone strip", type: "boolean", default: true },
    { key: "animate", label: "Blink", type: "boolean", default: true },
  ], draw: drawFacesPattern },
  { id: "photo-faces", name: "Photo Faces", category: "Reference", animated: true, params: [
    { key: "rows", label: "Faces down", type: "number", min: 1, max: 20, default: 3 },
    { key: "seed", label: "Variation", type: "number", min: 1, max: 999, default: 1 },
    { key: "gap", label: "Gap (px)", type: "number", min: 0, max: 200, default: 0 },
    { key: "shuffles", label: "New faces per loop (0 = never)", type: "number", min: 0, max: 20, default: 0 },
    { key: "background", label: "Fill colour", type: "color", default: "#7a7a7a" },
  ], draw: (ctx, w, h, s, info) =>
    drawPhotoFaces(ctx, w, h, {
      rows: Number(s.rows),
      seed: Math.round(Number(s.seed)),
      gap: Number(s.gap),
      shuffles: Math.max(0, Math.round(Number(s.shuffles))),
      background: String(s.background),
      progress: info.progress,
    }) },
  { id: "led-moving", name: "LED Moving Pattern", category: "LED", animated: true, params: [
    { key: "panelPxW", label: "Panel width (px)", type: "number", min: 4, max: 4096, default: 168 },
    { key: "panelPxH", label: "Panel height (px)", type: "number", min: 4, max: 4096, default: 168 },
    { key: "motion", label: "Motion", type: "boolean", default: true },
    { key: "showLabels", label: "Row / column labels", type: "boolean", default: true },
    { key: "showAlignment", label: "Alignment cross & circle", type: "boolean", default: true },
    { key: "showInfo", label: "Screen details", type: "boolean", default: true },
  ], draw: drawLedMoving },
  { id: LED_LAYOUT_PATTERN_ID, name: "LED Planner Layout Pattern", category: "LED", animated: true, params: [], draw: drawBlank },
  { id: "blank", name: "Black", category: "Greyscale & colour", animated: false, params: [], draw: drawBlank },
];

export const PATTERN_BY_ID = new Map(PATTERNS.map((p) => [p.id, p]));

/** What auto cycle steps through: everything that draws on its own (not the imported LED layout, not plain black). */
export const AUTO_CYCLE_IDS = PATTERNS.filter((p) => p.id !== LED_LAYOUT_PATTERN_ID && p.id !== "blank").map((p) => p.id);

export const patternName = (id: string): string => PATTERN_BY_ID.get(id)?.name ?? id;

/** The pattern's defaults with whatever the screen has set laid over them. */
export const resolveSettings = (def: PatternDef, settings: PatternSettings | undefined): Required<PatternSettings> => {
  const out: PatternSettings = {};
  def.params.forEach((p) => {
    const v = settings?.[p.key];
    out[p.key] = v === undefined || typeof v !== typeof p.default ? p.default : v;
  });
  return out as Required<PatternSettings>;
};
