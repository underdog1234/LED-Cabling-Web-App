// ---------------------------------------------------------------------------
// Fictional faces - drawn entirely from shapes and gradients, from a seed.
//
// A face is the thing an audience notices first when colour is off, so a
// screen check needs one. These are illustrations, not photographs: no real
// person is depicted, and every portrait carries a "fictional" label. The
// set always spans the full range of skin tones, lightest to deepest, so
// colour and gamma problems show on every complexion rather than one.
// ---------------------------------------------------------------------------

export type FacesOptions = {
  count: number;
  background: string;
  labels: boolean;
  animate: boolean;
  seed: number;
  toneStrip: boolean;
  progress: number;
};

/** Light to deep. Approximate sRGB values chosen to cover the range evenly. */
export const SKIN_TONES = [
  "#f6e0cf", "#efcfb6", "#e6bc98", "#d9a47c", "#c98e66", "#b67a52",
  "#a0663f", "#8a5434", "#73432a", "#5e3622", "#4a2a1b", "#3a2016",
];

const HAIR = ["#1c1410", "#2e1f16", "#4b3020", "#6b4226", "#a0703f", "#c9a063", "#8c8c8c", "#d9d4cc", "#7a2e1d", "#3b2a4a"];
const EYES = ["#3b2a1e", "#5a3b22", "#2f5d7c", "#4c6b3a", "#6f5a3a", "#2b2b2b"];
const SHIRTS = ["#1e3a8a", "#9f1239", "#065f46", "#7c2d12", "#4c1d95", "#334155", "#a16207", "#0e7490", "#be185d", "#3f6212"];
const NAMES = ["Avery", "Jordan", "Riley", "Sam", "Morgan", "Quinn", "Rowan", "Sky", "Kai", "Drew", "Emery", "Remy"];

// Small deterministic generator, so a given seed always draws the same people.
const mulberry32 = (seed: number) => {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
};

const shade = (hex: string, amount: number) => {
  const n = parseInt(hex.slice(1), 16);
  const f = (c: number) => Math.max(0, Math.min(255, Math.round(c + (amount < 0 ? c * amount : (255 - c) * amount))));
  const r = f((n >> 16) & 255);
  const g = f((n >> 8) & 255);
  const b = f(n & 255);
  return `rgb(${r},${g},${b})`;
};

type Person = {
  skin: string;
  hair: string;
  hairStyle: number;
  eyes: string;
  shirt: string;
  name: string;
  faceWidth: number;
  smile: number;
  blinkAt: number;
};

const people = (count: number, seed: number): Person[] => {
  const rand = mulberry32(seed * 7919 + 17);
  return Array.from({ length: count }, (_, i) => {
    // Tones spread evenly across the whole range for any face count.
    const tone = count === 1 ? 5 : Math.round((i * (SKIN_TONES.length - 1)) / (count - 1));
    return {
      skin: SKIN_TONES[(tone + (seed - 1)) % SKIN_TONES.length],
      hair: HAIR[Math.floor(rand() * HAIR.length)],
      hairStyle: Math.floor(rand() * 5),
      eyes: EYES[Math.floor(rand() * EYES.length)],
      shirt: SHIRTS[Math.floor(rand() * SHIRTS.length)],
      name: NAMES[(i + seed * 3) % NAMES.length],
      faceWidth: 0.86 + rand() * 0.14,
      smile: 0.4 + rand() * 0.6,
      blinkAt: rand(),
    };
  });
};

const drawPortrait = (ctx: CanvasRenderingContext2D, x: number, y: number, w: number, h: number, p: Person, opts: FacesOptions, index: number) => {
  ctx.save();
  ctx.beginPath();
  ctx.rect(x, y, w, h);
  ctx.clip();
  // Soft studio backdrop.
  const bg = ctx.createRadialGradient(x + w / 2, y + h * 0.4, 0, x + w / 2, y + h * 0.4, Math.max(w, h) * 0.75);
  bg.addColorStop(0, shade(opts.background, 0.25));
  bg.addColorStop(1, shade(opts.background, -0.35));
  ctx.fillStyle = bg;
  ctx.fillRect(x, y, w, h);

  const unit = Math.min(w, h * 0.82);
  const cx = x + w / 2;
  const headW = unit * 0.36 * p.faceWidth;
  const headH = unit * 0.46;
  const headY = y + h * 0.42;

  // Hair behind the head (long styles).
  ctx.fillStyle = p.hair;
  if (p.hairStyle === 2 || p.hairStyle === 4) {
    ctx.beginPath();
    ctx.ellipse(cx, headY + headH * 0.18, headW * 0.68, headH * 0.68, 0, 0, Math.PI * 2);
    ctx.fill();
  }
  // Shoulders and shirt.
  ctx.fillStyle = p.shirt;
  ctx.beginPath();
  ctx.ellipse(cx, y + h * 1.02, unit * 0.48, h * 0.3, 0, Math.PI, 0);
  ctx.fill();
  // Neck.
  ctx.fillStyle = shade(p.skin, -0.12);
  ctx.fillRect(cx - headW * 0.22, headY + headH * 0.32, headW * 0.44, headH * 0.42);
  // Ears.
  ctx.fillStyle = shade(p.skin, -0.05);
  [-1, 1].forEach((side) => {
    ctx.beginPath();
    ctx.ellipse(cx + side * headW * 0.5, headY + headH * 0.02, headW * 0.09, headH * 0.11, 0, 0, Math.PI * 2);
    ctx.fill();
  });
  // Head, lit from the upper left.
  const face = ctx.createRadialGradient(cx - headW * 0.15, headY - headH * 0.15, headW * 0.05, cx, headY, headH * 0.62);
  face.addColorStop(0, shade(p.skin, 0.12));
  face.addColorStop(0.7, p.skin);
  face.addColorStop(1, shade(p.skin, -0.18));
  ctx.fillStyle = face;
  ctx.beginPath();
  ctx.ellipse(cx, headY, headW * 0.5, headH * 0.5, 0, 0, Math.PI * 2);
  ctx.fill();
  // Cheeks.
  ctx.fillStyle = "rgba(220,90,90,0.12)";
  [-1, 1].forEach((side) => {
    ctx.beginPath();
    ctx.ellipse(cx + side * headW * 0.24, headY + headH * 0.1, headW * 0.1, headH * 0.06, 0, 0, Math.PI * 2);
    ctx.fill();
  });
  // Hair on top.
  ctx.fillStyle = p.hair;
  if (p.hairStyle !== 3) {
    ctx.beginPath();
    ctx.ellipse(cx, headY - headH * 0.2, headW * 0.54, headH * 0.36, 0, Math.PI, 0);
    ctx.fill();
    if (p.hairStyle === 1) {
      ctx.beginPath();
      ctx.ellipse(cx - headW * 0.12, headY - headH * 0.26, headW * 0.42, headH * 0.16, -0.25, 0, Math.PI * 2);
      ctx.fill();
    }
    if (p.hairStyle === 4) {
      ctx.beginPath();
      ctx.arc(cx, headY - headH * 0.58, headW * 0.18, 0, Math.PI * 2);
      ctx.fill();
    }
  } else {
    // Close crop: a thin cap.
    ctx.globalAlpha = 0.6;
    ctx.beginPath();
    ctx.ellipse(cx, headY - headH * 0.3, headW * 0.48, headH * 0.22, 0, Math.PI, 0);
    ctx.fill();
    ctx.globalAlpha = 1;
  }
  // Eyes - closed for a few frames once a loop when blinking.
  const eyeY = headY - headH * 0.02;
  const eyeDx = headW * 0.19;
  const blinkPhase = (opts.progress + p.blinkAt) % 1;
  const blinking = opts.animate && blinkPhase < 0.025;
  [-1, 1].forEach((side) => {
    const ex = cx + side * eyeDx;
    if (blinking) {
      ctx.strokeStyle = shade(p.skin, -0.5);
      ctx.lineWidth = Math.max(1, headW * 0.02);
      ctx.beginPath();
      ctx.moveTo(ex - headW * 0.08, eyeY);
      ctx.quadraticCurveTo(ex, eyeY + headH * 0.025, ex + headW * 0.08, eyeY);
      ctx.stroke();
    } else {
      ctx.fillStyle = "#f8f8f4";
      ctx.beginPath();
      ctx.ellipse(ex, eyeY, headW * 0.085, headH * 0.04, 0, 0, Math.PI * 2);
      ctx.fill();
      ctx.fillStyle = p.eyes;
      ctx.beginPath();
      ctx.arc(ex, eyeY, headH * 0.034, 0, Math.PI * 2);
      ctx.fill();
      ctx.fillStyle = "#0b0b0b";
      ctx.beginPath();
      ctx.arc(ex, eyeY, headH * 0.016, 0, Math.PI * 2);
      ctx.fill();
      ctx.fillStyle = "rgba(255,255,255,0.9)";
      ctx.beginPath();
      ctx.arc(ex - headH * 0.01, eyeY - headH * 0.012, headH * 0.007, 0, Math.PI * 2);
      ctx.fill();
    }
    // Brows.
    ctx.strokeStyle = shade(p.hair, -0.1);
    ctx.lineWidth = Math.max(1, headH * 0.018);
    ctx.lineCap = "round";
    ctx.beginPath();
    ctx.moveTo(ex - headW * 0.09, eyeY - headH * 0.075);
    ctx.quadraticCurveTo(ex, eyeY - headH * 0.1, ex + headW * 0.09, eyeY - headH * 0.075);
    ctx.stroke();
  });
  // Nose: a soft shadow line.
  ctx.strokeStyle = shade(p.skin, -0.25);
  ctx.lineWidth = Math.max(1, headH * 0.012);
  ctx.beginPath();
  ctx.moveTo(cx, eyeY + headH * 0.04);
  ctx.quadraticCurveTo(cx - headW * 0.05, eyeY + headH * 0.15, cx + headW * 0.03, eyeY + headH * 0.17);
  ctx.stroke();
  // Mouth.
  ctx.strokeStyle = shade(p.skin, -0.45);
  ctx.lineWidth = Math.max(1, headH * 0.014);
  ctx.beginPath();
  const my = headY + headH * 0.26;
  ctx.moveTo(cx - headW * 0.14, my);
  ctx.quadraticCurveTo(cx, my + headH * 0.07 * p.smile, cx + headW * 0.14, my);
  ctx.stroke();

  if (opts.labels) {
    const px = Math.max(9, Math.round(Math.min(w, h) * 0.05));
    ctx.font = `bold ${px}px Arial, Helvetica, sans-serif`;
    ctx.textAlign = "left";
    ctx.textBaseline = "top";
    const label = `${p.name} · tone ${SKIN_TONES.indexOf(p.skin) + 1}`;
    ctx.fillStyle = "rgba(0,0,0,0.55)";
    ctx.fillRect(x + px * 0.4, y + px * 0.4, ctx.measureText(label).width + px * 0.8, px * 1.5);
    ctx.fillStyle = "#ffffff";
    ctx.fillText(label, x + px * 0.8, y + px * 0.65);
    ctx.font = `${Math.max(8, Math.round(px * 0.7))}px Arial, Helvetica, sans-serif`;
    ctx.textBaseline = "bottom";
    ctx.fillStyle = "rgba(255,255,255,0.85)";
    ctx.fillText(`Fictional person ${index + 1} - illustration`, x + px * 0.6, y + h - px * 0.4);
  }
  ctx.restore();
};

const gridFor = (count: number, w: number, h: number) => {
  let best = { cols: 1, rows: count, score: Infinity };
  for (let cols = 1; cols <= count; cols += 1) {
    const rows = Math.ceil(count / cols);
    const cellAspect = w / cols / (h / rows);
    // Portrait-ish cells (about 0.8) look best for a head and shoulders.
    const score = Math.abs(Math.log(cellAspect / 0.8)) + (cols * rows - count) * 0.2;
    if (score < best.score) best = { cols, rows, score };
  }
  return best;
};

export const drawFaces = (ctx: CanvasRenderingContext2D, w: number, h: number, opts: FacesOptions) => {
  const strip = opts.toneStrip ? Math.max(8, Math.round(h * 0.07)) : 0;
  const areaH = h - strip;
  const count = Math.max(1, Math.min(24, opts.count));
  const { cols, rows } = gridFor(count, w, areaH);
  const list = people(count, Math.max(1, opts.seed));
  ctx.fillStyle = "#000000";
  ctx.fillRect(0, 0, w, h);
  const xs = Array.from({ length: cols + 1 }, (_, i) => Math.round((i * w) / cols));
  const ys = Array.from({ length: rows + 1 }, (_, i) => Math.round((i * areaH) / rows));
  list.forEach((p, i) => {
    const c = i % cols;
    const r = Math.floor(i / cols);
    drawPortrait(ctx, xs[c], ys[r], xs[c + 1] - xs[c], ys[r + 1] - ys[r], p, opts, i);
  });
  if (strip) {
    const sx = Array.from({ length: SKIN_TONES.length + 1 }, (_, i) => Math.round((i * w) / SKIN_TONES.length));
    SKIN_TONES.forEach((tone, i) => {
      ctx.fillStyle = tone;
      ctx.fillRect(sx[i], areaH, sx[i + 1] - sx[i], strip);
      if (opts.labels) {
        const px = Math.max(8, Math.round(strip * 0.35));
        ctx.font = `bold ${px}px Arial, Helvetica, sans-serif`;
        ctx.textAlign = "center";
        ctx.textBaseline = "middle";
        ctx.fillStyle = i < 5 ? "#000000" : "#ffffff";
        ctx.fillText(String(i + 1), (sx[i] + sx[i + 1]) / 2, areaH + strip / 2);
      }
    });
  }
};
