// ---------------------------------------------------------------------------
// Photo faces: real-looking portraits for checking skin tones on a display.
//
// The 24 portraits live in one atlas image (6 across, 4 down, square cells).
// The screen is divided into square cells, as many as fit at the chosen
// number of rows, each filled with a randomly chosen portrait. Portraits are
// never stretched: whatever the cells leave over round the edge is filled
// with grey, the same grey the photos were shot against.
// ---------------------------------------------------------------------------

import atlasUrl from "./assets/faces-atlas.jpg";

export const PHOTO_FACE_COUNT = 24;
const ATLAS_COLS = 6;

let atlas: HTMLImageElement | null = null;
let atlasReady: Promise<void> | null = null;

/** Starts loading the portraits; resolves once they can be drawn (or failed to load). */
export const loadPhotoFaces = (): Promise<void> => {
  if (!atlasReady) {
    atlasReady = new Promise((resolve) => {
      const img = new Image();
      img.onload = () => {
        atlas = img;
        resolve();
      };
      img.onerror = () => resolve();
      img.src = atlasUrl;
    });
  }
  return atlasReady;
};

/** Small seeded PRNG, so a given variation always picks the same faces. */
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

/**
 * Which portrait goes in each of `cells` cells: every portrait is used once
 * before any repeats, and no portrait sits next to itself where a fresh deck
 * starts.
 */
export const pickFaces = (cells: number, seed: number, total = PHOTO_FACE_COUNT): number[] => {
  const rand = mulberry32(seed * 7919 + 17);
  const out: number[] = [];
  while (out.length < cells) {
    const deck = Array.from({ length: total }, (_, i) => i);
    for (let i = deck.length - 1; i > 0; i -= 1) {
      const j = Math.floor(rand() * (i + 1));
      [deck[i], deck[j]] = [deck[j], deck[i]];
    }
    if (out.length && deck[0] === out[out.length - 1] && deck.length > 1) [deck[0], deck[1]] = [deck[1], deck[0]];
    out.push(...deck);
  }
  return out.slice(0, cells);
};

/** Square cells, `rows` down, as many across as fit whole, centred, in whole pixels. */
export const photoFaceGrid = (w: number, h: number, rows: number) => {
  const r = Math.max(1, Math.round(rows));
  let cell = Math.floor(h / r);
  let cols = Math.floor(w / Math.max(1, cell));
  // A screen narrower than one cell: one column, cells as wide as the screen.
  if (cols < 1) {
    cols = 1;
    cell = w;
  }
  const usedRows = Math.min(r, Math.max(1, Math.floor(h / Math.max(1, cell))));
  const x0 = Math.floor((w - cols * cell) / 2);
  const y0 = Math.floor((h - usedRows * cell) / 2);
  return { cell, cols, rows: usedRows, x0, y0 };
};

export type PhotoFaceOptions = {
  rows: number;
  seed: number;
  background: string;
  gap: number;
  /** Pick a fresh set of faces this many times per loop (0 = never). */
  shuffles: number;
  progress: number;
};

export const drawPhotoFaces = (ctx: CanvasRenderingContext2D, w: number, h: number, o: PhotoFaceOptions) => {
  ctx.fillStyle = o.background;
  ctx.fillRect(0, 0, w, h);
  if (!atlas) {
    void loadPhotoFaces();
    ctx.fillStyle = "#e5e7eb";
    ctx.font = `bold ${Math.max(12, Math.round(Math.min(w, h) * 0.05))}px Arial, Helvetica, sans-serif`;
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText("Loading faces…", w / 2, h / 2);
    return;
  }
  const g = photoFaceGrid(w, h, o.rows);
  const round = o.shuffles > 0 ? Math.min(o.shuffles - 1, Math.floor(o.progress * o.shuffles)) : 0;
  const faces = pickFaces(g.cols * g.rows, o.seed + round * 1009);
  const src = atlas.naturalWidth / ATLAS_COLS;
  const gap = Math.max(0, Math.min(Math.round(o.gap), Math.floor(g.cell / 4)));
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = "high";
  faces.forEach((face, i) => {
    const col = i % g.cols;
    const row = Math.floor(i / g.cols);
    const size = g.cell - gap * 2;
    if (size <= 0) return;
    ctx.drawImage(
      atlas!,
      (face % ATLAS_COLS) * src,
      Math.floor(face / ATLAS_COLS) * src,
      src,
      src,
      g.x0 + col * g.cell + gap,
      g.y0 + row * g.cell + gap,
      size,
      size,
    );
  });
};
