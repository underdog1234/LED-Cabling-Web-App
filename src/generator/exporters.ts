// ---------------------------------------------------------------------------
// Exports: PNG stills and WebM / MP4 loops, of the whole canvas, one screen,
// or several screens as separate files in one ZIP.
//
// Video is recorded the same way the LED planner records its Moving Test
// Pattern - a detached canvas drawn on a timer and captured with
// MediaRecorder - and MP4 goes through the planner's own delivery-settings
// encoder (testPattern/mp4Encode), so the files match what it produces.
// ---------------------------------------------------------------------------

import { zipSync } from "fflate";
import { MP4_RECORD_MARGIN_SECONDS } from "../testPattern/mp4Encode";
import { exportFileName, uniqueNames, type GeneratorConfig, type SubScreen } from "./model";
import { patternName } from "./patterns";
import { canvasPatternLabel, renderComposite, renderScreen, type RenderContext } from "./render";

export type ExportFormat = "png" | "webm" | "mp4";
export type ExportScope = "canvas" | "screen" | "selected";

export type ExportItem = {
  fileBase: (ext: string) => string;
  w: number;
  h: number;
  draw: (ctx: CanvasRenderingContext2D, time: number) => void;
};

export const downloadBlob = (blob: Blob, filename: string) => {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
};

export const canvasItem = (config: GeneratorConfig): ExportItem => {
  const cache = new Map<string, HTMLCanvasElement>();
  return {
    fileBase: (ext) => exportFileName(`${config.name}-Full-Canvas`, canvasPatternLabel(config, patternName), config.canvas.w, config.canvas.h, ext),
    w: config.canvas.w,
    h: config.canvas.h,
    draw: (ctx, time) => renderComposite(ctx, config, time, cache),
  };
};

export const screenItem = (config: GeneratorConfig, screen: SubScreen): ExportItem => {
  const rc: RenderContext = { loopSeconds: config.loopSeconds, canvas: { w: config.canvas.w, h: config.canvas.h } };
  return {
    fileBase: (ext) => exportFileName(screen.name, patternName(screen.pattern), screen.w, screen.h, ext),
    w: screen.w,
    h: screen.h,
    draw: (ctx, time) => {
      ctx.setTransform(1, 0, 0, 1, 0, 0);
      renderScreen(ctx, screen, time, rc);
    },
  };
};

export const itemsFor = (config: GeneratorConfig, scope: ExportScope, screenId: string | null, selectedIds: string[]): ExportItem[] => {
  if (scope === "canvas") return [canvasItem(config)];
  if (scope === "screen") {
    const s = config.screens.find((x) => x.id === screenId);
    return s ? [screenItem(config, s)] : [];
  }
  // Layer order, so the ZIP lists them the way the screen list does.
  return config.screens.filter((s) => selectedIds.includes(s.id)).map((s) => screenItem(config, s));
};

const toBlob = (canvas: HTMLCanvasElement): Promise<Blob> =>
  new Promise((resolve, reject) => canvas.toBlob((b) => (b ? resolve(b) : reject(new Error("PNG encoding failed"))), "image/png"));

export const renderPng = async (item: ExportItem, time: number): Promise<Blob> => {
  const c = document.createElement("canvas");
  c.width = item.w;
  c.height = item.h;
  const ctx = c.getContext("2d");
  if (!ctx) throw new Error("Could not create a canvas this size");
  item.draw(ctx, time);
  return toBlob(c);
};

export const pickVideoMimeType = (): string | null => {
  if (typeof MediaRecorder === "undefined") return null;
  for (const candidate of ["video/webm;codecs=vp9", "video/webm;codecs=vp8", "video/webm"]) {
    if (MediaRecorder.isTypeSupported?.(candidate)) return candidate;
  }
  return null;
};

/**
 * Records `seconds` of the item to WebM. The pattern is drawn at the file's
 * own frame rate from a clock starting at zero, so the recording begins on
 * the loop's first frame.
 */
export const recordWebm = (item: ExportItem, fps: number, seconds: number, onTick?: (elapsed: number) => void): Promise<Blob> => {
  const mimeType = pickVideoMimeType();
  if (!mimeType) return Promise.reject(new Error("This browser can't record video (no WebM/MediaRecorder support). Try Chrome, Edge or Firefox."));
  const canvas = document.createElement("canvas");
  canvas.width = item.w;
  canvas.height = item.h;
  const ctx = canvas.getContext("2d");
  if (!ctx) return Promise.reject(new Error("Could not create a recording canvas"));
  item.draw(ctx, 0);
  // ~6 bits per pixel, 8 to 80 Mbps - the planner's own rate, generous enough
  // for hard edges and small text.
  const videoBitsPerSecond = Math.min(80_000_000, Math.max(8_000_000, Math.round(item.w * item.h * 6)));
  const stream = canvas.captureStream(fps);
  const recorder = new MediaRecorder(stream, { mimeType, videoBitsPerSecond });
  const chunks: BlobPart[] = [];
  recorder.ondataavailable = (e) => {
    if (e.data.size > 0) chunks.push(e.data);
  };
  const start = performance.now();
  const drawId = window.setInterval(() => {
    const t = (performance.now() - start) / 1000;
    item.draw(ctx, t);
    onTick?.(t);
  }, 1000 / fps);
  return new Promise((resolve) => {
    recorder.onstop = () => {
      window.clearInterval(drawId);
      stream.getTracks().forEach((track) => track.stop());
      resolve(new Blob(chunks, { type: mimeType }));
    };
    recorder.start();
    setTimeout(() => recorder.stop(), seconds * 1000);
  });
};

export type VideoSettings = { fps: number; targetMbps: number; maxMbps: number };

export type ExportProgress = { label: string; ratio: number };

/**
 * Runs an export and hands back the files. One file downloads as itself;
 * several are packed into one ZIP.
 */
export const runExport = async (
  items: ExportItem[],
  format: ExportFormat,
  options: { time: number; loopSeconds: number; video: VideoSettings; zipName: string },
  onProgress: (p: ExportProgress) => void,
): Promise<{ blob: Blob; filename: string; files: string[] }> => {
  if (!items.length) throw new Error("Nothing to export - choose at least one screen.");
  const files: Array<{ name: string; data: Uint8Array }> = [];
  const names = uniqueNames(items.map((item) => item.fileBase(format)));
  for (let i = 0; i < items.length; i += 1) {
    const item = items[i];
    const prefix = items.length > 1 ? `${i + 1} of ${items.length}: ` : "";
    let blob: Blob;
    if (format === "png") {
      onProgress({ label: `${prefix}Rendering ${names[i]}`, ratio: i / items.length });
      blob = await renderPng(item, options.time);
    } else {
      const seconds = format === "mp4" ? options.loopSeconds + MP4_RECORD_MARGIN_SECONDS : options.loopSeconds;
      const webm = await recordWebm(item, options.video.fps, seconds, (t) =>
        onProgress({ label: `${prefix}Recording ${names[i]} - ${Math.min(seconds, t).toFixed(0)}s of ${seconds.toFixed(0)}s`, ratio: (i + Math.min(1, t / seconds) * (format === "mp4" ? 0.5 : 1)) / items.length }),
      );
      if (format === "mp4") {
        onProgress({ label: `${prefix}Loading the MP4 encoder`, ratio: (i + 0.5) / items.length });
        const { encodeWebmToMp4 } = await import("../testPattern/mp4Encode");
        blob = await encodeWebmToMp4(
          webm,
          {
            fps: options.video.fps,
            targetMbps: options.video.targetMbps,
            maxMbps: Math.max(options.video.targetMbps, options.video.maxMbps),
            width: item.w,
            height: item.h,
            loopSeconds: options.loopSeconds,
          },
          (r) => onProgress({ label: `${prefix}Encoding ${names[i]} - ${Math.round(r * 100)}%`, ratio: (i + 0.5 + r * 0.5) / items.length }),
        );
      } else {
        blob = webm;
      }
    }
    if (items.length === 1) return { blob, filename: names[0], files: names };
    files.push({ name: names[i], data: new Uint8Array(await blob.arrayBuffer()) });
  }
  onProgress({ label: "Packing ZIP", ratio: 1 });
  const zipped = zipSync(Object.fromEntries(files.map((f) => [f.name, [f.data, { level: 0 }]])));
  return { blob: new Blob([zipped as unknown as BlobPart], { type: "application/zip" }), filename: options.zipName, files: names };
};
