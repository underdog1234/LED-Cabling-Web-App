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
import { beepsFor, renderBeepSamples, renderBeepsWav, screenBeeps, startRecordingSound, type ScheduledBeep } from "./audio";
import { muxWebm, opusHead, type MuxPacket } from "./webmMux";
import { patternName } from "./patterns";
import { loadPhotoFaces } from "./photoFaces";
import { canvasPatternLabel, renderComposite, renderScreen, type RenderContext } from "./render";

export type ExportFormat = "png" | "webm" | "mp4";
export type ExportScope = "canvas" | "screen" | "selected";

export type ExportItem = {
  fileBase: (ext: string) => string;
  w: number;
  h: number;
  draw: (ctx: CanvasRenderingContext2D, time: number) => void;
  /** Sync tones the video carries, on the pattern clock. Empty for a silent file. */
  beeps: ScheduledBeep[];
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
    beeps: beepsFor(config.screens, config.loopSeconds),
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
    beeps: screenBeeps(screen, config.loopSeconds),
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

// --- frame-by-frame encoding (WebCodecs) ----------------------------------------
// Draws every frame of the loop at its exact time and encodes it, as fast as
// the computer allows - no real-time clock involved. So the file is always
// exactly one loop long, with every frame present and every beep on its
// frame, however busy the machine or whichever tab is in front. Browsers
// without WebCodecs fall back to the real-time recorder below.

const VIDEO_CODECS: Array<{ codec: string; mux: "V_VP9" | "V_VP8" }> = [
  { codec: "vp09.00.41.08", mux: "V_VP9" },
  { codec: "vp09.00.51.08", mux: "V_VP9" },
  { codec: "vp8", mux: "V_VP8" },
];

const pickEncoder = async (w: number, h: number, fps: number, bitrate: number) => {
  if (typeof VideoEncoder === "undefined" || typeof VideoFrame === "undefined") return null;
  for (const c of VIDEO_CODECS) {
    const config: VideoEncoderConfig = { codec: c.codec, width: w, height: h, bitrate, framerate: fps, bitrateMode: "variable", latencyMode: "quality" };
    try {
      const support = await VideoEncoder.isConfigSupported(config);
      if (support.supported) return { config, mux: c.mux };
    } catch {
      // Try the next codec.
    }
  }
  return null;
};

const encodeOpus = async (samples: Int16Array, sampleRate: number): Promise<{ packets: MuxPacket[]; head: Uint8Array; preSkip: number } | null> => {
  if (typeof AudioEncoder === "undefined" || typeof AudioData === "undefined") return null;
  const config: AudioEncoderConfig = { codec: "opus", sampleRate, numberOfChannels: 1, bitrate: 128_000 };
  try {
    if (!(await AudioEncoder.isConfigSupported(config)).supported) return null;
  } catch {
    return null;
  }
  const packets: MuxPacket[] = [];
  let description: Uint8Array | null = null;
  let failed: unknown = null;
  const encoder = new AudioEncoder({
    output: (chunk, meta) => {
      const data = new Uint8Array(chunk.byteLength);
      chunk.copyTo(data);
      packets.push({ track: 2, timestampUs: chunk.timestamp, key: true, data });
      const d = meta?.decoderConfig?.description;
      if (d && !description) description = d instanceof ArrayBuffer ? new Uint8Array(d) : new Uint8Array((d as ArrayBufferView).buffer.slice(0));
    },
    error: (e) => {
      failed = e;
    },
  });
  encoder.configure(config);
  const block = 960;
  for (let i = 0; i < samples.length; i += block) {
    const part = samples.slice(i, Math.min(samples.length, i + block));
    encoder.encode(new AudioData({ format: "s16", sampleRate, numberOfChannels: 1, numberOfFrames: part.length, timestamp: Math.round((i / sampleRate) * 1e6), data: part }));
  }
  await encoder.flush();
  encoder.close();
  if (failed) return null;
  const desc = description as Uint8Array | null;
  const preSkip = desc && desc.length >= 12 ? new DataView(desc.buffer, desc.byteOffset).getUint16(10, true) : 312;
  return { packets, head: desc && desc.length >= 19 ? desc : opusHead(1, sampleRate, preSkip), preSkip };
};

/**
 * Exactly `seconds` of the item as a WebM, frame by frame, or null when this
 * browser can't encode that way (the caller then records in real time).
 */
export const encodeWebmFrames = async (
  item: ExportItem,
  fps: number,
  seconds: number,
  loopSeconds: number,
  withSound: boolean,
  onTick?: (elapsed: number) => void,
): Promise<Blob | null> => {
  // Encoders want even dimensions; the file is never resized, so odd sizes record in real time instead.
  if (item.w % 2 || item.h % 2) return null;
  const bitrate = Math.min(80_000_000, Math.max(8_000_000, Math.round(item.w * item.h * 6)));
  const picked = await pickEncoder(item.w, item.h, fps, bitrate);
  if (!picked) return null;
  let audio: Awaited<ReturnType<typeof encodeOpus>> = null;
  if (withSound && item.beeps.length) {
    audio = await encodeOpus(renderBeepSamples(item.beeps, loopSeconds, seconds), 48000);
    if (!audio) return null;
  }
  const canvas = document.createElement("canvas");
  canvas.width = item.w;
  canvas.height = item.h;
  const ctx = canvas.getContext("2d");
  if (!ctx) return null;
  const packets: MuxPacket[] = [];
  let failed: unknown = null;
  const encoder = new VideoEncoder({
    output: (chunk) => {
      const data = new Uint8Array(chunk.byteLength);
      chunk.copyTo(data);
      packets.push({ track: 1, timestampUs: chunk.timestamp, key: chunk.type === "key", data });
    },
    error: (e) => {
      failed = e;
    },
  });
  encoder.configure(picked.config);
  const frames = Math.max(1, Math.round(seconds * fps));
  const frameUs = 1e6 / fps;
  for (let i = 0; i < frames; i += 1) {
    if (failed) break;
    const t = i / fps;
    item.draw(ctx, t);
    const frame = new VideoFrame(canvas, { timestamp: Math.round(i * frameUs), duration: Math.round(frameUs) });
    encoder.encode(frame, { keyFrame: i % fps === 0 });
    frame.close();
    onTick?.(t);
    // Keep the encoder's queue short, and let the page repaint the progress.
    while (encoder.encodeQueueSize > 4) await new Promise((r) => setTimeout(r, 1));
    if (i % 10 === 9) await new Promise((r) => setTimeout(r, 0));
  }
  if (!failed) await encoder.flush().catch((e) => (failed = e));
  encoder.close();
  if (failed) return null;
  const data = muxWebm([...packets, ...(audio?.packets ?? [])], {
    width: item.w,
    height: item.h,
    fps,
    durationSeconds: frames / fps,
    videoCodec: picked.mux,
    audio: audio ? { codecPrivate: audio.head, sampleRate: 48000, channels: 1, preSkip: audio.preSkip } : undefined,
  });
  return new Blob([data as unknown as BlobPart], { type: "video/webm" });
};

/**
 * Records `seconds` of the item to WebM. The pattern is drawn at the file's
 * own frame rate from a clock starting at zero, so the recording begins on
 * the loop's first frame.
 */
export const recordWebm = (item: ExportItem, fps: number, seconds: number, loopSeconds: number, withSound: boolean, onTick?: (elapsed: number) => void): Promise<Blob> => {
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
  // The sync tones play into the recording live; the audio starts a moment
  // after this, and the picture clock is held back by the same moment.
  const sound = withSound ? startRecordingSound(item.beeps, loopSeconds, seconds) : null;
  if (sound) stream.addTrack(sound.track);
  const withOpus = mimeType.replace(/codecs=([^;]+)/, "codecs=$1,opus");
  const recorder = new MediaRecorder(stream, { mimeType: sound && MediaRecorder.isTypeSupported?.(withOpus) ? withOpus : mimeType, videoBitsPerSecond });
  const chunks: BlobPart[] = [];
  recorder.ondataavailable = (e) => {
    if (e.data.size > 0) chunks.push(e.data);
  };
  let drawId = 0;
  return new Promise((resolve) => {
    recorder.onstop = () => {
      window.clearInterval(drawId);
      sound?.stop();
      stream.getTracks().forEach((track) => track.stop());
      resolve(new Blob(chunks, { type: mimeType }));
    };
    recorder.start();
    // The clock starts with the recording, on the loop's first frame.
    const start = performance.now() + (sound?.leadMs ?? 0);
    drawId = window.setInterval(() => {
      const t = Math.max(0, (performance.now() - start) / 1000);
      item.draw(ctx, t);
      onTick?.(t);
    }, 1000 / fps);
    setTimeout(() => recorder.stop(), seconds * 1000 + (sound?.leadMs ?? 0));
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
  // A still taken before the portraits arrive would show "Loading faces…".
  await loadPhotoFaces();
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
      const share = format === "mp4" ? 0.5 : 1;
      // Frame by frame where the browser can, exactly one loop; MP4 gets its
      // tones from an exact WAV instead of the recording.
      let webm = await encodeWebmFrames(item, options.video.fps, options.loopSeconds, options.loopSeconds, format === "webm", (t) =>
        onProgress({ label: `${prefix}Rendering ${names[i]} - ${t.toFixed(1)}s of ${options.loopSeconds}s`, ratio: (i + Math.min(1, t / options.loopSeconds) * share) / items.length }),
      );
      if (!webm) {
        // Real-time fallback; MP4 records a little longer and is cut back to one loop.
        const seconds = format === "mp4" ? options.loopSeconds + MP4_RECORD_MARGIN_SECONDS : options.loopSeconds;
        webm = await recordWebm(item, options.video.fps, seconds, options.loopSeconds, format === "webm", (t) =>
          onProgress({ label: `${prefix}Recording ${names[i]} - ${Math.min(seconds, t).toFixed(0)}s of ${seconds.toFixed(0)}s`, ratio: (i + Math.min(1, t / seconds) * share) / items.length }),
        );
      }
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
          item.beeps.length ? renderBeepsWav(item.beeps, options.loopSeconds, options.loopSeconds) : undefined,
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
