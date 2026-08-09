// ---------------------------------------------------------------------------
// Client-side WebM -> MP4 transcode for the Moving Test Pattern's downloadable
// video. MediaRecorder (used to capture the animated canvas - see App.tsx's
// downloadMovingTestPatternVideo) only reliably produces WebM across
// browsers; there's no standards-based way to record straight to MP4.
// ffmpeg.wasm is loaded lazily (only when the user actually clicks "Download
// MP4") and only ever runs in a background Web Worker, matching the existing
// lazy-WASM pattern already used for the NovaStar export's sql.js (see
// novastar/novaDb.ts).
// ---------------------------------------------------------------------------

import type { FFmpeg } from "@ffmpeg/ffmpeg";

let ffmpegPromise: Promise<FFmpeg> | null = null;

const loadFFmpeg = async (): Promise<FFmpeg> => {
  const { FFmpeg } = await import("@ffmpeg/ffmpeg");
  const { toBlobURL } = await import("@ffmpeg/util");
  const coreJsUrl = (await import("@ffmpeg/core?url")).default;
  const coreWasmUrl = (await import("@ffmpeg/core/wasm?url")).default;
  const ffmpeg = new FFmpeg();
  await ffmpeg.load({
    coreURL: await toBlobURL(coreJsUrl, "text/javascript"),
    wasmURL: await toBlobURL(coreWasmUrl, "application/wasm"),
  });
  return ffmpeg;
};

const getFFmpeg = (): Promise<FFmpeg> => {
  if (!ffmpegPromise) ffmpegPromise = loadFFmpeg();
  return ffmpegPromise;
};

/**
 * Transcodes a recorded WebM Blob into an H.264/AAC-less MP4 Blob (the test
 * pattern video has no audio track). `onProgress` receives 0..1 and is only
 * called once encoding itself starts (loading the ~30MB ffmpeg-core WASM
 * happens first and isn't reflected in it).
 */
export const encodeWebmToMp4 = async (webmBlob: Blob, onProgress?: (ratio: number) => void): Promise<Blob> => {
  const { fetchFile } = await import("@ffmpeg/util");
  const ffmpeg = await getFFmpeg();
  const onProgressEvent = ({ progress }: { progress: number }) => {
    if (onProgress) onProgress(Math.max(0, Math.min(1, progress)));
  };
  ffmpeg.on("progress", onProgressEvent);
  try {
    await ffmpeg.writeFile("in.webm", await fetchFile(webmBlob));
    // veryfast preset: WASM encoding is already far slower than native, so
    // trade a little compression efficiency for real-world completion time.
    // crf 18 keeps this pattern's sharp text/edges clean (default ~23 would
    // visibly block them, the same reasoning as the WebM recorder's own
    // generous bitrate floor).
    const code = await ffmpeg.exec(["-i", "in.webm", "-c:v", "libx264", "-preset", "veryfast", "-crf", "18", "-pix_fmt", "yuv420p", "-movflags", "+faststart", "out.mp4"]);
    if (code !== 0) throw new Error(`ffmpeg exited with code ${code}`);
    const data = await ffmpeg.readFile("out.mp4");
    // DOM lib types Blob's BlobPart as ArrayBufferView<ArrayBuffer> specifically,
    // while ffmpeg.wasm's readFile() returns a plain Uint8Array<ArrayBufferLike>.
    return new Blob([data as unknown as BlobPart], { type: "video/mp4" });
  } finally {
    ffmpeg.off("progress", onProgressEvent);
    await ffmpeg.deleteFile("in.webm").catch(() => {});
    await ffmpeg.deleteFile("out.mp4").catch(() => {});
  }
};
