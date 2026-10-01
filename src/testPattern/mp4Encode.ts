// ---------------------------------------------------------------------------
// Client-side WebM -> MP4 transcode for the Moving Test Pattern's downloadable
// video. MediaRecorder (used to capture the animated canvas - see App.tsx's
// recordMovingTestPatternWebm) only reliably produces WebM across browsers;
// there's no standards-based way to record straight to MP4. ffmpeg.wasm is
// loaded lazily (only when the user actually asks for an MP4) and only ever
// runs in a background Web Worker, matching the existing lazy-WASM pattern
// already used for the NovaStar export's sql.js (see novastar/novaDb.ts).
//
// THE OUTPUT IS A DELIVERY FILE FOR A MEDIA SERVER, not a preview: the
// settings below are fixed to what a server and an LED processor expect, and
// MP4_PROFILE says what they are in one place so the download dialog can show
// the same figures the encoder actually uses.
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

/** What the MP4 export is fixed to, in the words the settings are specified in. */
export const MP4_PROFILE = {
  container: "MP4 / H.264",
  profile: "High",
  /** The level ASKED for. A wall too big for it gets the lowest level that fits - see h264LevelFor. */
  preferredLevel: "4.2",
  pixelFormat: "yuv420p (8-bit 4:2:0)",
  scan: "Progressive",
  colour: "Rec.709, limited (TV) range",
  defaultFps: 60,
  defaultTargetMbps: 8,
  defaultMaxMbps: 12,
  /** Keyframes every half second, whatever the frame rate. */
  keyframeSeconds: 0.5,
  bFrames: 0,
} as const;

export type Mp4EncodeSettings = {
  /** Frames per second of the finished file - constant, whatever the recording managed. */
  fps: number;
  /** Rate-control target, in megabits per second. */
  targetMbps: number;
  /** Rate-control ceiling, in megabits per second. */
  maxMbps: number;
  /** Encoded width and height, for choosing the H.264 level. */
  width: number;
  height: number;
  /**
   * Exact length of the finished file, in seconds - one whole loop of the
   * pattern. The recording is deliberately made longer than this and cut back
   * to it: a recorder gives you the frames it happens to catch, and a file
   * even a few frames short of the loop jumps every time it repeats.
   */
  loopSeconds: number;
};

export const defaultMp4Settings = (width: number, height: number, loopSeconds: number): Mp4EncodeSettings => ({
  fps: MP4_PROFILE.defaultFps,
  targetMbps: MP4_PROFILE.defaultTargetMbps,
  maxMbps: MP4_PROFILE.defaultMaxMbps,
  width,
  height,
  loopSeconds,
});

/**
 * How much longer than one loop to record before cutting back.
 *
 * MediaRecorder's own clock does not line up with the one that stops it: a
 * 20-second recording came out 19.22 seconds long, which is three quarters of
 * a second of pattern missing every time the file repeats. The animation is
 * periodic, so ANY exact loop-length window of it loops seamlessly - the
 * margin just guarantees there is a whole one to cut from.
 */
export const MP4_RECORD_MARGIN_SECONDS = 1.5;

/** Keyframe interval in FRAMES for a frame rate - half a second, rounded to a whole frame. */
export const keyframeIntervalFor = (fps: number): number => Math.max(1, Math.round(fps * MP4_PROFILE.keyframeSeconds));

// H.264 levels, in the order they are tried: the macroblocks one frame may
// hold, and the macroblocks per second the decoder must get through. A level
// tag that the stream exceeds is worse than no file at all - a hardware
// decoder is entitled to refuse it - so the level follows the wall rather than
// the other way round.
const H264_LEVELS: Array<{ level: string; maxFrameMbs: number; maxMbsPerSecond: number }> = [
  { level: "4.2", maxFrameMbs: 8704, maxMbsPerSecond: 522240 },
  { level: "5.0", maxFrameMbs: 22080, maxMbsPerSecond: 589824 },
  { level: "5.1", maxFrameMbs: 36864, maxMbsPerSecond: 983040 },
  { level: "5.2", maxFrameMbs: 36864, maxMbsPerSecond: 2073600 },
  { level: "6.0", maxFrameMbs: 139264, maxMbsPerSecond: 4177920 },
  { level: "6.1", maxFrameMbs: 139264, maxMbsPerSecond: 8355840 },
  { level: "6.2", maxFrameMbs: 139264, maxMbsPerSecond: 16711680 },
];

/**
 * The H.264 level for a wall: 4.2 as asked for, or the lowest level above it
 * that the resolution and frame rate actually fit inside. An LED wall is
 * routinely wider than any broadcast format - 6216 x 1344 at 60fps needs level
 * 6.0 - and tagging that stream 4.2 would be a lie a decoder may act on.
 */
export const h264LevelFor = (width: number, height: number, fps: number): string => {
  const frameMbs = Math.ceil(width / 16) * Math.ceil(height / 16);
  const mbsPerSecond = frameMbs * Math.max(1, fps);
  const fits = H264_LEVELS.find((entry) => frameMbs <= entry.maxFrameMbs && mbsPerSecond <= entry.maxMbsPerSecond);
  return (fits ?? H264_LEVELS[H264_LEVELS.length - 1]).level;
};

/**
 * The ffmpeg command line for one export. Kept separate from running it so the
 * settings can be read, tested and shown to the user without an encode.
 *
 * `sourceIsFullRange` says how to read the recording's own levels: the browser
 * decides what MediaRecorder tags its WebM as, and getting this wrong shifts
 * every value in a test pattern - the one file where the numbers are the point.
 */
export const buildMp4Args = (settings: Mp4EncodeSettings, sourceIsFullRange: boolean): string[] => {
  const gop = keyframeIntervalFor(settings.fps);
  const level = h264LevelFor(settings.width, settings.height, settings.fps);
  const target = `${Math.round(settings.targetMbps * 1000)}k`;
  const max = `${Math.round(settings.maxMbps * 1000)}k`;
  // One second of ceiling: enough for the rate control to breathe over a
  // half-second GOP, tight enough that a server's buffer is predictable.
  const bufsize = `${Math.round(settings.maxMbps * 1000)}k`;
  return [
    "-i", "in.webm",
    "-an",
    // Exactly one loop, cut from a recording made longer on purpose.
    "-t", String(settings.loopSeconds),
    // Constant frame rate, whatever the recording managed: a wall big enough
    // to draw slower than real time still has to deliver a 60p file, so short
    // frames are repeated rather than the file being variable rate.
    "-fps_mode", "cfr",
    "-r", String(settings.fps),
    "-c:v", "libx264",
    // veryfast: WASM encoding is already far slower than native, so trade a
    // little compression efficiency for finishing this side of the show.
    "-preset", "veryfast",
    "-profile:v", "high",
    "-level:v", level,
    "-bf", "0",
    "-g", String(gop),
    "-keyint_min", String(gop),
    // Without this, a scene cut inserts its own keyframe and the GOP is no
    // longer the fixed half second a server seeks against.
    "-sc_threshold", "0",
    "-b:v", target,
    "-maxrate", max,
    "-bufsize", bufsize,
    // The colour conversion is spelled out rather than left to whatever the
    // recording was tagged as, then tagged to match on the way out.
    "-vf", `scale=in_range=${sourceIsFullRange ? "full" : "limited"}:out_range=limited,format=yuv420p`,
    "-color_primaries", "bt709",
    "-color_trc", "bt709",
    "-colorspace", "bt709",
    "-color_range", "tv",
    // write_colr puts the same tags in the MP4 itself, where a player looks
    // first; faststart moves the index to the front for streaming.
    "-movflags", "+faststart+write_colr",
    "out.mp4",
  ];
};

/**
 * Transcodes a recorded WebM Blob into an MP4 Blob (the test pattern video has
 * no audio track). `onProgress` receives 0..1 and is only called once encoding
 * itself starts - loading the ~30MB ffmpeg-core WASM happens first and isn't
 * reflected in it.
 */
export const encodeWebmToMp4 = async (
  webmBlob: Blob,
  settings: Mp4EncodeSettings,
  onProgress?: (ratio: number) => void,
): Promise<Blob> => {
  const { fetchFile } = await import("@ffmpeg/util");
  const ffmpeg = await getFFmpeg();
  const onProgressEvent = ({ progress }: { progress: number }) => {
    if (onProgress) onProgress(Math.max(0, Math.min(1, progress)));
  };
  // ffmpeg names the input's colour range in its stream line ("yuv420p(tv,
  // bt709)"). Reading it back is the only way to know what this browser's
  // recorder produced, and it decides whether the conversion below shifts the
  // levels or leaves them alone.
  let sourceIsFullRange = false;
  const onLog = ({ message }: { message: string }) => {
    if (/Video:\s*(vp0?9|vp8|av1)/i.test(message) && /\((pc|full)[,)]/i.test(message)) sourceIsFullRange = true;
  };
  ffmpeg.on("log", onLog);
  ffmpeg.on("progress", onProgressEvent);
  try {
    await ffmpeg.writeFile("in.webm", await fetchFile(webmBlob));
    // A first pass over the input alone, purely to read its stream line - it
    // decodes nothing, so it costs a moment rather than an encode.
    await ffmpeg.exec(["-i", "in.webm", "-t", "0", "-f", "null", "-"]).catch(() => 0);
    const code = await ffmpeg.exec(buildMp4Args(settings, sourceIsFullRange));
    if (code !== 0) throw new Error(`ffmpeg exited with code ${code}`);
    const data = await ffmpeg.readFile("out.mp4");
    // DOM lib types Blob's BlobPart as ArrayBufferView<ArrayBuffer> specifically,
    // while ffmpeg.wasm's readFile() returns a plain Uint8Array<ArrayBufferLike>.
    return new Blob([data as unknown as BlobPart], { type: "video/mp4" });
  } finally {
    ffmpeg.off("progress", onProgressEvent);
    ffmpeg.off("log", onLog);
    await ffmpeg.deleteFile("in.webm").catch(() => {});
    await ffmpeg.deleteFile("out.mp4").catch(() => {});
  }
};
