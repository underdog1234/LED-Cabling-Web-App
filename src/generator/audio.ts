// ---------------------------------------------------------------------------
// Sync tones: the beeps that go with the timecode and AV sync patterns.
//
// A pattern says where in its loop it beeps (`PatternDef.beeps`). The same
// list drives three things, so they can never disagree:
//  - live sound in the editor and the output window, scheduled with Web Audio
//    a moment ahead of the shared playback clock;
//  - the soundtrack of a WebM export, played into the recording as it runs;
//  - the soundtrack of an MP4 export, written as a WAV sample by sample and
//    muxed under the video, so every beep lands on its exact frame.
// ---------------------------------------------------------------------------

import type { GeneratorConfig, SubScreen } from "./model";
import { PATTERN_BY_ID, resolveSettings, type Beep } from "./patterns";

/** A beep placed on the pattern clock: it sounds whenever (time mod loop) equals `at`. */
export type ScheduledBeep = Beep;

const mod = (a: number, b: number) => ((a % b) + b) % b;

/** The beeps one screen makes, moved onto the shared clock (its animation offset taken out). */
export const screenBeeps = (screen: SubScreen, loopSeconds: number): ScheduledBeep[] => {
  if (!screen.visible) return [];
  const def = PATTERN_BY_ID.get(screen.pattern);
  if (!def?.beeps) return [];
  const loop = Math.max(0.001, loopSeconds);
  return def.beeps(resolveSettings(def, screen.settings[def.id]), loop).map((b) => ({ ...b, at: mod(b.at - (screen.phase || 0), loop) }))
    .sort((a, b) => a.at - b.at);
};

/** Every distinct beep the given screens make. Two screens beeping at the same moment make one beep. */
export const beepsFor = (screens: SubScreen[], loopSeconds: number): ScheduledBeep[] => {
  const seen = new Map<string, ScheduledBeep>();
  screens.forEach((s) =>
    screenBeeps(s, loopSeconds).forEach((b) => {
      const key = `${Math.round(b.at * 1000)}:${b.freq}`;
      if (!seen.has(key)) seen.set(key, b);
    }),
  );
  return [...seen.values()].sort((a, b) => a.at - b.at);
};

/** Every beep that starts in [from, to) on a clock that loops every `loopSeconds`, as absolute times. */
export const beepsBetween = (beeps: ScheduledBeep[], loopSeconds: number, from: number, to: number): Array<ScheduledBeep & { time: number }> => {
  const out: Array<ScheduledBeep & { time: number }> = [];
  if (!beeps.length || to <= from) return out;
  const loop = Math.max(0.001, loopSeconds);
  const first = Math.floor(from / loop);
  const last = Math.floor(to / loop);
  for (let k = first; k <= last; k += 1) {
    beeps.forEach((b) => {
      const time = k * loop + b.at;
      if (time >= from - 1e-9 && time < to - 1e-9) out.push({ ...b, time });
    });
  }
  return out.sort((a, b) => a.time - b.time);
};

// --- tone shape -----------------------------------------------------------------
// A sine with 1 ms ramps at each end: hard enough to have an unmistakable
// start, soft enough not to click.

const RAMP = 0.001;

const playTone = (ctx: BaseAudioContext, dest: AudioNode, when: number, b: ScheduledBeep) => {
  const osc = ctx.createOscillator();
  const gain = ctx.createGain();
  osc.type = "sine";
  osc.frequency.value = b.freq;
  const level = b.gain ?? 0.5;
  gain.gain.setValueAtTime(0, when);
  gain.gain.linearRampToValueAtTime(level, when + RAMP);
  gain.gain.setValueAtTime(level, when + Math.max(RAMP, b.duration - RAMP));
  gain.gain.linearRampToValueAtTime(0, when + b.duration);
  osc.connect(gain).connect(dest);
  osc.start(when);
  osc.stop(when + b.duration + 0.01);
};

/** A 16-bit mono WAV of `seconds` with every beep written in at its exact sample. */
export const renderBeepsWav = (beeps: ScheduledBeep[], loopSeconds: number, seconds: number, sampleRate = 48000): Uint8Array => {
  const n = Math.max(1, Math.round(seconds * sampleRate));
  const samples = new Int16Array(n);
  beepsBetween(beeps, loopSeconds, 0, seconds).forEach((b) => {
    const start = Math.round(b.time * sampleRate);
    const len = Math.round(b.duration * sampleRate);
    const ramp = Math.max(1, Math.round(RAMP * sampleRate));
    const level = (b.gain ?? 0.5) * 32767;
    for (let i = 0; i < len && start + i < n; i += 1) {
      const env = Math.min(1, i / ramp, (len - i) / ramp);
      const v = samples[start + i] + Math.round(Math.sin((2 * Math.PI * b.freq * i) / sampleRate) * level * env);
      samples[start + i] = Math.max(-32768, Math.min(32767, v));
    }
  });
  const out = new Uint8Array(44 + n * 2);
  const view = new DataView(out.buffer);
  const text = (at: number, s: string) => [...s].forEach((c, i) => view.setUint8(at + i, c.charCodeAt(0)));
  text(0, "RIFF");
  view.setUint32(4, 36 + n * 2, true);
  text(8, "WAVE");
  text(12, "fmt ");
  view.setUint32(16, 16, true);
  view.setUint16(20, 1, true);
  view.setUint16(22, 1, true);
  view.setUint32(24, sampleRate, true);
  view.setUint32(28, sampleRate * 2, true);
  view.setUint16(32, 2, true);
  view.setUint16(34, 16, true);
  text(36, "data");
  view.setUint32(40, n * 2, true);
  for (let i = 0; i < n; i += 1) view.setInt16(44 + i * 2, samples[i], true);
  return out;
};

// --- live sound -----------------------------------------------------------------

const LOOKAHEAD = 0.3;

/**
 * Plays the beeps of whatever is on screen, in step with the playback clock.
 * Call `tick` every animation frame; it schedules the next few hundred
 * milliseconds ahead, each beep once.
 */
export class SyncSound {
  private ctx: AudioContext | null = null;
  private scheduled = new Set<string>();
  /** Milliseconds to move the sound by: positive plays it later, to match a display that shows pictures late. */
  offsetMs = 0;
  enabled = false;

  /** Must be called from a click or key press the first time - browsers only start sound after one. */
  unlock() {
    if (typeof AudioContext === "undefined") return;
    if (!this.ctx) this.ctx = new AudioContext({ latencyHint: "interactive" });
    if (this.ctx.state === "suspended") void this.ctx.resume();
  }

  get blocked() {
    return this.enabled && (!this.ctx || this.ctx.state !== "running");
  }

  tick(config: GeneratorConfig, screens: SubScreen[], time: number, playing: boolean) {
    if (!this.enabled || !playing || !this.ctx || this.ctx.state !== "running") {
      this.scheduled.clear();
      return;
    }
    const beeps = beepsFor(screens, config.loopSeconds);
    if (!beeps.length) return;
    const ctx = this.ctx;
    // The browser's own output delay is taken off, so the beep leaves the
    // speaker when the frame it belongs to is drawn.
    const outputDelay = (ctx.outputLatency || 0) + (ctx.baseLatency || 0);
    const offset = this.offsetMs / 1000 - outputDelay;
    // Beeps that should leave the speaker from now on: on the pattern clock
    // those are the ones from (time - offset).
    const from = time - offset;
    beepsBetween(beeps, config.loopSeconds, from, from + LOOKAHEAD).forEach((b) => {
      const key = `${Math.round(b.time * 1000)}:${b.freq}`;
      if (this.scheduled.has(key)) return;
      this.scheduled.add(key);
      const when = ctx.currentTime + (b.time - time) + offset;
      if (when < ctx.currentTime - 0.005) return;
      playTone(ctx, ctx.destination, Math.max(ctx.currentTime, when), b);
    });
    if (this.scheduled.size > 256) {
      const keep = [...this.scheduled].slice(-64);
      this.scheduled = new Set(keep);
    }
  }

  hasBeeps(screens: SubScreen[], loopSeconds: number) {
    return beepsFor(screens, loopSeconds).length > 0;
  }

  /** Forget what was scheduled, e.g. after a jump in the clock. */
  reset() {
    this.scheduled.clear();
  }
}

/**
 * Live soundtrack for a WebM recording: the beeps played into a stream that
 * the recorder takes alongside the canvas. Returns null when there is nothing
 * to play.
 */
export const startRecordingSound = (
  beeps: ScheduledBeep[],
  loopSeconds: number,
  seconds: number,
): { track: MediaStreamTrack; stop: () => void; leadMs: number } | null => {
  if (!beeps.length || typeof AudioContext === "undefined") return null;
  const ctx = new AudioContext();
  const dest = ctx.createMediaStreamDestination();
  const leadMs = 50;
  const t0 = ctx.currentTime + leadMs / 1000;
  beepsBetween(beeps, loopSeconds, 0, seconds).forEach((b) => playTone(ctx, dest, t0 + b.time, b));
  const track = dest.stream.getAudioTracks()[0];
  if (!track) {
    void ctx.close();
    return null;
  }
  return { track, stop: () => void ctx.close(), leadMs };
};
