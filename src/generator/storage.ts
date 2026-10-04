// ---------------------------------------------------------------------------
// Persistence and the link to the output window.
//
// The working setup is autosaved; named setups and named resolution presets
// are kept in this browser too, and any setup can go to and from a JSON file.
// The output window gets the whole state once through localStorage (it may
// open before any message could reach it) and every change after that over a
// BroadcastChannel.
// ---------------------------------------------------------------------------

import { applyPlaylistStep, normalizeConfig, normalizePresets, playlistStepAt, type GeneratorConfig, type ResolutionPreset } from "./model";

const AUTOSAVE_KEY = "testPatternGenerator:config:v1";
const SAVED_KEY = "testPatternGenerator:saved:v1";
const PRESETS_KEY = "testPatternGenerator:presets:v1";
const OUTPUT_KEY = "testPatternGenerator:output:v1";
export const OUTPUT_CHANNEL = "testPatternGenerator:output";

const read = (key: string): unknown => {
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
};

const write = (key: string, value: unknown): boolean => {
  try {
    localStorage.setItem(key, JSON.stringify(value));
    return true;
  } catch {
    return false;
  }
};

export const loadAutosave = (): GeneratorConfig | null => {
  const raw = read(AUTOSAVE_KEY);
  return raw ? normalizeConfig(raw) : null;
};

export const saveAutosave = (config: GeneratorConfig) => write(AUTOSAVE_KEY, config);

export type SavedConfig = { name: string; savedAt: string; config: GeneratorConfig };

export const loadSavedConfigs = (): SavedConfig[] => {
  const raw = read(SAVED_KEY);
  if (!Array.isArray(raw)) return [];
  return raw.flatMap((entry) => {
    const e = entry as Record<string, unknown>;
    if (!e || typeof e.name !== "string") return [];
    return [{ name: e.name, savedAt: typeof e.savedAt === "string" ? e.savedAt : "", config: normalizeConfig(e.config) }];
  });
};

export const storeSavedConfigs = (list: SavedConfig[]) => write(SAVED_KEY, list);

export const loadCustomPresets = (): ResolutionPreset[] => normalizePresets(read(PRESETS_KEY));
export const storeCustomPresets = (list: ResolutionPreset[]) => write(PRESETS_KEY, list);

/** The clock both windows run patterns and playlists from. Wall-clock based, so separate windows agree. */
export type Playback = {
  playing: boolean;
  /** Date.now() when `anchorTime` was current. */
  anchor: number;
  /** Pattern clock, seconds, at `anchor`. */
  anchorTime: number;
  playlist: { active: boolean; loop: boolean; startTime: number };
};

export const defaultPlayback = (): Playback => ({ playing: true, anchor: Date.now(), anchorTime: 0, playlist: { active: false, loop: true, startTime: 0 } });

export const playbackTime = (p: Playback, now = Date.now()): number => (p.playing ? p.anchorTime + (now - p.anchor) / 1000 : p.anchorTime);

export type OutputState = { config: GeneratorConfig; playback: Playback };

export const publishOutputState = (state: OutputState, channel: BroadcastChannel | null) => {
  write(OUTPUT_KEY, state);
  try {
    channel?.postMessage(state);
  } catch {
    // The output window simply keeps showing the last state it had.
  }
};

export const readOutputState = (): OutputState | null => {
  const raw = read(OUTPUT_KEY) as { config?: unknown; playback?: Playback } | null;
  if (!raw || !raw.config) return null;
  return { config: normalizeConfig(raw.config), playback: raw.playback ?? defaultPlayback() };
};

export const openChannel = (): BroadcastChannel | null => {
  try {
    return typeof BroadcastChannel === "undefined" ? null : new BroadcastChannel(OUTPUT_CHANNEL);
  } catch {
    return null;
  }
};

/** The config as it should be shown at `now`: the editor's own, or with the current playlist step laid over it. */
export const effectiveConfig = (config: GeneratorConfig, playback: Playback, now = Date.now()): { config: GeneratorConfig; stepIndex: number | null } => {
  if (!playback.playlist.active || !config.playlist.length) return { config, stepIndex: null };
  const at = playlistStepAt(config.playlist, playbackTime(playback, now) - playback.playlist.startTime, playback.playlist.loop);
  if (!at) return { config, stepIndex: null };
  return { config: applyPlaylistStep(config, config.playlist[at.index]), stepIndex: at.index };
};
