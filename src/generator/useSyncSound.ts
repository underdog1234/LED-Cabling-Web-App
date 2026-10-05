// Live sync tones for a window that shows patterns (the editor or an output
// window). Each window keeps its own on/off and offset in this browser.
import { useEffect, useRef, useState } from "react";
import { SyncSound } from "./audio";
import type { GeneratorConfig, SubScreen } from "./model";

export type SoundFrame = { config: GeneratorConfig; screens: SubScreen[]; time: number; playing: boolean };

type Prefs = { enabled: boolean; offsetMs: number };

const readPrefs = (key: string, defaultOn: boolean): Prefs => {
  try {
    const raw = JSON.parse(localStorage.getItem(key) ?? "null") as Partial<Prefs> | null;
    return {
      enabled: typeof raw?.enabled === "boolean" ? raw.enabled : defaultOn,
      offsetMs: typeof raw?.offsetMs === "number" && Number.isFinite(raw.offsetMs) ? raw.offsetMs : 0,
    };
  } catch {
    return { enabled: defaultOn, offsetMs: 0 };
  }
};

export function useSyncSound(source: () => SoundFrame | null, key: string, defaultOn: boolean) {
  const sound = useRef<SyncSound | null>(null);
  if (!sound.current) sound.current = new SyncSound();
  const [prefs, setPrefs] = useState<Prefs>(() => readPrefs(key, defaultOn));
  const [blocked, setBlocked] = useState(false);
  const [active, setActive] = useState(false);
  const sourceRef = useRef(source);
  sourceRef.current = source;

  useEffect(() => {
    sound.current!.enabled = prefs.enabled;
    sound.current!.offsetMs = prefs.offsetMs;
    sound.current!.reset();
    try {
      localStorage.setItem(key, JSON.stringify(prefs));
    } catch {
      // Not remembered next time; nothing else depends on it.
    }
  }, [prefs, key]);

  // Browsers only allow sound after the person has clicked or pressed a key.
  useEffect(() => {
    const unlock = () => {
      if (sound.current!.enabled) sound.current!.unlock();
    };
    window.addEventListener("pointerdown", unlock, true);
    window.addEventListener("keydown", unlock, true);
    return () => {
      window.removeEventListener("pointerdown", unlock, true);
      window.removeEventListener("keydown", unlock, true);
    };
  }, []);

  useEffect(() => {
    let raf = 0;
    let lastCheck = 0;
    const tick = (ts: number) => {
      raf = requestAnimationFrame(tick);
      const f = sourceRef.current();
      const s = sound.current!;
      if (f) s.tick(f.config, f.screens, f.time, f.playing);
      if (ts - lastCheck > 400) {
        lastCheck = ts;
        setBlocked(s.blocked);
        setActive(!!f && s.hasBeeps(f.screens, f.config.loopSeconds));
      }
    };
    raf = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(raf);
  }, []);

  return {
    enabled: prefs.enabled,
    offsetMs: prefs.offsetMs,
    /** Something on screen beeps. */
    active,
    /** Sound is on but the browser is still waiting for a click. */
    blocked,
    setEnabled: (enabled: boolean) => {
      if (enabled) sound.current!.unlock();
      setPrefs((p) => ({ ...p, enabled }));
    },
    setOffsetMs: (offsetMs: number) => setPrefs((p) => ({ ...p, offsetMs })),
  };
}
