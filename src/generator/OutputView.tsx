// The output itself: the whole canvas, or one sub-screen, at its native
// resolution and nothing else. Used for the separate output window (opened
// with ?output=1, kept in step with the editor over a BroadcastChannel) and
// for the editor's own fullscreen output. Like the LED planner's Moving Test
// Pattern window it never silently scales: one canvas pixel lands on one
// display pixel, and anything else is reported on screen.
import { useEffect, useRef, useState } from "react";
import { computePixelMappingStatus, watchDevicePixelRatio } from "../testPattern/pixelMapping";
import mmsLogoUrl from "../testPattern/assets/mms-logo.png";
import { defaultDisplay, effectiveConfig, openChannel, playbackTime, readOutputState, type OutputDisplay, type OutputRequest, type OutputState } from "./storage";
import type { GeneratorConfig, Rect } from "./model";
import { renderComposite, renderScreen, loadLedModule } from "./render";
import { useSyncSound } from "./useSyncSound";

type Props = {
  /** Inline (fullscreen inside the editor): the editor's live state. Omit in the separate window. */
  stateRef?: React.MutableRefObject<OutputState>;
  /** null = the whole canvas. */
  screenId: string | null;
  onClose?: () => void;
  /** Inline only: change the display settings (the separate window asks over the channel instead). */
  onDisplay?: (patch: Partial<OutputDisplay>) => void;
};

const bounce = (t: number, speed: number, range: number) => {
  if (range <= 0) return 0;
  const period = range * 2;
  const x = (((t * speed) % period) + period) % period;
  return x <= range ? x : period - x;
};

// Wall-clock based, so the logo keeps moving while the pattern is paused.
const drawLogo = (ctx: CanvasRenderingContext2D, img: HTMLImageElement, area: Rect, t: number, seed: number) => {
  const lw = Math.max(16, Math.round(Math.min(area.w, area.h) * 0.18));
  const lh = lw * (img.naturalHeight / img.naturalWidth);
  if (lw > area.w || lh > area.h) return;
  // Each screen gets its own start point and pace, so they don't move in lockstep.
  const rx = area.w - lw;
  const ry = area.h - lh;
  const x = bounce(t + seed * 3.7, Math.max(20, rx / 9), rx);
  const y = bounce(t + seed * 2.3, Math.max(20, ry / 7), ry);
  ctx.save();
  ctx.beginPath();
  ctx.rect(area.x, area.y, area.w, area.h);
  ctx.clip();
  ctx.drawImage(img, area.x + x, area.y + y, lw, lh);
  ctx.restore();
};

/** Where the logo bounces: inside every visible sub-screen on the canvas, or the one screen being shown. */
const logoAreas = (config: GeneratorConfig, screenId: string | null): Array<{ rect: Rect; seed: number }> => {
  if (screenId) {
    const s = config.screens.find((x) => x.id === screenId);
    return s ? [{ rect: { x: 0, y: 0, w: s.w, h: s.h }, seed: 0 }] : [];
  }
  const { w, h } = config.canvas;
  return config.screens.flatMap((s, i) => {
    if (!s.visible) return [];
    const x0 = Math.max(0, s.x);
    const y0 = Math.max(0, s.y);
    const x1 = Math.min(w, s.x + s.w);
    const y1 = Math.min(h, s.y + s.h);
    return x1 > x0 && y1 > y0 ? [{ rect: { x: x0, y: y0, w: x1 - x0, h: y1 - y0 }, seed: i }] : [];
  });
};

export default function OutputView({ stateRef, screenId, onClose, onDisplay }: Props) {
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const ownState = useRef<OutputState | null>(stateRef ? null : readOutputState());
  const [, setVersion] = useState(0);
  const [dpr, setDpr] = useState(() => window.devicePixelRatio || 1);
  const [isFullscreen, setIsFullscreen] = useState(() => document.fullscreenElement != null);
  const [box, setBox] = useState({ w: 0, h: 0 });
  const logoRef = useRef<HTMLImageElement | null>(null);

  const current = () => (stateRef ? stateRef.current : ownState.current);
  const channelRef = useRef<BroadcastChannel | null>(null);
  const display: OutputDisplay = current()?.display ?? defaultDisplay();
  const { fit, logo, status: statusVisible } = display;
  /** Display changes always go through the editor, so its buttons stay in step. */
  const changeDisplay = (patch: Partial<OutputDisplay>) => {
    if (onDisplay) onDisplay(patch);
    else channelRef.current?.postMessage({ type: "display", patch } satisfies OutputRequest);
  };
  const changeRef = useRef(changeDisplay);
  changeRef.current = changeDisplay;
  const displayRef = useRef(display);
  displayRef.current = display;

  // Inline output shares the editor's sound setting; a separate window has
  // its own, off to begin with so the two windows don't both beep.
  const sound = useSyncSound(
    () => {
      const s = current();
      if (!s) return null;
      const now = Date.now();
      const { config } = effectiveConfig(s.config, s.playback, now);
      const screens = screenId ? config.screens.filter((x) => x.id === screenId) : config.screens;
      return { config, screens, time: playbackTime(s.playback, now), playing: s.playback.playing };
    },
    "testPatternGenerator:sound:editor",
    true,
    // The separate window beeps when the editor's "Beeps in output window" is on.
    stateRef ? undefined : { enabled: display.beeps, offsetMs: display.beepOffsetMs, onEnabled: (v) => changeDisplay({ beeps: v }) },
  );
  const soundRef = useRef(sound);
  soundRef.current = sound;

  // Separate window: follow the editor, and tell it whether this window is
  // fullscreen and 1:1 (the editor shows a warning when it isn't).
  useEffect(() => {
    if (stateRef) return;
    const channel = openChannel();
    if (!channel) return;
    channelRef.current = channel;
    channel.onmessage = (e) => {
      const data = e.data as OutputState | OutputRequest;
      if ("type" in data) return;
      ownState.current = { ...data, display: { ...defaultDisplay(), ...(data.display ?? {}) } };
      setVersion((v) => v + 1);
    };
    const id = window.setInterval(() => {
      channel.postMessage({ type: "status", fullscreen: document.fullscreenElement != null, oneToOne: oneToOneRef.current } satisfies OutputRequest);
    }, 1000);
    return () => {
      window.clearInterval(id);
      channelRef.current = null;
      channel.close();
    };
  }, [stateRef]);
  const oneToOneRef = useRef(true);

  useEffect(() => {
    const img = new Image();
    img.src = mmsLogoUrl;
    logoRef.current = img;
  }, []);

  const state = current();
  const target = state ? (screenId ? state.config.screens.find((s) => s.id === screenId) ?? null : null) : null;
  const outW = state ? (target ? target.w : state.config.canvas.w) : 0;
  const outH = state ? (target ? target.h : state.config.canvas.h) : 0;
  const hasLed = !!state?.config.screens.some((s) => s.ledLayout);

  useEffect(() => {
    if (hasLed) void loadLedModule();
  }, [hasLed]);

  useEffect(() => {
    if (!stateRef) document.title = target ? `Output - ${target.name}` : `Output - ${state?.config.name ?? "Test Pattern Generator"}`;
  }, [stateRef, target, state?.config.name]);

  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas || !outW || !outH) return;
    canvas.width = outW;
    canvas.height = outH;
    const ctx = canvas.getContext("2d")!;
    const cache = new Map<string, HTMLCanvasElement>();
    let raf = 0;
    const tick = () => {
      raf = requestAnimationFrame(tick);
      const s = current();
      if (!s) return;
      const now = Date.now();
      const { config } = effectiveConfig(s.config, s.playback, now);
      const time = playbackTime(s.playback, now);
      if (screenId) {
        const screen = config.screens.find((x) => x.id === screenId);
        if (!screen) return;
        ctx.setTransform(1, 0, 0, 1, 0, 0);
        renderScreen(ctx, screen, time, { loopSeconds: config.loopSeconds, canvas: { w: config.canvas.w, h: config.canvas.h } });
      } else {
        renderComposite(ctx, config, time, cache);
      }
      const img = logoRef.current;
      if (displayRef.current.logo && img?.complete && img.naturalWidth) {
        ctx.setTransform(1, 0, 0, 1, 0, 0);
        logoAreas(config, screenId).forEach(({ rect, seed }) => drawLogo(ctx, img, rect, now / 1000, seed));
      }
    };
    raf = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(raf);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [outW, outH, screenId]);

  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas || !outW) return;
    if (fit) {
      canvas.style.width = "100%";
      canvas.style.height = "100%";
      canvas.style.objectFit = "contain";
    } else {
      canvas.style.width = `${outW / dpr}px`;
      canvas.style.height = `${outH / dpr}px`;
      canvas.style.objectFit = "";
    }
  }, [outW, outH, dpr, fit]);

  useEffect(() => watchDevicePixelRatio(setDpr), []);

  // Inline in the editor, leaving fullscreen (Esc, or the browser's own
  // control) is leaving the output.
  useEffect(() => {
    if (!onClose) return;
    let was = document.fullscreenElement != null;
    const onChange = () => {
      const is = document.fullscreenElement != null;
      if (was && !is) onClose();
      was = is;
    };
    document.addEventListener("fullscreenchange", onChange);
    return () => document.removeEventListener("fullscreenchange", onChange);
  }, [onClose]);

  useEffect(() => {
    const onChange = () => setIsFullscreen(document.fullscreenElement != null);
    document.addEventListener("fullscreenchange", onChange);
    return () => document.removeEventListener("fullscreenchange", onChange);
  }, []);

  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const measure = () => {
      const r = canvas.getBoundingClientRect();
      setBox({ w: r.width, h: r.height });
    };
    measure();
    const ro = new ResizeObserver(measure);
    ro.observe(canvas);
    window.addEventListener("resize", measure);
    return () => {
      ro.disconnect();
      window.removeEventListener("resize", measure);
    };
  }, [outW, outH]);

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const k = e.key.toLowerCase();
      if (k === "h") changeRef.current({ status: !displayRef.current.status });
      if (k === "f") document.documentElement.requestFullscreen?.().catch(() => {});
      if (k === "l") changeRef.current({ logo: !displayRef.current.logo });
      if (k === " ") e.preventDefault();
      if (k === "s") soundRef.current.setEnabled(!soundRef.current.enabled);
      if (e.key === "Escape" && onClose) {
        if (document.fullscreenElement) document.exitFullscreen?.().catch(() => {});
        onClose();
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [onClose]);

  if (!state || (screenId && !target)) {
    return (
      <div className="fixed inset-0 flex items-center justify-center bg-slate-950 p-6 text-center text-slate-300">
        <div>
          <div className="mb-2">{screenId ? "That sub-screen no longer exists in the editor." : "No generator state found for this window."}</div>
          <div className="text-sm text-slate-500">Open the output from the Test Pattern Generator&apos;s Output menu.</div>
        </div>
      </div>
    );
  }

  const mapping = computePixelMappingStatus(outW, outH, box.w, box.h, dpr);
  const displayW = Math.round((window.screen?.width ?? 0) * dpr);
  const displayH = Math.round((window.screen?.height ?? 0) * dpr);
  const scaled = mapping === "scaled";
  oneToOneRef.current = !scaled && !fit;
  const enterFullscreen = () => document.documentElement.requestFullscreen?.().catch(() => {});

  return (
    <div
      className="fixed inset-0 z-50 overflow-auto bg-black"
      style={{ background: fit ? "#000" : state.config.canvas.background }}
      onClick={() => {
        if (!statusVisible) changeDisplay({ status: true });
      }}
    >
      <canvas ref={canvasRef} className="absolute left-0 top-0 block" />
      {statusVisible ? (
        <div
          className="fixed left-3 top-3 z-10 max-w-xs cursor-pointer space-y-2 font-mono text-xs"
          onClick={(e) => {
            e.stopPropagation();
            changeDisplay({ status: false });
          }}
        >
          {scaled ? (
            <div className={`rounded-lg border-2 px-3 py-2 font-sans text-sm font-bold shadow-lg ${fit ? "border-sky-500 bg-sky-950/90 text-sky-200" : "border-red-500 bg-red-950/90 text-red-200"}`}>
              {fit ? "Scaled to fit output" : "⚠ OUTPUT IS NOT BEING DISPLAYED 1:1"}
            </div>
          ) : null}
          <div className="space-y-1 rounded-lg border border-slate-600 bg-slate-950/85 px-3 py-2 text-slate-200 shadow-lg">
            <Row label="Output:" value={target ? target.name : `${state.config.name} (full canvas)`} />
            <Row label="Output Resolution:" value={`${outW} x ${outH}`} />
            <Row label="Display Resolution:" value={displayW && displayH ? `${displayW} x ${displayH} (approx.)` : "Unknown"} />
            <Row label="Device Pixel Ratio:" value={dpr.toFixed(2)} />
            <Row label="Fullscreen:" value={isFullscreen ? "Yes" : "No"} />
            {sound.enabled && sound.blocked && sound.active ? <Row label="Sound:" value="Click to allow" warn /> : null}
            <Row label="Browser -> Output:" value={fit ? "Scaled (fit)" : scaled ? "Scaled" : "1:1"} warn={scaled && !fit} />
            <div className="flex flex-wrap gap-1 pt-1 font-sans">
              <button type="button" className="rounded border border-slate-500 bg-slate-800 px-2 py-0.5 text-white" onClick={(e) => { e.stopPropagation(); changeDisplay({ fit: !fit }); }}>
                {fit ? "Native 1:1" : "Fit to output"}
              </button>
              <button type="button" className="rounded border border-slate-500 bg-slate-800 px-2 py-0.5 text-white" onClick={(e) => { e.stopPropagation(); changeDisplay({ logo: !logo }); }}>
                {logo ? "Hide logo" : "Bouncing logo"}
              </button>
              <button type="button" className="rounded border border-slate-500 bg-slate-800 px-2 py-0.5 text-white" onClick={(e) => { e.stopPropagation(); sound.setEnabled(!sound.enabled); }}>
                {sound.enabled ? "🔊 Beeps on" : "🔇 Beeps off"}
              </button>
              {onClose ? (
                <button type="button" className="rounded border border-slate-500 bg-slate-800 px-2 py-0.5 text-white" onClick={(e) => { e.stopPropagation(); document.exitFullscreen?.().catch(() => {}); onClose(); }}>
                  Close output
                </button>
              ) : null}
            </div>
            <div className="pt-1 text-[10px] text-slate-500">Click to hide · H toggles · F fullscreen · L logo (live only, never exported) · S beeps{onClose ? " · Esc closes" : ""}</div>
          </div>
        </div>
      ) : null}
      {!isFullscreen ? (
        <button
          type="button"
          className="fixed left-1/2 top-3 z-20 -translate-x-1/2 rounded-lg border-2 border-amber-400 bg-amber-500/90 px-5 py-2 text-sm font-bold text-black shadow-lg hover:bg-amber-400"
          onClick={(e) => {
            e.stopPropagation();
            enterFullscreen();
          }}
        >
          ⚠ NOT FULLSCREEN - click here to go fullscreen
        </button>
      ) : null}
    </div>
  );
}

const Row = ({ label, value, warn }: { label: string; value: string; warn?: boolean }) => (
  <div className="flex justify-between gap-4">
    <span className="text-slate-400">{label}</span>
    <span className={warn ? "font-semibold text-red-300" : "font-semibold text-white"}>{value}</span>
  </div>
);
