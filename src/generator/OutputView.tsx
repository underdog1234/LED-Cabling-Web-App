// The output itself: the whole canvas, or one sub-screen, at its native
// resolution and nothing else. Used for the separate output window (opened
// with ?output=1, kept in step with the editor over a BroadcastChannel) and
// for the editor's own fullscreen output. Like the LED planner's Moving Test
// Pattern window it never silently scales: one canvas pixel lands on one
// display pixel, and anything else is reported on screen.
import { useEffect, useRef, useState } from "react";
import { computePixelMappingStatus, watchDevicePixelRatio } from "../testPattern/pixelMapping";
import mmsLogoUrl from "../testPattern/assets/mms-logo.png";
import { effectiveConfig, openChannel, playbackTime, readOutputState, type OutputState } from "./storage";
import { renderComposite, renderScreen, loadLedModule } from "./render";

type Props = {
  /** Inline (fullscreen inside the editor): the editor's live state. Omit in the separate window. */
  stateRef?: React.MutableRefObject<OutputState>;
  /** null = the whole canvas. */
  screenId: string | null;
  onClose?: () => void;
};

const bounce = (t: number, speed: number, range: number) => {
  if (range <= 0) return 0;
  const period = range * 2;
  const x = (((t * speed) % period) + period) % period;
  return x <= range ? x : period - x;
};

export default function OutputView({ stateRef, screenId, onClose }: Props) {
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const ownState = useRef<OutputState | null>(stateRef ? null : readOutputState());
  const [, setVersion] = useState(0);
  const [dpr, setDpr] = useState(() => window.devicePixelRatio || 1);
  const [isFullscreen, setIsFullscreen] = useState(() => document.fullscreenElement != null);
  const [statusVisible, setStatusVisible] = useState(true);
  const [fit, setFit] = useState(false);
  const [logo, setLogo] = useState(false);
  const [box, setBox] = useState({ w: 0, h: 0 });
  const logoRef = useRef<HTMLImageElement | null>(null);

  const current = () => (stateRef ? stateRef.current : ownState.current);

  // Separate window: follow the editor.
  useEffect(() => {
    if (stateRef) return;
    const channel = openChannel();
    if (!channel) return;
    channel.onmessage = (e) => {
      ownState.current = e.data as OutputState;
      setVersion((v) => v + 1);
    };
    return () => channel.close();
  }, [stateRef]);

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
      if (logo && img?.complete && img.naturalWidth) {
        const lw = Math.max(16, Math.round(Math.min(outW, outH) * 0.12));
        const lh = lw * (img.naturalHeight / img.naturalWidth);
        const t = now / 1000;
        ctx.drawImage(img, bounce(t, (outW - lw) / 18, outW - lw), bounce(t, (outH - lh) / 14, outH - lh), lw, lh);
      }
    };
    raf = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(raf);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [outW, outH, screenId, logo]);

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
      if (k === "h") setStatusVisible((v) => !v);
      if (k === "f") document.documentElement.requestFullscreen?.().catch(() => {});
      if (k === "l") setLogo((v) => !v);
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

  return (
    <div
      className="fixed inset-0 z-50 overflow-auto bg-black"
      style={{ background: fit ? "#000" : state.config.canvas.background }}
      onClick={() => setStatusVisible(true)}
    >
      <canvas ref={canvasRef} className="absolute left-0 top-0 block" />
      {statusVisible ? (
        <div
          className="fixed left-3 top-3 z-10 max-w-xs cursor-pointer space-y-2 font-mono text-xs"
          onClick={(e) => {
            e.stopPropagation();
            setStatusVisible(false);
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
            <Row label="Browser -> Output:" value={fit ? "Scaled (fit)" : scaled ? "Scaled" : "1:1"} warn={scaled && !fit} />
            <div className="flex flex-wrap gap-1 pt-1 font-sans">
              <button type="button" className="rounded border border-slate-500 bg-slate-800 px-2 py-0.5 text-white" onClick={(e) => { e.stopPropagation(); setFit((v) => !v); }}>
                {fit ? "Native 1:1" : "Fit to output"}
              </button>
              <button type="button" className="rounded border border-slate-500 bg-slate-800 px-2 py-0.5 text-white" onClick={(e) => { e.stopPropagation(); setLogo((v) => !v); }}>
                {logo ? "Hide logo" : "Bouncing logo"}
              </button>
              {onClose ? (
                <button type="button" className="rounded border border-slate-500 bg-slate-800 px-2 py-0.5 text-white" onClick={(e) => { e.stopPropagation(); document.exitFullscreen?.().catch(() => {}); onClose(); }}>
                  Close output
                </button>
              ) : null}
            </div>
            <div className="pt-1 text-[10px] text-slate-500">Click to hide · H toggles · F fullscreen · L logo (live only, never exported){onClose ? " · Esc closes" : ""}</div>
          </div>
        </div>
      ) : null}
      {!isFullscreen ? (
        <button
          type="button"
          className="fixed bottom-4 right-4 z-10 rounded-lg border border-white/30 bg-black/60 px-4 py-2 text-sm font-bold tracking-widest text-white hover:bg-black/80"
          onClick={(e) => {
            e.stopPropagation();
            document.documentElement.requestFullscreen?.().catch(() => {});
          }}
        >
          ENTER FULLSCREEN
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
