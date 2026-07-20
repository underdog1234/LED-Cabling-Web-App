import React, { useEffect, useMemo, useRef, useState } from "react";
import { type Cell, type PanelTypeKey, normalizePanels } from "../App";
import { type TestPatternLayout, type TestPatternProject, DRAW_FPS, computeTestPatternLayout, drawTestPatternFrame } from "./drawTestPattern";

export const TEST_PATTERN_STORAGE_KEY = "ledCablingTestPattern:v1";

type StoredPayload = {
  formatVersion?: number;
  projectName?: string;
  surfaceName?: string;
  panelType?: PanelTypeKey;
  panels?: unknown;
};

const loadProject = (): TestPatternProject | null => {
  try {
    const raw = localStorage.getItem(TEST_PATTERN_STORAGE_KEY);
    if (!raw) return null;
    const data = JSON.parse(raw) as StoredPayload;
    const panels: Cell[] = normalizePanels(data.panels);
    if (!panels.length) return null;
    return {
      projectName: (data.projectName || "").trim(),
      surfaceName: (data.surfaceName || "").trim(),
      panelType: data.panelType && (data.panelType === "MG9" || data.panelType === "MT") ? data.panelType : "MG9",
      panels,
    };
  } catch {
    return null;
  }
};

// Pure full-screen live view: the canvas and nothing else. No header, no
// buttons, no text outside the LED canvas itself (the wall info/labels are
// drawn ON the canvas by drawTestPatternFrame). A click anywhere requests the
// browser's native fullscreen mode; pressing any key toggles between the
// default true 1:1 pixel mapping and an optional scaled-to-fit preview -
// there's no on-screen control for this so the "no text outside the canvas"
// rule holds either way.
export default function TestPatternView() {
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const loopStartRef = useRef(performance.now());
  const [scaledFit, setScaledFit] = useState(false);

  const project = useMemo(loadProject, []);
  const layout: TestPatternLayout | null = useMemo(() => (project ? computeTestPatternLayout(project) : null), [project]);

  // Animation loop, capped at DRAW_FPS. Uses setInterval rather than
  // requestAnimationFrame so it keeps running even if the tab is momentarily
  // backgrounded - browsers suspend rAF in hidden tabs, but timers keep firing.
  useEffect(() => {
    if (!layout) return;
    const canvas = canvasRef.current;
    const ctx = canvas?.getContext("2d");
    if (!canvas || !ctx) return;
    const id = window.setInterval(() => {
      drawTestPatternFrame(ctx, layout, (performance.now() - loopStartRef.current) / 1000);
    }, 1000 / DRAW_FPS);
    return () => window.clearInterval(id);
  }, [layout]);

  useEffect(() => {
    const onKeyDown = () => setScaledFit((v) => !v);
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, []);

  const requestFullscreen = () => {
    const el = document.documentElement;
    if (document.fullscreenElement) document.exitFullscreen?.();
    else el.requestFullscreen?.().catch(() => {});
  };

  if (!layout) {
    return (
      <div style={{ minHeight: "100vh", background: "#0f172a", color: "#cbd5e1", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "system-ui, -apple-system, sans-serif", padding: 24, textAlign: "center" }}>
        <div>
          <div style={{ fontSize: 15, marginBottom: 8 }}>No project data found for this tab.</div>
          <div style={{ fontSize: 13, color: "#94a3b8" }}>
            Open this page from the main app's <b>Video Test Pattern</b> button.
          </div>
          <div style={{ marginTop: 16 }}>
            <a href={location.pathname} style={{ color: "#38bdf8" }}>
              Back to the LED Cabling Planner
            </a>
          </div>
        </div>
      </div>
    );
  }

  // Default: true 1:1 pixel mapping - the canvas's CSS size is fixed to its
  // exact native pixel dimensions (never stretched/scaled), centred when it's
  // smaller than the viewport, scrollable when it's larger. The outer element
  // scrolls; the inner flex wrapper is at least 100% of the viewport so
  // centring still works when the content fits, and grows past 100% (with
  // scrollbars) when it doesn't.
  const canvasStyle: React.CSSProperties = scaledFit
    ? { width: "auto", height: "auto", maxWidth: "100vw", maxHeight: "100vh", display: "block", imageRendering: "pixelated", cursor: "pointer", flexShrink: 0 }
    : { width: `${layout.W}px`, height: `${layout.H}px`, display: "block", imageRendering: "pixelated", cursor: "pointer", flexShrink: 0 };

  return (
    <div style={{ width: "100vw", height: "100vh", margin: 0, padding: 0, background: "#000", overflow: "auto" }}>
      <div
        style={{ minWidth: "100%", minHeight: "100%", width: "max-content", display: "flex", alignItems: "center", justifyContent: "center" }}
        onClick={requestFullscreen}
      >
        <canvas ref={canvasRef} width={layout.W} height={layout.H} style={canvasStyle} />
      </div>
    </div>
  );
}
