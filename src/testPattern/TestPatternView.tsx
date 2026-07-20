import React, { useEffect, useMemo, useRef } from "react";
import { type Cell, type PanelTypeKey, normalizePanels } from "../App";
import { type TestPatternLayout, type TestPatternProject, DRAW_FPS, computeTestPatternLayout, drawTestPatternFrame } from "./drawTestPattern";

export const TEST_PATTERN_STORAGE_KEY = "ledCablingTestPattern:v1";

type StoredPayload = {
  formatVersion?: number;
  projectName?: string;
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
      projectName: (data.projectName || "Untitled Project").trim() || "Untitled Project",
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
// browser's native fullscreen mode.
export default function TestPatternView() {
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const loopStartRef = useRef(performance.now());

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

  return (
    <div style={{ width: "100vw", height: "100vh", margin: 0, padding: 0, background: "#000", overflow: "hidden", display: "flex", alignItems: "center", justifyContent: "center" }} onClick={requestFullscreen}>
      <canvas
        ref={canvasRef}
        width={layout.W}
        height={layout.H}
        style={{ width: "100%", height: "100%", objectFit: "contain", display: "block", imageRendering: "pixelated", cursor: "pointer" }}
      />
    </div>
  );
}
