import React, { useEffect, useMemo, useRef, useState } from "react";
import { type Cell, type PanelTypeKey, normalizePanels } from "../App";
import { type TestPatternLayout, type TestPatternProject, DRAW_FPS, computeTestPatternLayout, drawTestPatternFrame, drawBouncingLogo } from "./drawTestPattern";
import mmsLogoUrl from "./assets/mms-logo.png";

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
// browser's native fullscreen mode.
export default function TestPatternView() {
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const loopStartRef = useRef(performance.now());
  // Loaded once and reused every frame - a bouncing DVD-screensaver-style
  // logo, browser-live-view only (never the recorded video or PNG/PDF
  // exports, which all go through drawTestPatternFrame alone).
  const logoImgRef = useRef<HTMLImageElement | null>(null);
  // Anchored top-left, scaled to fit the available browser window while
  // preserving the LED canvas's own aspect ratio - never stretched, cropped,
  // centred, or auto-rotated between portrait/landscape. Recomputed on resize.
  const [viewport, setViewport] = useState({ w: window.innerWidth, h: window.innerHeight });

  const project = useMemo(loadProject, []);
  const layout: TestPatternLayout | null = useMemo(() => (project ? computeTestPatternLayout(project) : null), [project]);

  useEffect(() => {
    document.title = project?.projectName ? `Moving Test Pattern - ${project.projectName}` : "Moving Test Pattern";
  }, [project]);

  useEffect(() => {
    const img = new Image();
    img.src = mmsLogoUrl;
    logoImgRef.current = img;
  }, []);

  // Animation loop, capped at DRAW_FPS. Uses setInterval rather than
  // requestAnimationFrame so it keeps running even if the tab is momentarily
  // backgrounded - browsers suspend rAF in hidden tabs, but timers keep firing.
  //
  // The canvas's backing store (canvas.width/height, in real device pixels)
  // is sized to match exactly how many physical pixels it will actually be
  // displayed at - the CSS-fit-to-window size times devicePixelRatio - not
  // just the LED wall's own native pixel count. Content is still drawn in
  // the wall's native layout.W x layout.H coordinate space (unchanged), via
  // a scale transform, so this only affects rendering fidelity: without it,
  // a canvas.width == layout.W backing store gets stretched or shrunk by the
  // browser's CSS box sizing, which is exactly what makes thin lines/text/
  // arrows blur or look sub-pixel on high-DPI displays or when the window
  // doesn't match the wall's own resolution 1:1.
  useEffect(() => {
    if (!layout) return;
    const canvas = canvasRef.current;
    const ctx = canvas?.getContext("2d");
    if (!canvas || !ctx) return;
    const dpr = window.devicePixelRatio || 1;
    const cssScale = Math.min(viewport.w / layout.W, viewport.h / layout.H);
    const pixelScale = cssScale * dpr;
    canvas.width = Math.max(1, Math.round(layout.W * pixelScale));
    canvas.height = Math.max(1, Math.round(layout.H * pixelScale));
    const id = window.setInterval(() => {
      const t = (performance.now() - loopStartRef.current) / 1000;
      // Defensive: re-applied every frame rather than relying on it surviving
      // drawTestPatternFrame's own internal save/restore pairs untouched.
      ctx.setTransform(pixelScale, 0, 0, pixelScale, 0, 0);
      drawTestPatternFrame(ctx, layout, t);
      const logo = logoImgRef.current;
      if (logo && logo.complete && logo.naturalWidth) drawBouncingLogo(ctx, layout, t, logo);
    }, 1000 / DRAW_FPS);
    return () => window.clearInterval(id);
  }, [layout, viewport]);

  useEffect(() => {
    const onResize = () => setViewport({ w: window.innerWidth, h: window.innerHeight });
    window.addEventListener("resize", onResize);
    return () => window.removeEventListener("resize", onResize);
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
            Open this page from the main app's <b>Moving Test Pattern</b> button.
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

  // Scale to fit the available window on both axes (never past either one, so
  // nothing is cropped), using a single uniform factor so the aspect ratio is
  // always preserved and the canvas is never stretched. Rounded to whole
  // device pixels so the LED pixel grid stays crisp rather than blurring
  // across a fractional-pixel boundary.
  const scale = Math.min(viewport.w / layout.W, viewport.h / layout.H);
  const displayW = Math.max(1, Math.round(layout.W * scale));
  const displayH = Math.max(1, Math.round(layout.H * scale));

  // Backing-store pixel count (canvas.width/height) is owned imperatively by
  // the draw-loop effect above, not by React/JSX - it's sized in real device
  // pixels (CSS size x devicePixelRatio), not the wall's native resolution,
  // so no width/height props here (they'd fight the effect every render).
  // imageRendering stays at the browser default (smooth) since the backing
  // store is deliberately kept matched to its displayed CSS size - relying on
  // nearest-neighbour "pixelated" scaling here is exactly what used to make
  // thin lines/text/arrows blur or look jagged.
  return (
    <div style={{ position: "fixed", inset: 0, margin: 0, padding: 0, background: "#000", overflow: "hidden" }} onClick={requestFullscreen}>
      <canvas
        ref={canvasRef}
        style={{
          position: "absolute",
          top: 0,
          left: 0,
          width: displayW,
          height: displayH,
          display: "block",
          cursor: "pointer",
        }}
      />
    </div>
  );
}
