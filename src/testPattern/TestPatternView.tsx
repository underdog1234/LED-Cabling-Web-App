import React, { useEffect, useMemo, useRef, useState } from "react";
import { type Cell, type PanelTypeKey, normalizePanels } from "../App";
import { type TestPatternLayout, type TestPatternProject, DRAW_FPS, computeTestPatternLayout, drawTestPatternFrame, drawBouncingLogo } from "./drawTestPattern";
import { computePixelMappingStatus, watchDevicePixelRatio } from "./pixelMapping";
import { requestScreenDetails } from "./screenPlacement";
import TestPatternStatusOverlay, { type TestPatternStatusInfo } from "./TestPatternStatusOverlay";
import mmsLogoUrl from "./assets/mms-logo.png";

export const TEST_PATTERN_STORAGE_KEY = "ledCablingTestPattern:v1";

type StoredPayload = {
  formatVersion?: number;
  projectName?: string;
  surfaceName?: string;
  panelType?: PanelTypeKey;
  panels?: unknown;
  /** v2+. Absent on payloads written before per-sub-screen patterns existed, which then render as one whole-wall surface exactly as they used to. */
  subScreens?: unknown;
};

const normalizeStoredSubScreens = (raw: unknown): Array<{ id: string; name: string; color: string }> => {
  if (!Array.isArray(raw)) return [];
  return raw.flatMap((item) => {
    const entry = item as { id?: unknown; name?: unknown; color?: unknown };
    if (typeof entry?.id !== "string" || !entry.id) return [];
    return [{
      id: entry.id,
      name: typeof entry.name === "string" ? entry.name : "",
      color: typeof entry.color === "string" && /^#[0-9a-fA-F]{6}$/.test(entry.color) ? entry.color : "#ffffff",
    }];
  });
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
      subScreens: normalizeStoredSubScreens(data.subScreens),
    };
  } catch {
    return null;
  }
};

// Pure full-screen live view: the canvas and nothing else besides the status
// overlay and (when not fullscreen) the ENTER FULLSCREEN control - no header,
// no other buttons, no text baked into the page outside those (the wall
// info/labels are drawn ON the canvas by drawTestPatternFrame).
export default function TestPatternView() {
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const loopStartRef = useRef(performance.now());
  // Loaded once and reused every frame - a bouncing DVD-screensaver-style
  // logo, browser-live-view only (never the recorded video or PNG/PDF
  // exports, which all go through drawTestPatternFrame alone).
  const logoImgRef = useRef<HTMLImageElement | null>(null);

  const project = useMemo(loadProject, []);
  const layout: TestPatternLayout | null = useMemo(() => (project ? computeTestPatternLayout(project) : null), [project]);

  const [dpr, setDpr] = useState(() => window.devicePixelRatio || 1);
  const [isFullscreen, setIsFullscreen] = useState(() => document.fullscreenElement != null);
  // The canvas's ACTUAL rendered CSS box, measured (never assumed) - see the
  // ResizeObserver effect below and pixelMapping.ts's comment on why this is
  // read from the DOM rather than computed from the numbers we set.
  const [renderedCssBox, setRenderedCssBox] = useState({ w: 0, h: 0 });
  const [displayInfo, setDisplayInfo] = useState<{ w: number; h: number; exact: boolean } | null>(null);
  // null = not yet determined; true/false once requestScreenDetails resolves.
  const [screenDetailsAvailable, setScreenDetailsAvailable] = useState<boolean | null>(null);
  const [statusVisible, setStatusVisible] = useState(true);
  // Opt-in escape hatch for when the test pattern resolution is bigger than
  // the display can show 1:1 - deliberately OFF by default (native
  // resolution, never silently scaled), only ever engaged by the user
  // clicking "Fit to Output" in the status panel.
  const [fitToOutput, setFitToOutput] = useState(false);

  useEffect(() => {
    document.title = project?.projectName ? `Moving Test Pattern - ${project.projectName}` : "Moving Test Pattern";
  }, [project]);

  useEffect(() => {
    const img = new Image();
    img.src = mmsLogoUrl;
    logoImgRef.current = img;
  }, []);

  // Canvas backing store (canvas.width/height, in real device pixels) is
  // sized to EXACTLY the layout's Recommended Content Resolution
  // (contentPixelW x contentPixelH) - always, regardless of window size or
  // devicePixelRatio. Never scaled/fit to whatever's available - if the
  // display can't show this 1:1, the mismatch is reported (see
  // TestPatternStatusOverlay), never silently resolved by shrinking the
  // canvas. Drawing itself stays entirely in the wall's native W x H
  // coordinate space (unchanged) via a single vertical scale transform - a
  // no-op for non-MT walls (contentPixelH === H there) - matching exactly the
  // technique used by the PNG/WebM exports in App.tsx (see
  // drawTestPatternFrame's own comment on why this is the ONLY place MT's
  // content-resolution doubling needs to be implemented).
  useEffect(() => {
    if (!layout) return;
    const canvas = canvasRef.current;
    const ctx = canvas?.getContext("2d");
    if (!canvas || !ctx) return;
    canvas.width = Math.max(1, layout.contentPixelW);
    canvas.height = Math.max(1, layout.contentPixelH);
    const scaleY = layout.contentPixelH / layout.H;
    const id = window.setInterval(() => {
      const t = (performance.now() - loopStartRef.current) / 1000;
      // Defensive: re-applied every frame rather than relying on it surviving
      // drawTestPatternFrame's own internal save/restore pairs untouched.
      ctx.setTransform(1, 0, 0, scaleY, 0, 0);
      drawTestPatternFrame(ctx, layout, t);
      const logo = logoImgRef.current;
      if (logo && logo.complete && logo.naturalWidth) drawBouncingLogo(ctx, layout, t, logo);
    }, 1000 / DRAW_FPS);
    return () => window.clearInterval(id);
  }, [layout]);

  // CSS display size = backing-store size / devicePixelRatio, in CSS pixels -
  // this is what makes 1 canvas backing pixel land on 1 physical display
  // pixel, by construction, whatever dpr actually is. Deliberately NOT tied
  // to the window/viewport size - if this doesn't fit, the element overflows
  // (the wrapper below scrolls rather than clipping it invisibly) instead of
  // ever being shrunk to fit - UNLESS the user has explicitly opted into
  // Fit to Output (see the status panel), in which case it's deliberately
  // stretched to fill the viewport's width, aspect ratio preserved via the
  // canvas's own intrinsic width/height and an explicit "Scaled to fit
  // output" message - never silently.
  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas || !layout) return;
    if (fitToOutput) {
      canvas.style.width = "100%";
      canvas.style.height = "auto";
    } else {
      canvas.style.width = `${layout.contentPixelW / dpr}px`;
      canvas.style.height = `${layout.contentPixelH / dpr}px`;
    }
  }, [layout, dpr, fitToOutput]);

  // devicePixelRatio has no native change event - watchDevicePixelRatio uses
  // the standard matchMedia re-registration idiom (see pixelMapping.ts).
  useEffect(() => watchDevicePixelRatio(setDpr), []);

  // fullscreenchange is the one source of truth for fullscreen state - every
  // control below only ever requests/exits fullscreen, never sets this
  // directly, so a browser-initiated exit (e.g. Esc) is reflected correctly.
  useEffect(() => {
    const onChange = () => setIsFullscreen(document.fullscreenElement != null);
    document.addEventListener("fullscreenchange", onChange);
    return () => document.removeEventListener("fullscreenchange", onChange);
  }, []);

  // Best-effort automatic fullscreen on mount - covers this window having
  // just been opened+positioned on a secondary display by
  // openMovingTestPatternTab in the main app. Browsers may reject this (no
  // transient user activation in a window opened programmatically) - that's
  // an expected, handled outcome, not an error: the ENTER FULLSCREEN control
  // below stays visible whenever isFullscreen is false, so the user always
  // has an explicit, unmissable way to finish it themselves. Never silently
  // fails with no indication.
  useEffect(() => {
    document.documentElement.requestFullscreen?.().catch(() => {});
  }, []);

  // Toggle the status overlay with 'h' (default visible) - useful to hide it
  // during an actual deployed test if it would sit over wall content.
  // Clicking the panel itself also hides it, and clicking anywhere else in
  // the view brings it back (see the wrapper's onClick below) - the keyboard
  // toggle and the click behaviour are independent of each other.
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key.toLowerCase() === "h") setStatusVisible((prev) => !prev);
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, []);

  // ResizeObserver on the canvas is the single most reliable "did its
  // actually-rendered box change" signal - covers window resize, fullscreen
  // enter/exit, and anything else layout-related in one listener. Measured
  // via getBoundingClientRect (never assumed from the CSS we set), feeding
  // the Browser -> Content Canvas verification below.
  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const measure = () => {
      const rect = canvas.getBoundingClientRect();
      setRenderedCssBox({ w: rect.width, h: rect.height });
    };
    measure();
    const observer = new ResizeObserver(measure);
    observer.observe(canvas);
    window.addEventListener("resize", measure);
    return () => {
      observer.disconnect();
      window.removeEventListener("resize", measure);
    };
  }, [layout]);

  // Display Resolution: prefer the Window Management API's ScreenDetailed
  // (its own devicePixelRatio, unaffected by page zoom - see
  // screenDetails.d.ts) over a window.devicePixelRatio-based fallback, per
  // explicit instruction. Also the mechanism for "detect if this window is
  // subsequently moved to another display" - currentscreenchange fires when
  // it is. Entirely best-effort: silently falls back if unsupported, denied,
  // or the context isn't secure (see screenPlacement.ts).
  useEffect(() => {
    let cancelled = false;
    let details: ScreenDetails | null = null;
    let onScreenChange: (() => void) | null = null;
    const applyFromScreen = (screen: ScreenDetailed) => {
      if (cancelled) return;
      setDisplayInfo({ w: Math.round(screen.width * screen.devicePixelRatio), h: Math.round(screen.height * screen.devicePixelRatio), exact: true });
    };
    requestScreenDetails().then((result) => {
      if (cancelled) return;
      if (!result.ok) {
        setScreenDetailsAvailable(false);
        return;
      }
      setScreenDetailsAvailable(true);
      details = result.details;
      applyFromScreen(details.currentScreen);
      onScreenChange = () => applyFromScreen(details!.currentScreen);
      details.addEventListener("currentscreenchange", onScreenChange);
    });
    return () => {
      cancelled = true;
      if (details && onScreenChange) details.removeEventListener("currentscreenchange", onScreenChange);
    };
  }, []);

  // Fallback Display Resolution path - only takes effect once we know the
  // real Window Management API isn't available, so it never overwrites the
  // exact ScreenDetailed-derived value above. Recomputed on every dpr change
  // as the best available proxy for "this window may have moved to a
  // different display" without real screen enumeration.
  useEffect(() => {
    if (screenDetailsAvailable !== false) return;
    setDisplayInfo(window.screen ? { w: Math.round(window.screen.width * dpr), h: Math.round(window.screen.height * dpr), exact: false } : null);
  }, [screenDetailsAvailable, dpr]);

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

  const statusInfo: TestPatternStatusInfo = {
    physicalW: layout.W,
    physicalH: layout.H,
    contentW: layout.contentPixelW,
    contentH: layout.contentPixelH,
    isMtContent: layout.contentPixelH !== layout.H,
    displayW: displayInfo?.w ?? null,
    displayH: displayInfo?.h ?? null,
    displayResolutionIsExact: displayInfo?.exact ?? false,
    canvasW: layout.contentPixelW,
    canvasH: layout.contentPixelH,
    devicePixelRatio: dpr,
    isFullscreen,
    browserToContentCanvas: computePixelMappingStatus(layout.contentPixelW, layout.contentPixelH, renderedCssBox.w, renderedCssBox.h, dpr),
  };

  return (
    <div
      style={{ position: "fixed", inset: 0, margin: 0, padding: 0, background: "#000", overflow: "auto" }}
      onClick={() => {
        // Clicking anywhere in the view (other than the status panel itself,
        // which stops propagation) brings the panel back if it was hidden.
        setStatusVisible(true);
        if (isFullscreen) document.exitFullscreen?.().catch(() => {});
      }}
    >
      <canvas ref={canvasRef} style={{ position: "absolute", top: 0, left: 0, display: "block", cursor: isFullscreen ? "pointer" : "default" }} />
      <TestPatternStatusOverlay
        info={statusInfo}
        visible={statusVisible}
        onHide={() => setStatusVisible(false)}
        fitToOutput={fitToOutput}
        onToggleFitToOutput={() => setFitToOutput((prev) => !prev)}
      />
      {!isFullscreen ? (
        <button
          type="button"
          onClick={(event) => {
            event.stopPropagation();
            document.documentElement.requestFullscreen?.().catch(() => {});
          }}
          style={{
            position: "fixed",
            inset: 0,
            zIndex: 50,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            background: "rgba(0, 0, 0, 0.5)",
            color: "#fff",
            fontSize: 32,
            fontWeight: 800,
            letterSpacing: 2,
            border: "none",
            cursor: "pointer",
            fontFamily: "system-ui, -apple-system, sans-serif",
          }}
        >
          ENTER FULLSCREEN
        </button>
      ) : null}
    </div>
  );
}
