import React, { useEffect, useMemo, useRef, useState } from "react";
import { Button } from "../components/ui";
import { type Cell, type PanelTypeKey, normalizePanels } from "../App";
import {
  type TestPatternLayout,
  type TestPatternProject,
  LOOP_SECONDS,
  DRAW_FPS,
  computeTestPatternLayout,
  drawTestPatternFrame,
} from "./drawTestPattern";

export const TEST_PATTERN_STORAGE_KEY = "ledCablingTestPattern:v1";

type StoredPayload = {
  formatVersion?: number;
  projectName?: string;
  panelType?: PanelTypeKey;
  panels?: unknown;
};

const fileSafe = (name: string) => name.replace(/[<>:"/\\|?*\x00-\x1F]/g, "-").replace(/\s+/g, "-");

const downloadBlob = (blob: Blob, filename: string) => {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.setAttribute("download", filename);
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
};

const pickMimeType = (): string | null => {
  if (typeof MediaRecorder === "undefined") return null;
  const candidates = ["video/webm;codecs=vp9", "video/webm;codecs=vp8", "video/webm"];
  for (const candidate of candidates) {
    if (MediaRecorder.isTypeSupported && MediaRecorder.isTypeSupported(candidate)) return candidate;
  }
  return null;
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

export default function TestPatternView() {
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const loopStartRef = useRef(performance.now());
  const recorderRef = useRef<MediaRecorder | null>(null);
  const [isRecording, setIsRecording] = useState(false);
  const [recordSeconds, setRecordSeconds] = useState(0);

  const project = useMemo(loadProject, []);
  const layout: TestPatternLayout | null = useMemo(() => (project ? computeTestPatternLayout(project) : null), [project]);

  // Animation loop, capped at DRAW_FPS. Uses setInterval rather than
  // requestAnimationFrame so it keeps running (and keeps feeding the WebM
  // recorder) if the tab is backgrounded mid-recording - browsers suspend rAF
  // in hidden tabs, but timers keep firing.
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

  // Recording progress ticker (display only - the actual stop is a setTimeout).
  useEffect(() => {
    if (!isRecording) return;
    const id = setInterval(() => setRecordSeconds((s) => Math.min(LOOP_SECONDS, s + 0.25)), 250);
    return () => clearInterval(id);
  }, [isRecording]);

  const handleDownloadVideo = () => {
    if (!layout || isRecording) return;
    const canvas = canvasRef.current;
    if (!canvas) return;
    const mimeType = pickMimeType();
    if (!mimeType) {
      alert("This browser can't record video (no WebM/MediaRecorder support). Try Chrome, Edge or Firefox.");
      return;
    }
    const stream = canvas.captureStream(DRAW_FPS);
    const recorder = new MediaRecorder(stream, { mimeType });
    const chunks: BlobPart[] = [];
    recorder.ondataavailable = (event) => {
      if (event.data.size > 0) chunks.push(event.data);
    };
    recorder.onstop = () => {
      const blob = new Blob(chunks, { type: mimeType });
      downloadBlob(blob, `${fileSafe(layout.projectName)}-front-test-pattern.webm`);
      setIsRecording(false);
      recorderRef.current = null;
    };
    // Reset the loop phase to t=0 so the recording always starts at the
    // pattern's home position - this is what makes the downloaded file loop
    // seamlessly when played back on repeat.
    loopStartRef.current = performance.now();
    setRecordSeconds(0);
    recorder.start();
    recorderRef.current = recorder;
    setIsRecording(true);
    setTimeout(() => recorder.stop(), LOOP_SECONDS * 1000);
  };

  return (
    <div style={{ minHeight: "100vh", background: "#0f172a", color: "#f8fafc", padding: 16, fontFamily: "system-ui, -apple-system, sans-serif" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12, marginBottom: 12 }}>
        <div>
          <div style={{ fontSize: 18, fontWeight: 700 }}>{layout ? layout.projectName : "Animated Test Pattern"}</div>
          <div style={{ fontSize: 12, color: "#94a3b8" }}>
            RGB Checkerboard Slide + Moving Greyscale Gradient · loops every {LOOP_SECONDS}s · Front View
          </div>
        </div>
        {layout ? (
          <div style={{ display: "flex", gap: 12, alignItems: "center" }}>
            {isRecording ? (
              <span style={{ fontSize: 13, color: "#facc15" }}>
                Recording… {recordSeconds.toFixed(0)}/{LOOP_SECONDS}s
              </span>
            ) : null}
            <Button intent="primary" onClick={handleDownloadVideo} disabled={isRecording}>
              {isRecording ? "Recording…" : "Download Video (WebM)"}
            </Button>
          </div>
        ) : null}
      </div>

      {!layout ? (
        <div style={{ padding: 32, textAlign: "center", color: "#cbd5e1" }}>
          <div style={{ fontSize: 15, marginBottom: 8 }}>No project data found for this tab.</div>
          <div style={{ fontSize: 13, color: "#94a3b8" }}>
            Open this page from the main app's <b>"Animated Test Pattern"</b> button.
          </div>
          <div style={{ marginTop: 16 }}>
            <a href={location.pathname} style={{ color: "#38bdf8" }}>
              Back to the LED Cabling Planner
            </a>
          </div>
        </div>
      ) : (
        <div style={{ overflow: "auto", border: "1px solid #334155", borderRadius: 8, background: "#000" }}>
          <canvas
            ref={canvasRef}
            width={layout.W}
            height={layout.H}
            style={{ maxWidth: "100%", height: "auto", display: "block", imageRendering: "pixelated" }}
          />
        </div>
      )}
    </div>
  );
}
