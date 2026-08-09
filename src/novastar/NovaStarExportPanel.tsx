import React from "react";
import { Download } from "lucide-react";
import { Button, Card, CardContent, CardHeader, CardTitle } from "../components/ui";
import type { ExportSummary } from "./exportBuilder";

type Props = {
  hasProcessorModel: boolean;
  summary: ExportSummary | null;
  errors: string[];
  warnings: string[];
  onDownload: () => void;
  downloading: boolean;
};

export default function NovaStarExportPanel({ hasProcessorModel, summary, errors, warnings, onDownload, downloading }: Props) {
  const canDownload = hasProcessorModel && summary !== null && errors.length === 0;

  return (
    <Card className="border-slate-700 bg-slate-800 print-card no-print" collapsible defaultOpen={false}>
      <CardHeader>
        <CardTitle className="text-white [text-shadow:0_0_2px_black]">NovaStar Processor Configuration</CardTitle>
      </CardHeader>
      <CardContent className="space-y-3 text-sm text-white [text-shadow:0_0_2px_black]">
        {!hasProcessorModel ? (
          <div className="rounded-lg border border-slate-600 bg-slate-900/50 p-3 text-xs text-slate-300">
            Select a Processor Model in LED Wall Setup to configure and export a NovaStar file.
          </div>
        ) : null}

        {summary ? (
          <div className="grid gap-x-4 gap-y-1 rounded-lg border border-slate-700 bg-slate-900/50 p-3 text-xs md:grid-cols-2">
            <div>
              <span className="text-slate-400">Processor: </span>
              {summary.processorLabel}
            </div>
            <div>
              <span className="text-slate-400">Surface: </span>
              {summary.surfaceName || summary.projectName}
            </div>
            <div>
              <span className="text-slate-400">Output canvas: </span>
              {summary.outputCanvas.w}×{summary.outputCanvas.h}px
            </div>
            <div>
              <span className="text-slate-400">Total screen resolution: </span>
              {summary.totalScreenResolution.w}×{summary.totalScreenResolution.h}px
            </div>
            <div>
              <span className="text-slate-400">Panels: </span>
              {summary.panelCount}
            </div>
            <div>
              <span className="text-slate-400">Sub-screens: </span>
              {summary.subScreenCount}
            </div>
            <div className="md:col-span-2">
              <span className="text-slate-400">Ethernet outputs used: </span>
              {summary.ethernetOutputsUsed}
            </div>
            {summary.outputPixelLoads.length ? (
              <div className="md:col-span-2">
                <span className="text-slate-400">Pixel load per output: </span>
                {summary.outputPixelLoads.map((o) => `Port ${o.signalPort}: ${o.pixels.toLocaleString()}px`).join(" · ")}
              </div>
            ) : null}
            {summary.inputAssignments.length ? (
              <div className="md:col-span-2">
                <span className="text-slate-400">Input assignments: </span>
                {summary.inputAssignments.map((a) => `${a.name} → ${a.inputLabel ?? "unassigned"}`).join(" · ")}
              </div>
            ) : null}
          </div>
        ) : null}

        {errors.length ? (
          <div className="rounded-lg border border-rose-500 bg-rose-500/15 p-2 text-xs text-rose-200">
            {errors.map((e, i) => (
              <div key={i}>✕ {e}</div>
            ))}
          </div>
        ) : null}
        {warnings.length ? (
          <div className="rounded-lg border border-amber-400 bg-amber-500/15 p-2 text-xs text-amber-200">
            {warnings.map((w, i) => (
              <div key={i}>⚠ {w}</div>
            ))}
          </div>
        ) : null}

        <Button intent="primary" onClick={onDownload} disabled={!canDownload || downloading}>
          <Download className="h-4 w-4" />
          {downloading ? "Generating…" : "Download Processor Configuration"}
        </Button>
      </CardContent>
    </Card>
  );
}
