// Orchestrates turning the LED Cabling Tool's own layout state into a real,
// single-processor NovaStar .uprj file: validate against the selected
// processor's capacity, populate the embedded blank template's database,
// re-zip, and re-wrap in the envelope format from uprjFormat.ts.
//
// Mirrors the existing ImportResult shape from src/import/yesTechLayout.ts
// (ok/errors/warnings/summary) but in the export direction.

import { unzipSync, zipSync, type Zippable } from "fflate";
import { type Cell, type SubScreen, cellRect } from "../App";
import { resolutionOf, wallNativePositionsOf } from "../canvasView/canvasModel";
import { subScreenBBoxOf } from "../subScreens/subScreenModel";
import { activeBBox, type RectMm } from "../model/panels";
import { PROCESSOR_SPECS, type ProcessorModelId, type ProcessorInput } from "./processorModels";
import { decodeEnvelope, encodeSingleDeviceEnvelope, type ProjectMeta } from "./uprjFormat";
import { NovaDb, type CabinetInput, type InputLayerAssignment } from "./novaDb";
import { loadBlankTemplateBytes } from "./templateAsset";

/** Sentinel key for the "no sub-screens defined yet" whole-layout entry - matches OutputCanvasPanel's own convention. */
export const WHOLE_LAYOUT_KEY = "__whole__";

export type CanvasEntryInput = {
  /** subScreen.id, or WHOLE_LAYOUT_KEY when no sub-screens exist. */
  key: string;
  name: string;
  /** FK into the selected processor's input list, or null = unassigned. */
  interfacePk: number | null;
};

/**
 * "perEntry" assigns a separate input per sub-screen (or a single whole-
 * layout entry when no sub-screens exist) - the original behavior. "whole"
 * assigns a single input to the entire output canvas regardless of
 * sub-screen boundaries, for the common case of one video source feeding
 * the whole wall.
 */
export type InputMode = "perEntry" | "whole";

export type ExportBuilderInput = {
  processorModel: ProcessorModelId;
  projectName: string;
  surfaceName: string;
  outputCanvasW: number;
  outputCanvasH: number;
  wholeLayoutCanvasX: number;
  wholeLayoutCanvasY: number;
  grid: Cell[];
  subScreens: SubScreen[];
  inputMode: InputMode;
  /** Used when inputMode === "whole". FK into the selected processor's input list, or null = unassigned. */
  wholeCanvasInputId: number | null;
  /** Used when inputMode === "perEntry": one entry per canvas entry (sub-screens, or a single WHOLE_LAYOUT_KEY entry). */
  canvasInputs: CanvasEntryInput[];
};

export type OutputPixelLoad = { signalPort: number; pixels: number };

export type ExportSummary = {
  processorLabel: string;
  surfaceName: string;
  projectName: string;
  outputCanvas: { w: number; h: number };
  totalScreenResolution: { w: number; h: number };
  panelCount: number;
  subScreenCount: number;
  ethernetOutputsUsed: number;
  outputPixelLoads: OutputPixelLoad[];
  inputAssignments: { name: string; inputLabel: string | null }[];
};

export type ExportResult = {
  ok: boolean;
  errors: string[];
  warnings: string[];
  summary: ExportSummary;
  blob?: Blob;
  fileName?: string;
};

function fileSafe(name: string): string {
  return (name.trim() || "Untitled").replace(/[<>:"/\\|?*\x00-\x1F]/g, "-").replace(/\s+/g, "-");
}

function normalizeRotation(rotation: number | undefined): 0 | 90 | 180 | 270 {
  const r = ((Math.round((rotation ?? 0) / 90) * 90) % 360 + 360) % 360;
  return r as 0 | 90 | 180 | 270;
}

type CanvasEntry = {
  key: string;
  name: string;
  canvasX: number;
  canvasY: number;
  bboxMm: RectMm;
  cells: Cell[];
};

/** Same "whole layout vs. per-sub-screen" branching OutputCanvasPanel.tsx uses for its entries list. */
function buildCanvasEntries(input: ExportBuilderInput): CanvasEntry[] {
  const activeCells = input.grid.filter((c) => !c.isRemoved);
  if (input.subScreens.length) {
    return input.subScreens.map((screen) => ({
      key: screen.id,
      name: screen.name,
      canvasX: screen.canvasX,
      canvasY: screen.canvasY,
      bboxMm: subScreenBBoxOf(input.grid, screen.id, cellRect),
      cells: activeCells.filter((c) => c.subScreenId === screen.id),
    }));
  }
  return [
    {
      key: WHOLE_LAYOUT_KEY,
      name: "Whole Layout",
      canvasX: input.wholeLayoutCanvasX,
      canvasY: input.wholeLayoutCanvasY,
      bboxMm: activeBBox(activeCells.map(cellRect)),
      cells: activeCells,
    },
  ];
}

type ResolvedInputAssignment = {
  name: string;
  interfacePk: number | null;
  x: number;
  y: number;
  width: number;
  height: number;
};

/** Single source of truth for both modes, shared by validation and the actual DB write. */
function resolveInputAssignments(input: ExportBuilderInput, entries: CanvasEntry[]): ResolvedInputAssignment[] {
  if (input.inputMode === "whole") {
    return [
      {
        name: "Whole Canvas",
        interfacePk: input.wholeCanvasInputId,
        x: 0,
        y: 0,
        width: input.outputCanvasW,
        height: input.outputCanvasH,
      },
    ];
  }
  return entries.map((entry) => {
    const assignment = input.canvasInputs.find((a) => a.key === entry.key);
    const resolution = resolutionOf(entry.cells);
    return {
      name: entry.name,
      interfacePk: assignment?.interfacePk ?? null,
      x: entry.canvasX,
      y: entry.canvasY,
      width: resolution.w,
      height: resolution.h,
    };
  });
}

export function buildExportSummaryAndCabinets(input: ExportBuilderInput): {
  cabinets: CabinetInput[];
  summary: ExportSummary;
  errors: string[];
  warnings: string[];
} {
  const spec = PROCESSOR_SPECS[input.processorModel];
  const errors: string[] = [];
  const warnings: string[] = [];

  const entries = buildCanvasEntries(input);
  const entryByKey = new Map(entries.map((e) => [e.key, e]));

  const unassignedCount = input.subScreens.length
    ? input.grid.filter((c) => !c.isRemoved && !entryByKey.has(c.subScreenId ?? "")).length
    : 0;
  if (unassignedCount > 0) {
    warnings.push(`${unassignedCount} panel(s) are not assigned to a sub-screen and will be excluded from the export.`);
  }

  // Cabinet positions live in the LED screen's own native pixel grid - NOT
  // the output-canvas/sub-screen placement space finalCanvasPositionOf
  // computes. Confirmed against real NovaStar-generated project files: the
  // processor's cabinet-topology canvas is sized and positioned purely from
  // real (patched) panel pixels, tightly packed with no gaps, entirely
  // independent of where a sub-screen sits on the output canvas. See
  // wallNativePositionsOf's own comment for the full explanation.
  const wallPositions = wallNativePositionsOf(input.grid);

  const cabinets: CabinetInput[] = [];
  let missingPatchCount = 0;
  for (const entry of entries) {
    for (const cell of entry.cells) {
      if (cell.assignedPort == null || cell.sequence == null) {
        missingPatchCount += 1;
        continue;
      }
      const pos = wallPositions.get(cell.id);
      if (!pos) continue;
      cabinets.push({
        netPortIndex: cell.assignedPort - 1,
        cabinetIndex: cell.sequence - 1,
        x: pos.x,
        y: pos.y,
        width: pos.w,
        height: pos.h,
        angle: normalizeRotation(cell.rotation),
      });
    }
  }
  if (missingPatchCount > 0) {
    warnings.push(`${missingPatchCount} panel(s) have no assigned signal port and will be excluded from the export.`);
  }
  if (cabinets.length === 0) {
    errors.push("No panels are patched to a signal port - there is nothing to export.");
  }

  const usedPorts = new Set(cabinets.map((c) => c.netPortIndex + 1));
  for (const port of usedPorts) {
    if (port > spec.ethernetOutputCount) {
      errors.push(`Signal port ${port} exceeds ${spec.label}'s ${spec.ethernetOutputCount} available Ethernet outputs.`);
    }
  }

  const pixelsByPort = new Map<number, number>();
  for (const c of cabinets) {
    const port = c.netPortIndex + 1;
    pixelsByPort.set(port, (pixelsByPort.get(port) ?? 0) + c.width * c.height);
  }
  const outputPixelLoads: OutputPixelLoad[] = Array.from(pixelsByPort.entries())
    .sort((a, b) => a[0] - b[0])
    .map(([signalPort, pixels]) => ({ signalPort, pixels }));
  for (const { signalPort, pixels } of outputPixelLoads) {
    if (pixels > spec.maxPixelsPerPort) {
      errors.push(
        `Signal port ${signalPort} carries ${pixels.toLocaleString()} px, exceeding ${spec.label}'s ${spec.maxPixelsPerPort.toLocaleString()} px/port limit.`,
      );
    }
  }

  const totalPixels = cabinets.reduce((sum, c) => sum + c.width * c.height, 0);
  if (totalPixels > spec.maxTotalPixels) {
    errors.push(
      `Total pixel count (${totalPixels.toLocaleString()}) exceeds ${spec.label}'s ${spec.maxTotalPixels.toLocaleString()} px capacity.`,
    );
  }

  if (input.outputCanvasW > spec.maxCanvasWidth || input.outputCanvasH > spec.maxCanvasHeight) {
    errors.push(
      `Output canvas ${input.outputCanvasW}×${input.outputCanvasH} exceeds ${spec.label}'s ${spec.maxCanvasWidth}×${spec.maxCanvasHeight} maximum.`,
    );
  }

  const inputByPk = new Map<number, ProcessorInput>(spec.inputs.map((i) => [i.interfacePk, i]));
  const resolvedInputs = resolveInputAssignments(input, entries);
  const assignedInputs = resolvedInputs.filter((a) => a.interfacePk != null);
  if (assignedInputs.length > spec.maxInputLayers) {
    errors.push(`${assignedInputs.length} inputs assigned exceeds the ${spec.maxInputLayers}-input limit.`);
  }
  for (const assignment of assignedInputs) {
    if (!inputByPk.has(assignment.interfacePk as number)) {
      errors.push(`"${assignment.name}" is assigned an input that is not available on ${spec.label}.`);
    }
  }
  const entriesWithoutInput =
    input.inputMode === "whole"
      ? resolvedInputs.filter((a) => a.interfacePk == null && cabinets.length > 0)
      : resolvedInputs.filter((a, i) => entries[i].cells.length > 0 && a.interfacePk == null);
  if (entriesWithoutInput.length > 0) {
    const subject = input.inputMode === "whole" ? "The whole canvas has" : `${entriesWithoutInput.length} canvas entr${entriesWithoutInput.length === 1 ? "y has" : "ies have"}`;
    warnings.push(`${subject} no input assigned.`);
  }

  const summary: ExportSummary = {
    processorLabel: spec.label,
    surfaceName: input.surfaceName,
    projectName: input.projectName,
    outputCanvas: { w: input.outputCanvasW, h: input.outputCanvasH },
    // Pixel resolution (not the mm bounding box - that's bboxMm above,
    // used only for cabinet positioning), matching what Wall Details
    // already shows elsewhere in the app.
    totalScreenResolution: resolutionOf(input.grid.filter((c) => !c.isRemoved)),
    panelCount: input.grid.filter((c) => !c.isRemoved).length,
    subScreenCount: input.subScreens.length,
    ethernetOutputsUsed: usedPorts.size,
    outputPixelLoads,
    inputAssignments: resolvedInputs.map((a) => ({
      name: a.name,
      inputLabel: a.interfacePk != null ? (inputByPk.get(a.interfacePk)?.label ?? null) : null,
    })),
  };

  return { cabinets, summary, errors, warnings };
}

export async function buildNovaStarExport(input: ExportBuilderInput): Promise<ExportResult> {
  const { cabinets, summary, errors, warnings } = buildExportSummaryAndCabinets(input);
  if (errors.length > 0) {
    return { ok: false, errors, warnings, summary };
  }

  const templateBytes = await loadBlankTemplateBytes();
  const template = decodeEnvelope(templateBytes);
  const zipFiles = unzipSync(template.deviceZips[0]);

  const dbEntryKey = Object.keys(zipFiles).find((k) => k.endsWith("Userver.db"));
  if (!dbEntryKey) throw new Error("exportBuilder: embedded template zip has no Userver.db entry");
  const nodeFolder = dbEntryKey.slice(0, dbEntryKey.length - "Userver.db".length);

  const db = await NovaDb.open(zipFiles[dbEntryKey]);
  db.assertTemplateShape();
  // t_canvas / t_screen_splice_load are the LED screen's own native pixel
  // grid (matches summary.totalScreenResolution - see wallNativePositionsOf
  // above), NOT the output canvas: confirmed against a real NovaStar file
  // where these were 1680x840 (the wall's own resolution) while
  // t_logic_screen_general stayed at an unrelated, independently-set value.
  db.setCanvasResolution(summary.totalScreenResolution.w, summary.totalScreenResolution.h);
  db.setScreenSpliceLoad(summary.totalScreenResolution.w, summary.totalScreenResolution.h);
  db.setMainLogicScreen(input.outputCanvasW, input.outputCanvasH, input.surfaceName || input.projectName);
  db.replaceCabinetTopology(cabinets);

  const spec = PROCESSOR_SPECS[input.processorModel];
  const inputByPk = new Map(spec.inputs.map((i) => [i.interfacePk, i]));
  const layerAssignments: InputLayerAssignment[] = resolveInputAssignments(input, buildCanvasEntries(input))
    .filter((a): a is ResolvedInputAssignment & { interfacePk: number } => a.interfacePk != null && inputByPk.has(a.interfacePk))
    .map((a) => ({ name: a.name, interfacePk: a.interfacePk, x: a.x, y: a.y, width: a.width, height: a.height }));
  db.replaceInputLayers(layerAssignments);
  db.setDeviceName(input.surfaceName || input.projectName);

  const projectId = `led_cabling_tool-${fileSafe(input.projectName)}-${Date.now()}`;
  db.setProjectIdentity(projectId, input.projectName);
  const subcardManifest = db.readSubcardManifest();

  const newDbBytes = db.toBytes();
  db.close();

  const newZipFiles: Zippable = { ...zipFiles };
  newZipFiles[dbEntryKey] = newDbBytes;
  const bakKey = `${nodeFolder}Userver_Bak.db`;
  if (bakKey in newZipFiles) newZipFiles[bakKey] = newDbBytes;
  const subcardKey = `${nodeFolder}subcard.json`;
  if (subcardKey in newZipFiles) {
    newZipFiles[subcardKey] = new TextEncoder().encode(
      JSON.stringify(
        subcardManifest.map((s) => ({
          slotId: s.slotId,
          modelId: s.modelId,
          boardId: s.boardId,
          cardType: s.cardType,
          specialFunc: s.specialFunc,
        })),
      ),
    );
  }
  const newZipBytes = zipSync(newZipFiles);

  const templateNode = template.meta.nodeList[0];
  const meta: ProjectMeta = {
    projectId,
    projectName: input.projectName,
    nodeList: [
      {
        ...templateNode,
        // sn/deviceName are the only identity fields this tool changes -
        // modelId/ip/nodeVersion are preserved from the template verbatim
        // (see processorModels.ts's UNCONFIRMED note on modelId).
        sn: `virtual${Date.now()}`,
        deviceName: input.surfaceName || input.projectName,
      },
    ],
  };
  const fileBytes = encodeSingleDeviceEnvelope(meta, newZipBytes);
  // Cast needed because TS's DOM lib types Blob's BlobPart as
  // ArrayBufferView<ArrayBuffer> specifically, while our byte-manipulation
  // helpers return the more general Uint8Array<ArrayBufferLike> - Blob
  // itself accepts any Uint8Array at runtime.
  const blob = new Blob([fileBytes as unknown as BlobPart], { type: "application/octet-stream" });
  const fileName = `${fileSafe(input.projectName)}_${fileSafe(input.surfaceName || input.projectName)}_${spec.id === "VX1000_PRO" ? "VX1000Pro" : "VX2000Pro"}_Config.uprj`;

  return { ok: true, errors, warnings, summary, blob, fileName };
}
