// ---------------------------------------------------------------------------
// Opening an LED Cabling Planner project in the generator.
//
// The planner's Output Canvas becomes this generator's canvas: same canvas
// resolution, one sub-screen per planner sub-screen at the same X/Y and the
// same pixel footprint the planner quotes, each running the planner's own
// Moving Test Pattern for exactly those panels. Processor input assignments
// come across too. A project with no sub-screens comes in as one "Whole
// Layout" screen, as the planner shows it.
// ---------------------------------------------------------------------------

import { clampDimension, makeScreen, type GeneratorConfig, type LedLayoutSource, type SubScreen } from "./model";
import { LED_LAYOUT_PATTERN_ID } from "./patterns";
import { ledLayoutFor, loadLedModule } from "./render";

type PlannerFile = {
  projectName?: string;
  surfaceName?: string;
  panelType?: string;
  panels?: unknown[];
  subScreens?: Array<{ id?: string; name?: string; canvasX?: number; canvasY?: number; color?: string }>;
  outputCanvas?: { w?: number; h?: number };
  wholeLayoutCanvasPos?: { x?: number; y?: number };
  processorModel?: string;
  canvasInputs?: Record<string, number | null>;
  inputMode?: string;
  wholeCanvasInputId?: number | null;
};

export const looksLikePlannerProject = (data: unknown): boolean => {
  const d = data as PlannerFile | null;
  return !!d && typeof d === "object" && (Array.isArray(d.panels) || "wall" in (d as object)) && !("app" in (d as object));
};

export const importPlannerProject = async (
  data: PlannerFile,
  base: GeneratorConfig,
): Promise<GeneratorConfig> => {
  if (!Array.isArray(data.panels) || !data.panels.length) {
    throw new Error("This planner file has no panel list. Open it in the LED Cabling Planner and save it again, then import the new file.");
  }
  await loadLedModule();
  const panelType = data.panelType === "MT" ? "MT" : "MG9";
  const projectName = (data.projectName || "").trim();
  const subScreens = Array.isArray(data.subScreens) ? data.subScreens.filter((s) => typeof s?.id === "string") : [];
  const raw = data.panels as Array<Record<string, unknown>>;
  const isActive = (p: Record<string, unknown>) => !p?.isRemoved;
  const screens: SubScreen[] = [];
  const inputs = data.canvasInputs && typeof data.canvasInputs === "object" ? data.canvasInputs : {};

  const build = (index: number, name: string, color: string | undefined, panels: unknown[], sub: LedLayoutSource["subScreen"], x: number, y: number, inputKey: string) => {
    const source: LedLayoutSource = { projectName, surfaceName: name, panelType, panels, subScreen: sub };
    const layout = ledLayoutFor(source);
    if (!layout) return;
    const screen = makeScreen(index, {
      name,
      x: Math.round(Number(x) || 0),
      y: Math.round(Number(y) || 0),
      w: clampDimension(layout.W),
      h: clampDimension(layout.H),
      pattern: LED_LAYOUT_PATTERN_ID,
      ledLayout: source,
      input: typeof inputs[inputKey] === "number" ? (inputs[inputKey] as number) : null,
      aspectLock: true,
    });
    if (color && /^#[0-9a-fA-F]{6}$/.test(color)) screen.color = color.toLowerCase();
    // A planner sub-screen already names itself on its pattern; the label overlay would say it twice.
    screen.overlays = { border: false, label: false, resolution: false, crosshair: false, clock: false };
    screens.push(screen);
  };

  if (subScreens.length) {
    subScreens.forEach((sub, i) => {
      const panels = raw.filter((p) => p?.subScreenId === sub.id && isActive(p));
      if (!panels.length) return;
      build(i, sub.name || `Sub-screen ${i + 1}`, sub.color, panels, { id: sub.id!, name: sub.name || "", color: sub.color || "#ffffff" }, sub.canvasX ?? 0, sub.canvasY ?? 0, sub.id!);
    });
  } else {
    build(0, (data.surfaceName || "").trim() || "Whole Layout", undefined, raw.filter(isActive), null, data.wholeLayoutCanvasPos?.x ?? 0, data.wholeLayoutCanvasPos?.y ?? 0, "__whole__");
  }
  if (!screens.length) throw new Error("No active panels were found in this planner project.");

  return {
    ...base,
    name: projectName || base.name,
    canvas: {
      ...base.canvas,
      w: clampDimension(Number(data.outputCanvas?.w) || 1920),
      h: clampDimension(Number(data.outputCanvas?.h) || 1080),
    },
    screens,
    playlist: [],
    processor: {
      model: data.processorModel === "VX1000_PRO" || data.processorModel === "VX2000_PRO" ? data.processorModel : "",
      inputMode: data.inputMode === "whole" ? "whole" : "perEntry",
      wholeInput: typeof data.wholeCanvasInputId === "number" ? data.wholeCanvasInputId : null,
    },
  };
};
