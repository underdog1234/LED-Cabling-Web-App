// ---------------------------------------------------------------------------
// Import mapper for the sibling "YES TECH Layout Tool" (MG9 Creative Designer).
//
// Kept deliberately separate from the app's own project format so the external
// schema can evolve independently and additional formats can be added later.
// The mapper converts the external file into neutral panel descriptors; the app
// assigns ids and drops them into its panel list.
//
// External schema (formatVersion "1.x"):
//   { version, projectName, inventory, panels:[{type,x,y,rotation}], textLayout, ui }
//   - Coordinates are in the tool's units where 1 unit = 4mm (MM_TO_UNITS 0.25).
//   - A panel is 500x500mm (125 units); (x,y) is the panel CENTRE.
//   - type: MG9 (square) | MG12 (right triangle) | MG13 (quarter circle).
//   - rotation: 0/90/180/270, SVG-clockwise. No patching is stored.
// ---------------------------------------------------------------------------

export const YESTECH_UNIT_MM = 4; // 1 external unit = 4mm (MM_TO_UNITS 0.25)
export const YESTECH_PANEL_MM = 500;

const SUPPORTED_VERSIONS = ["1.0", "1.1", "1.2"];

// External type -> this app's (panelType, panelVariant).
const TYPE_MAP: Record<string, { panelType: "MG9"; panelVariant: "STANDARD" | "TRIANGLE" | "CURVED" }> = {
  MG9: { panelType: "MG9", panelVariant: "STANDARD" },
  MG12: { panelType: "MG9", panelVariant: "TRIANGLE" },
  MG13: { panelType: "MG9", panelVariant: "CURVED" },
};

export type ImportedPanel = {
  panelType: "MG9";
  panelVariant: "STANDARD" | "TRIANGLE" | "CURVED";
  /** Top-left corner, workspace millimetres. */
  x: number;
  y: number;
  rotation: number;
  _extra?: Record<string, unknown>;
};

export type ImportResult = {
  ok: boolean;
  /** Fatal error - nothing usable was parsed. */
  error?: string;
  projectName: string;
  panels: ImportedPanel[];
  /** Non-fatal notes surfaced to the user before applying the import. */
  warnings: string[];
  /** Panels that could not be imported, with a reason. */
  skipped: string[];
  /** Values that were changed to fit this app. */
  converted: string[];
  /** Preview summary. */
  summary: {
    sourceVersion: string | null;
    panelCount: number;
    typeCounts: Record<string, number>;
    widthM: number;
    heightM: number;
  };
  /** Top-level fields we don't model, kept so a later save doesn't lose data. */
  extra: Record<string, unknown>;
};

const normalizeRotation = (rotation: unknown): { value: number; wasInvalid: boolean } => {
  const raw = Number(rotation);
  if (!Number.isFinite(raw)) return { value: 0, wasInvalid: true };
  const snapped = ((Math.round(raw / 90) * 90) % 360 + 360) % 360;
  return { value: snapped, wasInvalid: snapped !== raw };
};

/**
 * Parse and map a YES TECH layout export. Never throws - a bad file returns
 * `{ ok:false, error }`. Import is intentionally un-patched: only type,
 * position, rotation and variant are carried across.
 */
export const parseYesTechLayout = (text: string): ImportResult => {
  const base: ImportResult = {
    ok: false,
    projectName: "Imported Layout",
    panels: [],
    warnings: [],
    skipped: [],
    converted: [],
    summary: { sourceVersion: null, panelCount: 0, typeCounts: {}, widthM: 0, heightM: 0 },
    extra: {},
  };

  let data: Record<string, unknown>;
  try {
    data = JSON.parse(text) as Record<string, unknown>;
  } catch {
    return { ...base, error: "The file is not valid JSON." };
  }
  if (!data || typeof data !== "object") {
    return { ...base, error: "The file does not contain a layout object." };
  }

  const sourceVersion = typeof data.version === "string" ? data.version : null;
  base.summary.sourceVersion = sourceVersion;
  if (sourceVersion && !SUPPORTED_VERSIONS.some((v) => sourceVersion.startsWith(v))) {
    base.warnings.push(
      `File was created by layout tool version ${sourceVersion}, which is newer than supported (${SUPPORTED_VERSIONS.join(", ")}). Importing what is compatible.`,
    );
  }

  const rawPanels = data.panels;
  if (!Array.isArray(rawPanels)) {
    return { ...base, error: "No panel list found (expected a top-level \"panels\" array)." };
  }

  const projectName = typeof data.projectName === "string" && data.projectName.trim() ? data.projectName.trim() : "Imported Layout";

  const knownTop = new Set(["version", "projectName", "panels", "inventory"]);
  const extra: Record<string, unknown> = {};
  Object.keys(data).forEach((key) => {
    if (!knownTop.has(key)) extra[key] = data[key];
  });
  if (Object.keys(extra).length) {
    base.warnings.push(`Kept ${Object.keys(extra).length} extra project field(s) (${Object.keys(extra).join(", ")}) so nothing is lost on save.`);
  }

  const panels: ImportedPanel[] = [];
  const typeCounts: Record<string, number> = {};
  let rotationFixes = 0;

  rawPanels.forEach((item, index) => {
    const raw = item as Record<string, unknown>;
    const type = String(raw?.type ?? "");
    const mapping = TYPE_MAP[type];
    if (!mapping) {
      base.skipped.push(`Panel ${index + 1}: unsupported type "${type || "(missing)"}".`);
      return;
    }
    const cx = Number(raw?.x);
    const cy = Number(raw?.y);
    if (!Number.isFinite(cx) || !Number.isFinite(cy)) {
      base.skipped.push(`Panel ${index + 1} (${type}): missing or invalid position.`);
      return;
    }
    const { value: rotation, wasInvalid } = normalizeRotation(raw?.rotation);
    if (wasInvalid) rotationFixes += 1;

    // Centre (units) -> top-left (mm).
    const centreXmm = cx * YESTECH_UNIT_MM;
    const centreYmm = cy * YESTECH_UNIT_MM;
    const xMm = centreXmm - YESTECH_PANEL_MM / 2;
    const yMm = centreYmm - YESTECH_PANEL_MM / 2;

    const known = new Set(["type", "x", "y", "rotation"]);
    const panelExtra: Record<string, unknown> = {};
    Object.keys(raw ?? {}).forEach((key) => {
      if (!known.has(key)) panelExtra[key] = raw[key];
    });

    panels.push({
      panelType: mapping.panelType,
      panelVariant: mapping.panelVariant,
      x: xMm,
      y: yMm,
      rotation,
      ...(Object.keys(panelExtra).length ? { _extra: panelExtra } : {}),
    });
    typeCounts[type] = (typeCounts[type] ?? 0) + 1;
  });

  if (!panels.length) {
    return { ...base, error: "No importable panels were found in the file.", skipped: base.skipped };
  }

  if (rotationFixes) base.converted.push(`Snapped ${rotationFixes} rotation value(s) to the nearest 90°.`);
  base.converted.push("Imported un-patched: signal and power patching are left for you to assign.");
  base.converted.push("Converted grid positions to the flexible mm workspace (panels are freely movable).");

  // Bounding box for the preview (mm -> m).
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  panels.forEach((p) => {
    minX = Math.min(minX, p.x);
    minY = Math.min(minY, p.y);
    maxX = Math.max(maxX, p.x + YESTECH_PANEL_MM);
    maxY = Math.max(maxY, p.y + YESTECH_PANEL_MM);
  });

  return {
    ...base,
    ok: true,
    projectName,
    panels,
    extra,
    summary: {
      sourceVersion,
      panelCount: panels.length,
      typeCounts,
      widthM: (maxX - minX) / 1000,
      heightM: (maxY - minY) / 1000,
    },
  };
};
