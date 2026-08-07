// Capability catalog for the two supported NovaStar processors.
//
// Port count / per-port pixel limit / total pixel capacity / max canvas are
// sourced from NovaStar's own public spec sheets (not the sample files):
//   - VX1000 Pro: https://oss.novastar.tech/uploads/2025/03/VX1000-Pro-All-in-One-Controller-Specifications-V1.0.1.pdf
//   - VX2000 Pro: https://oss.novastar.tech/uploads/2024/12/VX2000-Pro-All-in-One-Controller-Specifications-V1.0.1.pdf
//
// The `interfacePks` list for each model is the subset of interface rows
// that already exist in the embedded blank template's `t_interface_baseinfo`
// table (see novaDb.ts) - never a fabricated interface id. The embedded
// template's own input set (DP, HDMI 1-6, OPT1-1/1-2, OPT2-1/2-2, SDI) maps
// cleanly onto VX2000 Pro's published input list; VX1000 Pro is the same
// underlying template restricted to the smaller input set its spec sheet
// advertises (no DP, only 3 HDMI inputs). This is the closest we can get
// without a hardware-confirmed VX1000 Pro sample - see the UNCONFIRMED note
// below.
//
// IMPORTANT - unconfirmed identity: neither sample file contains a literal
// "VX1000"/"VX2000" string anywhere, only opaque numeric `modelId`s. The
// blank template's `modelId`/`t_subcard` triplet is known to produce a file
// NovaStar's own software will accept (it came from real, working software),
// but which *product name* that numeric id actually corresponds to has not
// been confirmed against real hardware. Both processor selections below
// currently reuse that same confirmed-valid triplet - capability differences
// are enforced entirely by this catalog's own numbers, not by the embedded
// modelId. If/when a hardware-confirmed sample for either model becomes
// available, update MODEL_IDENTITY below (search for UNCONFIRMED).

export type ProcessorModelId = "VX1000_PRO" | "VX2000_PRO";

/** The template's own numeric IDs - see the UNCONFIRMED note above. */
export const MODEL_IDENTITY = {
  /** UNCONFIRMED against real hardware - see file header. */
  chassisModelId: 25132,
  /** UNCONFIRMED against real hardware - see file header. */
  slotModelIds: [15729040, 15729034] as const,
};

export type ProcessorInput = {
  /** FK into t_interface_baseinfo.interface_pk in the embedded template DB. */
  interfacePk: number;
  /** Human-readable label, copied verbatim from the template's own row. */
  label: string;
};

// All video-capable inputs present in the embedded blank template, keyed by
// their real interface_pk (see novaDb.ts TEMPLATE_INTERFACES for the full
// dump this was drawn from). Excludes non-video rows (audio, RS-485 control,
// monitor loop-out, internal USB/mosaic sources, unnamed reserved rows).
const ALL_TEMPLATE_INPUTS: ProcessorInput[] = [
  { interfacePk: 3, label: "DP" },
  { interfacePk: 4, label: "HDMI 1" },
  { interfacePk: 5, label: "HDMI 2" },
  { interfacePk: 6, label: "HDMI 3" },
  { interfacePk: 7, label: "HDMI 4" },
  { interfacePk: 8, label: "HDMI 5" },
  { interfacePk: 9, label: "HDMI 6" },
  { interfacePk: 11, label: "OPT1-1" },
  { interfacePk: 12, label: "OPT1-2" },
  { interfacePk: 13, label: "OPT2-1" },
  { interfacePk: 14, label: "OPT2-2" },
  { interfacePk: 18, label: "SDI" },
];

function inputsByPk(pks: number[]): ProcessorInput[] {
  const byPk = new Map(ALL_TEMPLATE_INPUTS.map((i) => [i.interfacePk, i]));
  return pks.map((pk) => {
    const found = byPk.get(pk);
    if (!found) throw new Error(`processorModels: unknown template interfacePk ${pk}`);
    return found;
  });
}

export type ProcessorSpec = {
  id: ProcessorModelId;
  label: string;
  /** Number of Gigabit Ethernet LED outputs. */
  ethernetOutputCount: number;
  /** Pixel ceiling per Ethernet output (identical on both models per spec). */
  maxPixelsPerPort: number;
  /** Total pixel budget across all outputs combined. */
  maxTotalPixels: number;
  maxCanvasWidth: number;
  maxCanvasHeight: number;
  /** Physical video inputs this model exposes, restricted to real template rows. */
  inputs: ProcessorInput[];
  /**
   * Soft cap on simultaneous positioned inputs ("layers"). Derived from the
   * embedded template shipping exactly 12 pre-built spare layer rows on its
   * main screen (see novaDb.ts) - not from a published per-model spec, so
   * treat as a template limitation rather than a confirmed hardware limit.
   */
  maxInputLayers: number;
};

export const PROCESSOR_SPECS: Record<ProcessorModelId, ProcessorSpec> = {
  VX1000_PRO: {
    id: "VX1000_PRO",
    label: "NovaStar VX1000 Pro",
    ethernetOutputCount: 10,
    maxPixelsPerPort: 650_000,
    maxTotalPixels: 6_500_000,
    maxCanvasWidth: 10_240,
    maxCanvasHeight: 10_240,
    inputs: inputsByPk([4, 5, 6, 11, 12, 13, 14, 18]),
    maxInputLayers: 12,
  },
  VX2000_PRO: {
    id: "VX2000_PRO",
    label: "NovaStar VX2000 Pro",
    ethernetOutputCount: 20,
    maxPixelsPerPort: 650_000,
    maxTotalPixels: 13_100_000,
    maxCanvasWidth: 16_384,
    maxCanvasHeight: 8_192,
    inputs: inputsByPk([3, 4, 5, 6, 7, 8, 9, 11, 12, 13, 14, 18]),
    maxInputLayers: 12,
  },
};

export const PROCESSOR_MODEL_IDS: ProcessorModelId[] = ["VX1000_PRO", "VX2000_PRO"];
