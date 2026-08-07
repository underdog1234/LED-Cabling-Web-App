import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { parseUprj } from "./parseUprj";

function fixture(name: string): Uint8Array {
  const path = fileURLToPath(new URL(`./__fixtures__/${name}`, import.meta.url));
  return new Uint8Array(readFileSync(path));
}

describe("parseUprj golden comparison against the real 'with cables' sample", () => {
  it("reads back Device1's cabinet topology matching the values found during reverse engineering", async () => {
    // with-cables.uprj's first device (Device1) is a real, fully-cabled
    // project: 3192x1008 canvas, 33 cabinets across 3 Ethernet outputs, each
    // cabinet 192x192px. These exact numbers were confirmed directly against
    // the live SQLite database while reverse engineering the format (see the
    // plan) - re-deriving them here through parseUprj guards against
    // regressions in the read path independently of anything this tool
        // itself ever generates.
    const parsed = await parseUprj(fixture("with-cables.uprj"));

    expect(parsed.checksumValid).toBe(true);
    expect(parsed.canvas).toEqual({ width: 3192, height: 1008 });
    expect(parsed.cabinets).toHaveLength(33);
    expect(parsed.cabinets.every((c) => c.width === 192 && c.height === 192)).toBe(true);

    const byPort = new Map<number, typeof parsed.cabinets>();
    for (const c of parsed.cabinets) byPort.set(c.netPortIndex, [...(byPort.get(c.netPortIndex) ?? []), c]);
    expect([...byPort.keys()].sort()).toEqual([0, 1, 2]);
    expect(byPort.get(0)).toHaveLength(18);
    expect(byPort.get(1)).toHaveLength(7);
    expect(byPort.get(2)).toHaveLength(8);

    // Starting cabinet (cabinet_index 0) of the first output, confirmed at x=168,y=840.
    const startOfPort0 = byPort.get(0)!.find((c) => c.cabinetIndex === 0)!;
    expect(startOfPort0).toMatchObject({ x: 168, y: 840 });

    expect(parsed.subcard).toEqual([
      { slotId: 100, modelId: 25132, boardId: 0, cardType: 7, specialFunc: 0 },
      { slotId: 0, modelId: 15729040, boardId: 0, cardType: 2, specialFunc: 0 },
      { slotId: 1, modelId: 15729034, boardId: 0, cardType: 3, specialFunc: 0 },
    ]);
  });
});
