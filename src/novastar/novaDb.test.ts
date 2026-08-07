import { describe, it, expect, beforeEach } from "vitest";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { decodeEnvelope } from "./uprjFormat";
import { unzipSync } from "fflate";
import { NovaDb } from "./novaDb";

function fixture(name: string): Uint8Array {
  const path = fileURLToPath(new URL(`./__fixtures__/${name}`, import.meta.url));
  return new Uint8Array(readFileSync(path));
}

function extractUserverDb(zipBytes: Uint8Array): Uint8Array {
  const files = unzipSync(zipBytes);
  const key = Object.keys(files).find((k) => k.endsWith("Userver.db"));
  if (!key) throw new Error("Userver.db not found in zip");
  return files[key];
}

describe("NovaDb against the real blank template", () => {
  let dbBytes: Uint8Array;

  beforeEach(() => {
    const env = decodeEnvelope(fixture("blank.uprj"));
    dbBytes = extractUserverDb(env.deviceZips[0]);
  });

  it("opens the blank template and passes its own shape assertion", async () => {
    const db = await NovaDb.open(dbBytes);
    expect(() => db.assertTemplateShape()).not.toThrow();
    db.close();
  });

  it("round-trips canvas resolution", async () => {
    const db = await NovaDb.open(dbBytes);
    db.setCanvasResolution(3840, 1152);
    expect(db.readCanvasResolution()).toEqual({ width: 3840, height: 1152 });
    db.close();
  });

  it("round-trips the main logic screen resolution and name", async () => {
    const db = await NovaDb.open(dbBytes);
    db.setMainLogicScreen(1920, 1080, "Front Wall");
    expect(db.readMainLogicScreen()).toEqual({ width: 1920, height: 1080, name: "Front Wall" });
    db.close();
  });

  it("inserts cabinet topology rows, grouping consistently by the input net_port_index", async () => {
    const db = await NovaDb.open(dbBytes);
    db.replaceCabinetTopology([
      { netPortIndex: 0, cabinetIndex: 0, x: 0, y: 0, width: 168, height: 168, angle: 0 },
      { netPortIndex: 0, cabinetIndex: 1, x: 168, y: 0, width: 168, height: 168, angle: 0 },
      { netPortIndex: 1, cabinetIndex: 0, x: 336, y: 0, width: 168, height: 168, angle: 90 },
    ]);
    const rows = db.readCabinetTopology();
    expect(rows).toHaveLength(3);
    // net_port_index is allocated fresh per export (not a stable "port 0/1"
    // label - see replaceCabinetTopology's comment), so assert grouping and
    // field values rather than exact output port-index numbers.
    const port0Rows = rows.filter((r) => r.cabinetIndex === 0 && r.x === 0);
    const port1Rows = rows.filter((r) => r.angle === 90);
    expect(port0Rows).toHaveLength(1);
    expect(port1Rows).toHaveLength(1);
    expect(port0Rows[0].netPortIndex).not.toBe(port1Rows[0].netPortIndex);
    const sameGroupRow = rows.find((r) => r.cabinetIndex === 1 && r.x === 168)!;
    expect(sameGroupRow.netPortIndex).toBe(port0Rows[0].netPortIndex);
    expect(sameGroupRow).toMatchObject({ y: 0, width: 168, height: 168, angle: 0 });
    expect(port1Rows[0]).toMatchObject({ cabinetIndex: 0, x: 336, y: 0, angle: 90 });
    db.close();
  });

  it("replacing cabinet topology twice clears the first set", async () => {
    const db = await NovaDb.open(dbBytes);
    db.replaceCabinetTopology([{ netPortIndex: 0, cabinetIndex: 0, x: 0, y: 0, width: 168, height: 168, angle: 0 }]);
    db.replaceCabinetTopology([{ netPortIndex: 0, cabinetIndex: 0, x: 10, y: 10, width: 168, height: 168, angle: 0 }]);
    const rows = db.readCabinetTopology();
    expect(rows).toHaveLength(1);
    expect(rows[0].x).toBe(10);
    db.close();
  });

  it("assigns input layers to spare template rows and clears unused ones", async () => {
    const db = await NovaDb.open(dbBytes);
    db.replaceInputLayers([
      { name: "Centre Screen", interfacePk: 4, x: 0, y: 0, width: 1920, height: 1080 },
      { name: "Side Screen", interfacePk: 11, x: 1920, y: 0, width: 960, height: 540 },
    ]);
    const layers = db.readInputLayers();
    expect(layers).toHaveLength(2);
    expect(layers[0]).toMatchObject({ name: "Centre Screen", interfacePk: 4, width: 1920, height: 1080 });
    expect(layers[1]).toMatchObject({ name: "Side Screen", interfacePk: 11 });
    db.close();
  });

  it("rejects more input assignments than the template has spare layers for", async () => {
    const db = await NovaDb.open(dbBytes);
    const tooMany = Array.from({ length: 13 }, (_, i) => ({
      name: `L${i}`,
      interfacePk: 4,
      x: 0,
      y: 0,
      width: 100,
      height: 100,
    }));
    expect(() => db.replaceInputLayers(tooMany)).toThrow();
    db.close();
  });

  it("sets device name and project identity", async () => {
    const db = await NovaDb.open(dbBytes);
    db.setDeviceName("Front Wall");
    db.setProjectIdentity("my-project-id", "My Project");
    db.close();
  });

  it("reads back the subcard manifest matching the template's known triplet", async () => {
    const db = await NovaDb.open(dbBytes);
    const manifest = db.readSubcardManifest();
    expect(manifest).toEqual([
      { slotId: 100, modelId: 25132, boardId: 0, cardType: 7, specialFunc: 0 },
      { slotId: 0, modelId: 15729040, boardId: 0, cardType: 2, specialFunc: 0 },
      { slotId: 1, modelId: 15729034, boardId: 0, cardType: 3, specialFunc: 0 },
    ]);
    db.close();
  });

  it("exports bytes that reopen successfully after edits", async () => {
    const db = await NovaDb.open(dbBytes);
    db.setCanvasResolution(2560, 1440);
    const bytes = db.toBytes();
    db.close();

    const reopened = await NovaDb.open(bytes);
    expect(reopened.readCanvasResolution()).toEqual({ width: 2560, height: 1440 });
    reopened.close();
  });
});
