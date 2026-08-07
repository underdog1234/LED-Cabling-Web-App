// Reads a .uprj file back into a plain structured object, for tests and an
// optional on-screen debug view. Deliberately reuses NovaDb's own read
// methods rather than querying SQL directly here, so "what a field means"
// stays defined in exactly one place (novaDb.ts).

import { unzipSync } from "fflate";
import { decodeEnvelope } from "./uprjFormat";
import { NovaDb } from "./novaDb";

export type ParsedUprj = {
  projectId: string;
  projectName: string;
  deviceName: string;
  /** Always 1 for files this tool produces - see exportBuilder.ts. */
  deviceCount: number;
  checksumValid: boolean;
  canvas: { width: number; height: number };
  mainLogicScreen: { width: number; height: number; name: string };
  cabinets: {
    netPortIndex: number;
    cabinetIndex: number;
    x: number;
    y: number;
    width: number;
    height: number;
    angle: number;
  }[];
  inputLayers: { name: string; interfacePk: number; x: number; y: number; width: number; height: number }[];
  subcard: { slotId: number; modelId: number; boardId: number; cardType: number; specialFunc: number }[];
};

export async function parseUprj(bytes: Uint8Array): Promise<ParsedUprj> {
  const env = decodeEnvelope(bytes);
  const zipFiles = unzipSync(env.deviceZips[0]);
  const dbKey = Object.keys(zipFiles).find((k) => k.endsWith("Userver.db"));
  if (!dbKey) throw new Error("parseUprj: no Userver.db found in the device zip");

  const db = await NovaDb.open(zipFiles[dbKey]);
  try {
    return {
      projectId: env.meta.projectId,
      projectName: env.meta.projectName,
      deviceName: env.meta.nodeList[0]?.deviceName ?? "",
      deviceCount: env.meta.nodeList.length,
      checksumValid: env.declaredMd5 === env.actualMd5,
      canvas: db.readCanvasResolution(),
      mainLogicScreen: db.readMainLogicScreen(),
      cabinets: db.readCabinetTopology().map((c) => ({
        netPortIndex: c.netPortIndex,
        cabinetIndex: c.cabinetIndex,
        x: c.x,
        y: c.y,
        width: c.width,
        height: c.height,
        angle: c.angle,
      })),
      inputLayers: db.readInputLayers(),
      subcard: db.readSubcardManifest(),
    };
  } finally {
    db.close();
  }
}
