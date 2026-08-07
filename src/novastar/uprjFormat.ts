// Byte-level codec for NovaStar's ".uprj" ("NovaProject") container format.
//
// Reverse-engineered from three real sample files (a blank single-processor
// template, a 4-internal-slot project, and a 2-device spliced-wall project).
// All three decode with the exact same envelope shape, which is the basis
// for trusting this structure rather than treating it as coincidence.
//
// Byte layout (all multi-byte integers little-endian):
//
//   "NovaProject" 0x00 0x00 "<" 0x00        - 15-byte fixed magic
//   JSON1                                    - UTF-8 JSON, no length prefix;
//                                              parse by scanning to the
//                                              matching closing brace.
//                                              Shape: {"version":"1.0.0","md5":"<32 hex chars>"}
//   0x01 0x00 <uint16 LE: len(JSON2)>        - 4-byte separator
//   JSON2                                    - UTF-8 JSON project metadata:
//                                              {"projectId","projectName","nodeList":[...]}
//                                              nodeList has one entry per
//                                              physical device in the file.
//   for each entry in nodeList, in order:
//     0x02 0x00 <uint32 LE: len(zip)>        - 6-byte separator
//     <zip bytes>                            - that device's data, a
//                                              standard ZIP archive
//
// The "md5" field inside JSON1 is the MD5 hash of every byte in the file
// *after* JSON1 (i.e. starting at the 4-byte separator and running to EOF).
// This was confirmed by hashing that exact byte range on all 3 samples and
// getting an exact match against the stored value - it is the file's only
// integrity check; there is no additional checksum beyond standard ZIP
// per-entry CRC32.
//
// A device's ZIP archive contains `<nodeFolder>/Userver.db` (a real SQLite 3
// database - the live configuration), `<nodeFolder>/Userver_Bak.db` (backup
// copy of the same schema), and `<nodeFolder>/subcard.json` (a plain-JSON
// mirror of the DB's `t_subcard` table). Some real-world exports bundle
// unrelated leftover application data alongside these - that is not part of
// the NovaStar format and must be ignored by any consumer of this module.

import { md5Hex } from "./md5";

const MAGIC_PREFIX = new Uint8Array([
  // "NovaProject"
  0x4e, 0x6f, 0x76, 0x61, 0x50, 0x72, 0x6f, 0x6a, 0x65, 0x63, 0x74,
  0x00, 0x00, 0x3c /* "<" */, 0x00,
]);

export type NodePart = {
  slotId: number;
  modelId: number;
  boardId: number;
  version: string;
  brief: string;
  section: unknown[];
};

export type NodeListEntry = {
  modelId: number;
  sn: string;
  ip: string;
  deviceName: string;
  fileType: number[];
  nodeVersion: {
    version: string;
    brief: string;
    database: { version: string };
    parts: NodePart[];
  };
};

export type ProjectMeta = {
  projectId: string;
  projectName: string;
  nodeList: NodeListEntry[];
};

export type NovaProjectEnvelope = {
  /** JSON1.version - always "1.0.0" on every sample seen; preserved as-is. */
  formatVersion: string;
  meta: ProjectMeta;
  /** One ZIP archive's raw bytes per meta.nodeList entry, same order. */
  deviceZips: Uint8Array[];
};

function u16le(n: number): [number, number] {
  return [n & 0xff, (n >> 8) & 0xff];
}
function u32le(n: number): [number, number, number, number] {
  return [n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, (n >> 24) & 0xff];
}
function readU16le(bytes: Uint8Array, offset: number): number {
  return bytes[offset] | (bytes[offset + 1] << 8);
}
function readU32le(bytes: Uint8Array, offset: number): number {
  return (
    (bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24)) >>> 0
  );
}

function concatBytes(chunks: (Uint8Array | number[])[]): Uint8Array {
  const arrays = chunks.map((c) => (c instanceof Uint8Array ? c : new Uint8Array(c)));
  const total = arrays.reduce((sum, a) => sum + a.length, 0);
  const out = new Uint8Array(total);
  let offset = 0;
  for (const a of arrays) {
    out.set(a, offset);
    offset += a.length;
  }
  return out;
}

/** Scans forward from `start` (which must point at a '{') to the index just past the matching '}'. */
function findJsonExtent(bytes: Uint8Array, start: number): number {
  const OPEN = 0x7b; // "{"
  const CLOSE = 0x7d; // "}"
  let depth = 0;
  for (let i = start; i < bytes.length; i++) {
    if (bytes[i] === OPEN) depth++;
    else if (bytes[i] === CLOSE) {
      depth--;
      if (depth === 0) return i + 1;
    }
  }
  throw new Error("uprjFormat: unterminated JSON block");
}

/**
 * Builds a single-device .uprj envelope. Always emits exactly one
 * nodeList entry / one zip blob - this tool never generates a
 * multi-processor file (see exportBuilder.ts for why).
 */
export function encodeSingleDeviceEnvelope(meta: ProjectMeta, deviceZip: Uint8Array): Uint8Array {
  if (meta.nodeList.length !== 1) {
    throw new Error("uprjFormat: encodeSingleDeviceEnvelope requires exactly one nodeList entry");
  }
  const json2Bytes = new TextEncoder().encode(JSON.stringify(meta));
  const sep1 = new Uint8Array([0x01, 0x00, ...u16le(json2Bytes.length)]);
  const sep2 = new Uint8Array([0x02, 0x00, ...u32le(deviceZip.length)]);

  const payloadAfterJson1 = concatBytes([sep1, json2Bytes, sep2, deviceZip]);
  const md5 = md5Hex(payloadAfterJson1);
  const json1Bytes = new TextEncoder().encode(JSON.stringify({ version: "1.0.0", md5 }));

  return concatBytes([MAGIC_PREFIX, json1Bytes, payloadAfterJson1]);
}

/** Decodes any .uprj envelope (single- or multi-device) for parsing/tests. */
export function decodeEnvelope(bytes: Uint8Array): NovaProjectEnvelope & { declaredMd5: string; actualMd5: string } {
  for (let i = 0; i < MAGIC_PREFIX.length; i++) {
    if (bytes[i] !== MAGIC_PREFIX[i]) {
      throw new Error("uprjFormat: missing NovaProject magic header");
    }
  }
  const json1Start = MAGIC_PREFIX.length;
  const json1End = findJsonExtent(bytes, json1Start);
  const json1 = JSON.parse(new TextDecoder().decode(bytes.slice(json1Start, json1End))) as {
    version: string;
    md5: string;
  };

  const afterJson1 = bytes.slice(json1End);
  const actualMd5 = md5Hex(afterJson1);

  if (afterJson1[0] !== 0x01 || afterJson1[1] !== 0x00) {
    throw new Error("uprjFormat: unexpected separator before JSON2");
  }
  const json2Len = readU16le(afterJson1, 2);
  const json2Start = 4;
  const json2End = json2Start + json2Len;
  const meta = JSON.parse(new TextDecoder().decode(afterJson1.slice(json2Start, json2End))) as ProjectMeta;

  const deviceZips: Uint8Array[] = [];
  let cursor = json2End;
  for (let i = 0; i < meta.nodeList.length; i++) {
    if (afterJson1[cursor] !== 0x02 || afterJson1[cursor + 1] !== 0x00) {
      throw new Error(`uprjFormat: unexpected separator before device zip #${i}`);
    }
    const zipLen = readU32le(afterJson1, cursor + 2);
    const zipStart = cursor + 6;
    const zipEnd = zipStart + zipLen;
    deviceZips.push(afterJson1.slice(zipStart, zipEnd));
    cursor = zipEnd;
  }

  return { formatVersion: json1.version, meta, deviceZips, declaredMd5: json1.md5, actualMd5 };
}
