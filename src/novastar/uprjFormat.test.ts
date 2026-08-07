import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { decodeEnvelope, encodeSingleDeviceEnvelope, type ProjectMeta } from "./uprjFormat";
import { md5Hex } from "./md5";

function fixture(name: string): Uint8Array {
  const path = fileURLToPath(new URL(`./__fixtures__/${name}`, import.meta.url));
  return new Uint8Array(readFileSync(path));
}

describe("md5Hex", () => {
  it("matches known RFC 1321 test vectors", () => {
    const enc = new TextEncoder();
    expect(md5Hex(enc.encode(""))).toBe("d41d8cd98f00b204e9800998ecf8427e");
    expect(md5Hex(enc.encode("abc"))).toBe("900150983cd24fb0d6963f7d28e17f72");
    expect(md5Hex(enc.encode("message digest"))).toBe("f96b697d7cb7938d525a2f31aaf161d0");
    expect(md5Hex(enc.encode("abcdefghijklmnopqrstuvwxyz"))).toBe("c3fcd3d76192e4007dfb496cca67e13b");
  });
});

describe("decodeEnvelope on real sample files", () => {
  it("decodes blank.uprj (single device) with a verified checksum", () => {
    const env = decodeEnvelope(fixture("blank.uprj"));
    expect(env.actualMd5).toBe(env.declaredMd5);
    expect(env.meta.nodeList).toHaveLength(1);
    expect(env.deviceZips).toHaveLength(1);
    expect(env.meta.nodeList[0].deviceName).toBe("Device3");
  });

  it("decodes piev3.uprj (single device, 4 internal slots) with a verified checksum", () => {
    const env = decodeEnvelope(fixture("piev3.uprj"));
    expect(env.actualMd5).toBe(env.declaredMd5);
    expect(env.meta.nodeList).toHaveLength(1);
    expect(env.deviceZips).toHaveLength(1);
  });

  it("decodes with-cables.uprj (two independent devices) with a verified checksum", () => {
    const env = decodeEnvelope(fixture("with-cables.uprj"));
    expect(env.actualMd5).toBe(env.declaredMd5);
    expect(env.meta.nodeList).toHaveLength(2);
    expect(env.deviceZips).toHaveLength(2);
    expect(env.meta.nodeList[0].deviceName).toBe("Device1");
    expect(env.meta.nodeList[1].deviceName).toBe("Device3");
  });
});

describe("encodeSingleDeviceEnvelope round-trip", () => {
  it("re-encodes blank.uprj's own meta/zip and decodes back identically", () => {
    const original = decodeEnvelope(fixture("blank.uprj"));
    const meta: ProjectMeta = original.meta;
    const rebuilt = encodeSingleDeviceEnvelope(meta, original.deviceZips[0]);

    const decoded = decodeEnvelope(rebuilt);
    expect(decoded.actualMd5).toBe(decoded.declaredMd5);
    expect(decoded.meta).toEqual(meta);
    expect(decoded.deviceZips[0]).toEqual(original.deviceZips[0]);
  });

  it("rejects a multi-device meta", () => {
    const original = decodeEnvelope(fixture("with-cables.uprj"));
    expect(() => encodeSingleDeviceEnvelope(original.meta, original.deviceZips[0])).toThrow();
  });
});
