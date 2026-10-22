import { describe, it, expect } from "vitest";
import { PNG_EXPORT_DPI, pixelsPerMetre, withPngDpi } from "./export/pngDpi";

// A PNG says how big its pixels are meant to be on paper in a pHYs chunk.
// Canvas never writes one, so every export arrived with no resolution and
// readers fell back to 72 DPI. These cover the chunk being written correctly,
// and - just as important - the pixels being left completely alone.

const SIG = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

const chunk = (type: string, data: number[] = []) => {
  const body = [...type].map((c) => c.charCodeAt(0)).concat(data);
  const len = data.length;
  // CRC is not checked by the reader below, so any four bytes will do here.
  return [(len >>> 24) & 255, (len >>> 16) & 255, (len >>> 8) & 255, len & 255, ...body, 0, 0, 0, 0];
};

const png = (...chunks: number[][]) => new Uint8Array([...SIG, ...chunks.flat()]);

const readChunks = (bytes: Uint8Array) => {
  const out: Array<{ type: string; data: Uint8Array }> = [];
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let pos = 8;
  while (pos + 8 <= bytes.length) {
    const len = view.getUint32(pos);
    const type = String.fromCharCode(...bytes.subarray(pos + 4, pos + 8));
    out.push({ type, data: bytes.subarray(pos + 8, pos + 8 + len) });
    pos += 12 + len;
  }
  return out;
};

describe("pixelsPerMetre", () => {
  it("converts DPI the way PNG stores it", () => {
    expect(pixelsPerMetre(300)).toBe(11811);
    expect(pixelsPerMetre(72)).toBe(2835);
  });
});

describe("withPngDpi", () => {
  const sample = () => png(chunk("IHDR", [0, 0, 0, 2, 0, 0, 0, 2, 8, 6, 0, 0, 0]), chunk("IDAT", [1, 2, 3]), chunk("IEND"));

  it("writes the resolution straight after IHDR, where the spec wants it", () => {
    const types = readChunks(withPngDpi(sample(), 300)).map((c) => c.type);
    expect(types).toEqual(["IHDR", "pHYs", "IDAT", "IEND"]);
  });

  it("records the asked-for DPI in both axes, in metres", () => {
    const phys = readChunks(withPngDpi(sample(), 300)).find((c) => c.type === "pHYs")!;
    const view = new DataView(phys.data.buffer, phys.data.byteOffset, phys.data.byteLength);
    expect(phys.data).toHaveLength(9);
    expect(view.getUint32(0)).toBe(11811);
    expect(view.getUint32(4)).toBe(11811);
    expect(phys.data[8]).toBe(1); // unit specifier: metres
  });

  it("leaves every other chunk byte for byte as it was", () => {
    const before = readChunks(sample()).filter((c) => c.type !== "pHYs");
    const after = readChunks(withPngDpi(sample(), 300)).filter((c) => c.type !== "pHYs");
    expect(after.map((c) => [c.type, [...c.data]])).toEqual(before.map((c) => [c.type, [...c.data]]));
  });

  it("replaces a resolution already there rather than adding a second", () => {
    const already = png(
      chunk("IHDR", [0, 0, 0, 2, 0, 0, 0, 2, 8, 6, 0, 0, 0]),
      chunk("pHYs", [0, 0, 11, 19, 0, 0, 11, 19, 1]), // 72 DPI
      chunk("IDAT", [1, 2, 3]),
      chunk("IEND"),
    );
    const out = readChunks(withPngDpi(already, 300));
    expect(out.filter((c) => c.type === "pHYs")).toHaveLength(1);
    const phys = out.find((c) => c.type === "pHYs")!;
    expect(new DataView(phys.data.buffer, phys.data.byteOffset, phys.data.byteLength).getUint32(0)).toBe(11811);
  });

  it("hands back anything it does not recognise, rather than losing the export", () => {
    const notPng = new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8, 9]);
    expect(withPngDpi(notPng, 300)).toBe(notPng);
    const noHeader = new Uint8Array(SIG);
    expect(withPngDpi(noHeader, 300)).toBe(noHeader);
  });

  it("defaults to the app's own export resolution", () => {
    const phys = readChunks(withPngDpi(sample())).find((c) => c.type === "pHYs")!;
    const view = new DataView(phys.data.buffer, phys.data.byteOffset, phys.data.byteLength);
    expect(view.getUint32(0)).toBe(pixelsPerMetre(PNG_EXPORT_DPI));
    expect(PNG_EXPORT_DPI).toBe(300);
  });
});
