// A PNG carries its intended print resolution in a `pHYs` chunk. Canvas's own
// toDataURL / toBlob never writes one, so every PNG this app exported arrived
// with no resolution at all - and a reader with nothing to go on assumes 72
// (sometimes 96) DPI. Dropped into a document or a print queue, a test pattern
// came out at roughly four times the size it should be.
//
// This writes the chunk in afterwards. It deliberately does NOT touch a single
// pixel: a test pattern is mapped one-for-one onto the LED wall's own
// resolution, so resampling it to "make it 300 DPI" would destroy the very
// thing it exists to show. The pixels are the pixels; this only says how big
// they are meant to be on paper.

/** Print resolution tagged onto exported PNGs. */
export const PNG_EXPORT_DPI = 300;

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
const MM_PER_INCH = 25.4;

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

const crc32 = (bytes: Uint8Array): number => {
  let c = 0xffffffff;
  for (let i = 0; i < bytes.length; i += 1) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
};

/** DPI as PNG stores it: whole pixels per metre. */
export const pixelsPerMetre = (dpi: number) => Math.round((dpi / MM_PER_INCH) * 1000);

const isPng = (bytes: Uint8Array) =>
  bytes.length > 8 && PNG_SIGNATURE.every((b, i) => bytes[i] === b);

const buildPhys = (dpi: number): Uint8Array => {
  const chunk = new Uint8Array(21); // 4 length + 4 type + 9 data + 4 crc
  const view = new DataView(chunk.buffer);
  view.setUint32(0, 9);
  chunk.set([0x70, 0x48, 0x59, 0x73], 4); // "pHYs"
  const perMetre = pixelsPerMetre(dpi);
  view.setUint32(8, perMetre);
  view.setUint32(12, perMetre);
  chunk[16] = 1; // unit specifier: 1 = metre
  view.setUint32(17, crc32(chunk.subarray(4, 17)));
  return chunk;
};

/**
 * The same PNG with its print resolution set to `dpi`.
 *
 * The chunk goes immediately after IHDR, which the spec requires of pHYs
 * (before the first IDAT) and which is also where any existing one would be,
 * so a file that already carries a resolution has it replaced rather than
 * duplicated. Anything that is not a PNG is handed back untouched - this is on
 * a download path, and an export that failed to be tagged is still an export
 * worth giving to the person.
 */
export const withPngDpi = (bytes: Uint8Array, dpi = PNG_EXPORT_DPI): Uint8Array => {
  if (!isPng(bytes) || !(dpi > 0)) return bytes;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const out: Uint8Array[] = [bytes.subarray(0, 8)];
  let pos = 8;
  let wrote = false;
  while (pos + 8 <= bytes.length) {
    const length = view.getUint32(pos);
    const end = pos + 12 + length;
    if (end > bytes.length) break; // truncated - leave the rest as it is
    const type = String.fromCharCode(bytes[pos + 4], bytes[pos + 5], bytes[pos + 6], bytes[pos + 7]);
    if (type === "pHYs") {
      // Drop the old one; the replacement is written after IHDR either way.
      pos = end;
      continue;
    }
    out.push(bytes.subarray(pos, end));
    if (type === "IHDR") {
      out.push(buildPhys(dpi));
      wrote = true;
    }
    pos = end;
  }
  if (pos < bytes.length) out.push(bytes.subarray(pos));
  if (!wrote) return bytes; // no IHDR found - not a PNG we understand
  const total = out.reduce((sum, part) => sum + part.length, 0);
  const merged = new Uint8Array(total);
  let at = 0;
  out.forEach((part) => {
    merged.set(part, at);
    at += part.length;
  });
  return merged;
};
