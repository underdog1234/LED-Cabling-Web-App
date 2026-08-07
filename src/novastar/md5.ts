// Minimal RFC 1321 MD5 implementation.
//
// The .uprj container format's integrity field is an MD5 hash (see
// uprjFormat.ts) and the Web Crypto API's SubtleCrypto intentionally does
// not implement MD5 (it's considered broken for security purposes, but
// NovaStar uses it here purely as a length/corruption check, not a security
// control). Rather than add a dependency for one hash function, this is a
// standard, self-contained implementation.

function rotl(x: number, c: number): number {
  return (x << c) | (x >>> (32 - c));
}

const S = [
  7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 5, 9, 14, 20, 5, 9, 14, 20, 5, 9, 14, 20, 5, 9, 14, 20,
  4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 6, 10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21, 6, 10, 15,
  21,
];

const K = new Int32Array([
  -680876936, -389564586, 606105819, -1044525330, -176418897, 1200080426, -1473231341, -45705983, 1770035416,
  -1958414417, -42063, -1990404162, 1804603682, -40341101, -1502002290, 1236535329, -165796510, -1069501632,
  643717713, -373897302, -701558691, 38016083, -660478335, -405537848, 568446438, -1019803690, -187363961,
  1163531501, -1444681467, -51403784, 1735328473, -1926607734, -378558, -2022574463, 1839030562, -35309556,
  -1530992060, 1272893353, -155497632, -1094730640, 681279174, -358537222, -722521979, 76029189, -640364487,
  -421815835, 530742520, -995338651, -198630844, 1126891415, -1416354905, -57434055, 1700485571, -1894986606,
  -1051523, -2054922799, 1873313359, -30611744, -1560198380, 1309151649, -145523070, -1120210379, 718787259,
  -343485551,
]);

/** Returns the lowercase hex MD5 digest of the given bytes. */
export function md5Hex(bytes: Uint8Array): string {
  const msgLen = bytes.length;
  const withPad = new Uint8Array((((msgLen + 8) >> 6) + 1) << 6);
  withPad.set(bytes);
  withPad[msgLen] = 0x80;
  const bitLenLow = (msgLen * 8) >>> 0;
  const bitLenHigh = Math.floor((msgLen * 8) / 0x100000000) >>> 0;
  const view = new DataView(withPad.buffer);
  view.setUint32(withPad.length - 8, bitLenLow, true);
  view.setUint32(withPad.length - 4, bitLenHigh, true);

  let a0 = 0x67452301;
  let b0 = 0xefcdab89;
  let c0 = 0x98badcfe;
  let d0 = 0x10325476;

  const M = new Int32Array(16);
  for (let chunkStart = 0; chunkStart < withPad.length; chunkStart += 64) {
    for (let j = 0; j < 16; j++) {
      M[j] = view.getInt32(chunkStart + j * 4, true);
    }
    let A = a0;
    let B = b0;
    let C = c0;
    let D = d0;

    for (let i = 0; i < 64; i++) {
      let F: number;
      let g: number;
      if (i < 16) {
        F = (B & C) | (~B & D);
        g = i;
      } else if (i < 32) {
        F = (D & B) | (~D & C);
        g = (5 * i + 1) % 16;
      } else if (i < 48) {
        F = B ^ C ^ D;
        g = (3 * i + 5) % 16;
      } else {
        F = C ^ (B | ~D);
        g = (7 * i) % 16;
      }
      F = (F + A + K[i] + M[g]) | 0;
      A = D;
      D = C;
      C = B;
      B = (B + rotl(F, S[i])) | 0;
    }

    a0 = (a0 + A) | 0;
    b0 = (b0 + B) | 0;
    c0 = (c0 + C) | 0;
    d0 = (d0 + D) | 0;
  }

  const out = new Uint8Array(16);
  const outView = new DataView(out.buffer);
  outView.setInt32(0, a0, true);
  outView.setInt32(4, b0, true);
  outView.setInt32(8, c0, true);
  outView.setInt32(12, d0, true);

  return Array.from(out)
    .map((b) => b.toString(16).padStart(2, "0"))
    .join("");
}
