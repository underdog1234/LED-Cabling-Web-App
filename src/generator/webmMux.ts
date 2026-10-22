// ---------------------------------------------------------------------------
// A small WebM (Matroska) writer for frames encoded with WebCodecs.
//
// Every frame and audio packet arrives with its exact timestamp, so the file
// says precisely how long it is and where every frame sits - unlike a
// real-time MediaRecorder capture, which is only as good as the browser's
// timers were while it ran. One cluster per video keyframe, a Duration in
// the header and a Cues index at the end, so players show the right length
// and can seek.
// ---------------------------------------------------------------------------

export type MuxPacket = { track: 1 | 2; timestampUs: number; key: boolean; data: Uint8Array };

export type MuxOptions = {
  width: number;
  height: number;
  fps: number;
  durationSeconds: number;
  videoCodec: "V_VP9" | "V_VP8";
  /** Opus audio, when there is any: its OpusHead header and the encoder's pre-skip. */
  audio?: { codecPrivate: Uint8Array; sampleRate: number; channels: number; preSkip: number };
};

const concat = (parts: Uint8Array[]): Uint8Array => {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let at = 0;
  parts.forEach((p) => {
    out.set(p, at);
    at += p.length;
  });
  return out;
};

const idBytes = (id: number): Uint8Array => {
  const bytes: number[] = [];
  let v = id;
  while (v > 0) {
    bytes.unshift(v & 0xff);
    v = Math.floor(v / 256);
  }
  return new Uint8Array(bytes);
};

/** Element size as an 8-byte EBML vint: simple, and big enough for any file. */
const size8 = (n: number): Uint8Array => {
  const out = new Uint8Array(8);
  out[0] = 0x01;
  let v = n;
  for (let i = 7; i >= 1; i -= 1) {
    out[i] = v & 0xff;
    v = Math.floor(v / 256);
  }
  return out;
};

const el = (id: number, ...children: Uint8Array[]): Uint8Array => {
  const body = concat(children);
  return concat([idBytes(id), size8(body.length), body]);
};

const uint = (id: number, value: number): Uint8Array => {
  const bytes: number[] = [];
  let v = Math.max(0, Math.round(value));
  do {
    bytes.unshift(v & 0xff);
    v = Math.floor(v / 256);
  } while (v > 0);
  return el(id, new Uint8Array(bytes));
};

const float = (id: number, value: number): Uint8Array => {
  const b = new Uint8Array(8);
  new DataView(b.buffer).setFloat64(0, value);
  return el(id, b);
};

const str = (id: number, value: string): Uint8Array => el(id, new TextEncoder().encode(value));

const simpleBlock = (p: MuxPacket, clusterMs: number): Uint8Array => {
  const head = new Uint8Array(4);
  head[0] = 0x80 | p.track;
  const rel = Math.round(p.timestampUs / 1000) - clusterMs;
  new DataView(head.buffer).setInt16(1, rel);
  head[3] = p.key ? 0x80 : 0x00;
  return el(0xa3, head, p.data);
};

/** OpusHead for mono or stereo, as WebM's CodecPrivate wants it. */
export const opusHead = (channels: number, sampleRate: number, preSkip: number): Uint8Array => {
  const b = new Uint8Array(19);
  b.set(new TextEncoder().encode("OpusHead"), 0);
  const v = new DataView(b.buffer);
  v.setUint8(8, 1);
  v.setUint8(9, channels);
  v.setUint16(10, preSkip, true);
  v.setUint32(12, sampleRate, true);
  v.setInt16(16, 0, true);
  v.setUint8(18, 0);
  return b;
};

export const muxWebm = (packets: MuxPacket[], o: MuxOptions): Uint8Array => {
  const ebml = el(
    0x1a45dfa3,
    uint(0x4286, 1),
    uint(0x42f7, 1),
    uint(0x42f2, 4),
    uint(0x42f3, 8),
    str(0x4282, "webm"),
    uint(0x4287, 4),
    uint(0x4285, 2),
  );
  const info = el(0x1549a966, uint(0x2ad7b1, 1_000_000), float(0x4489, o.durationSeconds * 1000), str(0x4d80, "Test Pattern Generator"), str(0x5741, "Test Pattern Generator"));
  const video = el(
    0xae,
    uint(0xd7, 1),
    uint(0x73c5, 1),
    uint(0x83, 1),
    str(0x86, o.videoCodec),
    uint(0x23e383, Math.round(1e9 / o.fps)),
    el(0xe0, uint(0xb0, o.width), uint(0xba, o.height)),
  );
  const tracks = [video];
  if (o.audio) {
    tracks.push(
      el(
        0xae,
        uint(0xd7, 2),
        uint(0x73c5, 2),
        uint(0x83, 2),
        str(0x86, "A_OPUS"),
        el(0x63a2, o.audio.codecPrivate),
        uint(0x56aa, Math.round((o.audio.preSkip / 48000) * 1e9)),
        uint(0x56bb, 80_000_000),
        el(0xe1, float(0xb5, o.audio.sampleRate), uint(0x9f, o.audio.channels)),
      ),
    );
  }
  const tracksEl = el(0x1654ae6b, ...tracks);

  // Clusters: a new one at every video keyframe, packets in time order.
  const sorted = [...packets].sort((a, b) => a.timestampUs - b.timestampUs || a.track - b.track);
  const clusters: Array<{ ms: number; blocks: Uint8Array[] }> = [];
  sorted.forEach((p) => {
    const ms = Math.round(p.timestampUs / 1000);
    const last = clusters[clusters.length - 1];
    if (!last || (p.track === 1 && p.key) || ms - last.ms > 30_000) clusters.push({ ms, blocks: [] });
    clusters[clusters.length - 1].blocks.push(simpleBlock(p, clusters[clusters.length - 1].ms));
  });
  const clusterEls = clusters.map((c) => el(0x1f43b675, uint(0xe7, c.ms), ...c.blocks));

  // Cues point at each cluster, counted from the start of the segment's data.
  let offset = info.length + tracksEl.length;
  const cuePoints = clusterEls.map((c, i) => {
    const point = el(0xbb, uint(0xb3, clusters[i].ms), el(0xb7, uint(0xf7, 1), uint(0xf1, offset)));
    offset += c.length;
    return point;
  });
  const cues = el(0x1c53bb6b, ...cuePoints);
  return concat([ebml, el(0x18538067, info, tracksEl, ...clusterEls, cues)]);
};
