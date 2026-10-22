import { describe, it, expect } from "vitest";
import {
  MP4_PROFILE,
  MP4_RECORD_MARGIN_SECONDS,
  buildMp4Args,
  defaultMp4Settings,
  h264LevelFor,
  keyframeIntervalFor,
  type Mp4EncodeSettings,
} from "./testPattern/mp4Encode";

// The MP4 goes on a media server, so these settings are a contract rather than
// a preference: every one of them is pinned here, because a file that silently
// drifts to variable frame rate or the wrong colour range is one nobody notices
// until the wall is up.

const settings = (over: Partial<Mp4EncodeSettings> = {}): Mp4EncodeSettings => ({
  ...defaultMp4Settings(1920, 1080, 20),
  ...over,
});

const valueAfter = (args: string[], flag: string) => args[args.indexOf(flag) + 1];

describe("the delivery settings", () => {
  it("starts where the spec says: 60fps, 8 Mbps target, 12 Mbps max", () => {
    expect(defaultMp4Settings(1920, 1080, 20)).toMatchObject({ fps: 60, targetMbps: 8, maxMbps: 12 });
  });

  it("cuts the file to exactly one loop", () => {
    // A recorder hands back the frames it happened to catch - 19.22s of a 20s
    // loop, in the case this came from - and a file short of the loop jumps
    // every time it repeats. The recording is made longer and cut back.
    expect(valueAfter(buildMp4Args(settings({ loopSeconds: 20 }), false), "-t")).toBe("20");
    expect(MP4_RECORD_MARGIN_SECONDS).toBeGreaterThan(0);
  });

  it("keyframes every half second, whatever the frame rate", () => {
    expect(keyframeIntervalFor(60)).toBe(30);
    expect(keyframeIntervalFor(50)).toBe(25);
    expect(keyframeIntervalFor(30)).toBe(15);
    expect(keyframeIntervalFor(25)).toBe(13); // 12.5 rounded - still half a second to the frame
  });

  it("writes H.264 High, 8-bit 4:2:0, progressive, no B-frames", () => {
    const args = buildMp4Args(settings(), false);
    expect(valueAfter(args, "-c:v")).toBe("libx264");
    expect(valueAfter(args, "-profile:v")).toBe("high");
    expect(valueAfter(args, "-pix_fmt") ?? "").not.toBe("yuv422p");
    expect(args.join(" ")).toContain("format=yuv420p");
    expect(valueAfter(args, "-bf")).toBe("0");
  });

  it("forces a constant frame rate, so a wall that draws slowly still delivers 60p", () => {
    const args = buildMp4Args(settings(), false);
    expect(valueAfter(args, "-fps_mode")).toBe("cfr");
    expect(valueAfter(args, "-r")).toBe("60");
  });

  it("fixes the GOP at the keyframe interval, scene cuts included", () => {
    const args = buildMp4Args(settings({ fps: 50 }), false);
    expect(valueAfter(args, "-g")).toBe("25");
    expect(valueAfter(args, "-keyint_min")).toBe("25");
    expect(valueAfter(args, "-sc_threshold")).toBe("0");
  });

  it("asks for the target bitrate with a ceiling and a buffer", () => {
    const args = buildMp4Args(settings({ targetMbps: 8, maxMbps: 12 }), false);
    expect(valueAfter(args, "-b:v")).toBe("8000k");
    expect(valueAfter(args, "-maxrate")).toBe("12000k");
    expect(valueAfter(args, "-bufsize")).toBe("12000k");
  });

  it("tags Rec.709 limited range, in the stream and in the file", () => {
    const args = buildMp4Args(settings(), false);
    expect(valueAfter(args, "-color_primaries")).toBe("bt709");
    expect(valueAfter(args, "-color_trc")).toBe("bt709");
    expect(valueAfter(args, "-colorspace")).toBe("bt709");
    expect(valueAfter(args, "-color_range")).toBe("tv");
    expect(valueAfter(args, "-movflags")).toContain("write_colr");
  });

  it("converts from whatever the recording actually was, not from a guess", () => {
    expect(valueAfter(buildMp4Args(settings(), false), "-vf")).toBe("scale=in_range=limited:out_range=limited,format=yuv420p");
    expect(valueAfter(buildMp4Args(settings(), true), "-vf")).toBe("scale=in_range=full:out_range=limited,format=yuv420p");
  });

  it("carries no audio track", () => {
    expect(buildMp4Args(settings(), false)).toContain("-an");
  });
});

describe("h264LevelFor", () => {
  it("uses the asked-for 4.2 for anything that fits in it", () => {
    expect(MP4_PROFILE.preferredLevel).toBe("4.2");
    expect(h264LevelFor(1920, 1080, 60)).toBe("4.2");
    expect(h264LevelFor(1280, 720, 60)).toBe("4.2");
  });

  it("steps up for a wall 4.2 cannot carry, rather than tagging a lie", () => {
    // 6216 x 1344 at 60fps - the eagletc wall - is 32,640 macroblocks a frame
    // and 1.96M a second: past 4.2 on both counts, and past 5.1 on rate.
    expect(h264LevelFor(6216, 1344, 60)).toBe("5.2");
    expect(h264LevelFor(6216, 1344, 30)).toBe("5.1");
    // Slowing a wall down can bring it back inside a lower level.
    expect(h264LevelFor(3840, 2160, 60)).toBe("5.2");
    expect(h264LevelFor(3840, 2160, 30)).toBe("5.1");
  });

  it("never runs out: the biggest level is the floor", () => {
    expect(h264LevelFor(30000, 20000, 60)).toBe("6.2");
  });
});
