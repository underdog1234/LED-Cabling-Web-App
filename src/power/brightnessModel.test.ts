import { describe, it, expect } from "vitest";
import {
  BARCO_EXAMPLE_BLACK_LEVEL_RATIO,
  BARCO_TYPICAL_BRIGHTNESS_PERCENT,
  BARCO_TYPICAL_CONTENT_PERCENT,
  defaultBlackLevelW,
  panelWattsAt,
  ruleOfThumbWatts,
  scalePanelPower,
} from "./brightnessModel";

// Barco's own worked example, reproduced number for number. If these drift,
// the model has stopped matching the source it is built on.
//
//   Max 100W, black level 30W -> scalable range 70W
//   70% brightness, full white:      (70 x 0.70) + 30        = 79W
//   70% brightness, typical content: (70 x 0.70 x 0.33) + 30 = 46.2W
//   Old divide-by-three:                                       33W
describe("the Barco worked example", () => {
  const maxW = 100;
  const blackLevelW = 30;

  it("full white at 70% brightness is 79W", () => {
    expect(panelWattsAt({ maxW, blackLevelW, brightnessPercent: 70, contentPercent: 100 })).toBeCloseTo(79, 6);
  });

  it("typical content at 70% brightness is 46.17W, which the article prints as 46.2", () => {
    // 70 x 0.70 x 0.33 = 16.17, + 30 = 46.17. The article rounds that to one
    // decimal in its table; the model keeps the exact value and the UI does
    // its own rounding, so this asserts both.
    const watts = panelWattsAt({ maxW, blackLevelW, brightnessPercent: 70, contentPercent: 33 });
    expect(watts).toBeCloseTo(46.17, 6);
    expect(Number(watts.toFixed(1))).toBe(46.2);
  });

  it("the old rule of thumb would have said 33W", () => {
    expect(ruleOfThumbWatts(maxW)).toBeCloseTo(33.333, 3);
    // ...which is BELOW the black level plus any real content, and that is
    // exactly the article's complaint about it.
    expect(ruleOfThumbWatts(maxW)).toBeLessThan(
      panelWattsAt({ maxW, blackLevelW, brightnessPercent: 70, contentPercent: 33 }),
    );
  });

  it("uses the article's own default percentages", () => {
    expect(BARCO_TYPICAL_BRIGHTNESS_PERCENT).toBe(70);
    expect(BARCO_TYPICAL_CONTENT_PERCENT).toBe(33);
    expect(BARCO_EXAMPLE_BLACK_LEVEL_RATIO).toBe(0.3);
    expect(defaultBlackLevelW(100)).toBe(30);
  });
});

describe("the black level is a floor, not a fraction", () => {
  it("full brightness on full white is the measured maximum", () => {
    expect(panelWattsAt({ maxW: 175, blackLevelW: 52.5, brightnessPercent: 100, contentPercent: 100 })).toBeCloseTo(175, 6);
  });

  it("zero brightness still draws the black level - the drivers are still on", () => {
    expect(panelWattsAt({ maxW: 175, blackLevelW: 52.5, brightnessPercent: 0, contentPercent: 100 })).toBeCloseTo(52.5, 6);
  });

  it("full black content still draws the black level", () => {
    expect(panelWattsAt({ maxW: 175, blackLevelW: 52.5, brightnessPercent: 100, contentPercent: 0 })).toBeCloseTo(52.5, 6);
  });

  it("never drops below the black level however the dials are set", () => {
    for (const brightness of [0, 1, 25, 70, 100]) {
      for (const content of [0, 10, 33, 100]) {
        const watts = panelWattsAt({ maxW: 175, blackLevelW: 52.5, brightnessPercent: brightness, contentPercent: content });
        expect(watts).toBeGreaterThanOrEqual(52.5);
        expect(watts).toBeLessThanOrEqual(175);
      }
    }
  });

  it("halving the brightness does NOT halve the power, because of that floor", () => {
    const full = panelWattsAt({ maxW: 100, blackLevelW: 30, brightnessPercent: 100, contentPercent: 100 });
    const half = panelWattsAt({ maxW: 100, blackLevelW: 30, brightnessPercent: 50, contentPercent: 100 });
    expect(full).toBeCloseTo(100, 6);
    expect(half).toBeCloseTo(65, 6); // not 50
  });

  it("with no black level at all it collapses to plain proportional scaling", () => {
    expect(panelWattsAt({ maxW: 100, blackLevelW: 0, brightnessPercent: 50, contentPercent: 100 })).toBeCloseTo(50, 6);
  });
});

describe("nonsense inputs cannot produce nonsense power", () => {
  it("clamps a black level entered above the maximum", () => {
    expect(panelWattsAt({ maxW: 100, blackLevelW: 250, brightnessPercent: 50, contentPercent: 100 })).toBeCloseTo(100, 6);
  });

  it("clamps percentages outside 0-100", () => {
    expect(panelWattsAt({ maxW: 100, blackLevelW: 30, brightnessPercent: 400, contentPercent: 100 })).toBeCloseTo(100, 6);
    expect(panelWattsAt({ maxW: 100, blackLevelW: 30, brightnessPercent: -50, contentPercent: 100 })).toBeCloseTo(30, 6);
  });

  it("survives a zero-watt panel", () => {
    expect(panelWattsAt({ maxW: 0, blackLevelW: 0, brightnessPercent: 70, contentPercent: 33 })).toBe(0);
  });
});

describe("scalePanelPower keeps amps consistent with watts", () => {
  // MG9: 175W / 0.77A - whatever voltage basis that implies, the scaled pair
  // has to keep it, or the amp checks and the watt totals would disagree.
  const spec = { maxW: 175, maxA: 0.77 };

  it("at full brightness and full white it is the catalogue figure", () => {
    const scaled = scalePanelPower(spec, 52.5, 100, 100);
    expect(scaled.maxW).toBeCloseTo(175, 6);
    expect(scaled.maxA).toBeCloseTo(0.77, 6);
  });

  it("holds the panel's own watts-per-amp at every setting", () => {
    const scaled = scalePanelPower(spec, 52.5, 70, 33);
    expect(scaled.maxW / scaled.maxA).toBeCloseTo(spec.maxW / spec.maxA, 6);
    expect(scaled.avgW / scaled.avgA).toBeCloseTo(spec.maxW / spec.maxA, 6);
  });

  it("the design figure is full white, so it is never below the typical one", () => {
    const scaled = scalePanelPower(spec, 52.5, 70, 33);
    expect(scaled.maxW).toBeGreaterThan(scaled.avgW);
    expect(scaled.maxA).toBeGreaterThan(scaled.avgA);
  });

  it("design and typical meet when the content is full white", () => {
    const scaled = scalePanelPower(spec, 52.5, 70, 100);
    expect(scaled.maxW).toBeCloseTo(scaled.avgW, 6);
  });
});
