// Power draw as a function of configured brightness.
//
// Method and figures from Barco, "The truth about the power consumption of LED
// walls" (2020-04-15, barco.com/it/news/news-insights/2020-04-15-ledtalks-
// truth-about-power-consumption). Their point, in short:
//
//   "The only power consumptions that can be measured are the maximum, the
//   black level and the standby power consumption."
//
//   - MAXIMUM is the wall at full brightness showing full white.
//   - BLACK LEVEL is "the power needed to run the electronics without showing
//     content on the LED wall. Because even though the diodes are off, the
//     drivers and receiver cards are still consuming energy."
//
// Only the difference between those two scales with brightness. The black
// level is a fixed floor that does not, which is the whole reason the old
// rule of thumb is wrong: "The traditional industry rules of thumb (dividing
// the maximum power consumption by three) usually don't take into account the
// average use and black level power consumptions." Barco note it matters more
// the finer the pitch, since more pixels need more drivers.
//
// Their worked example, which the tests below reproduce exactly:
//
//   Max 100W, black level 30W, so the scalable range is 70W.
//   At 70% brightness, full white:        (70 x 0.70) + 30        = 79W
//   At 70% brightness, typical content:   (70 x 0.70 x 0.33) + 30 = 46.2W
//   The old divide-by-three would have said 33W.
//
// Two figures come out of this and they are NOT interchangeable:
//
//   - the DESIGN figure (full white at the configured brightness) is what
//     cable, outlets and breakers have to carry, because a white frame is
//     always one cue away; and
//   - the TYPICAL figure (with the content factor) is what it costs to run.
//
// Sizing anything off the typical figure would under-size the supply, so the
// two are kept apart here and the caller has to say which it wants.

export type BrightnessPowerInputs = {
  /** Full white at full brightness - the panel's measured maximum, in watts. */
  maxW: number;
  /** Drivers and receiver cards with every diode off, in watts. */
  blackLevelW: number;
  /** Configured brightness, as a percentage of the panel's maximum. */
  brightnessPercent: number;
  /**
   * What the content actually draws, as a percentage of the configured
   * brightness. 100 = full white, the design case. Barco's own default for
   * mixed content is 33, and they are explicit that it varies by what is on
   * screen: a white-background lobby slide is higher, a mostly-black control
   * room lower.
   */
  contentPercent: number;
};

/** Barco's example black level, 30W against a 100W maximum. A placeholder until a panel is measured. */
export const BARCO_EXAMPLE_BLACK_LEVEL_RATIO = 0.3;
/** "On average a customer configures the display at 70% of its maximum brightness." */
export const BARCO_TYPICAL_BRIGHTNESS_PERCENT = 70;
/** "the typical content on the wall consumes only 33% of that power". */
export const BARCO_TYPICAL_CONTENT_PERCENT = 33;

const clamp = (value: number, min: number, max: number) => Math.min(Math.max(value, min), max);

/**
 * Watts for one panel at a given brightness and content level.
 *
 * `(max - black) x brightness x content + black` - the black level is added
 * back outside the scaling because it does not scale.
 */
export const panelWattsAt = (inputs: BrightnessPowerInputs): number => {
  const maxW = Math.max(0, inputs.maxW);
  // A black level above the maximum is a typo, not a panel; clamping keeps the
  // scalable range from going negative and the answer from exceeding the
  // measured maximum.
  const blackW = clamp(inputs.blackLevelW, 0, maxW);
  const brightness = clamp(inputs.brightnessPercent, 0, 100) / 100;
  const content = clamp(inputs.contentPercent, 0, 100) / 100;
  return (maxW - blackW) * brightness * content + blackW;
};

export type ScaledPanelPower = {
  /** Full white at the configured brightness. Size cable, outlets and breakers off this. */
  maxW: number;
  maxA: number;
  /** The configured brightness with the content factor applied. Running cost, not sizing. */
  avgW: number;
  avgA: number;
};

/**
 * A panel's catalogue power figures rescaled for a configured brightness.
 *
 * Amps follow watts at the panel's own stored voltage basis (maxW / maxA), so
 * the pair stays internally consistent instead of being scaled by two
 * different routes.
 */
export const scalePanelPower = (
  spec: { maxW: number; maxA: number },
  blackLevelW: number,
  brightnessPercent: number,
  contentPercent: number,
): ScaledPanelPower => {
  const designW = panelWattsAt({ maxW: spec.maxW, blackLevelW, brightnessPercent, contentPercent: 100 });
  const typicalW = panelWattsAt({ maxW: spec.maxW, blackLevelW, brightnessPercent, contentPercent });
  const ampsPerWatt = spec.maxW > 0 ? spec.maxA / spec.maxW : 0;
  return {
    maxW: designW,
    maxA: designW * ampsPerWatt,
    avgW: typicalW,
    avgA: typicalW * ampsPerWatt,
  };
};

/** The black level to start from for a panel whose own figure has not been measured yet. */
export const defaultBlackLevelW = (maxW: number): number =>
  Math.round(Math.max(0, maxW) * BARCO_EXAMPLE_BLACK_LEVEL_RATIO * 100) / 100;

/**
 * What the old industry rule of thumb would have said, for comparison in the
 * UI - Barco's article exists to argue against it, and seeing both is the
 * quickest way to understand why the numbers moved.
 */
export const ruleOfThumbWatts = (maxW: number): number => Math.max(0, maxW) / 3;
