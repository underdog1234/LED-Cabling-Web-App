import { describe, it, expect } from "vitest";
import { fitTextToWidth, pdfSafeText, STOCK_CATALOG } from "./App";

// jsPDF's built-in fonts are WinAnsi only. Hand one a character outside that
// and the line comes out as spaced-out nonsense - or the export throws and
// takes the whole report with it. Every string in the PDF crosses one
// boundary, and this is what guards it.

describe("pdfSafeText", () => {
  it("writes three-phase the way jsPDF can set it", () => {
    // The adaptor's catalogue name printed as "3 2 A  3 |  P D L ..." before
    // this existed.
    expect(pdfSafeText("32A 3Φ PDL - 32A 3Φ Ceeform Power Adaptor"))
      .toBe("32A 3Ph PDL - 32A 3Ph Ceeform Power Adaptor");
  });

  it("keeps every catalogue name settable", () => {
    Object.values(STOCK_CATALOG).forEach((item) => {
      const safe = pdfSafeText(item.name);
      expect({ code: item.code, outside: safe.match(/[^ -ÿ]/g) }).toEqual({ code: item.code, outside: null });
      expect(safe.length).toBeGreaterThan(0);
    });
  });

  it("leaves ordinary text, accents and symbols WinAnsi already covers alone", () => {
    expect(pdfSafeText("MG9 LED Panel")).toBe("MG9 LED Panel");
    expect(pdfSafeText("Café 200mm x 100mm")).toBe("Cafe 200mm x 100mm"); // combining mark dropped
    expect(pdfSafeText("Café °C ½")).toBe("Café °C ½");
  });

  it("turns the marks the app itself prints into something readable", () => {
    expect(pdfSafeText("↓ 1 → 2")).toBe("v 1 -> 2");
    expect(pdfSafeText("1920 × 1080")).toBe("1920 x 1080");
    expect(pdfSafeText("‘quoted’ — dash")).toBe("'quoted' - dash");
  });

  it("drops what it cannot render rather than corrupting the line around it", () => {
    // A screen name someone typed an emoji into still prints the words.
    expect(pdfSafeText("Stage Left \u{1f3a4} Tower")).toBe("Stage Left Tower");
    expect(pdfSafeText("\u{1f50c} P1 (3)")).toBe("P1 (3)");
  });
});

// Headings and the summary block are drawn at fixed positions in two columns,
// so a long project name had nowhere to go: it ran off the right edge of the
// page in the title and across the second column in the block below. Reported
// from a real report - "7582 Investment Services Group Limited - Shaw and
// Partners Appreciation Event".

// A stand-in for the PDF engine: every glyph one unit wide per point of size.
const measure = (text: string, size: number) => text.length * size * 0.5;

describe("fitTextToWidth", () => {
  it("leaves text that already fits exactly as it is", () => {
    expect(fitTextToWidth("Short title", 200, 18, 9, measure)).toEqual({ text: "Short title", size: 18 });
  });

  it("shrinks before it cuts", () => {
    // 40 chars needs 360 units at 18pt; at 12pt it is 240, which fits 250.
    const text = "x".repeat(40);
    const fitted = fitTextToWidth(text, 250, 18, 9, measure);
    expect(fitted.text).toBe(text);
    expect(fitted.size).toBeLessThan(18);
    expect(measure(fitted.text, fitted.size)).toBeLessThanOrEqual(250);
  });

  it("never shrinks past the floor, and cuts what is left over", () => {
    const text = "7582 Investment Services Group Limited - Shaw and Partners Appreciation Event";
    const fitted = fitTextToWidth(text, 120, 18, 9, measure);
    expect(fitted.size).toBe(9);
    expect(fitted.text.endsWith("...")).toBe(true);
    expect(fitted.text.startsWith("7582 Investment")).toBe(true);
    expect(measure(fitted.text, fitted.size)).toBeLessThanOrEqual(120);
  });

  it("keeps whatever it can rather than giving up on an impossible width", () => {
    const fitted = fitTextToWidth("A long name", 2, 18, 9, measure);
    expect(fitted.text).toBe("...");
    expect(fitted.size).toBe(9);
  });

  it("has nothing to do with an empty line or a width of zero", () => {
    expect(fitTextToWidth("", 100, 18, 9, measure)).toEqual({ text: "", size: 18 });
    expect(fitTextToWidth("anything", 0, 18, 9, measure)).toEqual({ text: "anything", size: 18 });
  });

  it("fits the reported title inside an A4 landscape page", () => {
    // 277mm of usable width, and helvetica bold at 18pt runs about 0.55mm per
    // character - close enough to show the real title needed shrinking.
    const real = (text: string, size: number) => text.length * size * 0.55;
    const title = "7582 Investment Services Group Limited - Shaw and Partners Appreciation Event - Panel Layout - Back View";
    expect(real(title, 18)).toBeGreaterThan(277); // this is the overflow
    const fitted = fitTextToWidth(title, 277, 18, 9, real);
    expect(real(fitted.text, fitted.size)).toBeLessThanOrEqual(277);
  });
});
