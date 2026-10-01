import { describe, it, expect } from "vitest";
import { peakUsage } from "./rentman/availabilityPeak";
import type { EquipmentAvailabilityProject } from "./rentman/rentmanClient";

// "Can I have 320 panels from the 4th to the 9th?" is a question about the
// worst DAY in that window, not about the sum of everything that touches it.
// These are the cases that were being answered wrongly.

const booking = (
  projectName: string,
  quantity: number,
  planPeriodStart: string,
  planPeriodEnd: string,
): EquipmentAvailabilityProject => ({
  projectName,
  projectNumber: projectName,
  status: "Confirmed",
  quantity,
  planPeriodStart,
  planPeriodEnd,
});

describe("peakUsage", () => {
  it("does not add up jobs that never clash", () => {
    // The bug: 100 out on the 4th and 100 out on the 8th was reported as 200
    // unavailable, though the first lot is back on the shelf days before the
    // second goes out.
    const result = peakUsage(
      [booking("Job A", 100, "2026-10-04", "2026-10-05"), booking("Job B", 100, "2026-10-08", "2026-10-09")],
      "2026-10-01",
      "2026-10-10",
    );
    expect(result.peak).toBe(100);
    expect(result.total).toBe(200);
    expect(result.overstated).toBe(true);
  });

  it("adds up jobs that do clash, on the day they clash", () => {
    const result = peakUsage(
      [booking("Job A", 100, "2026-10-04", "2026-10-07"), booking("Job B", 60, "2026-10-06", "2026-10-09")],
      "2026-10-01",
      "2026-10-10",
    );
    expect(result).toMatchObject({ peak: 160, peakStart: "2026-10-06", peakEnd: "2026-10-07", total: 160, overstated: false });
    expect(result.peakBookings.map((b) => b.projectName).sort()).toEqual(["Job A", "Job B"]);
  });

  it("counts two jobs sharing a single day as clashing", () => {
    // One derigs in the morning, the other rigs in the afternoon: the same
    // panels cannot do both, and a planner wants that flagged.
    const result = peakUsage(
      [booking("Out", 40, "2026-10-01", "2026-10-05"), booking("In", 40, "2026-10-05", "2026-10-08")],
      "2026-10-01",
      "2026-10-10",
    );
    expect(result).toMatchObject({ peak: 80, peakStart: "2026-10-05", peakEnd: "2026-10-05" });
  });

  it("only counts the days a booking shares with this job", () => {
    // A month-long sub-rental that happens to cover this window counts once,
    // and a booking outside the window does not count at all.
    const result = peakUsage(
      [booking("Long run", 50, "2026-09-01", "2026-10-31"), booking("Before", 999, "2026-08-01", "2026-08-09")],
      "2026-10-04",
      "2026-10-06",
    );
    expect(result).toMatchObject({ peak: 50, peakStart: "2026-10-04", peakEnd: "2026-10-06" });
    expect(result.peakBookings.map((b) => b.projectName)).toEqual(["Long run"]);
  });

  it("reports the whole run the peak holds for, not just its first day", () => {
    const result = peakUsage(
      [booking("A", 10, "2026-10-01", "2026-10-10"), booking("B", 10, "2026-10-03", "2026-10-07")],
      "2026-10-01",
      "2026-10-31",
    );
    expect(result).toMatchObject({ peak: 20, peakStart: "2026-10-03", peakEnd: "2026-10-07" });
  });

  it("holds the peak across a change that does not lower it", () => {
    // B leaves on the same day C arrives, so 20 runs straight through.
    const result = peakUsage(
      [
        booking("A", 10, "2026-10-01", "2026-10-10"),
        booking("B", 10, "2026-10-02", "2026-10-04"),
        booking("C", 10, "2026-10-05", "2026-10-06"),
      ],
      "2026-10-01",
      "2026-10-31",
    );
    expect(result).toMatchObject({ peak: 20, peakStart: "2026-10-02", peakEnd: "2026-10-06" });
  });

  it("says nothing is out when nothing overlaps the window", () => {
    expect(peakUsage([booking("Elsewhere", 80, "2026-01-01", "2026-01-05")], "2026-10-01", "2026-10-10"))
      .toMatchObject({ peak: 0, peakStart: null, peakEnd: null, overstated: false });
    expect(peakUsage([], "2026-10-01", "2026-10-10")).toMatchObject({ peak: 0, total: 0 });
  });

  it("survives dates it cannot read, and quantities it should ignore", () => {
    const result = peakUsage(
      [
        booking("No dates", 50, "", ""),
        booking("Zero", 0, "2026-10-04", "2026-10-05"),
        booking("Real", 25, "2026-10-04", "2026-10-05"),
      ],
      "2026-10-01",
      "2026-10-10",
    );
    expect(result.peak).toBe(25);
  });

  it("treats a booking that ends before it starts as a single day", () => {
    // Bad data should not quietly remove demand from the answer.
    const result = peakUsage([booking("Backwards", 30, "2026-10-05", "2026-10-01")], "2026-10-01", "2026-10-10");
    expect(result).toMatchObject({ peak: 30, peakStart: "2026-10-05", peakEnd: "2026-10-05" });
  });

  it("reads a plan period with a time on it", () => {
    const result = peakUsage(
      [booking("Morning out", 12, "2026-10-04T08:00:00Z", "2026-10-04T18:00:00Z")],
      "2026-10-01",
      "2026-10-10",
    );
    expect(result).toMatchObject({ peak: 12, peakStart: "2026-10-04", peakEnd: "2026-10-04" });
  });
});
