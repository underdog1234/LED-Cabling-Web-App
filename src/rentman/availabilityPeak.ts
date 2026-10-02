import type { EquipmentAvailabilityProject } from "./rentmanClient";

// ---------------------------------------------------------------------------
// How much of an item is really unavailable while this job is out.
//
// The question is "can I have 320 panels from the 4th to the 9th", and the
// answer is NOT the sum of every booking that touches those dates. Two jobs
// of 100 panels, one on the 4th and one on the 8th, are 100 panels of demand
// twice - never 200 at once - and the 100 that came back on the 5th are on the
// shelf again for the 8th. Adding them made the tool report a shortage that
// does not exist, and the fix is to ask a different question: on the worst
// single DAY inside the window, how much is out?
//
// Days rather than hours on purpose. Kit is booked by the day, and two jobs
// sharing a day are competing for the same panels even when one derigs in the
// morning and the other rigs in the afternoon - a planner wants that flagged,
// not smoothed away by a clock.
//
// A CANCELLED job is not demand at all. Rentman keeps the equipment lines on a
// cancelled job rather than deleting them, so they keep coming back from the
// date-overlap query and used to be counted like any other booking - kit held
// against a job that is not happening. Those are listed but never counted (see
// isCancelledStatus).
// ---------------------------------------------------------------------------

const DAY_MS = 86_400_000;

/**
 * Whether a booking's status means the job is off.
 *
 * Rentman has no status on the project itself, only on its subproject(s), and
 * this account's own status list (GET /statuses) is: Pending, Canceled,
 * Confirmed, Prepped, On location, Returned, Inquiry, Concept. Note the API's
 * `name` is the US spelling "Canceled" while Rentman's UI shows the display
 * name "Cancelled" - matching either exactly would miss half of it, which is
 * why this normalises and tests the stem.
 *
 * Only cancelled is excluded. A Pending or Inquiry job is kit somebody is
 * seriously expecting to send out, and quietly freeing it would be the same
 * class of mistake in the other direction.
 */
export const isCancelledStatus = (status: string | null | undefined): boolean =>
  (status ?? "").toLowerCase().replace(/[^a-z]/g, "").startsWith("cancel");

/** The day a timestamp falls on, as a whole number of days since the epoch. Null when it cannot be read. */
export const dayNumber = (value: string | null | undefined): number | null => {
  if (!value) return null;
  const ms = Date.parse(value);
  if (!Number.isFinite(ms)) return null;
  return Math.floor(ms / DAY_MS);
};

const dayToIso = (day: number): string => new Date(day * DAY_MS).toISOString().slice(0, 10);

export type PeakUsage = {
  /** The most of this item booked out on any ONE day inside the window - what is really unavailable. */
  peak: number;
  /** First and last day of the run where usage sits at that peak, as YYYY-MM-DD. Null when nothing overlaps. */
  peakStart: string | null;
  peakEnd: string | null;
  /** The bookings that are out on the peak day - the ones to argue with. */
  peakBookings: EquipmentAvailabilityProject[];
  /** Every overlapping booking that COUNTS, added together - which is what this used to report. Kept so the two can be shown side by side. */
  total: number;
  /** True when adding the counted bookings up overstates the real demand - i.e. they do not all clash. */
  overstated: boolean;
  /** Bookings left out because their job is cancelled. Listed so they can be seen, never counted. */
  cancelled: EquipmentAvailabilityProject[];
  /** What those cancelled bookings would have added to the sum had they been counted. */
  cancelledQuantity: number;
};

/**
 * Peak concurrent demand for one item across [from, to], both days inclusive.
 *
 * A booking is counted on every day from its plan-period start to its end, and
 * only on the days that fall inside the window - a job that runs for a month
 * either side of this one still only takes up the days it shares with it.
 */
export const peakUsage = (
  bookings: EquipmentAvailabilityProject[],
  from: string,
  to: string,
): PeakUsage => {
  // Cancelled jobs come back from the date query like any other booking -
  // Rentman keeps their equipment lines - so they are split off here, before
  // anything is counted, rather than being subtracted again further down.
  const cancelled = bookings.filter((booking) => isCancelledStatus(booking.status));
  const counted = bookings.filter((booking) => !isCancelledStatus(booking.status));
  const quantityOf = (list: EquipmentAvailabilityProject[]) =>
    list.reduce((sum, booking) => sum + (Number(booking.quantity) || 0), 0);
  const total = quantityOf(counted);
  const cancelledQuantity = quantityOf(cancelled);
  const empty: PeakUsage = {
    peak: 0,
    peakStart: null,
    peakEnd: null,
    peakBookings: [],
    total,
    overstated: total > 0,
    cancelled,
    cancelledQuantity,
  };
  const windowStart = dayNumber(from);
  const windowEnd = dayNumber(to);
  if (windowStart === null || windowEnd === null || windowEnd < windowStart) return empty;

  // Each booking clipped to the window, as a run of whole days.
  const spans = counted
    .map((booking) => {
      const start = dayNumber(booking.planPeriodStart);
      const end = dayNumber(booking.planPeriodEnd);
      const quantity = Number(booking.quantity) || 0;
      if (start === null || end === null || quantity <= 0) return null;
      const first = Math.max(start, windowStart);
      // A booking whose end reads before its start is a single day, not a
      // negative one - bad data should not silently drop demand.
      const last = Math.min(Math.max(end, start), windowEnd);
      if (last < first) return null;
      return { booking, quantity, first, last };
    })
    .filter((span): span is { booking: EquipmentAvailabilityProject; quantity: number; first: number; last: number } => span !== null);
  if (!spans.length) return { ...empty, overstated: false };

  // Sweep the day boundaries: usage only changes where a booking starts or
  // the day after one ends, so there is no need to walk a long window day by
  // day just to find the maximum.
  const deltas = new Map<number, number>();
  spans.forEach((span) => {
    deltas.set(span.first, (deltas.get(span.first) ?? 0) + span.quantity);
    deltas.set(span.last + 1, (deltas.get(span.last + 1) ?? 0) - span.quantity);
  });
  const boundaries = [...deltas.keys()].sort((a, b) => a - b);
  let running = 0;
  let peak = 0;
  let peakStartDay: number | null = null;
  let peakEndDay: number | null = null;
  for (let i = 0; i < boundaries.length; i += 1) {
    running += deltas.get(boundaries[i]) ?? 0;
    if (running <= 0) continue;
    const dayFrom = boundaries[i];
    // This level holds until the next boundary, or to the end of the window.
    const dayTo = (i + 1 < boundaries.length ? boundaries[i + 1] - 1 : windowEnd);
    if (dayTo < windowStart || dayFrom > windowEnd) continue;
    if (running > peak) {
      peak = running;
      peakStartDay = Math.max(dayFrom, windowStart);
      peakEndDay = Math.min(dayTo, windowEnd);
    } else if (running === peak && peakEndDay !== null && dayFrom === peakEndDay + 1) {
      // The same level carrying on across a boundary is one run, not two.
      peakEndDay = Math.min(dayTo, windowEnd);
    }
  }
  if (peakStartDay === null || peakEndDay === null) return { ...empty, overstated: false };

  const peakBookings = spans
    .filter((span) => span.first <= peakEndDay! && span.last >= peakStartDay!)
    .map((span) => span.booking);
  return {
    peak,
    peakStart: dayToIso(peakStartDay),
    peakEnd: dayToIso(peakEndDay),
    peakBookings,
    total,
    overstated: total > peak,
    cancelled,
    cancelledQuantity,
  };
};
