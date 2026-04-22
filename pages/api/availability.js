const { DateTime } = require("luxon");
const { calculateAvailability } = require("../../lib/scheduler");

const TIMEZONE = process.env.TIMEZONE || "Asia/Tokyo";
const ADVANCE_BOOKING_DAYS = Math.max(0, Number(process.env.ADVANCE_BOOKING_DAYS || 7));

function bookingWindowFromNow() {
  const today = DateTime.now().setZone(TIMEZONE).startOf("day");
  const earliestDate = today.plus({ days: ADVANCE_BOOKING_DAYS });
  const latestDate = today.plus({ months: 1 }).endOf("month").startOf("day");
  return { today, earliestDate, latestDate };
}

export default async function handler(req, res) {
  if (req.method !== "GET") {
    return res.status(405).json({ error: "Method Not Allowed" });
  }

  try {
    const venueId = String(req.query.venueId || "").trim();
    const date = req.query.date;
    const requestedDays = Math.max(1, Number(req.query.days || 1));

    if (!venueId) {
      return res.status(400).json({ error: "Missing query param: venueId" });
    }

    if (!date) {
      return res.status(400).json({ error: "Missing query param: date (format: YYYY-MM-DD)" });
    }

    const baseDate = DateTime.fromISO(date, { zone: TIMEZONE });
    if (!baseDate.isValid) {
      return res.status(400).json({ error: "Invalid date format. Use YYYY-MM-DD" });
    }

    const { today, earliestDate, latestDate } = bookingWindowFromNow();
    const requestDate = baseDate.startOf("day");
    if (requestDate < today) {
      return res.status(400).json({
        error: `Cannot book dates before today (${today.toISODate()})`,
      });
    }

    if (requestDate < earliestDate) {
      return res.status(400).json({
        error: `Earliest bookable date is ${earliestDate.toISODate()} (advance ${ADVANCE_BOOKING_DAYS} days)`,
      });
    }

    if (requestDate > latestDate) {
      return res.status(400).json({
        error: `Latest bookable start date is ${latestDate.toISODate()}`,
      });
    }

    const maxDaysByWindow = Math.floor(latestDate.diff(requestDate, "days").days) + 1;
    const days = Math.max(1, Math.min(requestedDays, maxDaysByWindow));

    const data = await calculateAvailability(venueId, date, days);
    return res.status(200).json({
      ...data,
      bookingRules: {
        advanceBookingDays: ADVANCE_BOOKING_DAYS,
        earliestDate: earliestDate.toISODate(),
        latestDate: latestDate.toISODate(),
      },
      range: {
        ...data.range,
        requestedDays,
        appliedDays: days,
        endDate: requestDate.plus({ days: days - 1 }).toISODate(),
      },
    });
  } catch (error) {
    return res.status(500).json({
      error: "Failed to calculate availability",
      message: error.message,
    });
  }
}
