const { DateTime } = require("luxon");
const { calculateAvailability } = require("../../lib/scheduler");

const TIMEZONE = process.env.TIMEZONE || "Asia/Tokyo";

export default async function handler(req, res) {
  if (req.method !== "GET") {
    return res.status(405).json({ error: "Method Not Allowed" });
  }

  try {
    const venueId = String(req.query.venueId || "").trim();
    const date = req.query.date;
    const days = Math.max(1, Math.min(14, Number(req.query.days || 1)));

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

    const data = await calculateAvailability(venueId, date, days);
    return res.status(200).json(data);
  } catch (error) {
    return res.status(500).json({
      error: "Failed to calculate availability",
      message: error.message,
    });
  }
}
