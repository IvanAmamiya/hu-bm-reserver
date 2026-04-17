const { loadVenues } = require("../../lib/scheduler");

export default function handler(req, res) {
  try {
    const venues = loadVenues().map((venue) => ({
      id: venue.id,
      name: venue.name,
      calendarsCount: venue.embedUrls.length,
    }));

    res.status(200).json({ venues });
  } catch (error) {
    res.status(500).json({
      error: "Failed to load venues",
      message: error.message,
    });
  }
}
