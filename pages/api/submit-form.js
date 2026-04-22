import crypto from "crypto";

const FORM_SUBMIT_WEBHOOK_URL = process.env.FORM_SUBMIT_WEBHOOK_URL || "";
const FORM_SUBMIT_WEBHOOK_TOKEN = process.env.FORM_SUBMIT_WEBHOOK_TOKEN || "";

function normalizeSelected(rawSelected) {
  if (!Array.isArray(rawSelected)) {
    return [];
  }

  return rawSelected
    .map((item) => ({
      date: String(item?.date || "").trim(),
      start: String(item?.start || "").trim(),
      end: String(item?.end || "").trim(),
    }))
    .filter((item) => item.date && item.start && item.end);
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method Not Allowed" });
  }

  if (!FORM_SUBMIT_WEBHOOK_URL) {
    return res.status(501).json({
      error: "FORM_SUBMIT_WEBHOOK_URL is not configured",
      message: "Please set FORM_SUBMIT_WEBHOOK_URL in .env to enable auto submission.",
    });
  }

  try {
    const body = req.body || {};
    const venueId = String(body.venueId || "").trim();
    const venueName = String(body.venueName || "").trim() || "Unknown Venue";
    const organizationName = String(body.organizationName || "").trim();
    const applicantName = String(body.applicantName || "").trim();
    const selected = normalizeSelected(body.selected);

    if (!venueId) {
      return res.status(400).json({ error: "Missing venueId" });
    }

    if (!organizationName && !applicantName) {
      return res.status(400).json({ error: "Missing organizationName/applicantName. Please input at least one." });
    }

    if (selected.length === 0) {
      return res.status(400).json({ error: "No selected slots. Please choose at least one slot." });
    }

    const requestId = crypto.randomUUID();
    const payload = {
      requestId,
      submittedAt: new Date().toISOString(),
      source: "hu-bm-reserver",
      venue: {
        id: venueId,
        name: venueName,
      },
      applicant: {
        organizationName,
        applicantName,
      },
      selected,
      selectedCount: selected.length,
    };

    const headers = {
      "Content-Type": "application/json",
    };
    if (FORM_SUBMIT_WEBHOOK_TOKEN) {
      headers.Authorization = `Bearer ${FORM_SUBMIT_WEBHOOK_TOKEN}`;
    }

    const webhookResponse = await fetch(FORM_SUBMIT_WEBHOOK_URL, {
      method: "POST",
      headers,
      body: JSON.stringify(payload),
    });

    if (!webhookResponse.ok) {
      const text = await webhookResponse.text();
      return res.status(502).json({
        error: "Webhook submission failed",
        message: `Webhook responded with ${webhookResponse.status}`,
        detail: text.slice(0, 600),
      });
    }

    return res.status(200).json({
      ok: true,
      requestId,
      selectedCount: selected.length,
    });
  } catch (error) {
    return res.status(500).json({
      error: "Failed to auto submit form",
      message: error.message,
    });
  }
}
