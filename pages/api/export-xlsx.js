const { exportXlsxFromSelection, ORGANIZATION_NAME } = require("../../lib/scheduler");

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method Not Allowed" });
  }

  try {
    const selected = Array.isArray(req.body?.selected) ? req.body.selected : [];
    const venueName = String(req.body?.venueName || "").trim() || "Unknown Venue";
    const organizationName = String(req.body?.organizationName || "").trim() || ORGANIZATION_NAME;
    const applicantName = String(req.body?.applicantName || "").trim() || "";

    const { fileName, fileBuffer } = await exportXlsxFromSelection({
      selected,
      venueName,
      organizationName,
      applicantName,
    });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="schedule.xlsx"; filename*=UTF-8''${encodeURIComponent(fileName)}`
    );
    res.setHeader("X-Export-File-Encoded", encodeURIComponent(fileName));

    return res.status(200).send(fileBuffer);
  } catch (error) {
    return res.status(500).json({
      error: "Failed to export xlsx",
      message: error.message,
    });
  }
}
