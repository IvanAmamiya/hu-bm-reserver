const { exportXlsxFromSelection, ORGANIZATION_NAME } = require("../../lib/scheduler");

export const config = {
  api: {
    bodyParser: {
      sizeLimit: "10mb",
    },
  },
};

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method Not Allowed" });
  }

  try {
    console.log("[API导出] POST /api/export-xlsx 起序");
    console.log("[API导出] Content-Type:", req.headers["content-type"]);
    const body = req.body || {};
    console.log("[API导出] body 类型:", typeof body, "keys:", Object.keys(body));
    
    const rawSelected = body.selected;
    console.log("[API导出] rawSelected 类型:", typeof rawSelected, "长度:", Array.isArray(rawSelected) ? rawSelected.length : (typeof rawSelected === "string" ? rawSelected.length : "N/A"));
    
    let selected = Array.isArray(rawSelected) ? rawSelected : [];
    if (!Array.isArray(rawSelected) && typeof rawSelected === "string") {
      try {
        console.log("[API导出] 开始解析 JSON 字符串");
        const parsed = JSON.parse(rawSelected);
        console.log("[API导出] 解析成功、数据个数:", parsed.length);
        if (Array.isArray(parsed)) {
          selected = parsed;
        }
      } catch (e) {
        console.log("[API导出] JSON 解析失败:", e.message);
        selected = [];
      }
    }

    const venueName = String(body.venueName || "").trim() || "Unknown Venue";
    const organizationName = String(body.organizationName || "").trim() || ORGANIZATION_NAME;
    const applicantName = String(body.applicantName || "").trim() || "";

    console.log("[API导出] 导出参数:", {
      selectedCount: selected.length,
      venueName,
      organizationName,
      applicantName,
    });

    const { fileName, fileBuffer } = await exportXlsxFromSelection({
      selected,
      venueName,
      organizationName,
      applicantName,
    });

    console.log("[API导出] 导出成功，文件名:", fileName, "大小:", fileBuffer?.length, "bytes");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="schedule.xlsx"; filename*=UTF-8''${encodeURIComponent(fileName)}`
    );
    res.setHeader("X-Export-File-Encoded", encodeURIComponent(fileName));

    return res.status(200).send(fileBuffer);
  } catch (error) {
    console.log("[API导出] 错误：", error.message, "stack:", error.stack);
    return res.status(500).json({
      error: "Failed to export xlsx",
      message: error.message,
    });
  }
}
