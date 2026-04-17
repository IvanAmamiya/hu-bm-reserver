const venueSelect = document.getElementById("venueSelect");
const organizationInput = document.getElementById("organizationInput");
const applicantInput = document.getElementById("applicantInput");
const dateInput = document.getElementById("dateInput");
const daysInput = document.getElementById("daysInput");
const loadBtn = document.getElementById("loadBtn");
const exportBtn = document.getElementById("exportBtn");
const statusEl = document.getElementById("status");
const slotsEl = document.getElementById("slots");

const API_BASE_URL = String(window.APP_CONFIG?.API_BASE_URL || "").trim().replace(/\/$/, "");

function apiUrl(path) {
  return API_BASE_URL ? `${API_BASE_URL}${path}` : path;
}

let currentVenue = null;

function todayISO() {
  return new Date().toISOString().slice(0, 10);
}

dateInput.value = todayISO();

function setStatus(text) {
  statusEl.textContent = text;
}

function selectedVenueName() {
  const option = venueSelect.options[venueSelect.selectedIndex];
  return option ? option.text : "";
}

function readableErrorMessage(error, fallback) {
  const raw = String(error?.message || "");
  if (/Failed to fetch/i.test(raw) || /NetworkError/i.test(raw)) {
    return `${fallback}：后端服务不可达。请确认 API 正在运行，并检查 config.js 的 API_BASE_URL。`;
  }
  return `${fallback}：${raw || "未知错误"}`;
}

async function loadVenues() {
  setStatus("正在加载体育馆列表...");
  loadBtn.disabled = true;
  exportBtn.disabled = true;
  venueSelect.innerHTML = "";

  try {
    const response = await fetch(apiUrl("/api/venues"));
    if (!response.ok) {
      throw new Error("加载体育馆失败");
    }

    const data = await response.json();
    const venues = Array.isArray(data.venues) ? data.venues : [];
    if (venues.length === 0) {
      throw new Error("未配置体育馆");
    }

    venues.forEach((venue, index) => {
      const option = document.createElement("option");
      option.value = venue.id;
      option.textContent = venue.name;
      if (index === 0) {
        option.selected = true;
      }
      venueSelect.appendChild(option);
    });

    currentVenue = {
      id: venueSelect.value,
      name: selectedVenueName(),
    };
    loadBtn.disabled = false;
    setStatus("请选择日期并查询空余时间");
  } catch (error) {
    setStatus(readableErrorMessage(error, "体育馆加载失败"));
  }
}

function renderSlots(days) {
  slotsEl.innerHTML = "";
  let selectableCount = 0;

  days.forEach((day, dayIndex) => {
    const dayCard = document.createElement("article");
    dayCard.className = "day-card";

    const title = document.createElement("div");
    title.className = "day-title";
    title.textContent = `${day.date} (${day.window.start}-${day.window.end})`;
    dayCard.appendChild(title);

    if (!day.commonFree || day.commonFree.length === 0) {
      const empty = document.createElement("div");
      empty.className = "empty";
      empty.textContent = "无共同空余时间";
      dayCard.appendChild(empty);
    } else {
      const list = document.createElement("div");
      list.className = "slot-list";

      day.commonFree.forEach((slot, slotIndex) => {
        selectableCount += 1;
        const item = document.createElement("label");
        item.className = "slot-item";

        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.className = "slot-checkbox";
        checkbox.value = JSON.stringify({
          date: day.date,
          start: slot.start,
          end: slot.end,
        });
        checkbox.id = `slot-${dayIndex}-${slotIndex}`;

        const text = document.createElement("span");
        text.textContent = `${slot.start}-${slot.end}`;

        item.appendChild(checkbox);
        item.appendChild(text);
        list.appendChild(item);
      });

      dayCard.appendChild(list);
    }

    slotsEl.appendChild(dayCard);
  });

  exportBtn.disabled = selectableCount === 0;
}

async function loadAvailability() {
  const venueId = venueSelect.value;
  const date = dateInput.value;
  const days = Math.max(1, Math.min(14, Number(daysInput.value || 7)));

  if (!venueId) {
    setStatus("请先选择体育馆");
    return;
  }
  if (!date) {
    setStatus("请先选择日期");
    return;
  }

  const venueName = selectedVenueName();
  currentVenue = { id: venueId, name: venueName };
  setStatus(`正在查询 ${venueName} 空余时间...`);
  loadBtn.disabled = true;
  exportBtn.disabled = true;

  try {
    const response = await fetch(
      apiUrl(`/api/availability?venueId=${encodeURIComponent(venueId)}&date=${encodeURIComponent(date)}&days=${days}`)
    );
    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(errorData.error || "查询失败");
    }

    const data = await response.json();
    lastAvailabilityData = Array.isArray(data.data) ? data.data : [];
    renderSlots(lastAvailabilityData);
    setStatus(`查询完成：${data.venue?.name || venueName}，共 ${lastAvailabilityData.length} 天，勾选后可导出`);
  } catch (error) {
    slotsEl.innerHTML = "";
    setStatus(readableErrorMessage(error, "查询失败"));
  } finally {
    loadBtn.disabled = false;
  }
}

function selectedSlots() {
  return Array.from(document.querySelectorAll(".slot-checkbox:checked")).map((node) =>
    JSON.parse(node.value)
  );
}

function safeFilePart(value, fallback) {
  const text = String(value || "").trim();
  if (!text) {
    return fallback;
  }
  return text.replace(/[\\/:*?"<>|]/g, "_");
}

function xlsxFileName(organizationName, facilityName) {
  const org = safeFilePart(organizationName, "団体名");
  const facility = safeFilePart(facilityName, "施設名");
  return `${org}_${facility}_予定表.xlsx`;
}

function setCellValue(sheet, address, value) {
  if (!sheet[address]) {
    sheet[address] = { t: "s", v: String(value) };
    return;
  }
  sheet[address].t = "s";
  sheet[address].v = String(value);
  delete sheet[address].w;
}

function parseMonthNumber(text) {
  const match = String(text || "").match(/(\d{1,2})/);
  if (!match) {
    return null;
  }
  const value = Number(match[1]);
  return Number.isInteger(value) && value >= 1 && value <= 12 ? value : null;
}

function setTemplateMonthHeader(sheet, excelRow, monthDate) {
  setCellValue(sheet, `B${excelRow}`, monthDate.year);
  setCellValue(sheet, `C${excelRow}`, `${monthDate.month} 月`);
}

function findWeekdayHeaderRows(sheet) {
  const ref = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const rows = [];

  for (let r = ref.s.r; r <= ref.e.r; r += 1) {
    const sunday = sheet[XLSX.utils.encode_cell({ r, c: 1 })]?.v;
    const saturday = sheet[XLSX.utils.encode_cell({ r, c: 7 })]?.v;
    if (String(sunday || "").includes("日曜日") && String(saturday || "").includes("土曜日")) {
      rows.push(r);
    }
  }

  return rows;
}

function buildTemplateEntryMap(sheet) {
  const weekdayRows = findWeekdayHeaderRows(sheet);
  const entryMap = {};

  for (const headerRow of weekdayRows) {
    const infoRow = headerRow - 1;
    const year = Number(sheet[XLSX.utils.encode_cell({ r: infoRow, c: 1 })]?.v);
    const month = parseMonthNumber(sheet[XLSX.utils.encode_cell({ r: infoRow, c: 2 })]?.v);
    if (!year || !month) {
      continue;
    }

    const firstOfMonth = DateTime.fromObject({ year, month, day: 1 }, { zone: TIMEZONE });
    const sundayOffset = firstOfMonth.weekday % 7;
    const firstCalendarDate = firstOfMonth.minus({ days: sundayOffset });

    for (let week = 0; week < 6; week += 1) {
      const dateRow = headerRow + 1 + week * 2;
      const writeRow = dateRow + 1;

      for (let dow = 0; dow < 7; dow += 1) {
        const dateValue = firstCalendarDate.plus({ days: week * 7 + dow }).toISODate();
        const writeCell = XLSX.utils.encode_cell({ r: writeRow, c: 1 + dow });
        entryMap[dateValue] = writeCell;
      }
    }
  }

  return entryMap;
}

function fillTemplateSheet2(workbook, selected, venueName, organizationName, applicantName) {
  const sheet2Name = workbook.SheetNames[1];
  if (!sheet2Name) {
    throw new Error("模板缺少 Sheet2");
  }

  const sheet = workbook.Sheets[sheet2Name];
  const sorted = [...selected].sort((a, b) => String(a.date).localeCompare(String(b.date)));
  const validDates = sorted
    .map((item) => DateTime.fromISO(String(item.date || ""), { zone: TIMEZONE }))
    .filter((dt) => dt.isValid);

  const uniqueMonthKeys = [...new Set(validDates.map((dt) => dt.toFormat("yyyy-LL")))].sort();
  const monthDates = uniqueMonthKeys.map((key) => DateTime.fromFormat(`${key}-01`, "yyyy-LL-dd", { zone: TIMEZONE }));
  const fallbackMonth = DateTime.now().setZone(TIMEZONE).startOf("month");
  const month1 = monthDates[0] || fallbackMonth;
  const month2 = monthDates[1] || month1.plus({ months: 1 });

  const orgApplicantText = `使用団体名：${organizationName}\r\n申請者氏名：${applicantName}`;

  setTemplateMonthHeader(sheet, 3, month1);
  setTemplateMonthHeader(sheet, 18, month2);
  setCellValue(sheet, "E3", `利用希望施設：${venueName}`);
  setCellValue(sheet, "G3", orgApplicantText);
  setCellValue(sheet, "E18", `利用希望施設：${venueName}`);
  setCellValue(sheet, "G18", orgApplicantText);

  const entryMap = buildTemplateEntryMap(sheet);
  Object.values(entryMap).forEach((cell) => setCellValue(sheet, cell, ""));

  const grouped = new Map();
  for (const item of selected) {
    const date = String(item.date || "").trim();
    const slot = `${String(item.start || "").trim()}-${String(item.end || "").trim()}`;
    if (!date || slot === "-") {
      continue;
    }
    if (!grouped.has(date)) {
      grouped.set(date, []);
    }
    grouped.get(date).push(slot);
  }

  for (const [date, slots] of grouped.entries()) {
    const targetCell = entryMap[date];
    if (!targetCell) {
      continue;
    }
    setCellValue(sheet, targetCell, slots.join("\n"));
  }
}

async function exportXlsx() {
  const selected = selectedSlots();
  const organizationName = (organizationInput.value || "").trim();
  const applicantName = (applicantInput.value || "").trim();
  const venueName = currentVenue?.name || selectedVenueName();

  if (!organizationName) {
    setStatus("请先填写团体名（或个人名）");
    return;
  }
  if (!applicantName) {
    setStatus("请先填写申请者姓名");
    return;
  }
  if (!venueName) {
    setStatus("请先选择体育馆");
    return;
  }
  if (selected.length === 0) {
    setStatus("请至少勾选一个时段再导出");
    return;
  }

  exportBtn.disabled = true;
  setStatus("正在生成 xlsx...");

  try {
    const response = await fetch(apiUrl("/api/export-xlsx"), {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        selected,
        venueId: currentVenue?.id,
        venueName,
        organizationName,
        applicantName,
      }),
    });

    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(errorData.error || "导出失败");
    }

    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;

    const disposition = response.headers.get("Content-Disposition") || "";
    const utf8NameMatch = disposition.match(/filename\*=UTF-8''([^;]+)/i);
    const plainNameMatch = disposition.match(/filename="?([^";]+)"?/i);
    const fileName = utf8NameMatch
      ? decodeURIComponent(utf8NameMatch[1])
      : plainNameMatch
        ? plainNameMatch[1]
        : "団体名_施設名_予定表.xlsx";

    anchor.download = fileName;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    window.URL.revokeObjectURL(url);

    setStatus(`导出成功：${fileName}`);
  } catch (error) {
    setStatus(readableErrorMessage(error, "导出失败"));
  } finally {
    exportBtn.disabled = false;
  }
}

loadBtn.addEventListener("click", loadAvailability);
exportBtn.addEventListener("click", exportXlsx);
venueSelect.addEventListener("change", () => {
  currentVenue = {
    id: venueSelect.value,
    name: selectedVenueName(),
  };
});

loadVenues();
