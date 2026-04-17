const { DateTime } = luxon;

const venueSelect = document.getElementById("venueSelect");
const organizationInput = document.getElementById("organizationInput");
const applicantInput = document.getElementById("applicantInput");
const dateInput = document.getElementById("dateInput");
const daysInput = document.getElementById("daysInput");
const loadBtn = document.getElementById("loadBtn");
const exportBtn = document.getElementById("exportBtn");
const statusEl = document.getElementById("status");
const slotsEl = document.getElementById("slots");

const TIMEZONE = window.APP_CONFIG?.TIMEZONE || "Asia/Tokyo";
const START_HOUR = Number(window.APP_CONFIG?.START_HOUR ?? 9);
const END_HOUR = Number(window.APP_CONFIG?.END_HOUR ?? 21);
const CORS_PROXY = String(window.APP_CONFIG?.CORS_PROXY || "").trim();
const TEMPLATE_PATH = String(window.APP_CONFIG?.TEMPLATE_PATH || "./template.xlsx").trim();
const VENUES = Array.isArray(window.APP_CONFIG?.VENUES) ? window.APP_CONFIG.VENUES : [];
const PROXY_CANDIDATES = [
  CORS_PROXY,
  "https://api.allorigins.win/raw?url={url}",
  "https://cors.isomorphic-git.org/{url}",
].filter(Boolean);

let currentVenue = null;
let lastAvailabilityData = [];

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

function embedToIcalUrl(embedUrl) {
  const parsed = new URL(embedUrl);
  const src = parsed.searchParams.get("src");
  if (!src) {
    throw new Error(`无效日历地址（缺少 src）：${embedUrl}`);
  }

  return `https://calendar.google.com/calendar/ical/${encodeURIComponent(src)}/public/basic.ics`;
}

function proxiedUrl(url, proxyPattern) {
  if (!proxyPattern) {
    return url;
  }

  if (proxyPattern.includes("{url}")) {
    if (proxyPattern.includes("url={url}")) {
      return proxyPattern.replace("{url}", encodeURIComponent(url));
    }
    return proxyPattern.replace("{url}", url);
  }

  return `${proxyPattern}${encodeURIComponent(url)}`;
}

async function fetchCalendarText(icalUrl) {
  const tried = [];

  for (const proxy of PROXY_CANDIDATES) {
    const targetUrl = proxiedUrl(icalUrl, proxy);
    tried.push(proxy);

    try {
      const response = await fetch(targetUrl);
      if (!response.ok) {
        continue;
      }
      return response.text();
    } catch (error) {
      // Try next proxy endpoint.
    }
  }

  const primary = tried[0] || "(none)";
  throw new Error(`日历拉取失败（可能是跨域/CORS限制）。已尝试代理：${primary}`);
}

function readableErrorMessage(error, fallback) {
  const raw = String(error?.message || "");
  if (/Failed to fetch/i.test(raw) || /NetworkError/i.test(raw)) {
    return `${fallback}：网络或跨域访问失败。请稍后重试，或更换 docs/config.js 中的 CORS_PROXY。`;
  }
  return `${fallback}：${raw || "未知错误"}`;
}

function parseIcsEvents(icsText) {
  const jcal = ICAL.parse(icsText);
  const component = new ICAL.Component(jcal);
  const vevents = component.getAllSubcomponents("vevent") || [];
  return vevents.map((item) => new ICAL.Event(item));
}

function mergeIntervals(intervals) {
  if (intervals.length === 0) {
    return [];
  }

  const sorted = [...intervals].sort((a, b) => a.start.toMillis() - b.start.toMillis());
  const merged = [sorted[0]];

  for (let i = 1; i < sorted.length; i += 1) {
    const current = sorted[i];
    const last = merged[merged.length - 1];

    if (current.start <= last.end) {
      if (current.end > last.end) {
        last.end = current.end;
      }
    } else {
      merged.push({ ...current });
    }
  }

  return merged;
}

function invertIntervals(busyIntervals, dayStart, dayEnd) {
  if (busyIntervals.length === 0) {
    return [{ start: dayStart, end: dayEnd }];
  }

  const free = [];
  let cursor = dayStart;

  for (const busy of busyIntervals) {
    if (busy.start > cursor) {
      free.push({ start: cursor, end: busy.start });
    }
    if (busy.end > cursor) {
      cursor = busy.end;
    }
  }

  if (cursor < dayEnd) {
    free.push({ start: cursor, end: dayEnd });
  }

  return free;
}

function intersectTwoIntervalLists(listA, listB) {
  const result = [];
  let i = 0;
  let j = 0;

  while (i < listA.length && j < listB.length) {
    const start = listA[i].start > listB[j].start ? listA[i].start : listB[j].start;
    const end = listA[i].end < listB[j].end ? listA[i].end : listB[j].end;

    if (end > start) {
      result.push({ start, end });
    }

    if (listA[i].end < listB[j].end) {
      i += 1;
    } else {
      j += 1;
    }
  }

  return result;
}

function formatIntervals(intervals) {
  return intervals.map((interval) => ({
    start: interval.start.toFormat("HH:mm"),
    end: interval.end.toFormat("HH:mm"),
  }));
}

function clipInterval(start, end, dayStart, dayEnd) {
  if (end <= dayStart || start >= dayEnd) {
    return null;
  }

  const clippedStart = start < dayStart ? dayStart : start;
  const clippedEnd = end > dayEnd ? dayEnd : end;

  if (clippedEnd <= clippedStart) {
    return null;
  }

  return { start: clippedStart, end: clippedEnd };
}

function eventDurationMillis(event) {
  if (event.endDate && event.startDate) {
    return Math.max(0, event.endDate.toJSDate().getTime() - event.startDate.toJSDate().getTime());
  }
  return 0;
}

function recurringIntervalsForDay(event, dayStart, dayEnd) {
  const intervals = [];
  const durationMillis = eventDurationMillis(event);
  if (durationMillis <= 0) {
    return intervals;
  }

  const iterator = event.iterator();
  let count = 0;
  while (count < 2000) {
    const occurrence = iterator.next();
    if (!occurrence) {
      break;
    }
    count += 1;

    const occurrenceStart = DateTime.fromJSDate(occurrence.toJSDate(), { zone: TIMEZONE });
    if (occurrenceStart >= dayEnd) {
      break;
    }

    const occurrenceEnd = occurrenceStart.plus({ milliseconds: durationMillis });
    const interval = clipInterval(occurrenceStart, occurrenceEnd, dayStart, dayEnd);
    if (interval) {
      intervals.push(interval);
    }
  }

  return intervals;
}

function eventIntervalsForDay(event, dayStart, dayEnd) {
  const status = String(event?.component?.getFirstPropertyValue("status") || "").toUpperCase();
  if (status === "CANCELLED") {
    return [];
  }

  if (event.isRecurring()) {
    return recurringIntervalsForDay(event, dayStart, dayEnd);
  }

  const start = DateTime.fromJSDate(event.startDate.toJSDate(), { zone: TIMEZONE });
  const end = DateTime.fromJSDate(event.endDate.toJSDate(), { zone: TIMEZONE });
  const interval = clipInterval(start, end, dayStart, dayEnd);
  return interval ? [interval] : [];
}

async function calculateAvailability(venue, baseDateISO, days) {
  const baseDate = DateTime.fromISO(baseDateISO, { zone: TIMEZONE });
  if (!baseDate.isValid) {
    throw new Error("日期格式错误，请使用 YYYY-MM-DD");
  }

  const icsTexts = await Promise.all(venue.embedUrls.map((url) => fetchCalendarText(embedToIcalUrl(url))));
  const calendarsEvents = icsTexts.map((text) => parseIcsEvents(text));

  const result = [];

  for (let dayOffset = 0; dayOffset < days; dayOffset += 1) {
    const currentDay = baseDate.plus({ days: dayOffset });
    const dayStart = currentDay.set({ hour: START_HOUR, minute: 0, second: 0, millisecond: 0 });
    const dayEnd = currentDay.set({ hour: END_HOUR, minute: 0, second: 0, millisecond: 0 });

    const freeByCalendar = calendarsEvents.map((events) => {
      const busyIntervals = events.flatMap((event) => eventIntervalsForDay(event, dayStart, dayEnd));
      const mergedBusy = mergeIntervals(busyIntervals);
      return invertIntervals(mergedBusy, dayStart, dayEnd);
    });

    let commonFree = freeByCalendar[0] || [];
    for (let i = 1; i < freeByCalendar.length; i += 1) {
      commonFree = intersectTwoIntervalLists(commonFree, freeByCalendar[i]);
    }

    result.push({
      date: currentDay.toISODate(),
      timezone: TIMEZONE,
      window: {
        start: `${String(START_HOUR).padStart(2, "0")}:00`,
        end: `${String(END_HOUR).padStart(2, "0")}:00`,
      },
      commonFree: formatIntervals(commonFree),
    });
  }

  return result;
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

function loadVenues() {
  venueSelect.innerHTML = "";
  if (VENUES.length === 0) {
    throw new Error("未配置场馆，请在 config.js 中设置 APP_CONFIG.VENUES");
  }

  VENUES.forEach((venue, index) => {
    const option = document.createElement("option");
    option.value = venue.id;
    option.textContent = venue.name;
    if (index === 0) {
      option.selected = true;
    }
    venueSelect.appendChild(option);
  });

  currentVenue = VENUES.find((item) => item.id === venueSelect.value) || null;
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

  const venue = VENUES.find((item) => item.id === venueId);
  if (!venue) {
    setStatus("场馆配置不存在");
    return;
  }

  currentVenue = venue;
  setStatus(`正在查询 ${venue.name} 空余时间...`);
  loadBtn.disabled = true;
  exportBtn.disabled = true;

  try {
    lastAvailabilityData = await calculateAvailability(venue, date, days);
    renderSlots(lastAvailabilityData);
    setStatus(`查询完成：${venue.name}，共 ${lastAvailabilityData.length} 天，勾选后可导出`);
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
    const templateResponse = await fetch(TEMPLATE_PATH);
    if (!templateResponse.ok) {
      throw new Error("读取模板文件失败，请确认 template.xlsx 已部署");
    }

    const arrayBuffer = await templateResponse.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array", cellStyles: true });
    fillTemplateSheet2(workbook, selected, venueName, organizationName, applicantName);

    const fileName = xlsxFileName(organizationName, venueName);
    const output = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const blob = new Blob([output], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);

    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);

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
  currentVenue = VENUES.find((item) => item.id === venueSelect.value) || null;
});

try {
  loadVenues();
  setStatus("请选择日期并查询空余时间");
} catch (error) {
  setStatus(`初始化失败：${error.message}`);
}
