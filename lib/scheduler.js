const axios = require("axios");
const ical = require("node-ical");
const { DateTime } = require("luxon");
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");

const TIMEZONE = process.env.TIMEZONE || "Asia/Tokyo";
const START_HOUR = Number(process.env.START_HOUR || 9);
const END_HOUR = Number(process.env.END_HOUR || 21);
const ORGANIZATION_NAME = process.env.ORGANIZATION_NAME || "団体名";
const HOLIDAY_ICS_URL =
  process.env.HOLIDAY_ICS_URL ||
  "https://calendar.google.com/calendar/ical/ja.japanese%23holiday%40group.v.calendar.google.com/public/basic.ics";
const EXCLUDE_HOLIDAYS = !["false", "0", "no"].includes(
  String(process.env.EXCLUDE_HOLIDAYS || "true").toLowerCase()
);
const VENUES_FILE = path.join(process.cwd(), "config", "venues.json");
const TEMPLATE_FILE = path.join(process.cwd(), "団体名_施設名_予定表.xlsx");

function shouldPersistExportCopy() {
  if (["false", "0", "no"].includes(String(process.env.SAVE_EXPORT_COPY || "").toLowerCase())) {
    return false;
  }

  // Serverless runtimes usually do not allow writing project directories.
  if (process.env.VERCEL || process.env.AWS_LAMBDA_FUNCTION_NAME || process.env.NETLIFY) {
    return false;
  }

  return true;
}

function parseEmbedUrls() {
  const raw = process.env.CALENDAR_EMBED_URLS || "";
  return raw
    .split(",")
    .map((url) => url.trim())
    .filter(Boolean);
}

function defaultVenuesFromEnv() {
  const urls = parseEmbedUrls();
  return urls.map((url, index) => ({
    id: `gym-${index + 1}`,
    name: `体育馆${index + 1}`,
    embedUrls: [url],
  }));
}

function loadVenues() {
  if (fs.existsSync(VENUES_FILE)) {
    const raw = fs.readFileSync(VENUES_FILE, "utf8");
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) {
      return parsed
        .filter((venue) => venue && venue.id && venue.name)
        .map((venue) => ({
          id: String(venue.id),
          name: String(venue.name),
          embedUrls: Array.isArray(venue.embedUrls) ? venue.embedUrls.filter(Boolean).map(String) : [],
        }));
    }
  }

  return defaultVenuesFromEnv();
}

function embedToIcalUrl(embedUrl) {
  const parsed = new URL(embedUrl);
  const src = parsed.searchParams.get("src");
  if (!src) {
    throw new Error(`Invalid Google Calendar embed URL (missing src): ${embedUrl}`);
  }

  return `https://calendar.google.com/calendar/ical/${encodeURIComponent(src)}/public/basic.ics`;
}

async function fetchCalendarIcs(icalUrl) {
  const response = await axios.get(icalUrl, {
    timeout: 12000,
    responseType: "text",
  });

  return ical.sync.parseICS(response.data);
}

function eventToInterval(event, dayStart, dayEnd) {
  if (!event || event.type !== "VEVENT" || !event.start || !event.end) {
    return null;
  }

  if (String(event.status || "").toUpperCase() === "CANCELLED") {
    return null;
  }

  const eventStart = DateTime.fromJSDate(event.start, { zone: TIMEZONE });
  const eventEnd = DateTime.fromJSDate(event.end, { zone: TIMEZONE });

  if (eventEnd <= dayStart || eventStart >= dayEnd) {
    return null;
  }

  const clippedStart = eventStart < dayStart ? dayStart : eventStart;
  const clippedEnd = eventEnd > dayEnd ? dayEnd : eventEnd;

  if (clippedEnd <= clippedStart) {
    return null;
  }

  return {
    start: clippedStart,
    end: clippedEnd,
  };
}

function recurringEventIntervalsForDay(event, dayStart, dayEnd) {
  if (!event || !event.rrule || !event.start || !event.end) {
    return [];
  }

  const durationMillis = event.end.getTime() - event.start.getTime();
  if (durationMillis <= 0) {
    return [];
  }

  let occurrences = [];
  try {
    occurrences = event.rrule.between(dayStart.toJSDate(), dayEnd.toJSDate(), true) || [];
  } catch {
    return [];
  }

  const intervals = [];
  for (const occurrenceStartDate of occurrences) {
    const occurrenceStart = DateTime.fromJSDate(occurrenceStartDate, { zone: TIMEZONE });
    const occurrenceEnd = occurrenceStart.plus({ milliseconds: durationMillis });

    const interval = eventToInterval(
      {
        type: "VEVENT",
        status: event.status,
        start: occurrenceStart.toJSDate(),
        end: occurrenceEnd.toJSDate(),
      },
      dayStart,
      dayEnd
    );

    if (interval) {
      intervals.push(interval);
    }
  }

  return intervals;
}

function eventIntervalsForDay(event, dayStart, dayEnd) {
  if (event?.rrule) {
    return recurringEventIntervalsForDay(event, dayStart, dayEnd);
  }

  const interval = eventToInterval(event, dayStart, dayEnd);
  return interval ? [interval] : [];
}

function eventToBooking(event, dayStart, dayEnd) {
  const interval = eventToInterval(event, dayStart, dayEnd);
  if (!interval) {
    return null;
  }

  return {
    start: interval.start,
    end: interval.end,
    summary: String(event.summary || event.description || "（无标题）").trim() || "（无标题）",
  };
}

function recurringEventBookingsForDay(event, dayStart, dayEnd) {
  if (!event || !event.rrule || !event.start || !event.end) {
    return [];
  }

  const durationMillis = event.end.getTime() - event.start.getTime();
  if (durationMillis <= 0) {
    return [];
  }

  let occurrences = [];
  try {
    occurrences = event.rrule.between(dayStart.toJSDate(), dayEnd.toJSDate(), true) || [];
  } catch {
    return [];
  }

  const bookings = [];
  for (const occurrenceStartDate of occurrences) {
    const occurrenceStart = DateTime.fromJSDate(occurrenceStartDate, { zone: TIMEZONE });
    const occurrenceEnd = occurrenceStart.plus({ milliseconds: durationMillis });
    const booking = eventToBooking(
      {
        ...event,
        type: "VEVENT",
        start: occurrenceStart.toJSDate(),
        end: occurrenceEnd.toJSDate(),
      },
      dayStart,
      dayEnd
    );
    if (booking) {
      bookings.push(booking);
    }
  }

  return bookings;
}

function eventBookingsForDay(event, dayStart, dayEnd) {
  if (event?.rrule) {
    return recurringEventBookingsForDay(event, dayStart, dayEnd);
  }

  const booking = eventToBooking(event, dayStart, dayEnd);
  return booking ? [booking] : [];
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

function collectEvents(parsedIcs) {
  return Object.values(parsedIcs).filter((item) => item && item.type === "VEVENT");
}

function holidayNamesForDay(events, currentDay) {
  const dayStart = currentDay.startOf("day");
  const dayEnd = dayStart.plus({ days: 1 });
  const names = [];

  for (const event of events) {
    if (!event || event.type !== "VEVENT" || !event.start || !event.end) {
      continue;
    }
    if (String(event.status || "").toUpperCase() === "CANCELLED") {
      continue;
    }

    const eventStart = DateTime.fromJSDate(event.start, { zone: TIMEZONE });
    const eventEnd = DateTime.fromJSDate(event.end, { zone: TIMEZONE });
    if (eventEnd <= dayStart || eventStart >= dayEnd) {
      continue;
    }

    const summary = String(event.summary || "祝日").trim() || "祝日";
    if (!names.includes(summary)) {
      names.push(summary);
    }
  }

  return names;
}

async function calculateAvailability(venueId, date, days) {
  const baseDate = DateTime.fromISO(date, { zone: TIMEZONE });
  if (!baseDate.isValid) {
    throw new Error("Invalid date format. Use YYYY-MM-DD");
  }

  const venues = loadVenues();
  if (venues.length === 0) {
    throw new Error("No venues configured. Add config/venues.json or CALENDAR_EMBED_URLS");
  }

  const venue = venues.find((item) => item.id === venueId);
  if (!venue) {
    throw new Error("Invalid venueId");
  }

  if (!Array.isArray(venue.embedUrls) || venue.embedUrls.length === 0) {
    throw new Error(`Facility '${venue.name}' is not configured yet. Please add calendar embed URL(s) in config/venues.json`);
  }

  const calendarUrls = venue.embedUrls.map(embedToIcalUrl);
  const calendarsData = await Promise.all(calendarUrls.map((url) => fetchCalendarIcs(url)));
  const calendarEventsList = calendarsData.map((icsData) => collectEvents(icsData));

  let holidayEvents = [];
  if (EXCLUDE_HOLIDAYS) {
    try {
      const holidayIcs = await fetchCalendarIcs(HOLIDAY_ICS_URL);
      holidayEvents = collectEvents(holidayIcs);
    } catch {
      holidayEvents = [];
    }
  }

  const result = [];

  for (let dayOffset = 0; dayOffset < days; dayOffset += 1) {
    const currentDay = baseDate.plus({ days: dayOffset });
    const dayStart = currentDay.set({ hour: START_HOUR, minute: 0, second: 0, millisecond: 0 });
    const dayEnd = currentDay.set({ hour: END_HOUR, minute: 0, second: 0, millisecond: 0 });
    const holidayNames = EXCLUDE_HOLIDAYS ? holidayNamesForDay(holidayEvents, currentDay) : [];
    const isHoliday = holidayNames.length > 0;

    if (isHoliday) {
      result.push({
        date: currentDay.toISODate(),
        timezone: TIMEZONE,
        window: {
          start: `${String(START_HOUR).padStart(2, "0")}:00`,
          end: `${String(END_HOUR).padStart(2, "0")}:00`,
        },
        isHoliday: true,
        holidayNames,
        commonFree: [],
      });
      continue;
    }

    const freeByCalendar = calendarEventsList.map((events) => {
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
      isHoliday: false,
      holidayNames: [],
      commonFree: formatIntervals(commonFree),
    });
  }

  return {
    venue: {
      id: venue.id,
      name: venue.name,
    },
    calendars: venue.embedUrls,
    range: {
      startDate: baseDate.toISODate(),
      days,
    },
    data: result,
    note: "Google Calendar must be publicly accessible for /public/basic.ics to work.",
  };
}

async function calculateBookingsBoard(date, days) {
  const baseDate = DateTime.fromISO(date, { zone: TIMEZONE });
  if (!baseDate.isValid) {
    throw new Error("Invalid date format. Use YYYY-MM-DD");
  }

  const venues = loadVenues();
  if (venues.length === 0) {
    throw new Error("No venues configured. Add config/venues.json or CALENDAR_EMBED_URLS");
  }

  const readyVenues = venues.filter((venue) => Array.isArray(venue.embedUrls) && venue.embedUrls.length > 0);
  const notReadyVenues = venues
    .filter((venue) => !Array.isArray(venue.embedUrls) || venue.embedUrls.length === 0)
    .map((venue) => ({ id: venue.id, name: venue.name, ready: false }));

  const venueEventsMap = new Map();
  await Promise.all(
    readyVenues.map(async (venue) => {
      const calendarUrls = venue.embedUrls.map(embedToIcalUrl);
      const calendarsData = await Promise.all(calendarUrls.map((url) => fetchCalendarIcs(url)));
      const events = calendarsData.flatMap((icsData) => collectEvents(icsData));
      venueEventsMap.set(venue.id, events);
    })
  );

  const data = [];
  for (let dayOffset = 0; dayOffset < days; dayOffset += 1) {
    const currentDay = baseDate.plus({ days: dayOffset });
    const dayStart = currentDay.set({ hour: START_HOUR, minute: 0, second: 0, millisecond: 0 });
    const dayEnd = currentDay.set({ hour: END_HOUR, minute: 0, second: 0, millisecond: 0 });

    const venuesForDay = readyVenues.map((venue) => {
      const events = venueEventsMap.get(venue.id) || [];
      const bookings = events
        .flatMap((event) => eventBookingsForDay(event, dayStart, dayEnd))
        .sort((a, b) => a.start.toMillis() - b.start.toMillis())
        .map((booking) => ({
          start: booking.start.toFormat("HH:mm"),
          end: booking.end.toFormat("HH:mm"),
          summary: booking.summary,
        }));

      return {
        id: venue.id,
        name: venue.name,
        ready: true,
        bookings,
      };
    });

    data.push({
      date: currentDay.toISODate(),
      venues: [...venuesForDay, ...notReadyVenues],
    });
  }

  return {
    range: {
      startDate: baseDate.toISODate(),
      days,
      endDate: baseDate.plus({ days: days - 1 }).toISODate(),
    },
    data,
  };
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
  if (!Number.isInteger(value) || value < 1 || value > 12) {
    return null;
  }

  return value;
}

function setTemplateMonthHeader(sheet, excelRow, monthDate) {
  setCellValue(sheet, `B${excelRow}`, monthDate.year);
  setCellValue(sheet, `C${excelRow}`, `${monthDate.month} 月`);
}

function parseWeekdayNumber(text) {
  const value = String(text || "");
  if (value.includes("月")) {
    return 1;
  }
  if (value.includes("火")) {
    return 2;
  }
  if (value.includes("水")) {
    return 3;
  }
  if (value.includes("木")) {
    return 4;
  }
  if (value.includes("金")) {
    return 5;
  }
  if (value.includes("土")) {
    return 6;
  }
  if (value.includes("日")) {
    return 7;
  }
  return null;
}

function toExcelDateSerial(dateTime) {
  const utcMillis = Date.UTC(dateTime.year, dateTime.month - 1, dateTime.day);
  return Math.round(utcMillis / 86400000 + 25569);
}

function setCellNumber(sheet, address, value) {
  if (!sheet[address]) {
    sheet[address] = { t: "n", v: Number(value) };
    return;
  }

  sheet[address].t = "n";
  sheet[address].v = Number(value);
  delete sheet[address].w;
}

function fillTemplateCalendarDates(sheet, monthHeaderRow, monthDate) {
  const monthHeaderIndex = monthHeaderRow - 1;
  const weekdayRow = monthHeaderIndex + 1;
  const firstColumnWeekday = parseWeekdayNumber(sheet[XLSX.utils.encode_cell({ r: weekdayRow, c: 1 })]?.v);
  if (!firstColumnWeekday) {
    return;
  }

  const firstOfMonth = monthDate.startOf("month");
  const offset = (firstOfMonth.weekday - firstColumnWeekday + 7) % 7;
  const firstCalendarDate = firstOfMonth.minus({ days: offset });

  for (let week = 0; week < 6; week += 1) {
    const dateRow = weekdayRow + 1 + week * 2;
    for (let dow = 0; dow < 7; dow += 1) {
      const dateValue = firstCalendarDate.plus({ days: week * 7 + dow });
      const cell = XLSX.utils.encode_cell({ r: dateRow, c: 1 + dow });
      const existing = sheet[cell];
      if (!templateCellToISODate(existing)) {
        continue;
      }
      setCellNumber(sheet, cell, toExcelDateSerial(dateValue));
    }
  }
}

function findWeekdayHeaderRows(sheet) {
  const ref = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const rows = [];

  for (let r = ref.s.r; r <= ref.e.r; r += 1) {
    let weekdayCount = 0;
    for (let c = 1; c <= 7; c += 1) {
      const cellText = sheet[XLSX.utils.encode_cell({ r, c })]?.v;
      if (parseWeekdayNumber(cellText)) {
        weekdayCount += 1;
      }
    }

    if (weekdayCount >= 6) {
      rows.push(r);
    }
  }

  return rows;
}

function templateCellToISODate(cell) {
  if (!cell) {
    return null;
  }

  if (typeof cell.v === "number") {
    const parsed = XLSX.SSF.parse_date_code(cell.v);
    if (!parsed || !parsed.y || !parsed.m || !parsed.d) {
      return null;
    }

    return DateTime.fromObject(
      { year: parsed.y, month: parsed.m, day: parsed.d },
      { zone: TIMEZONE }
    ).toISODate();
  }

  if (cell.v instanceof Date) {
    return DateTime.fromJSDate(cell.v, { zone: TIMEZONE }).toISODate();
  }

  if (typeof cell.v === "string") {
    const fromIso = DateTime.fromISO(cell.v, { zone: TIMEZONE });
    if (fromIso.isValid) {
      return fromIso.toISODate();
    }
  }

  return null;
}

function buildTemplateEntryMap(sheet) {
  const weekdayRows = findWeekdayHeaderRows(sheet);
  const entryMap = {};

  for (const headerRow of weekdayRows) {
    for (let week = 0; week < 6; week += 1) {
      const dateRow = headerRow + 1 + week * 2;
      const writeRow = dateRow + 1;

      for (let dow = 0; dow < 7; dow += 1) {
        const dateCell = sheet[XLSX.utils.encode_cell({ r: dateRow, c: 1 + dow })];
        const dateValue = templateCellToISODate(dateCell);
        if (!dateValue) {
          continue;
        }
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
    throw new Error("Template missing Sheet2");
  }

  const sheet = workbook.Sheets[sheet2Name];
  if (!sheet) {
    throw new Error("Template Sheet2 not found");
  }

  const sorted = [...selected].sort((a, b) => String(a.date).localeCompare(String(b.date)));
  const validDates = sorted
    .map((item) => DateTime.fromISO(String(item.date || ""), { zone: TIMEZONE }))
    .filter((dt) => dt.isValid);

  const uniqueMonthKeys = [...new Set(validDates.map((dt) => dt.toFormat("yyyy-LL")))].sort();
  const monthDates = uniqueMonthKeys.map((key) => DateTime.fromFormat(`${key}-01`, "yyyy-LL-dd", { zone: TIMEZONE }));
  const fallbackMonth = DateTime.now().setZone(TIMEZONE).startOf("month");
  const month1 = monthDates[0] || fallbackMonth;
  const month2 = monthDates[1] || month1.plus({ months: 1 });

  // Build org/applicant text only including non-empty fields
  const textLines = [];
  if (organizationName && organizationName.trim()) {
    textLines.push(`使用団体名：${organizationName.trim()}`);
  }
  if (applicantName && applicantName.trim()) {
    textLines.push(`申請者氏名：${applicantName.trim()}`);
  }
  const orgApplicantText = textLines.join("\r\n");

  setTemplateMonthHeader(sheet, 3, month1);
  setTemplateMonthHeader(sheet, 18, month2);
  fillTemplateCalendarDates(sheet, 3, month1);
  fillTemplateCalendarDates(sheet, 18, month2);
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

async function exportXlsxFromSelection({ selected, venueName, organizationName, applicantName }) {
  if (!Array.isArray(selected) || selected.length === 0) {
    throw new Error("No selected slots. Please choose at least one slot.");
  }
  const normalizedOrganizationName = String(organizationName || "").trim();
  const normalizedApplicantName = String(applicantName || "").trim();
  if (!normalizedOrganizationName && !normalizedApplicantName) {
    throw new Error("Missing organizationName/applicantName. Please input at least one.");
  }

  if (!fs.existsSync(TEMPLATE_FILE)) {
    throw new Error("Missing 団体名_施設名_予定表.xlsx in project root");
  }

  const workbook = XLSX.readFile(TEMPLATE_FILE, { cellStyles: true });
  fillTemplateSheet2(
    workbook,
    selected,
    venueName,
    normalizedOrganizationName || ORGANIZATION_NAME,
    normalizedApplicantName
  );

  const fileName = xlsxFileName(normalizedOrganizationName || ORGANIZATION_NAME, venueName || "Unknown Venue");
  const fileBuffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

  if (shouldPersistExportCopy()) {
    try {
      const exportDir = path.join(process.cwd(), "exports");
      fs.mkdirSync(exportDir, { recursive: true });
      fs.writeFileSync(path.join(exportDir, fileName), fileBuffer);
    } catch {
      // Ignore local persistence errors so API download still succeeds.
    }
  }

  return {
    fileName,
    fileBuffer,
  };
}

module.exports = {
  loadVenues,
  calculateAvailability,
  calculateBookingsBoard,
  exportXlsxFromSelection,
  ORGANIZATION_NAME,
};
