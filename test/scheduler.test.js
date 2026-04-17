const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("path");
const fs = require("fs");
const axios = require("axios");
const ical = require("node-ical");

const { calculateAvailability } = require("../lib/scheduler");

const originalExistsSync = fs.existsSync;
const originalReadFileSync = fs.readFileSync;
const originalAxiosGet = axios.get;
const originalParseICS = ical.sync.parseICS;

const venuesFileSuffix = path.join("config", "venues.json");

function isVenuesFile(targetPath) {
  return String(targetPath).endsWith(venuesFileSuffix);
}

function mockVenues(venues) {
  fs.existsSync = (targetPath) => {
    if (isVenuesFile(targetPath)) {
      return true;
    }
    return originalExistsSync(targetPath);
  };

  fs.readFileSync = (targetPath, encoding) => {
    if (isVenuesFile(targetPath)) {
      return JSON.stringify(venues);
    }
    return originalReadFileSync(targetPath, encoding);
  };
}

function restoreAll() {
  fs.existsSync = originalExistsSync;
  fs.readFileSync = originalReadFileSync;
  axios.get = originalAxiosGet;
  ical.sync.parseICS = originalParseICS;
}

test.afterEach(() => {
  restoreAll();
});

test("calculateAvailability returns expected free slots for a normal event", async () => {
  mockVenues([
    {
      id: "test-venue",
      name: "Test Venue",
      embedUrls: ["https://calendar.google.com/calendar/embed?src=test%40example.com"],
    },
  ]);

  axios.get = async () => ({ data: "BEGIN:VCALENDAR" });
  ical.sync.parseICS = () => ({
    ev1: {
      type: "VEVENT",
      start: new Date("2026-04-20T10:00:00+09:00"),
      end: new Date("2026-04-20T12:00:00+09:00"),
    },
  });

  const result = await calculateAvailability("test-venue", "2026-04-20", 1);

  assert.equal(result.data.length, 1);
  assert.deepEqual(result.data[0].commonFree, [
    { start: "09:00", end: "10:00" },
    { start: "12:00", end: "21:00" },
  ]);
});

test("calculateAvailability handles recurring full-day busy event", async () => {
  mockVenues([
    {
      id: "test-venue",
      name: "Test Venue",
      embedUrls: ["https://calendar.google.com/calendar/embed?src=test%40example.com"],
    },
  ]);

  axios.get = async () => ({ data: "BEGIN:VCALENDAR" });
  ical.sync.parseICS = () => ({
    evRecurring: {
      type: "VEVENT",
      start: new Date("2026-04-01T09:00:00+09:00"),
      end: new Date("2026-04-01T21:00:00+09:00"),
      rrule: {
        between: () => [new Date("2026-04-20T09:00:00+09:00")],
      },
    },
  });

  const result = await calculateAvailability("test-venue", "2026-04-20", 1);

  assert.equal(result.data.length, 1);
  assert.deepEqual(result.data[0].commonFree, []);
});

test("calculateAvailability throws on invalid date format", async () => {
  mockVenues([
    {
      id: "test-venue",
      name: "Test Venue",
      embedUrls: ["https://calendar.google.com/calendar/embed?src=test%40example.com"],
    },
  ]);

  await assert.rejects(async () => {
    await calculateAvailability("test-venue", "2026/04/20", 1);
  }, /Invalid date format/);
});
