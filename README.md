# Calendar Common Free Time API + Web Selector

This app extracts free time from gym calendars and lets you choose by venue first, then select slots in a web page and export to XLSX.

## What it does

- Loads gym venues from `config/venues.json`.
- Converts venue Google Calendar embed URLs to public iCal URLs.
- Computes free slots for the selected venue.
- Daily window defaults to **09:00-21:00** in **Asia/Tokyo**.
- Provides a web page for selecting free slots.
- Exports selected slots into an `.xlsx` file.

## Quick start

1. Install dependencies:

```bash
npm install
```

2. Create env file:

```bash
copy .env.example .env
```

3. Start API:

```bash
npm start
```

## Web UI

Open:

`http://localhost:3000`

Flow:

1. Choose gym venue.
2. Choose start date and days.
3. Click "查询空余时间".
4. Select desired time slots.
5. Click "导出到 XLSX".

The file will be downloaded by browser and also saved on server in `exports/`.

## API

### Health check

`GET /health`

### Query common availability

`GET /api/availability?venueId=gym-group-calendar&date=2026-04-18&days=1`

- `venueId` (required): gym/venue id from `GET /api/venues`
- `date` (required): `YYYY-MM-DD`
- `days` (optional): number of days to compute, default `1`, max `14`

Example:

```bash
curl "http://localhost:3000/api/availability?venueId=gym-group-calendar&date=2026-04-18&days=3"
```

Response field `commonFree` shows free slots in `HH:mm` for the selected venue.

### List venues

`GET /api/venues`

### Export selected slots to xlsx

`POST /api/export-xlsx`

Body example:

```json
{
	"selected": [
		{ "date": "2026-04-18", "start": "10:00", "end": "11:30" },
		{ "date": "2026-04-19", "start": "14:00", "end": "15:00" }
	],
	"venueId": "gym-group-calendar",
	"venueName": "东体育馆",
	"organizationName": "团体名（或个人名）",
	"applicantName": "申请者姓名"
}
```

Returns an xlsx attachment for browser download, and also saves a copy to local `exports/` folder.

Export behavior:

- Uses `団体名_施設名_予定表.xlsx` in project root as template.
- Fills **Sheet2 (`入力用`)** with selected slots.
- Fills `使用団体名` and `申請者氏名` from export inputs.
- Saves output to `exports/`.
- Month blocks are aligned to selected slot months automatically.
- If selected slots span more than 2 months, extra months are returned in `droppedMonths`.

Filename format:

`団体名_施設名_予定表.xlsx`

- `団体名` comes from request body `organizationName` or env `ORGANIZATION_NAME`.
- `施設名` comes from request body `venueName` (selected venue name from web UI).

Response headers include:

- `Content-Disposition` with download filename.
- `X-Export-File-Encoded` for building local output path under `exports/`.

## Notes

- Venue source file: `config/venues.json`.
- Uses `https://calendar.google.com/calendar/ical/{calendarId}/public/basic.ics`.
- Calendars must be public or accessible through public iCal, otherwise the API cannot read events.
- If you need private calendars, switch to Google Calendar API with OAuth/service account.

## GitHub Pages Deployment

This repo now includes a static site in `docs/` and workflow at `.github/workflows/deploy-pages.yml`.

Steps:

1. Push this repository to GitHub (branch `main` or `master`).
2. In GitHub repository settings, ensure **Pages** source is **GitHub Actions**.
3. Edit `docs/config.js` and set:

```js
window.APP_CONFIG = {
	API_BASE_URL: "https://your-backend-domain.example.com"
};
```

4. Push changes again. GitHub Actions will deploy `docs/` to Pages automatically.

Important:

- GitHub Pages can host only frontend static files.
- API server (`src/server.js`) must run on a backend host (Render, Railway, VPS, etc.).
- CORS is enabled in server code so Pages frontend can call the API cross-origin.
