# Gym Scheduler (Next.js)

This project now uses Next.js with integrated backend APIs.

## Features

- Venue-first booking flow.
- Availability calculation from Google Calendar iCal feeds.
- Recurring event expansion (RRULE supported).
- Export selected slots into template Sheet2 (`入力用`) as `.xlsx`.
- Single app deployment (frontend + backend in one service).

## Tech Structure

- Page UI: `pages/index.js`
- API routes:
  - `pages/api/health.js`
  - `pages/api/venues.js`
  - `pages/api/availability.js`
  - `pages/api/export-xlsx.js`
- Shared business logic: `lib/scheduler.js`
- Venue config: `config/venues.json`
- Export template: `団体名_施設名_予定表.xlsx`

## Local Run

Node requirement:

- Use Node.js 20 LTS (recommended for stable Next.js build).

1. Install dependencies:

```bash
npm install
```

2. Optional env setup:

```bash
copy .env.example .env
```

3. Run dev server:

```bash
npm run dev
```

4. Open:

`http://localhost:3000`

## API

- `GET /api/health`
- `GET /api/venues`
- `GET /api/availability?venueId=<id>&date=YYYY-MM-DD&days=1`
- `POST /api/export-xlsx`
- `POST /api/submit-form`

Export request body:

```json
{
  "venueId": "gym-group-calendar",
  "venueName": "东体育馆",
  "organizationName": "团体名（或个人名）",
  "applicantName": "申请者姓名",
  "selected": [
    { "date": "2026-04-20", "start": "09:00", "end": "12:00" }
  ]
}
```

Auto submit request body (`/api/submit-form`):

```json
{
  "venueId": "gym-group-calendar",
  "venueName": "东体育馆",
  "organizationName": "团体名（或个人名）",
  "applicantName": "申请者姓名",
  "selected": [
    { "date": "2026-04-20", "start": "09:00", "end": "12:00" }
  ]
}
```

Environment setup for auto submit:

- `FORM_SUBMIT_WEBHOOK_URL`: your Power Automate / webhook endpoint
- `FORM_SUBMIT_WEBHOOK_TOKEN`: optional bearer token

## Deploy Notes

This is now a server app, not GitHub Pages static hosting.

Recommended targets:

- Vercel
- Railway
- Render
- VPS (Node 18+)

## Legacy Static Files

`docs/` and `public/` legacy static files still exist for reference, but production path should use Next.js routes.
