# Gym Scheduler Frontend (GitHub Pages)

This project is now fully frontend-deployable.

Features:

- Reads Google Calendar public iCal data in browser.
- Calculates free slots (09:00-21:00, Asia/Tokyo).
- Lets user choose venue/date/slots.
- Fills template Sheet2 (`入力用`) and downloads `.xlsx` directly in browser.
- No backend API required for normal usage.

## Frontend Files

- App entry: `docs/index.html`
- Frontend logic: `docs/app.js`
- Frontend config: `docs/config.js`
- Export template: `docs/template.xlsx`

`docs/` is the single deployment source for GitHub Pages.
When changing behavior for production, edit `docs/*` first.

## Usage

1. Open deployed page.
2. Fill:
	 - 团体名（或个人名）
	 - 申请者姓名
3. Select venue/date/days and click 查询空余时间.
4. Check slots and click 导出到 XLSX.

Downloaded filename format:

`団体名（個人名）_施設名_予定表.xlsx`

Template fill behavior:

- Uses `template.xlsx`.
- Writes to Sheet2 (`入力用`).
- Fills:
	- `利用希望施設`
	- `使用団体名`
	- `申請者氏名`
- Month blocks auto-align to selected months.

## Config

Edit `docs/config.js`:

```js
window.APP_CONFIG = {
	TIMEZONE: "Asia/Tokyo",
	START_HOUR: 9,
	END_HOUR: 21,
	CORS_PROXY: "https://api.allorigins.win/raw?url={url}",
	TEMPLATE_PATH: "./template.xlsx",
	VENUES: [
		{
			id: "gym-group-calendar",
			name: "东体育馆",
			embedUrls: ["https://calendar.google.com/calendar/u/0/embed?..."],
		}
	]
};
```

Notes:

- `VENUES` is now maintained in frontend config.
- `CORS_PROXY` is needed because browser direct access to Google iCal may be blocked by CORS.

## GitHub Pages Deploy

This repo includes workflow: `.github/workflows/deploy-pages.yml`.

1. Push to `main` (or `master`).
2. In GitHub Settings -> Pages, choose Source = GitHub Actions.
3. Workflow deploys `docs/` automatically.

Quick commands:

```bash
git add .
git commit -m "refactor: github pages ready"
git push origin main
```

After push:

1. Open GitHub Actions and confirm `Deploy GitHub Pages` is green.
2. Open repository Settings -> Pages and verify Source = GitHub Actions.
3. Visit `https://<your-github-username>.github.io/<your-repo-name>/`.

Open site URL pattern:

`https://<your-github-username>.github.io/<your-repo-name>/`

### Troubleshooting: only README is shown

If your deployed site shows README text instead of the app:

1. Confirm you are opening the GitHub Pages URL, not the repository homepage.
2. In repository Settings -> Pages, set Source to GitHub Actions.
3. Check Actions tab and ensure the workflow `Deploy GitHub Pages` succeeded.
4. This repo includes root `index.html` and `404.html` redirects to `docs/` as fallback.

## Limitations

- Frontend-only mode cannot save files to server disk; files are browser-downloaded only.
- CORS proxy availability/rate limits may affect calendar fetch stability.

If you see `查询失败：Failed to fetch`, it is usually a CORS/proxy issue. Update `docs/config.js` -> `CORS_PROXY` and retry.
