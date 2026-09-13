# The Boss MPG Tracker

A React + Vite app for tracking diesel and DEF fill-ups. The supplied spreadsheet history is preloaded the first time the app opens. Data is stored in the browser with localStorage.

## Run in VS Code

```bash
npm install
npm run dev
```

## GitHub

```bash
git init
git add .
git commit -m "Initial MPG tracker"
git branch -M main
git remote add origin YOUR_GITHUB_REPOSITORY_URL
git push -u origin main
```

## Vercel

Import the GitHub repository in Vercel. It will detect Vite automatically. Build command: `npm run build`; output directory: `dist`.

## Notes

- MPG is distance since the previous fill divided by gallons.
- DEF efficiency is shown as miles per DEF gallon.
- Import accepts this workbook layout or a simple sheet with Date, Odometer, Gallons, Cost, and Type columns.
- Export downloads a backup workbook.
- Browser storage is device-specific until OneDrive sync is connected.

## OneDrive sync

1. Register a single-page application in Microsoft Entra admin center.
2. Add the deployed app URL as a SPA redirect URI, for example `https://your-app.vercel.app`.
3. Grant Microsoft Graph delegated permission `Files.ReadWrite`.
4. Copy `.env.example` to `.env.local` and set `VITE_MICROSOFT_CLIENT_ID` to the application client ID.
5. Add the same `VITE_MICROSOFT_CLIENT_ID` as a Vercel environment variable and redeploy.

Use **Connect OneDrive** in the app to sign in. The tracker stores its backup as `the-boss-mpg-data.json` in the signed-in user's OneDrive root. Existing cloud data is loaded on sign-in; later saves sync automatically.
