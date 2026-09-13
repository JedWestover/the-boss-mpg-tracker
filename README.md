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
- Browser storage is device-specific. Add a hosted database later if you want syncing between devices.
