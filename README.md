# fin. — personal finance dashboard

A Next.js app that stores all your data **live in the Excel file** (`data/budget.xlsx`) and deploys to Vercel with one push.

## Project structure

```
fin-app/
├── pages/
│   ├── index.jsx          # The entire app UI (React)
│   └── api/
│       ├── budget.js      # GET/POST — reads & writes budget.xlsx
│       └── suggest.js     # POST — Claude AI analysis
├── data/
│   └── budget.xlsx        # ← YOUR EXCEL FILE. This is the source of truth.
├── package.json
├── next.config.js
└── vercel.json
```

## How the Excel sync works

Every edit you make in the UI (adding an expense, changing a budget line, etc.) is:
1. Saved to React state immediately
2. Written back to `data/budget.xlsx` via the `/api/budget` POST endpoint within ~1.5 seconds

The Excel file grows new sheets prefixed with `App - ` (e.g. `App - Budget`, `App - Expenses`). Your original sheets (`Master Cashflow`, `Budget Analysis`, etc.) are untouched.

## Local development

```bash
npm install
npm run dev
```

Open http://localhost:3000

## Deploy to Vercel

### 1. Push to GitHub
```bash
git init
git add .
git commit -m "fin. initial"
git remote add origin https://github.com/YOUR_USERNAME/fin-app.git
git push -u origin main
```

### 2. Import on Vercel
- Go to vercel.com → New Project → Import your repo
- Framework: **Next.js** (auto-detected)
- Root directory: leave as `/`
- Click Deploy

### 3. (Optional) AI Suggestions
Add your Anthropic key in Vercel → Project → Settings → Environment Variables:
```
ANTHROPIC_API_KEY = sk-ant-...
```

Then redeploy. The "Get AI analysis" button on the Suggestions tab will start working.

## ⚠️ Important: Excel on Vercel

Vercel's filesystem is **read-only in production** for the deployed container. This means:
- The app reads `budget.xlsx` on every GET request ✅
- **Writes to Excel only work in local dev** (the file is on your machine) ⚠️

**For full Excel sync in production**, choose one of:
- **Option A (easiest)**: Run locally (`npm run dev`) — writes go to your actual file in real time
- **Option B**: Host on Railway/Render/Fly.io (persistent filesystem) instead of Vercel
- **Option C**: Replace Excel sync with Supabase/PlanetScale DB (ask for the migration)

The app works fully on Vercel as a UI — all data persists in the browser's localStorage as a fallback. The Excel file will be the initial seed data.
