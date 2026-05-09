# METFRAA PEB Costing Engine

Internal B2B web application for generating industrial quotations for Pre-Engineered Buildings (PEB).
Pulls live pricing from a OneDrive-hosted Excel master sheet, walks the user through a 4-step wizard, and exports the finished quotation as PDF, Excel, or HTML saved back to OneDrive.

## What's in this package

```
metfraa-costing-engine/
├── server.js                  Backend (Express) — also serves the frontend
├── package.json
├── render.yaml                Render blueprint (1-click deploy)
├── .env.example               Environment variable template
├── public/
│   └── index.html             Frontend (vanilla JS, no framework)
└── docs/
    ├── SETUP.md               First-time setup (Azure, OneDrive, env vars)
    ├── DEPLOYMENT.md          How to deploy to Render / Docker / locally
    └── MASTER_SHEET_SCHEMA.md Excel column spec + rows to add
```

## Quick start

1. **Set up your master sheet** — see `docs/MASTER_SHEET_SCHEMA.md`
2. **Configure environment** — copy `.env.example` to `.env` and fill in
3. **Install + run locally:**
   ```bash
   npm install
   npm start
   ```
   Open http://localhost:3000

4. **Deploy** — see `docs/DEPLOYMENT.md`

## Key features

- 4-step wizard: Project Info → Build Cost → Preview → Finalize
- Live master data sync from OneDrive (with manual ⟳ Refresh button)
- In-app rate editor (⚙ Settings) — edit master rates, push back to OneDrive
- Auto-banded categories: Distance (km) and Erection Height (m) self-select bands
- Three export formats: Excel, PDF, OneDrive HTML
- Optional admin password for the rate editor

## Architecture

- **Frontend:** vanilla HTML/CSS/JS — no React, no Vue, no build step
- **Backend:** Node.js + Express on Render
- **Storage:** Microsoft OneDrive via Graph API (app-only auth via Azure App Registration)
- **Same-origin deployment:** backend serves the frontend, so no CORS setup needed

## Tech notes

- Reads support multiple rate column names (`Rate`, `rate`, `Rate (₹)`, `Price`) so legacy sheets work unchanged
- Writes auto-detect which column the sheet uses and write back to the same one
- Filename sanitization protects Graph API path validation against slashes in reference IDs
