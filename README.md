# METFRAA PEB Costing Engine

Internal B2B web app for generating PEB quotations. Backend serves the frontend (single Render service).

## Features

- 4-step wizard: Project Info → Build Cost → Preview → Save & Download
- Live OneDrive master data sync with manual ⟳ Refresh
- Load existing quotations by Ref or Project Name
- Excel + PDF download
- Save PDF + JSON back to OneDrive (re-loadable later)
- ⚙ Settings page with master rate editor (password-gated)
- Hides "Secondary" parameter from the sidebar as a safety net

## Quick start

```bash
cp .env.example .env
# fill in the 4 Azure values + ADMIN_PASSWORD
npm install
npm start
```

Open http://localhost:3000

## Deploy

Push to GitHub → connect to Render via Blueprint → fill env vars → done.
See `docs/DEPLOYMENT.md` for details.

## File layout

```
.
├── server.js            Backend (Express 5)
├── package.json
├── public/
│   └── index.html       Frontend (vanilla JS)
├── render.yaml          Render blueprint
├── .env.example
└── docs/
    ├── DEPLOYMENT.md
    └── SETUP.md
```
