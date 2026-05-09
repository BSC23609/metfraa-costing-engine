# Setup Guide

This document walks through everything you need to do **once**, before deploying.

## 1. Azure App Registration

Your existing app registration should still work. Skip this section if you're reusing the same `TENANT_ID` / `CLIENT_ID` / `CLIENT_SECRET` from your previous deployment.

If you're creating a fresh app:

1. Go to [Azure Portal](https://portal.azure.com) → **Microsoft Entra ID** → **App registrations** → **+ New registration**
2. Name it `Metfraa Costing Engine`, leave defaults, click Register
3. From the **Overview** page, copy:
   - **Application (client) ID** → this is `CLIENT_ID`
   - **Directory (tenant) ID** → this is `TENANT_ID`
4. Go to **Certificates & secrets** → **+ New client secret** → copy the **Value** immediately → this is `CLIENT_SECRET`
5. Go to **API permissions** → **+ Add a permission** → **Microsoft Graph** → **Application permissions** → add:
   - `Files.ReadWrite.All`
   - `User.Read.All`
6. Click **Grant admin consent for [your tenant]** — you must do this or the API calls will return 403

## 2. OneDrive folder structure

In the OneDrive of the user whose email you'll set as `TARGET_USER_EMAIL`, create:

```
OneDrive/
└── Metfraa_Costing_App/
    ├── Master_Data/
    │   └── Master_Cost_DB.xlsx
    └── Generated_Costings/
```

These exact paths are hardcoded in the backend — see `server.js` constants `MASTER_PATH` and `SAVE_FOLDER`.

## 3. Master_Cost_DB.xlsx schema

See `docs/MASTER_SHEET_SCHEMA.md` for the full column spec and required rows.

## 4. Environment variables

Copy `.env.example` to `.env` for local dev, or set these in your Render dashboard:

| Variable | Required | What it is |
|---|---|---|
| `TENANT_ID` | yes | Azure Directory (tenant) ID |
| `CLIENT_ID` | yes | Azure Application (client) ID |
| `CLIENT_SECRET` | yes | Client secret value (NOT the secret ID) |
| `TARGET_USER_EMAIL` | yes | Email of the OneDrive owner |
| `ADMIN_PASSWORD` | no | If set, the Settings/rate editor requires this password |
| `PORT` | no | Auto-set by Render. Override only for local dev |

## 5. Verify

Once your `.env` is set up, run locally:

```bash
npm install
npm start
```

Open `http://localhost:3000`. You should see "⏳ Syncing Database..." disappear and the four-step wizard become active. If it stays at "❌ Connection Failed", check the server console for the actual Graph API error.
