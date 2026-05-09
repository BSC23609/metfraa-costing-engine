# Deployment Guide

The package is set up to run as a **single service** — the Express backend serves the frontend from `/public` on the same origin. One URL, no CORS, no separate static host.

## Option A: Render (recommended)

### Via Blueprint (one-click)

1. Push this folder to a GitHub repo
2. In Render, click **New +** → **Blueprint** → connect your repo
3. Render will detect `render.yaml` and prompt for the `sync: false` env vars — paste in your Azure values
4. Click Apply. First deploy takes ~2 minutes.

### Via manual setup

1. **New +** → **Web Service** → connect your repo
2. Settings:
   - **Runtime:** Node
   - **Build Command:** `npm install`
   - **Start Command:** `npm start`
   - **Plan:** Free is fine (note: free plan sleeps after 15 min idle, ~30s cold start)
3. Add environment variables (see SETUP.md table)
4. Deploy

Your app will be live at `https://<service-name>.onrender.com`.

### Custom domain (optional)

In Render dashboard → your service → **Settings** → **Custom Domains** → add your domain and update DNS as instructed.

## Option B: Local

```bash
cp .env.example .env
# fill in .env
npm install
npm start
```

App runs on `http://localhost:3000` (or whatever `PORT` you set).

## Option C: Docker

Create a `Dockerfile` (not included by default):

```dockerfile
FROM node:18-alpine
WORKDIR /app
COPY package*.json ./
RUN npm install --omit=dev
COPY . .
EXPOSE 3000
CMD ["npm", "start"]
```

Build and run:

```bash
docker build -t metfraa-costing .
docker run -p 3000:3000 --env-file .env metfraa-costing
```

## Health check

Once deployed, hit `/api/health` — should return `{ "ok": true, "time": "..." }`.

## Troubleshooting

**"❌ Connection Failed" in the app**
- Server console will show the underlying Graph API error
- Most common: missing admin consent on the app registration permissions, or wrong `TARGET_USER_EMAIL`

**"No matching Option IDs found" when saving rates**
- The `Option ID` column in your Excel is empty for the rows you tried to update — fill them in

**Settings button does nothing / loops password prompt**
- `ADMIN_PASSWORD` env var has whitespace or different from what you're typing — check the Render dashboard value exactly

**Cold start is slow (Render free plan)**
- Upgrade to a paid plan, or hit `/api/health` from an uptime monitor every 10 minutes to keep the dyno warm
