# Master_Cost_DB.xlsx — v5 Auto-Bootstrap

**You don't need to edit Excel manually anymore.** The app handles Distance and Paint setup automatically.

## What happens on first run

1. The app boots and calls `/api/master-data`
2. The server checks if your sheet has `Distance` and `Paint` parameters
3. If missing, it auto-writes the default rows below back to OneDrive
4. Required columns (`Min`, `Max`) get added to the header if not already there
5. From then on, everything reads from your sheet normally

The console log will show: `🔧 Bootstrapping: added 8 Distance band rows` and `🔧 Bootstrapping: added 3 Paint type rows` on first hit.

## What gets auto-written

### Distance (8 banded rows)

Used by Step 1's "Distance Band" auto-display. Rates editable from Settings.

| ID | Name | Rate (₹/MT) | Min | Max |
|---|---|---|---|---|
| DIST001 | 0–100 km    | 1500 | 0   | 100 |
| DIST002 | 101–200 km  | 2500 | 101 | 200 |
| DIST003 | 201–300 km  | 3500 | 201 | 300 |
| DIST004 | 301–400 km  | 4500 | 301 | 400 |
| DIST005 | 401–500 km  | 5000 | 401 | 500 |
| DIST006 | 501–600 km  | 5500 | 501 | 600 |
| DIST007 | 601–700 km  | 6000 | 601 | 700 |
| DIST008 | 701–800 km  | 6500 | 701 | 800 |

### Paint (3 starter rows)

Used by Step 1's "Paint Type" dropdown. Rates editable, **and** new paint types can be added from Settings → "+ Add Paint Type" button.

| ID | Name | Rate (₹/MT) |
|---|---|---|
| PAINT001 | Primer + Enamel | 3800 |
| PAINT002 | Primer + Epoxy  | 4700 |
| PAINT003 | PU              | 0 (placeholder — edit later) |

## Erection — still hardcoded

Erection charges live in the app code, not Excel. Auto-applied based on:
- **Step 1 dropdown:** Safety → Standard ₹9,000 / Intermediate ₹12,000 / Advanced ₹14,000
- **Step 1 auto-band:** Building Height → ≤10m: ₹0 / ≤15m: ₹1500 / ≤20m: ₹3000

To change Erection rates, you need a code change. Tell me when you want them moved to Excel and I'll wire it up.

## What you can do from Settings

| Action | Distance | Paint | Other params |
|---|---|---|---|
| Edit rate | ✅ | ✅ | ✅ |
| Add new option | ❌ (bands fixed) | ✅ + Add Paint Type | ❌ (manual Excel) |
| Delete option | ❌ | ✅ 🗑 Delete button per row | ❌ |
| Edit Min/Max | ❌ (Excel only) | n/a | n/a |

## Hidden parameters

`Secondary` is hidden server-side. To show or hide other parameters, edit `HIDDEN_PARAMETERS` in `server.js` line ~165.
