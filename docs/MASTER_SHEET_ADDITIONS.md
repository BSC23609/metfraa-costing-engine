# Rows to Add to Master_Cost_DB.xlsx

You need to add new rows to your master sheet for **Distance** and **Paint** parameters. **Erection charges are hardcoded in the app** — they don't go in Excel.

## Schema reminder

Your sheet currently has these columns: `Parameter Name | Subcategory | Option ID | Option Name | Rate | Type | Unit`

Two **new optional columns** need to be added if not already present: **`Min`** and **`Max`** (for the Distance bands). Leave them blank for all your existing rows.

## Updated header row

| Parameter Name | Subcategory | Option ID | Option Name | Rate | Type | Unit | Min | Max |
|---|---|---|---|---|---|---|---|---|

## 1. Distance (8 banded rows)

The system reads `Distance from Vichoor (km)` on Step 1, finds the row whose `Min ≤ km ≤ Max`, and auto-selects it on Step 2. User can override.

| Parameter Name | Subcategory | Option ID | Option Name | Rate | Type | Unit | Min | Max |
|---|---|---|---|---|---|---|---|---|
| Distance | Distance | DIST001 | 0–100 km    | 1500 | RATE | MT | 0   | 100 |
| Distance | Distance | DIST002 | 101–200 km  | 2500 | RATE | MT | 101 | 200 |
| Distance | Distance | DIST003 | 201–300 km  | 3500 | RATE | MT | 201 | 300 |
| Distance | Distance | DIST004 | 301–400 km  | 4500 | RATE | MT | 301 | 400 |
| Distance | Distance | DIST005 | 401–500 km  | 5000 | RATE | MT | 401 | 500 |
| Distance | Distance | DIST006 | 501–600 km  | 5500 | RATE | MT | 501 | 600 |
| Distance | Distance | DIST007 | 601–700 km  | 6000 | RATE | MT | 601 | 700 |
| Distance | Distance | DIST008 | 701–800 km  | 6500 | RATE | MT | 701 | 800 |

## 2. Paint (3 rows)

User picks one option in Step 2. PU rate is `0` for now — edit later from the Settings page in the app.

| Parameter Name | Subcategory | Option ID | Option Name | Rate | Type | Unit | Min | Max |
|---|---|---|---|---|---|---|---|---|
| Paint | Paint Type | PAINT001 | Primer + Enamel | 3800 | RATE | MT |  |  |
| Paint | Paint Type | PAINT002 | Primer + Epoxy  | 4700 | RATE | MT |  |  |
| Paint | Paint Type | PAINT003 | PU              | 0    | RATE | MT |  |  |

## 3. Erection — NOT in Excel

Erection charges live in the app code, not the master sheet. They appear automatically in the final quotation based on:
- **Step 1 dropdown:** Safety Type → Standard ₹9000 / Intermediate ₹12000 / Advanced ₹14000
- **Step 1 auto-band:** Building Height → 0–10m: ₹0 / 10–15m: ₹1500 / 15–20m: ₹3000

Both summed together, multiplied by Steel Tonnage. Two separate lines under "Erection" in the preview/PDF/Excel.

To change these rates later, you'll need a code change — they're not editable from the Settings rate editor. If you want them moved to Excel later, ping me.

## How qty defaults work

For any parameter with `Unit = MT`, the qty field in Step 2 will auto-fill to your **Steel Tonnage** value from Step 1. User can override per-parameter. So you don't have to type "100" eight times for Distance/Paint/etc — set tonnage on Step 1 once, it propagates.
