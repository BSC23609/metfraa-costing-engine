# Rows to Add to Master_Cost_DB.xlsx

You need to add new rows for **Distance** and **Paint**. **Erection charges are hardcoded in the app** — they don't go in Excel.

## Important: where these appear in the app

Distance and Paint are **project-wide charges** — picked once on **Step 1**, applied to the whole project. They do **NOT** appear as parameters in the Step 2 sidebar. The Excel rows below feed the Step 1 dropdowns:

- **Paint Type** → dropdown on Step 1, populated from your Paint rows
- **Distance Band** → readonly display on Step 1, auto-matched to the km you enter

You can still edit their rates from the **⚙ Settings** page in the app.

## Schema reminder

Your sheet currently has columns: `Parameter Name | Subcategory | Option ID | Option Name | Rate | Type | Unit`

Add **two new columns** if not already present: **`Min`** and **`Max`** (used by Distance bands). Leave them blank for everything else.

## Updated header row

| Parameter Name | Subcategory | Option ID | Option Name | Rate | Type | Unit | Min | Max |
|---|---|---|---|---|---|---|---|---|

## 1. Distance (8 banded rows)

The system reads `Distance from Vichoor (km)` on Step 1, finds the row whose `Min ≤ km ≤ Max`, and auto-displays the matched rate. Charged per MT × Steel Tonnage.

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

Populates the Paint Type dropdown on Step 1. PU rate is `0` for now — edit later from the Settings page in the app.

| Parameter Name | Subcategory | Option ID | Option Name | Rate | Type | Unit | Min | Max |
|---|---|---|---|---|---|---|---|---|
| Paint | Paint Type | PAINT001 | Primer + Enamel | 3800 | RATE | MT |  |  |
| Paint | Paint Type | PAINT002 | Primer + Epoxy  | 4700 | RATE | MT |  |  |
| Paint | Paint Type | PAINT003 | PU              | 0    | RATE | MT |  |  |

## 3. Erection — NOT in Excel

Erection charges live in the app code, not the master sheet. They appear automatically in the final quotation based on:
- **Step 1 dropdown:** Safety Type → Standard ₹9,000 / Intermediate ₹12,000 / Advanced ₹14,000
- **Step 1 auto-band:** Building Height → 0–10m: ₹0 / 10–15m: ₹1500 / 15–20m: ₹3000

To change Erection rates later, you'll need a code change. If you want them moved to Excel, ping me.

## Step 2 sidebar after this update

The sidebar in Step 2 will only show Excel parameters that are NOT Distance, Paint, or Secondary. Typically: Raw Material, Sheet Cost, and any other parameters you have.

## Final quotation layout

The final preview/PDF/Excel will show:
1. Your Step 2 parameters (Raw Material, etc.) with their subcategories
2. **Distance** (auto-injected from Step 1, charged per MT × tonnage)
3. **Paint** (auto-injected from Step 1, charged per MT × tonnage, only if you picked one)
4. **Erection** (auto-injected from Step 1, two lines: Safety + Height Surcharge, both per MT × tonnage)
5. Subtotal → Margin → GST → Grand Total
