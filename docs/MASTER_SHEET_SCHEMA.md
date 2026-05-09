# Master_Cost_DB.xlsx — Schema

The single sheet (first sheet of the workbook) drives everything in the app.

## Column spec

| Column | Required | Notes |
|---|---|---|
| `Parameter Name` | yes | Top-level category. Becomes a sidebar item in Step 2. |
| `Subcategory` | yes | Row label inside that category. |
| `Option ID` | yes | Unique identifier. Used for write-back matching. Anything goes — `STL001`, `DIST001`, etc. |
| `Option Name` | yes | What appears in the dropdown. |
| `Rate` | yes | Numeric. Aliases `rate`, `Rate (₹)`, `Price` also accepted on read. |
| `Min` | optional | Lower bound for banded categories (Distance km, Erection Height m). Leave blank otherwise. |
| `Max` | optional | Upper bound for banded categories. Leave blank otherwise. |

## Required rows for the four "fixed variable" categories

These power the Distance / Paint / Erection logic. Add them once; edit rates later from the in-app Settings page.

### Distance (8 bands — auto-selected from km input on Step 1)

| Parameter Name | Subcategory | Option ID | Option Name | Rate | Min | Max |
|---|---|---|---|---|---|---|
| Distance | Distance | DIST001 | 0–100 km    | 1500 | 0   | 100 |
| Distance | Distance | DIST002 | 101–200 km  | 2500 | 101 | 200 |
| Distance | Distance | DIST003 | 201–300 km  | 3500 | 201 | 300 |
| Distance | Distance | DIST004 | 301–400 km  | 4500 | 301 | 400 |
| Distance | Distance | DIST005 | 401–500 km  | 5000 | 401 | 500 |
| Distance | Distance | DIST006 | 501–600 km  | 5500 | 501 | 600 |
| Distance | Distance | DIST007 | 601–700 km  | 6000 | 601 | 700 |
| Distance | Distance | DIST008 | 701–800 km  | 6500 | 701 | 800 |

### Paint (PU left at 0 — edit later via Settings)

| Parameter Name | Subcategory | Option ID | Option Name | Rate | Min | Max |
|---|---|---|---|---|---|---|
| Paint | Paint Type | PAINT001 | Primer + Enamel | 3800 |  |  |
| Paint | Paint Type | PAINT002 | Primer + Epoxy  | 4700 |  |  |
| Paint | Paint Type | PAINT003 | PU              | 0    |  |  |

### Erection — Safety (manual pick in Step 2)

| Parameter Name | Subcategory | Option ID | Option Name | Rate | Min | Max |
|---|---|---|---|---|---|---|
| Erection | Safety | ERSAF001 | Standard     | 9000  |  |  |
| Erection | Safety | ERSAF002 | Intermediate | 12000 |  |  |
| Erection | Safety | ERSAF003 | Advanced     | 14000 |  |  |

### Erection — Height (auto-selected from m input on Step 1)

| Parameter Name | Subcategory | Option ID | Option Name | Rate | Min | Max |
|---|---|---|---|---|---|---|
| Erection | Height | ERHGT001 | Up to 10 m | 0    | 0     | 10 |
| Erection | Height | ERHGT002 | Up to 15 m | 1500 | 10.01 | 15 |
| Erection | Height | ERHGT003 | Up to 20 m | 3000 | 15.01 | 20 |

## Raw materials

Your existing rows are untouched. Just keep the same shape — Parameter Name + Subcategory + Option ID + Option Name + Rate. Leave `Min` and `Max` blank.

## How banding works

When the user enters a value in Step 1 (e.g. `155` km, `12` m), the frontend finds the row whose `Min ≤ value ≤ Max` and auto-selects it in Step 2. Bands are evaluated inclusively on both ends, so make sure your ranges don't overlap. The Excel rows above use `10 → 10.01` style boundaries to avoid this.

## Adding new banded items later

Any new `Parameter Name` row with `Min`/`Max` filled in becomes a banded option automatically. The frontend's auto-pick currently only runs for the two specific paths `Distance/Distance` and `Erection/Height`. If you add a third banded category later, ping the dev to wire up its trigger field on Step 1.
