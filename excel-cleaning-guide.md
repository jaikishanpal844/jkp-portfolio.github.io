# Excel Data Cleaning Guide — E-Commerce Dataset

A concise, practical guide for cleaning a messy e-commerce table in Excel and building quick analysis on top of it. Practice files: **ecommerce_messy.csv** (raw) and **ecommerce_cleaned.csv** (finished).

---

## 1. Bring the Data into Excel as Columns

### Option A — Text to Columns (comma-delimited CSV)

1. Open `ecommerce_messy.csv` in Excel (**File → Open**).  
   Excel usually auto-detects the comma delimiter and splits columns for you.
2. If the data lands in a single column, select that column → **Data → Text to Columns**.
3. Choose **Delimited → Next**, tick **Comma** → **Next → Finish**.

### Option B — Power Query (recommended for repeatable refresh)

1. **Data → Get Data → From File → From Text/CSV**.
2. Select `ecommerce_messy.csv`, click **Transform Data**.
3. Power Query opens the editor — you can apply every cleaning step below here before loading.
4. Click **Close & Load** when done. The result lands on a worksheet.

---

## 2. Cleaning Steps & Formulas

Work in a **helper column** (e.g., column J onward) while keeping raw data intact.

### 2.1 Remove Extra Spaces

```
=TRIM(A2)          ← strips leading, trailing, and double-internal spaces
=TRIM(CLEAN(A2))   ← also removes non-printable characters (line breaks, etc.)
```

### 2.2 Standardise Text Case

| Goal | Formula |
|---|---|
| Title Case (names) | `=PROPER(TRIM(A2))` |
| UPPER CASE (region, status) | `=UPPER(TRIM(B2))` |
| lower case | `=LOWER(TRIM(B2))` |

### 2.3 Fix Product Names ("(blank)", mixed case, extra spaces)

```
=IF(TRIM(E2)="(blank)","",PROPER(TRIM(E2)))
```

This converts `" whey powder "` → `"Whey Powder"` and turns `"(blank)"` into an empty cell.

### 2.4 Convert Currency Text to Numbers

The messy file stores prices as `$35.00` (text). Strip the `$` and convert:

```
=VALUE(SUBSTITUTE(F2,"$",""))
```

Or use **NUMBERVALUE** with locale-safe parsing:

```
=NUMBERVALUE(SUBSTITUTE(F2,"$",""),".","")
```

### 2.5 Standardise Dates

Mixed formats like `03/02/2024`, `2024/02/02`, `02-01-2024` all need to become real Excel dates.

**If already recognised as dates** — format the column as `YYYY-MM-DD` via  
**Format Cells → Number → Date → Custom → `yyyy-mm-dd`**.

**If stored as text** — use DATEVALUE on a known format, or Power Query's  
*Transform → Data Type → Date* which auto-detects formats.

```
=DATEVALUE(TEXT(B2,"yyyy-mm-dd"))   ← last resort; prefer Power Query
```

### 2.6 Standardise Status Column

```
=PROPER(TRIM(I2))   ← "COMPLETED", "completed" → "Completed"
```

---

## 3. Handling Blank & "(blank)" Rows

### Blank Customer Name rows

Filter or flag them so they can be reviewed before deletion:

```
=IF(TRIM(C2)="","MISSING","OK")
```

To delete all blank-customer rows: **Data → Filter → uncheck (Blanks) in Customer_Name column → select visible rows → Delete Row → Clear Filter**.

### "(blank)" Product rows

Use the formula from §2.3. After replacing, filter on empty Product cells and decide whether to impute or delete.

---

## 4. Build an Excel Table & Run Quick Analysis

### 4.1 Convert to an Excel Table

1. Click anywhere inside the cleaned data.  
2. **Insert → Table** (or `Ctrl + T`). Tick "My table has headers" → **OK**.  
3. Give the table a name in the **Table Design** tab, e.g., `tbl_Orders`.

Tables auto-expand, use structured references like `tbl_Orders[Total_Amount]`, and make pivot tables easy to refresh.

### 4.2 Create a Pivot Table

1. Click inside `tbl_Orders` → **Insert → PivotTable → New Worksheet → OK**.
2. Drag fields:
   - **Rows**: Region  
   - **Values**: Total_Amount (set to **Sum**), Order_ID (set to **Count**)
3. Right-click a value → **Value Field Settings** to change aggregation.

Sample output:

| Region | Total Revenue | Order Count |
|---|---|---|
| East | 245.00 | 5 |
| North | 244.00 | 5 |
| South | 159.00 | 4 |
| West | 105.00 | 1 |

### 4.3 Quick KPI Formulas (outside the pivot)

```excel
Total Revenue    =SUM(tbl_Orders[Total_Amount])
Order Count      =COUNTA(tbl_Orders[Order_ID])
Avg Order Value  =AVERAGE(tbl_Orders[Total_Amount])
Max Single Order =MAX(tbl_Orders[Total_Amount])
Max Order (East) =MAXIFS(tbl_Orders[Total_Amount],tbl_Orders[Region],"East")
```

To find **which region had the highest total revenue**, first build a small SUMIF summary table (see §4.4), then use `=INDEX(A2:A5,MATCH(MAX(B2:B5),B2:B5,0))`.

Add a SUMIF summary table and use `=INDEX(A2:A5,MATCH(MAX(B2:B5),B2:B5,0))`:

| | A (Region) | B (Revenue) |
|---|---|---|
| 2 | East | `=SUMIF(tbl_Orders[Region],A2,tbl_Orders[Total_Amount])` |
| 3 | North | `=SUMIF(tbl_Orders[Region],A3,tbl_Orders[Total_Amount])` |
| 4 | South | `=SUMIF(tbl_Orders[Region],A4,tbl_Orders[Total_Amount])` |
| 5 | West | `=SUMIF(tbl_Orders[Region],A5,tbl_Orders[Total_Amount])` |
| 6 | **Max Region** | `=INDEX(A2:A5,MATCH(MAX(B2:B5),B2:B5,0))` |

### 4.4 Add a Bar Chart

1. Select your region-revenue summary table (A1:B5).  
2. **Insert → Recommended Charts → Clustered Bar → OK**.  
3. Add a title: *"Revenue by Region"*.

---

## 5. Quick-Reference Formula Cheat Sheet

| Task | Formula |
|---|---|
| Remove spaces | `=TRIM(A2)` |
| Remove hidden chars | `=CLEAN(A2)` |
| Title Case | `=PROPER(A2)` |
| Upper Case | `=UPPER(A2)` |
| Replace `$` | `=SUBSTITUTE(A2,"$","")` |
| Text → Number | `=VALUE(A2)` |
| Text → Date | `=DATEVALUE(A2)` |
| Date → Text | `=TEXT(A2,"yyyy-mm-dd")` |
| Flag blanks | `=IF(A2="","MISSING","OK")` |
| Revenue by region | `=SUMIF(Region_col,criteria,Amount_col)` |
| Highest region | `=INDEX(range,MATCH(MAX(totals),totals,0))` |

---

*Practice files:*  
- [`ecommerce_messy.csv`](ecommerce_messy.csv) — raw data with mixed cases, `$` prices, varied date formats, blank names, `(blank)` product.  
- [`ecommerce_cleaned.csv`](ecommerce_cleaned.csv) — fully cleaned reference version.
