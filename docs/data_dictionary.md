# Data dictionary — the CMMS exports the pipeline expects

Save each report in `inputs/` under the file name shown. Column names must match (the scripts strip
surrounding whitespace but are otherwise literal). Extra columns are ignored. The CMMS caps exports at
2,000 rows per pull: for the large reports, pull in chunks (by site or date) and concatenate them into
one file with a single header row before saving.

## ClosedWorkOrders.csv  (CMMS: closed work order report, all sites, all time)
Rows: one per closed work order. Nov 2023 → snapshot. ~18k rows (chunk by site/quarter).
Date format: `dd/mm/yyyy HH.MM.SS` (set `FMT_CLOSED_WO` in config if this changes).

| Column | Used for |
|---|---|
| Priority, Asset Name, Asset Code, Asset Description, Asset Category, Site Name | completeness, hierarchy level, site cuts |
| Description, Work Order Code, Maintenance Type, Status | categories, status checks, test-record detection |
| DateCreated, Suggested Completion, Date Completed, Days Open | backfilling, lateness, PM-compliance proxy, run-rate window |
| Requested By, Assigned Users, Completed by | G-5.14 fields, Guest requester |
| Scheduled Maintenance ID | PM linkage (G15/V07) |
| Notes, Problem, Cause, Solution, Completion Notes | G-5.15/5.16 reporting standard, placeholders |
| Est Hours, Actual Hours, Total Labour Costs, Misc Costs, Total Misc Costs | labour rules, booked-hours run rate |
| Parts Used, Parts Codes, Total Parts Costs, Total All Costs | parts consumption (AM-002 G-8.14) |
| Acc Code, Acc Description, Charge Dept Code, Charge Dept Description | cost coding, site/dept consistency |

## OpenWorkOrderList.csv  (CMMS: open work orders, all sites)
Rows: one per non-closed WO. Date format `m/d/yyyy H:MM` (`FMT_OPEN_WO`).

| Column | Used for |
|---|---|
| WO Code, Work Order Description, Status, Asset, Category, Type, Priority | backlog composition; asset code is parsed from the parenthesised token at the end of `Asset` |
| Date Created, Due Date, Assigned Users, Est Hours, Days Late | age, overdue, G-5.12 flag, ready hours |

## ScheduledMaintenanceList.csv  (CMMS: scheduled maintenance list, active and paused)
Rows: one per schedule (~3k). Two header variants accepted:
`SM Status 1 = Active, 0 = Paused` / `SM Status`, and `Suggested Completion in Days After Work Order creation` / `Suggested Completion`.

| Column | Used for |
|---|---|
| When | trigger class and AM-001 7.3 interval check |
| SM Status | active (1) / paused (0) |
| Code, Priority, Type, Description, Site, Assigned Users, Time Estimated Hours | template rules S01–S11 |
| Assets | comma-separated `Name (CODE)` list — exploded to schedule × asset for coverage and overlap |
| Suggested Completion | days-to-complete vs priority window (S06) |

## AllAssets.csv  (CMMS: asset list, all sites — this is the ~5-hour pull)
Rows: one per asset (~27k). Date format `m/d/yyyy H:MM` (`FMT_ASSETS`).

| Column | Used for |
|---|---|
| Name, Code, Asset Status, Site, Is Site, Category, Description | identity, naming (G-8.4), equipment/location split |
| Asset Location | parent code parsed from the trailing `(CODE)` → hierarchy depth, orphans, via-parent PM coverage |
| Make, Model, Serial Number, Manufacturer Part Number | G-13.2 attributes, placeholders, duplicates, spreadsheet-corrupted serials |
| Asset Criticality | Strategy §7 group |
| Created | creation-batch analysis, "created in last 24 months" |
| Notes, Last Price Currency | carried through only |

## AssetWarranty.csv  (CMMS: asset warranty report)
Columns: `Asset Code, Asset Name, Serial Number, Expiry Date, Status`. Reproduced in full on the Warranty sheet.

## StockList.csv  (CMMS: stock / parts list)
Columns: `Stock Item, Part Code, Description, Category, Inventory Code, UNSPC Code, Location Name, Account Code, Account Description, Charge Dept Code, Charge Dept Description, Part URL, Aisle, Row, Bin, Min. Qty, Qty on Hand, Last Price, Total Value, BOM Groups`. One row per part code per storage location.

## PartsUsage.csv  (CMMS: parts usage / issues, 2019 → snapshot)
Columns: `Date Used, WO Code, Part Name, Part Code, Quantity Used, Stock On Hand, Inventory Cost`. A trailing total row is ignored.

## UserGroups.csv  (CMMS: user group list)
Columns: `Full Name, City, User Name` — only `Full Name` (the group name) is used.

## UsersBySite.csv  (CMMS: users by site, with groups and status) — PERSONAL DATA
Sectioned layout: a `Site Name` header row, the site name on the next row, a column header row
(`Full Name, User Name, User Title, User Status, Email Address, Groups, Last Login, Hourly Rate`), then users.
Users appear once per site they are assigned to; the script de-duplicates on `User Name`. Only aggregates are written out.

## FailureCodeProblemsCausesActions.csv  (CMMS: failure code configuration)
Columns: `Problem Code, Problem Description, Cause Code, Cause Description, Action Code, Action Description` — three independent lists side by side, different lengths, blanks where a list is shorter.

## MaintenanceHoursByType.csv  (optional — CMMS: hours by maintenance type, trailing 12 months)
Used only as a reconciliation note in the WO workbook. Not parsed; keep for the audit trail.

## Also useful but not consumed by the pipeline
- `WeeklyHoursLoggedperSite.pdf` — charts only; the same figures are computed from `ClosedWorkOrders.csv`.
- Governing documents (Strategy, AM-001, AM-002) — the clause references in the rules point at these revisions.
