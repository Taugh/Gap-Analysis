# CMMS Gap Analysis — reproducible pipeline

Scores IREN's CMMS exports against the Asset Management Governance Framework
(Strategy Rev 3.0, AM-001 Rev 3.0, AM-002 Rev 1.0) and rebuilds the four gap-analysis
workbooks, the charts and the Word report from a folder of CSV exports.

To re-baseline: pull the reports listed in `docs/data_dictionary.md`, save them under the
fixed file names in `inputs/`, set `SNAPSHOT_DATE` in `config.py`, and run one command.

```
python run_all.py
```

## What you need

| Component | Used for | Install |
|---|---|---|
| Python 3.10+ with `pandas`, `numpy`, `openpyxl`, `matplotlib` | scoring, workbooks, charts | `pip install -r requirements.txt` |
| LibreOffice (`soffice` on PATH) | recalculating workbook formulas and checking for `#NAME?`/`#REF!` errors | https://www.libreoffice.org — or skip with `--no-recalc` (the workbooks still work; Excel calculates on open) |
| Node.js 18+ with the `docx` package | the Word report only | `npm install` in this folder — or skip with `--no-report` |

## Folder layout

```
cmms-gap-analysis/
├── config.py            snapshot date, file names, date formats, site codes — the only file you normally edit
├── run_all.py           runs every step in order; --only / --no-recalc / --no-report
├── requirements.txt     Python dependencies
├── package.json         Node dependency (docx) for the report
├── inputs/              the CMMS exports, saved under the fixed names in docs/data_dictionary.md
├── src/                 one script per step (s01 … s09); each says what it reads and writes at the top
├── report/              build_report.js — the Word report (narrative + tables + charts)
├── tools/               recalc.py and LibreOffice helpers (from Anthropic's xlsx skill)
├── work/                intermediate files (pickles, facts.json) — safe to delete; regenerated on every run
├── outputs/             the deliverables
└── docs/                data dictionary, maintenance notes, rule catalogue
```

## What comes out

| File | Contents |
|---|---|
| `outputs/WorkOrder_Gap_Analysis.xlsx` | Closed WO history scored against 35 rules; §10 KPI block; traceability; bad-actor candidates; open backlog and backlog-at-booked-pace by site; permissions mapping (53 groups → 8 buckets) with aggregate user counts; failure-code review; heat maps; every record with a 1/0/blank flag per rule |
| `outputs/PM_Schedule_Gap_Analysis.xlsx` | 3,005 schedules scored against 15 rules; equipment PM coverage (Direct / Via parent / Paused only / None) by category, site and criticality group; Group A/B exceptions; schedule × asset references |
| `outputs/Asset_Register_Gap_Analysis.xlsx` | 27,193 assets scored against 20 rules; Strategy G-13.2 traceability; heat maps by site, category and creation batch; warranty; manufacturer variants |
| `outputs/Inventory_Gap_Analysis.xlsx` | Stock lines scored against 17 AM-002 rules; §13 KPI block; G-12.1 field-by-field traceability; parts usage |
| `outputs/charts/*.png` | the four report figures |
| `outputs/Facts_Summary.md` and `work/facts.json` | every headline figure, read from the recalculated workbooks — use these to update the report narrative |
| `outputs/CMMS_Gap_Analysis_Report_Rev<rev>.docx` | the leadership report |

Every count in the workbooks is a formula over the flag columns on that workbook's `Data`
sheet, so a rule can be adjusted (in `src/`) and everything downstream recomputes.

## Re-running after a new pull — step by step

1. Pull each report listed in `docs/data_dictionary.md` and save it in `inputs/` under the exact file name shown.
   The scheduled-maintenance report has been seen with two header variants; both are accepted.
2. Set `SNAPSHOT_DATE` in `config.py` to the date the exports were taken. Everything date-relative
   (backlog age, the 12-week booked-hours window, "created in the last 24 months", login recency) uses it.
3. `python run_all.py`. About two minutes on a laptop; LibreOffice recalculation is the slow part.
4. Read `outputs/Facts_Summary.md` and update the narrative in `report/build_report.js` — see
   `docs/maintenance.md` for which passages carry figures. Bump `REPORT_REVISION` in `config.py`
   and re-run `node report/build_report.js` (or `python run_all.py --only` nothing and just the report step).
5. Open the Word report and let it update the table of contents when prompted.

## What is and is not automatic

Automatic on every run: all rule scoring, all workbook tables and heat maps, the coverage join,
the bad-actor list, the permissions mapping (first-pass bucket proposals), the failure-code
review, the four charts, and the facts summary.

Manual: the report is a narrative document. Its tables and figures regenerate, but the prose
carries figures from the September 2026 baseline and must be revised against
`outputs/Facts_Summary.md`. A few "Read Me" and traceability cells in the workbooks also carry
baseline figures in prose — `docs/maintenance.md` lists where.

## Personal data

`inputs/UsersBySite.csv` contains names and email addresses. `src/s05_users.py` reads it and writes
only aggregate counts; no individual appears in any workbook or in the report. Treat `inputs/` as
internal and do not include it when sharing the project.

## Governing documents and rule references

Every rule carries the clause it traces to (Strategy, AM-001, AM-002). The rule text, impact,
likely root cause and recommendation live in the `RULES` dictionaries in each `src/s0*.py` file
and appear on the `Gap Rules` sheet of the corresponding workbook. If the framework documents are
revised, update the clause references there and the "Governing documents" paragraph in each
script's Read Me block.
