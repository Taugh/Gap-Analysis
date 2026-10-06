"""
Central configuration for the CMMS gap analysis pipeline.

Edit SNAPSHOT_DATE to the day the exports were pulled; every date-relative
rule (backlog age, run-rate window, "created in the last 24 months", etc.)
is computed from it. Input file names are fixed — save each CMMS export under
the name shown in INPUTS.
"""
from pathlib import Path
import pandas as pd

ROOT = Path(__file__).resolve().parent
INPUT_DIR = ROOT / "inputs"
WORK_DIR = ROOT / "work"        # intermediate pickles / csv
OUT_DIR = ROOT / "outputs"      # workbooks, charts, report
REPORT_DIR = ROOT / "report"
for d in (WORK_DIR, OUT_DIR):
    d.mkdir(exist_ok=True)

# ---- the day the exports were taken (YYYY-MM-DD). Used for all "as at" logic.
SNAPSHOT_DATE = "2026-09-28"
SNAPSHOT = pd.Timestamp(SNAPSHOT_DATE)
SNAPSHOT_LABEL = SNAPSHOT.strftime("%d %b %Y")          # e.g. 28 Sep 2026
PREPARED_LABEL = SNAPSHOT.strftime("%d %b %Y")

# ---- run-rate window for booked hours: the last N complete ISO weeks before the snapshot
RUN_RATE_WEEKS = 12

# ---- report revision shown on the Word document
REPORT_REVISION = "0.1"

# ---- input files (CMMS report → file name). Column expectations are in docs/data_dictionary.md
INPUTS = {
    "closed_wo":      INPUT_DIR / "ClosedWorkOrders.csv",
    "open_wo":        INPUT_DIR / "OpenWorkOrderList.csv",
    "schedules":      INPUT_DIR / "ScheduledMaintenanceList.csv",
    "assets":         INPUT_DIR / "AllAssets.csv",
    "warranty":       INPUT_DIR / "AssetWarranty.csv",
    "stock":          INPUT_DIR / "StockList.csv",
    "parts_usage":    INPUT_DIR / "PartsUsage.csv",
    "user_groups":    INPUT_DIR / "UserGroups.csv",
    "users_by_site":  INPUT_DIR / "UsersBySite.csv",
    "failure_codes":  INPUT_DIR / "FailureCodeProblemsCausesActions.csv",
    "hours_by_type":  INPUT_DIR / "MaintenanceHoursByType.csv",   # optional; used for reconciliation note only
}

# ---- intermediate files
WORK = {
    "wo_scored":     WORK_DIR / "wo_scored.pkl",
    "bad_actors":    WORK_DIR / "bad_actors.pkl",
    "assets_scored": WORK_DIR / "assets_scored.pkl",
    "sm":            WORK_DIR / "sm.pkl",
    "sm_assets":     WORK_DIR / "sm_assets.pkl",
    "assets_cov":    WORK_DIR / "assets_cov.pkl",
    "group_counts":  WORK_DIR / "group_counts.csv",
    "user_facts":    WORK_DIR / "user_facts.json",
    "facts":         WORK_DIR / "facts.json",
}

# ---- outputs
OUTPUTS = {
    "wo":     OUT_DIR / "WorkOrder_Gap_Analysis.xlsx",
    "pm":     OUT_DIR / "PM_Schedule_Gap_Analysis.xlsx",
    "assets": OUT_DIR / "Asset_Register_Gap_Analysis.xlsx",
    "inv":    OUT_DIR / "Inventory_Gap_Analysis.xlsx",
    "report": OUT_DIR / f"CMMS_Gap_Analysis_Report_Rev{REPORT_REVISION}.docx",
    "charts": OUT_DIR / "charts",
}
OUTPUTS["charts"].mkdir(exist_ok=True)

# ---- date formats used by the CMMS exports (these differ between reports)
FMT_CLOSED_WO = "%d/%m/%Y %H.%M.%S"   # ClosedWorkOrders: 18/03/2024 00.21.16
FMT_OPEN_WO = "%m/%d/%Y %H:%M"        # OpenWorkOrderList: 8/7/2026 19:47
FMT_ASSETS = "%m/%d/%Y %H:%M"         # AllAssets Created: 11/10/2022 17:19
FMT_USERS = "%m/%d/%Y %H:%M"          # UsersBySite Last Login

# ---- equipment vs location split for the asset register (by category prefix)
LOCATION_CATEGORY_PREFIXES = (
    "Locations And Facilities", "Buildings", "Road", "Structural", "Civil", "Culvert", "HSE", "Water",
    "Electrical Infrastructure - General", "Mechanical Infrastructure - General", "Piping",
)

# ---- site code prefixes used to check asset naming (Strategy G-8.4)
SITE_CODES = {
    "Mackenzie": "MKZ", "Canal Flats": "CF", "Prince George": "PG", "Childress": "CHI", "Sweetwater 1": "SWT",
    "Childress - HPC": "CHI", "Prince George - HPC": "PG", "Prince George - HPC - New": "PG", "Prince George - Miners": "PG",
}

# ---- expected charge department per site (WO rule V04)
SITE_DEPT = {
    "Childress": "US-CHL00", "Childress - HPC": "US-CHL00", "Mackenzie": "CA-MCK00", "Prince George": "CA-PGE00",
    "Prince George - HPC": "CA-PGE00", "Prince George - Miners": "CA-PGE00", "Canal Flats": "CA-CFL00", "Sweetwater 1": "US-SWT00",
}
