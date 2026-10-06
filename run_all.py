#!/usr/bin/env python3
"""
Run the whole CMMS gap analysis from the files in inputs/.

    python run_all.py              # everything: workbooks, recalc, charts, facts, report
    python run_all.py --no-report  # skip the Word report (no Node required)
    python run_all.py --no-recalc  # skip LibreOffice recalculation (formulas stay un-cached)
    python run_all.py --only s01 s02   # run selected steps

Steps run in dependency order:
  s01_wo_score     ClosedWorkOrders.csv                  -> work/wo_scored.pkl, work/bad_actors.pkl
  s02_assets       AllAssets.csv, AssetWarranty.csv      -> outputs/Asset_Register_Gap_Analysis.xlsx, work/assets_scored.pkl
  s03_pm_prepare   ScheduledMaintenanceList.csv          -> work/sm*.pkl, work/assets_cov.pkl
  s04_pm_build                                           -> outputs/PM_Schedule_Gap_Analysis.xlsx
  s05_users        UsersBySite.csv                       -> work/group_counts.csv, work/user_facts.json
  s06_inventory    StockList.csv, PartsUsage.csv         -> outputs/Inventory_Gap_Analysis.xlsx
  s07_wo_build     OpenWorkOrderList.csv, UserGroups.csv, FailureCode*.csv -> outputs/WorkOrder_Gap_Analysis.xlsx
  recalc           LibreOffice recalculates every workbook and checks for formula errors
  s08_charts                                             -> outputs/charts/*.png
  s09_facts                                              -> work/facts.json, outputs/Facts_Summary.md
  report           report/build_report.js (Node + docx)  -> outputs/CMMS_Gap_Analysis_Report_Rev<rev>.docx
"""
import argparse, subprocess, sys, json, time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
from config import INPUTS, OUTPUTS, REPORT_REVISION

STEPS = ["s01_wo_score", "s02_assets", "s03_pm_prepare", "s04_pm_build", "s05_users", "s06_inventory", "s07_wo_build"]

def run(cmd, cwd=ROOT, env=None):
    t = time.time()
    print(f"\n>>> {' '.join(str(c) for c in cmd)}")
    r = subprocess.run([str(c) for c in cmd], cwd=cwd, env=env)
    if r.returncode != 0:
        sys.exit(f"step failed: {cmd[1] if len(cmd) > 1 else cmd}")
    print(f"    done in {time.time()-t:.0f}s")

def check_inputs():
    missing = [k for k, p in INPUTS.items() if not p.exists() and k != "hours_by_type"]
    if missing:
        sys.exit("Missing input files:\n  " + "\n  ".join(f"{k}: {INPUTS[k].name}" for k in missing) + "\nSee docs/data_dictionary.md for how to pull each report.")

def recalc():
    for key in ("wo", "pm", "assets", "inv"):
        p = OUTPUTS[key]
        print(f"\n>>> recalc {p.name}")
        r = subprocess.run([sys.executable, str(ROOT / "tools" / "recalc.py"), str(p), "300"], capture_output=True, text=True, cwd=ROOT)
        try:
            res = json.loads(r.stdout)
        except json.JSONDecodeError:
            print(r.stdout, r.stderr); sys.exit("recalc did not return JSON — is LibreOffice installed? Use --no-recalc to skip.")
        print(f"    {res.get('status')}  formulas={res.get('total_formulas')}  errors={res.get('total_errors')}")
        if res.get("status") == "errors_found":
            print(json.dumps(res.get("error_summary"), indent=1)); sys.exit("formula errors — fix before delivering")
        if "error" in res:
            print(res); sys.exit("recalc error")

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-report", action="store_true"); ap.add_argument("--no-recalc", action="store_true")
    ap.add_argument("--only", nargs="*", help="run only these step names (e.g. s01 s07)")
    a = ap.parse_args()
    check_inputs()
    steps = STEPS if not a.only else [s for s in STEPS if any(s.startswith(o) for o in a.only)]
    for s in steps:
        run([sys.executable, ROOT / "src" / f"{s}.py"])
    if a.only:
        return
    if not a.no_recalc:
        recalc()
    run([sys.executable, ROOT / "src" / "s08_charts.py"])
    run([sys.executable, ROOT / "src" / "s09_facts.py"])
    if not a.no_report:
        import os
        env = dict(os.environ, REPORT_REVISION=REPORT_REVISION)
        run(["node", ROOT / "report" / "build_report.js"], env=env)
    print("\nAll outputs are in", OUTPUTS["wo"].parent)

if __name__ == "__main__":
    main()
