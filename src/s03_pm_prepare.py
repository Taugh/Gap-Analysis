"""Step 3 — parse the scheduled maintenance list, explode asset references, compute PM coverage.
Inputs : ScheduledMaintenanceList.csv, work/assets_scored.pkl, work/wo_scored.pkl
Outputs: work/sm.pkl, work/sm_assets.pkl, work/assets_cov.pkl
"""
import sys, re
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import pandas as pd, numpy as np
from config import INPUTS, WORK

df = pd.read_csv(INPUTS["schedules"], dtype=str, keep_default_na=False)
df.columns = [c.strip() for c in df.columns]
for c in df.columns:
    df[c] = df[c].str.strip()
# tolerate the two header variants seen in exports
df = df.rename(columns={"SM Status 1 = Active, 0 = Paused": "SM Status",
                        "Suggested Completion in Days After Work Order creation": "Suggested Completion"})
required = {"When", "SM Status", "Code", "Priority", "Assets", "Assigned Users", "Time Estimated Hours", "Type", "Description", "Site", "Suggested Completion"}
missing = required - set(df.columns)
if missing:
    raise SystemExit(f"ScheduledMaintenanceList.csv is missing columns: {sorted(missing)}")

a = pd.read_pickle(WORK["assets_scored"])
reg = a.drop_duplicates("Code").set_index("Code")

# ---- explode schedule × asset (asset code is the parenthesised token at the end of each asset name)
rows = []
for r in df.itertuples():
    codes = re.findall(r"\(([^()]+)\)", r.Assets)
    codes = [c for c in codes if c in reg.index or re.match(r"^[A-Z]{2,4}(-[A-Z0-9]+)+$", c)]
    for c in codes:
        rows.append((r.Code, getattr(r, "_2"), r.Type, r.Priority, r.When, r.Site, c))
ex = pd.DataFrame(rows, columns=["SM", "Active", "Type", "Priority", "When", "Site", "Asset"])
ex["level"] = ex.Asset.map(reg["Record Level"])
ex["crit"] = ex.Asset.map(reg["Asset Criticality"]).fillna("")
ex["cat"] = ex.Asset.map(reg["Category"])

# ---- coverage: Direct / Via parent / Paused only / None
act = set(ex[ex.Active == "1"].Asset); anyS = set(ex.Asset)
parent = dict(zip(a["Code"], a["Parent Code"]))
def cov(c):
    if c in act: return "Direct"
    p = parent.get(c, ""); hops = 0
    while p and hops < 8:
        if p in act: return "Via parent"
        p = parent.get(p, ""); hops += 1
    return "Paused only" if c in anyS else "None"
a["PM coverage"] = a["Code"].map(cov)

df.to_pickle(WORK["sm"]); ex.to_pickle(WORK["sm_assets"]); a.to_pickle(WORK["assets_cov"])
eq = a[a["Record Level"] == "Equipment"]
print(f"schedules {len(df):,} (active {(df['SM Status']=='1').sum():,}); asset refs {len(ex):,}; equipment coverage:")
print(eq["PM coverage"].value_counts().to_string())
