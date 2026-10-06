"""Step 8 — the four report figures (PNG) from the intermediate files.
Outputs: outputs/charts/fig_wo_trend.png, fig_serial_batch.png, fig_pm_coverage.png, fig_crit_site.png
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import pandas as pd, numpy as np, matplotlib
matplotlib.use("Agg"); import matplotlib.pyplot as plt
from config import WORK, OUTPUTS, SNAPSHOT

CH = OUTPUTS["charts"]
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 9, "axes.spines.top": False, "axes.spines.right": False, "axes.edgecolor": "#c9c8c2",
                     "axes.linewidth": 0.8, "xtick.color": "#52514e", "ytick.color": "#52514e", "axes.labelcolor": "#52514e", "text.color": "#0b0b0b"})
S1, S2, S3, S4, GRAY, SURF = "#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#9a9891", "#ffffff"
def grid(ax, axis): ax.grid(axis=axis, color="#e6e5e0", lw=0.8); ax.set_axisbelow(True)

# 1 — closure quality by year
wo = pd.read_pickle(WORK["wo_scored"]); c = wo[(wo.Status == "Closed, Completed") & (wo["Year Completed"] >= 2024)]
r01 = [x for x in wo.columns if x.startswith("R01")][0]
g = c.groupby("Year Completed").agg(quality=(r01, lambda s: 1 - s.mean()), notes=("Completion Notes", lambda s: (s != "").mean()), hours=("Actual Hours", lambda s: (s != "").mean()))
fig, ax = plt.subplots(figsize=(6.2, 3.0), dpi=200); x = g.index.astype(str)
for col, lab, colr in [("notes", "Completion notes present", S1), ("hours", "Actual hours recorded", S2), ("quality", "Meets AM-001 reporting standard", S3)]:
    ax.plot(x, g[col] * 100, color=colr, lw=2, marker="o", ms=6, mec=SURF, mew=1.5, label=lab)
    ax.annotate(f"{g[col].iloc[-1]*100:.0f}%", (x[-1], g[col].iloc[-1] * 100), xytext=(6, 0), textcoords="offset points", va="center", fontsize=8)
ax.axhline(95, color=GRAY, lw=1, ls=(0, (4, 3))); ax.text(x[0], 96.5, "Target 95%", fontsize=8, color="#52514e")
ax.set_ylim(0, 105); ax.set_ylabel("% of completed work orders"); grid(ax, "y"); ax.legend(frameon=False, fontsize=8, loc="lower right")
ax.set_title(f"Work order closure quality by year completed ({SNAPSHOT.year} year to date)", fontsize=10, loc="left"); fig.tight_layout(); fig.savefig(CH / "fig_wo_trend.png"); plt.close()

# 2 — serial completeness by creation batch (batches with 100+ equipment records)
a = pd.read_pickle(WORK["assets_scored"]); a7 = [x for x in a.columns if x.startswith("A07")][0]
eq = a[a["Record Level"] == "Equipment"]
b = eq.groupby("Creation batch").agg(n=("Code", "size"), blank=(a7, "mean")).reset_index(); b = b[b.n >= 100]
b["key"] = b["Creation batch"].str.replace("Bulk load ", "").replace("Manual / small batch", "zzz"); b = b.sort_values("key")
b["label"] = [("Hand-entered records" if "Manual" in x else x) + f"  (n={n:,})" for x, n in zip(b["Creation batch"], b["n"])]
fig, ax = plt.subplots(figsize=(6.2, 3.6), dpi=200); ax.barh(b["label"], b["blank"] * 100, color=S1, height=0.55)
for i, v in enumerate(b["blank"] * 100): ax.text(v + 1, i, f"{v:.0f}%", va="center", fontsize=8)
ax.invert_yaxis(); ax.set_xlim(0, 110); ax.set_xlabel("% of equipment records with no serial number"); grid(ax, "x")
ax.set_title("Serial-number gap by asset creation batch", fontsize=10, loc="left"); fig.tight_layout(); fig.savefig(CH / "fig_serial_batch.png"); plt.close()

# 3 — PM coverage by category
ac = pd.read_pickle(WORK["assets_cov"]); eq = ac[ac["Record Level"] == "Equipment"]
top = eq["Category"].value_counts().head(14).index
g = eq[eq.Category.isin(top)].groupby(["Category", "PM coverage"]).size().unstack(fill_value=0).reindex(columns=["Direct", "Via parent", "Paused only", "None"], fill_value=0).loc[top]
g = g.div(g.sum(axis=1), axis=0) * 100
fig, ax = plt.subplots(figsize=(6.2, 4.2), dpi=200); left = np.zeros(len(g))
for col, colr in zip(g.columns, [S1, "#8fb8ea", S4, S2]):
    ax.barh([f"{c} (n={eq[eq.Category==c].shape[0]:,})" for c in g.index], g[col], left=left, color=colr, height=0.6, label=col, edgecolor=SURF, linewidth=1.5); left += g[col].values
ax.invert_yaxis(); ax.set_xlim(0, 100); ax.set_xlabel("% of equipment records"); ax.legend(frameon=False, fontsize=8, ncol=4, loc="upper center", bbox_to_anchor=(0.5, -0.14)); grid(ax, "x")
ax.set_title("PM schedule coverage, 14 largest equipment categories", fontsize=10, loc="left"); fig.tight_layout(); fig.savefig(CH / "fig_pm_coverage.png"); plt.close()

# 4 — criticality assignment by site
g = eq.assign(has=eq["Asset Criticality"] != "").groupby("Site")["has"].agg(["mean", "size"]).sort_values("size", ascending=False)
fig, ax = plt.subplots(figsize=(6.2, 3.0), dpi=200); ax.barh([f"{s} (n={n:,})" for s, n in zip(g.index, g["size"])], g["mean"] * 100, color=S1, height=0.55)
for i, v in enumerate(g["mean"] * 100): ax.text(v + 1, i, f"{v:.0f}%", va="center", fontsize=8)
ax.invert_yaxis(); ax.set_xlim(0, 110); ax.set_xlabel("% of equipment with a criticality group assigned"); grid(ax, "x")
ax.set_title("Criticality assignment by site", fontsize=10, loc="left"); fig.tight_layout(); fig.savefig(CH / "fig_crit_site.png"); plt.close()
print("charts written to", CH)
