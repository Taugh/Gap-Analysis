"""Step 5 — aggregate the users-by-site export against AM-001 Section 6.
Personal data: this script reads names and emails but writes ONLY aggregate counts.
Inputs : UsersBySite.csv
Outputs: work/group_counts.csv, work/user_facts.json
"""
import sys, csv, json
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import pandas as pd, numpy as np
from collections import Counter
from config import INPUTS, WORK, SNAPSHOT, FMT_USERS

rows = []; site = None; hdr = None
for r in csv.reader(open(INPUTS["users_by_site"], encoding="utf-8-sig")):
    if not r: continue
    if r[0] == "Site Name": hdr = "next"; continue
    if hdr == "next": site = r[0]; hdr = "h"; continue
    if len(r) > 1 and r[1] == "Full Name": continue
    if len(r) >= 8 and (r[1] or r[2]): rows.append([site] + r[1:9])
df = pd.DataFrame(rows, columns=["Site", "Full Name", "User Name", "Title", "Status", "Email", "Groups", "Last Login", "Hourly Rate"])
df["Groups"] = df["Groups"].str.replace("Group - Can Import, Add and Modify Parts", "Group - Can Import Add and Modify Parts")
df["key"] = np.where(df["User Name"].str.strip() != "", df["User Name"].str.lower().str.strip(), df["Full Name"].str.lower().str.strip() + "|" + df["Site"])
u = df.groupby("key").agg(name=("Full Name", "first"), title=("Title", "first"), status=("Status", "first"),
                          groups=("Groups", lambda s: sorted(set(g.strip() for x in s for g in x.split(",") if g.strip()))),
                          sites=("Site", lambda s: sorted(set(s))), login=("Last Login", "first")).reset_index()
u["n_groups"] = u.groups.str.len()
is_sys = u.name.str.contains("SysUser", case=False); is_guest = u.name.eq("Guest")
is_test = u.name.str.contains(r"test|New User|User released|^Fiix$", case=False, regex=True) & ~is_sys
h = u[~is_sys & ~is_guest]; ha = h[h.status == "Active"]
ld = pd.to_datetime(ha.login, format=FMT_USERS, errors="coerce")
F = {
 "accounts": len(u), "system": int(is_sys.sum()), "guest_sites": int(df[df["Full Name"] == "Guest"].Site.nunique()),
 "test": int(is_test.sum()), "test_active": int((is_test & (u.status == "Active")).sum()),
 "human": len(h), "human_active": int((h.status == "Active").sum()), "human_inactive": int((h.status == "Inactive").sum()),
 "avg_groups": round(float(ha.n_groups.mean()), 2), "multi_group": int((ha.n_groups > 1).sum()), "multi_group_pct": int(round((ha.n_groups > 1).mean() * 100)), "max_groups": int(ha.n_groups.max()),
 "addon_users": int(ha.groups.apply(lambda g: any(x.startswith("Group -") for x in g)).sum()),
 "admins_total": int(u.groups.apply(lambda g: "Administrators" in g).sum()), "admins_active_human": int(ha.groups.apply(lambda g: "Administrators" in g).sum()),
 "admins_system": int((is_sys & u.groups.apply(lambda g: "Administrators" in g)).sum()),
 "cmms_mgmt_active": int(ha.groups.apply(lambda g: "CMMS Management" in g).sum()),
 "never_login": int((ha.login == "").sum()), "login_gt90": int(((SNAPSHOT - ld).dt.days > 90).sum()), "login_gt365": int(((SNAPSHOT - ld).dt.days > 365).sum()),
 "inactive_with_groups": int(((h.status == "Inactive") & (h.n_groups > 0)).sum()),
 "leaders_active": int(ha.groups.apply(lambda g: "Maintenance - Leaders" in g).sum()), "technicians_active": int(ha.groups.apply(lambda g: "Technicians" in g).sum()),
 "technicians_mgmt": int(ha[ha.groups.apply(lambda g: "Technicians" in g)].title.str.contains("Manager|Head|Vice|Director|Officer", case=False).sum()),
 "masterdata_rights": int(ha.groups.apply(lambda g: any(("Import" in x or "Modify" in x) and x.startswith("Group -") for x in g)).sum()),
 "vendor_active": int(ha.title.str.contains("Hunton|Service -", case=False).sum()),
 "multi_site_users": int((u.sites.str.len() > 1).sum()),
}
json.dump(F, open(WORK["user_facts"], "w"), indent=1)
c_all = Counter(g for gs in u.groups for g in gs); c_act = Counter(g for gs in ha.groups for g in gs)
pd.DataFrame({"all": c_all, "active_human": c_act}).fillna(0).astype(int).to_csv(WORK["group_counts"])
print(f"accounts {F['accounts']}, active people {F['human_active']}, multi-group {F['multi_group_pct']}%")
