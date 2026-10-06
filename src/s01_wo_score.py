"""Step 1 — score the closed work order export against the AM-001 / Strategy rules.
Inputs : ClosedWorkOrders.csv
Outputs: work/wo_scored.pkl (records + flags), work/bad_actors.pkl (AM-005 trigger candidates)
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from config import *
import pandas as pd, numpy as np, re

df = pd.read_csv(INPUTS['closed_wo'], dtype=str, keep_default_na=False)
for c in df.columns:
    df[c] = df[c].str.strip()

fmt = FMT_CLOSED_WO
created = pd.to_datetime(df['DateCreated'], format=fmt)
completed = pd.to_datetime(df['Date Completed'], format=fmt)
suggested = pd.to_datetime(df['Suggested Completion'], format=fmt, errors='coerce')
df['Year Completed'] = completed.dt.year
df['Month Completed'] = completed.dt.to_period('M').astype(str)
df['Assets on WO'] = df['Asset Name'].str.count(',') + 1

mt = df['Maintenance Type']
status = df['Status']
is_completed = status.eq('Closed, Completed')
is_pm = mt.isin(['Preventive', 'Inspection'])
is_fail = mt.isin(['Corrective', 'Breakdown', 'Damage', 'Warranty'])
ah = pd.to_numeric(df['Actual Hours'], errors='coerce')

PLACEHOLDER = re.compile(
    r'^(completed?|done|complete[d]?\.?|job (complete|done)\.?|pm completed|n/?a|none|ok|good|see attachment|'
    r'#name\?|completed -|completed by|not active\.?|overlapping|duplicate( wo)?|not required\.?|no longer need)$',
    re.I)
notes = df['Completion Notes']
notes_norm = notes.str.lower().str.replace(r'\s+', ' ', regex=True).str.strip()
placeholder_hit = notes_norm.str.match(PLACEHOLDER) | ((notes_norm != '') & (notes_norm.str.len() < 12))
df['Note Placeholder'] = np.where(placeholder_hit, notes_norm, '')

site_dept = SITE_DEPT
expected_dept = df['Site Name'].map(site_dept)

# ---- Gap flags: 1 = record fails the rule, 0 = passes, '' = rule not applicable
def flag(cond, applicable=None):
    out = cond.astype(int)
    if applicable is not None:
        out = out.where(applicable, other=np.nan)
    return out

flags = {}
# Completeness — required for every closed WO
flags['G01 Priority missing']            = flag(df['Priority'].eq(''))
flags['G02 Maint Type missing']          = flag(mt.eq(''))
flags['G03 Site missing']                = flag(df['Site Name'].isin(['', '(No Site)']))
flags['G04 Assigned Users missing']      = flag(df['Assigned Users'].eq(''))
flags['G05 Completed By missing']        = flag(df['Completed by'].eq(''))
flags['G06 Account Code missing']        = flag(df['Acc Code'].eq(''))
flags['G07 Charge Dept missing']         = flag(df['Charge Dept Code'].eq(''))
flags['G08 Asset Description missing']   = flag(df['Asset Description'].eq(''))
flags['G09 Target Date missing']         = flag(df['Suggested Completion'].eq(''))
# Completeness — conditional
flags['G10 Completion Notes missing']    = flag(notes.eq(''), is_completed)
flags['G11 Completion Notes placeholder']= flag(placeholder_hit, is_completed)
flags['G12 Actual Hours missing']        = flag(df['Actual Hours'].eq(''), is_completed)
flags['G13 Actual Hours zero']           = flag(ah.eq(0), is_completed)
flags['G14 Est Hours missing']           = flag(df['Est Hours'].eq(''), is_completed)
flags['G15 Sched Maint ID missing (PM/Insp)'] = flag(df['Scheduled Maintenance ID'].eq(''), is_pm)
flags['G16 Cause missing (failure WO)']  = flag(df['Cause'].eq(''), is_fail)
flags['G17 Solution missing (failure WO)'] = flag(df['Solution'].eq(''), is_fail)
flags['G18 Requested By missing (failure WO)'] = flag(df['Requested By'].eq('') | df['Requested By'].eq('Guest'), is_fail)
# Validity
flags['V01 Completed before Created (>1 day)'] = flag((created - completed).dt.total_seconds() > 86400)
flags['V02 Target before Created']       = flag(suggested < created, suggested.notna())
flags['V03 Actual Hours > 24 (review)']  = flag(ah > 24, ah.notna())
flags['V04 Charge Dept != Site']         = flag(df['Charge Dept Code'].ne(expected_dept) & df['Charge Dept Code'].ne('') & expected_dept.notna())
flags['V05 Rejected but hours booked']   = flag(df['Actual Hours'].ne(''), status.eq('Rejected'))
flags['V06 Closed but not Completed/Incomplete/Rejected'] = flag(~status.isin(['Closed, Completed', 'Closed, Incomplete', 'Rejected']))
flags['V07 Sched Maint ID on non-PM type'] = flag(df['Scheduled Maintenance ID'].ne(''), ~is_pm & mt.ne(''))
flags['V08 Days Open > 90']              = flag(df['Days Open'].astype(int) > 90)
flags['V09 Completed after Target (late)'] = flag(completed > suggested, is_completed & suggested.notna())
flags['V10 Closed Incomplete/Rejected as "not active"'] = flag(notes_norm.str.startswith('not active') | notes_norm.str.contains('not installed|does not exist|under construction'), status.isin(['Closed, Incomplete','Rejected']))

# ---- AM-001 derived checks
is_highest = df['Priority'].eq('Highest - This Week')
flags['V11 Highest priority on non-Breakdown work (AM-001 5.5)'] = flag(is_highest & ~mt.isin(['Breakdown', 'Damage']), is_highest)
flags['V12 Category should be a flag, not a type (AM-001 G-5.10)'] = flag(mt.isin(['Safety', 'Warranty', 'Damage']))
flags['V13 Category = Other (AM-001 G-5.11 cap 2%)'] = flag(mt.eq('Other'))
# Reporting standard composite (AM-001 G-5.14 / G-5.15 / G-5.16), evaluated on Closed, Completed
meets = (df['Asset Code'].ne('') & mt.ne('') & df['Priority'].ne('') & df['Assigned Users'].ne('')
         & df['Actual Hours'].ne('') & ah.fillna(0).gt(0)
         & notes.ne('') & ~placeholder_hit
         & (~is_fail | (df['Cause'].ne('') & df['Solution'].ne(''))))
flags['R01 Fails AM-001 reporting standard (G-5.14..16 composite)'] = flag(~meets, is_completed)
# PM compliance proxy: schedule-generated PM/Inspection completed on or before target date
is_sched_pm = is_pm & df['Scheduled Maintenance ID'].ne('')
flags['R02 Scheduled PM not completed by target (PM compliance proxy)'] = flag(~(is_completed & (completed <= suggested)), is_sched_pm & suggested.notna())

# ---- Strategy-derived: failure WO written against a location/building/area record rather than Level-3 equipment (Strategy 8, G-8.1)
LOC_CATS = ('Locations And Facilities', 'Buildings', 'Road', 'Structural', 'Civil', 'Culvert', 'HSE', 'Water', 'Electrical Infrastructure - General', 'Mechanical Infrastructure - General', 'HPC')
is_loc = df['Asset Category'].str.startswith(LOC_CATS) | df['Asset Code'].str.contains('-STORAGE-|-DATACENTERS$|-MOBILE-OTHER-|-ELEC-SUPPLY$', regex=True)
df['Record Level'] = np.where(is_loc, 'Location / Area', 'Equipment')
flags['H01 Failure WO on location/area record, not equipment (Strategy 8)'] = flag(is_loc, is_fail)

F = pd.DataFrame(flags)
df = pd.concat([df, F], axis=1)
core = [c for c in F.columns if c.startswith('G')]
df['Gap Count'] = F[core].fillna(0).sum(axis=1).astype(int)
df['Gaps'] = F[core].apply(lambda r: '; '.join(c[:3] for c, v in r.items() if v == 1), axis=1)

df.to_pickle(WORK['wo_scored'])

# ---- AM-005 bad-actor candidates: corrective/breakdown/damage/warranty WOs per asset in any rolling 12 months
cm = df[is_fail & status.ne('Rejected')].assign(dc=completed).sort_values('dc')
rows = []
for code, g in cm.groupby('Asset Code'):
    d = g['dc'].values; best = 0
    for i in range(len(d)):
        best = max(best, int(((d >= d[i]) & (d <= d[i] + pd.Timedelta(days=365).to_timedelta64())).sum()))
    rows.append({'Asset Code': code, 'Asset Name': g['Asset Name'].iloc[0][:80], 'Asset Category': g['Asset Category'].iloc[0], 'Site Name': g['Site Name'].iloc[0],
                 'Record Level': g['Record Level'].iloc[0], 'Failure WOs (all)': len(g), 'Max failure WOs in any 12 months': best,
                 'First': g['dc'].min().date(), 'Last': g['dc'].max().date(),
                 'Trigger if Group A (>=2)': 'Yes' if best >= 2 else '', 'Trigger if Group B/C/D (>=3)': 'Yes' if best >= 3 else '',
                 'Cause recorded on any WO': 'Yes' if g['Cause'].ne('').any() else 'No'})
ba = pd.DataFrame(rows).sort_values(['Max failure WOs in any 12 months', 'Failure WOs (all)'], ascending=False)
ba.to_pickle(WORK['bad_actors'])
print('bad actor candidates >=3:', (ba['Max failure WOs in any 12 months'] >= 3).sum(), ' equipment-level:', ((ba['Max failure WOs in any 12 months'] >= 3) & (ba['Record Level'] == 'Equipment')).sum())

n = len(df)
print('records', n)
summary = pd.DataFrame({
    'applicable': F.notna().sum(),
    'fails': F.eq(1).sum(),
})
summary['pct'] = (summary.fails / summary.applicable * 100).round(1)
print(summary.to_string())
print('\nrecords with >=1 core gap:', (df['Gap Count'] > 0).sum(), (df['Gap Count'] > 0).mean().round(3))
print(df['Gap Count'].value_counts().sort_index().to_string())
print('\nplaceholders top'); print(df.loc[df['Note Placeholder'] != '', 'Note Placeholder'].value_counts().head(15).to_string())
