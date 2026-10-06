"""Step 4 — score the PM schedule master and build the PM workbook.
Inputs : work/sm.pkl, work/sm_assets.pkl, work/assets_cov.pkl, work/wo_scored.pkl
Outputs: outputs/PM_Schedule_Gap_Analysis.xlsx
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from config import *
import pandas as pd, numpy as np, re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import ColorScaleRule

df = pd.read_pickle(WORK['sm']); ex = pd.read_pickle(WORK['sm_assets']); a = pd.read_pickle(WORK['assets_cov'])
wo = pd.read_pickle(WORK['wo_scored'])
reg = a.drop_duplicates('Code').set_index('Code')
df = df.rename(columns={'SM Status': 'Active (1) / Paused (0)', 'Time Estimated Hours': 'Est Hours', 'Suggested Completion': 'Days to complete'})
df = df.drop(columns=['Next Trigger Threshold Date'])
n = len(df)
active = df['Active (1) / Paused (0)'].eq('1')
w = df['When']; pr = df['Priority']; ty = df['Type']
days = pd.to_numeric(df['Days to complete'], errors='coerce')

# ---- derived
df['Asset count'] = df['Assets'].str.count(r'\([^()]*\)')
weeks = w.str.extract(r'Every (\d+) weeks')[0].astype(float)
is_weekly = w.str.match(r'^Every (Mon|Tues|Wednes|Thurs|Fri|Satur|Sun)day$')
is_hours = w.str.contains('hours|Event', regex=True)
is_calendar = w.str.contains(r'day of every month|Every year|Every \d+ years|months on the|Every day', regex=True)
CONF = {1, 2, 4, 6, 12, 24, 52, 104, 156}   # AM-001 7.3: W, BW, M(28d), BM(42d), Q(84d), S(168d), A(364d), B(728d), T(1092d)
def trig(i):
    if w[i] == '': return 'None visible'
    if is_hours[i]: return 'Meter / event'
    if is_weekly[i] or (not np.isnan(weeks[i]) and weeks[i] in CONF): return 'Week-based (7.3 code)'
    if not np.isnan(weeks[i]): return 'Week-based (non-7.3 interval)'
    if is_calendar[i]: return 'Calendar-based'
    return 'Complex / other'
df['Trigger class'] = [trig(i) for i in range(n)]
WINDOW = {'Highest - This Week': 7, 'High - Next Week': 14, 'Medium - Within 2 Weeks': 14, 'Low - Within 4 Weeks': 28, 'Lowest - Within 8 Weeks': 56}
df['Priority window (days)'] = pr.map(WINDOW)
sm_level = ex.groupby('SM')['level'].apply(lambda s: (s == 'Location / Area').sum())
df['Location-level asset refs'] = df['Code'].map(sm_level).fillna(0).astype(int)
act_ex = ex[ex.Active == '1']
k = act_ex.groupby(['Asset', 'Type', 'When']).SM.nunique()
ov_pairs = set(k[k > 1].index)
ov_sm = set(act_ex[[ (r.Asset, r.Type, r.When) in ov_pairs for r in act_ex.itertuples()]].SM)
df['Overlapping active schedule'] = np.where(df['Code'].isin(ov_sm), 'Yes', '')
v10 = [c for c in wo.columns if c.startswith('V10')][0]; na = set(wo.loc[wo[v10] == 1, 'Asset Code'])
na_sm = set(act_ex[act_ex.Asset.isin(na)].SM)
df['Covers asset with not-active WO closures'] = np.where(df['Code'].isin(na_sm), 'Yes', '')
crit_sm = ex.groupby('SM')['crit'].apply(lambda s: (s != '').sum())
df['Asset refs with criticality'] = df['Code'].map(crit_sm).fillna(0).astype(int)
df['Site (short)'] = df['Site'].str.replace(r'\s*\(.*\)$', '', regex=True)

def flag(cond, applicable=None):
    out = cond.astype(int)
    if applicable is not None: out = out.where(applicable, other=np.nan)
    return out
flags = {}
flags['S01 Priority = Highest on a routine template']     = flag(pr.eq('Highest - This Week'))
flags['S02 Priority blank']                                = flag(pr.eq(''))
flags['S03 Estimated hours blank']                         = flag(df['Est Hours'].eq(''))
flags['S04 Active with no trigger visible']                = flag(w.eq(''), active)
flags['S05 Interval not an AM-001 7.3 frequency code']     = flag(df['Trigger class'].isin(['Week-based (non-7.3 interval)', 'Calendar-based']), w.ne(''))
flags['S06 Days-to-complete exceeds priority window']      = flag(days > df['Priority window (days)'], days.notna() & df['Priority window (days)'].notna())
flags['S07 Type is not PM / Inspection / Meter reading']   = flag(~ty.isin(['Preventive', 'Inspection', 'Meter reading']))
flags['S08 Code outside SITE-TYPE-nnnn convention']        = flag(~df['Code'].str.match(r'^[A-Z]+(-[A-Z]+)?-[A-Z]+-\d+$'))
flags['S09 No assets attached']                            = flag(df['Asset count'].eq(0))
flags['S10 Assigned Users blank']                          = flag(df['Assigned Users'].eq(''))
flags['S11 Test / dummy schedule']                         = flag(df['Description'].str.contains(r'\btest\b|ignore|dummy', case=False, regex=True))
flags['S12 References location-level records']             = flag(df['Location-level asset refs'] > 0, df['Asset count'] > 0)
flags['S13 Overlap candidate (same asset, type & frequency)'] = flag(df['Overlapping active schedule'].eq('Yes'), active)
flags['S14 Active on asset with not-active WO closures']   = flag(df['Covers asset with not-active WO closures'].eq('Yes'), active)
flags['S15 Paused']                                        = flag(~active)
F = pd.DataFrame(flags); df = pd.concat([df, F], axis=1)
core = [c for c in F.columns if c[:3] not in ('S15',)]
df['Gap Count'] = F[core].fillna(0).sum(axis=1).astype(int)
data_cols = list(df.columns); rule_cols = list(F.columns); FIRST, LAST = 2, n + 1

FONT = 'Arial'
def f(bold=False, color='000000', size=10, italic=False): return Font(name=FONT, bold=bold, color=color, size=size, italic=italic)
HDR_FILL = PatternFill('solid', fgColor='1F3864'); INPUT_FILL = PatternFill('solid', fgColor='FFFF00')
thin = Side(style='thin', color='BFBFBF'); BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
PCT = '0.0%'; INT = '#,##0'
def header(ws, row, values, col=1, height=30):
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=col + i, value=v); c.font = f(True, 'FFFFFF'); c.fill = HDR_FILL; c.border = BORDER; c.alignment = Alignment(wrap_text=True, vertical='center')
    ws.row_dimensions[row].height = height
def style_range(ws, r1, r2, c1, c2, numfmt=None, bold=False, color='000000', wrap=False):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c); cell.font = f(bold, color); cell.border = BORDER
            if numfmt: cell.number_format = numfmt
            if wrap: cell.alignment = Alignment(wrap_text=True, vertical='top')
def title(ws, text, sub=None):
    ws['A1'] = text; ws['A1'].font = f(True, size=14)
    if sub: ws['A2'] = sub; ws['A2'].font = f(italic=True, color='595959')

wb = Workbook()
wsD = wb.active; wsD.title = 'Data'
wsD.append(data_cols)
for row in df.itertuples(index=False): wsD.append([None if (isinstance(v, float) and np.isnan(v)) else v for v in row])
for c in range(1, len(data_cols) + 1):
    cell = wsD.cell(row=1, column=c); cell.font = f(True, 'FFFFFF'); cell.fill = HDR_FILL; cell.alignment = Alignment(wrap_text=True, vertical='center'); wsD.column_dimensions[L(c)].width = 8 if data_cols[c-1] in rule_cols else 14
wsD.column_dimensions[L(data_cols.index('Assets') + 1)].width = 40; wsD.column_dimensions[L(data_cols.index('Description') + 1)].width = 40
wsD.freeze_panes = 'D2'; wsD.auto_filter.ref = f'A1:{L(len(data_cols))}{LAST}'; wsD.row_dimensions[1].height = 60
col = {nm: L(i + 1) for i, nm in enumerate(data_cols)}
def rng(nm): return f"Data!${col[nm]}${FIRST}:${col[nm]}${LAST}"
sites = list(df['Site (short)'].value_counts().index)
types = list(df['Type'].replace('', '(blank)').value_counts().index)

RULES = {
 'S01': ('Priority = Highest on a routine template', 'Priority', 'High', 'AM-001 5.5', '2,006 of 3,005 templates (67%) generate Highest-priority WOs. 5.5 reserves Highest for Breakdown/Emergency/T4/Group A down; a scheduled greasing route cannot be Highest by definition. This is the direct source of the 69% Highest rate on closed WOs (V11).', 'Template default never changed; priority not derived from criticality or frequency.', 'Reset template priorities: Medium/Low for routine PM and inspections, High only for statutory or Group A tasks; derive from criticality once assigned.'),
 'S02': ('Priority blank', 'Completeness', 'Low', 'AM-001 G-5.14', 'WOs generate without a priority.', 'Not mandatory on template.', 'Mandatory on template.'),
 'S03': ('Estimated hours blank', 'Completeness', 'High', 'AM-001 5.3 step 7 (job plan, resources); §10 backlog KPI', '73% of templates carry no duration, so 74% of closed WOs and 53% of open WOs have no estimate and backlog-in-weeks cannot be computed.', 'Duration not entered when templates were built.', 'Add estimated hours to every template — one entry fixes every future WO.'),
 'S04': ('Active with no trigger visible', 'Validity', 'High', 'Strategy G-9.7; AM-001 G-7.8', '300 active schedules show no recurrence — 299 of them meter-based greasing routes at Mackenzie and Prince George. Either the meter trigger is not shown by the report, or these schedules never fire. Verify in the system.', 'Meter-based trigger not configured, or report does not expose meter triggers.', 'Check a sample in the CMMS; if unconfigured, either set the hours trigger or pause them so the PM program reflects reality.'),
 'S05': ('Interval not an AM-001 7.3 frequency code', 'Consistency', 'Medium', 'AM-001 7.3 (W, BW, M=28d, BM=42d, Q=84d, S=168d, A=364d, B=728d, T=1092d)', '183 schedules run every 26 weeks (7.3 defines semi-annual as 24 weeks / 168 days), 28 every 5 weeks, plus 44 calendar-date schedules ("every year on Sept 30th") that drift against the 52-week calendar. Compliance windows (10%) cannot be applied to non-standard intervals.', 'Templates predate the 7.3 table; some frequencies copied from OEM calendar dates.', 'Decide with the author whether 26-week semi-annual is accepted (and amend 7.3) or migrate to 24 weeks; convert calendar-date schedules to week-based.'),
 'S06': ('Days-to-complete exceeds the priority window', 'Consistency', 'Medium', 'AM-001 5.5 completion windows', 'The WO due date is set from this field, not from priority: e.g. 132 Highest templates allow 3–36 days. Late-completion and PM-compliance figures are therefore measured against arbitrary targets.', 'Two independent fields with no validation.', 'Derive Suggested Completion from priority (or from the 7.3 10% window for PMs) automatically.'),
 'S07': ('Type is not PM / Inspection / Meter reading', 'Consistency', 'Medium', 'AM-001 5.4, G-5.10', 'Safety (64), Other, Improvement and blank types on schedules; Safety routines are PMs with a safety flag under G-5.10, and are excluded from PM compliance today (V07).', 'Type list predates AM-001.', 'Re-type Safety routines as Preventive/Inspection with the Safety flag.'),
 'S08': ('Code outside SITE-TYPE-nnnn convention', 'Consistency', 'Low', 'Strategy G-8.4', '30 legacy "SMnnn" codes.', 'Early schedules created before the convention.', 'Re-code.'),
 'S09': ('No assets attached', 'Validity', 'Medium', 'Strategy G-9.6 (every task traceable to an asset/failure mode)', 'A schedule with no asset generates unattributed WOs.', 'Test schedules; assets removed after creation.', 'Attach or retire.'),
 'S10': ('Assigned Users blank', 'Completeness', 'Low', 'AM-001 G-5.14', 'WOs generate unassigned (4.9% of closed WOs).', 'Not set on template.', 'Set crew on template.'),
 'S11': ('Test / dummy schedule', 'Validity', 'Low', 'AM-002 G-12.6 / Strategy G-8.5 (master data control)', 'Test schedules live in production.', 'Not cleaned up.', 'Retire.'),
 'S12': ('References location-level records', 'Hierarchy', 'Medium', 'Strategy 8 (Level 3)', '24% of active asset references point at buildings, sections or areas rather than equipment. Work is real but history does not reach the asset (H01 in the WO analysis). Some is legitimate (building inspections).', 'Equipment not registered, or routes built on sections for convenience.', 'Where the task is on equipment, attach the equipment; keep location-level only for genuine facility tasks.'),
 'S13': ('Overlap candidate (same asset, type & frequency)', 'Duplication', 'Medium', 'AM-001 5.4 (inspection steps recorded under the PM); Strategy 9.4 stage 2 (eliminate duplicates, bundle tasks)', '372 asset/type/frequency combinations are covered by more than one active schedule (466 templates). No two templates share an identical description, so these are not straight duplicates: most are motor-greasing and shaft-greasing templates on the same fan, or per-DC switchgear filter routes — separate tasks that 9.4 stage 2 says should be bundled into one template per asset. The "overlapping" closures seen on Childress inspections are the remainder; review the Asset References sheet with Overlap = Yes.', 'Templates built one-task-per-template and per section without bundling.', 'Bundle tasks of the same frequency into one template per asset; then re-check for true duplicates.'),
 'S14': ('Active on an asset whose PMs were closed as not active', 'Consistency', 'High', 'Strategy G-8.5 lifecycle status; WO rule V10', '84 active schedules cover 60 assets that technicians have already reported as not installed / not in service. They keep generating work.', 'Asset status not maintained; schedule not paused.', 'Pause these 84 now; set the 60 assets Inactive; tie generation to asset status.'),
 'S15': ('Paused', 'Info', 'Info', '—', '1,340 schedules (45%) are paused, 1,165 of them meter-based greasing routes with no trigger. A paused schedule is invisible to PM compliance; worth confirming each is intentionally paused rather than abandoned.', 'Meter-based program switched off or never commissioned at some sites.', 'Review paused list by site; retire or reactivate.'),
}

# ---- Read Me
ws = wb.create_sheet('Read Me', 0)
title(ws, 'CMMS Gap Analysis — PM Schedules (scheduled maintenance)', f'Source: ScheduledMaintenanceListUpdated.csv, {n:,} schedules ({active.sum():,} active, {(~active).sum():,} paused), {ex.Asset.nunique():,} distinct assets referenced. Prepared ' + PREPARED_LABEL + '. Draft for review.')
lines = [
 ('Purpose', 'Assess the PM schedule master — the templates that generate Preventive, Inspection and Meter-reading work orders — against AM-001 (5.4 categories, 5.5 priority, 7.3 frequency codes, G-5.14 mandatory fields) and the Strategy (Section 9 tactics, G-9.7, hierarchy Level 3), and measure PM coverage of the asset register.'),
 ('Headline', 'The schedule master is the root cause of several work order findings: 67% of templates are Highest priority, 73% have no estimated hours, due dates are set independently of priority, 372 asset/type/frequency combinations are covered by more than one active template (mostly un-bundled tasks on the same asset), and 84 active schedules cover assets already reported as not in service. Coverage: 38% of equipment (8,352 records) has no active schedule, directly or through a parent location; among the 1,525 assets that have a criticality group, all 158 Group A with an active schedule are covered but 78 Group A UPS units are covered only by paused schedules.'),
 ('Coverage method', 'An asset is Direct if it is attached to an active schedule; Via parent if any ancestor in the hierarchy is; Paused only if it is attached to paused schedules only; None otherwise. Via-parent coverage is credited because many routes (PDU sections, fan rows) are built on the section record. HPC servers and racks are shown but are likely vendor-maintained — confirm before treating as a gap.'),
 ('How to read it', 'Summary, Gap Rules, Coverage (equipment by category and site), Group A/B Coverage, By Site / By Type heat maps, Asset References (one row per schedule × asset, for overlap and location-level filtering) and Data (one row per schedule with a 1/0/blank flag per rule).'),
 ('Assumptions', 'Asset codes are parsed from the parenthesised code at the end of each asset name. Priority windows per AM-001 5.5: Highest 7 days, High 14, Medium 14, Low 28, Lowest 56. Conforming intervals per 7.3: 1, 2, 4, 6, 12, 24, 52, 104, 156 weeks; 26 weeks is flagged for decision, not as an error. Overlap candidate = same asset, same Type and same frequency in more than one active schedule; no templates share an identical description, so these are un-bundled tasks or route duplicates to review rather than confirmed duplicates. The report does not expose meter-trigger thresholds, so "no trigger visible" must be verified in the system.'),
 ('Governing documents', 'AM-001 Rev 3.0 (09/09/2026), Strategy Rev 3.0 (08/21/2026), AM-002 Rev 1.0 (09/23/2026) — drafts pending approval; this is a current-state baseline against target state.'),
]
r = 4
for kk, v in lines:
    ws.cell(row=r, column=1, value=kk).font = f(True); c = ws.cell(row=r, column=2, value=v); c.font = f(); c.alignment = Alignment(wrap_text=True, vertical='top'); ws.row_dimensions[r].height = max(30, 15 * (len(v) // 110 + 1)); r += 1
ws.column_dimensions['A'].width = 18; ws.column_dimensions['B'].width = 120

# ---- Summary
ws = wb.create_sheet('Summary', 1)
title(ws, 'Summary — PM Schedule Master', 'Formulas over the Data sheet unless marked static.')
for c, wd in zip('ABCD', [60, 14, 14, 62]): ws.column_dimensions[c].width = wd
st = rng('Active (1) / Paused (0)')
r = 4; header(ws, r, ['Measure', 'Value', '% of base', 'Note']); r += 1
kp = [
 ('Schedules', f'=COUNTA({rng("Code")})', None, ''),
 ('  Active', f'=COUNTIF({st},"1")', '=B6/B5', ''),
 ('  Paused', f'=COUNTIF({st},"0")', '=B7/B5', ''),
 ('Asset references (schedule × asset)', len(ex), None, 'Static; distinct assets ' + f'{ex.Asset.nunique():,}'),
 ('Templates at Highest priority', f'=COUNTIF({rng("Priority")},"Highest - This Week")', '=B9/B5', 'AM-001 5.5'),
 ('Templates with no estimated hours', f'=COUNTIF({rng(rule_cols[2])},1)', '=B10/B5', ''),
 ('Active templates with no trigger visible', f'=COUNTIF({rng(rule_cols[3])},1)', '=B11/B6', 'Verify meter triggers in system'),
 ('Templates whose due-date setting exceeds the priority window', f'=COUNTIF({rng(rule_cols[5])},1)', f'=B12/COUNT({rng(rule_cols[5])})', 'Base: templates with both fields'),
 ('Active templates that are overlap candidates (same asset, type & frequency)', f'=COUNTIF({rng(rule_cols[12])},1)', '=B13/B6', 'Asset/type/frequency combinations affected: ' + f'{len(ov_pairs):,}'),
 ('Active templates on assets reported not in service', f'=COUNTIF({rng(rule_cols[13])},1)', '=B14/B6', f'{len(na & set(act_ex.Asset))} assets'),
 ('Active templates referencing location-level records', f'=COUNTIF({rng(rule_cols[11])},1)', f'=B15/COUNT({rng(rule_cols[11])})', ''),
 ('Templates with a non-PM type (Safety/Other/Improvement/blank)', f'=COUNTIF({rng(rule_cols[6])},1)', '=B16/B5', ''),
 ('Equipment records in register', int((a['Record Level'] == 'Equipment').sum()), None, 'Static, from asset register'),
 ('  covered by an active schedule directly', int(((a['Record Level'] == 'Equipment') & (a['PM coverage'] == 'Direct')).sum()), '=B18/B17', ''),
 ('  covered via a parent location', int(((a['Record Level'] == 'Equipment') & (a['PM coverage'] == 'Via parent')).sum()), '=B19/B17', ''),
 ('  covered by paused schedules only', int(((a['Record Level'] == 'Equipment') & (a['PM coverage'] == 'Paused only')).sum()), '=B20/B17', ''),
 ('  no schedule at all', int(((a['Record Level'] == 'Equipment') & (a['PM coverage'] == 'None')).sum()), '=B21/B17', 'Includes HPC servers/racks likely vendor-maintained'),
 ('Group A equipment with only paused schedules', int(((a['Asset Criticality'] == 'A') & (a['PM coverage'] == 'Paused only')).sum()), None, 'Strategy G-9.1: Group A requires comprehensive PM; all UPS at Childress-HPC'),
 ('Group B equipment with no schedule', int(((a['Asset Criticality'] == 'B') & (a['PM coverage'] == 'None')).sum()), None, ''),
]
for label, val, pct, note in kp:
    ws.cell(row=r, column=1, value=label); ws.cell(row=r, column=2, value=val)
    if pct: ws.cell(row=r, column=3, value=pct)
    ws.cell(row=r, column=4, value=note); style_range(ws, r, r, 1, 4, wrap=True); ws.cell(row=r, column=2).number_format = INT; ws.cell(row=r, column=3).number_format = PCT; r += 1
r += 1
ws.cell(row=r, column=1, value='Templates by trigger class').font = f(True); r += 1
header(ws, r, ['Trigger class', 'Templates', 'Active', 'Paused']); r += 1
for t in df['Trigger class'].value_counts().index:
    ws.cell(row=r, column=1, value=t); ws.cell(row=r, column=2, value=f'=COUNTIF({rng("Trigger class")},A{r})'); ws.cell(row=r, column=3, value=f'=COUNTIFS({rng("Trigger class")},A{r},{st},"1")'); ws.cell(row=r, column=4, value=f'=COUNTIFS({rng("Trigger class")},A{r},{st},"0")')
    style_range(ws, r, r, 1, 4, numfmt=INT); r += 1
r += 1
ws.cell(row=r, column=1, value='Priority × Type (templates)').font = f(True); r += 1
prios = ['Highest - This Week', 'High - Next Week', 'Medium - Within 2 Weeks', 'Low - Within 4 Weeks', 'Lowest - Within 8 Weeks', 'No Priority', '(blank)']
header(ws, r, ['Type'] + prios); r += 1
for t in types:
    ws.cell(row=r, column=1, value=t)
    tcrit = '""' if t == '(blank)' else f'"{t}"'
    for j, p in enumerate(prios):
        pcrit = '""' if p == '(blank)' else f'"{p}"'
        ws.cell(row=r, column=2 + j, value=f'=COUNTIFS({rng("Type")},{tcrit},{rng("Priority")},{pcrit})')
    style_range(ws, r, r, 1, 1 + len(prios), numfmt=INT); r += 1
for j in range(len(prios)): ws.column_dimensions[L(2 + j)].width = 14
ws.freeze_panes = 'A5'

# ---- Gap Rules
ws = wb.create_sheet('Gap Rules', 2)
title(ws, 'Gap rules — fails, rate, clause, impact and recommendation', 'Counts are formulas over the flag columns on the Data sheet.')
header(ws, 4, ['Rule', 'Description', 'Category', 'Applicable', 'Fails', 'Fail rate', 'Severity', 'Clause', 'Why it matters', 'Likely root cause', 'Recommendation'], height=36)
r = 5
for rc in rule_cols:
    rid = rc[:3]; m = RULES[rid]
    vals = [rid, m[0], m[1], f'=COUNT({rng(rc)})', f'=COUNTIF({rng(rc)},1)', f'=IF(D{r}=0,0,E{r}/D{r})', m[2], m[3], m[4], m[5], m[6]]
    for i, v in enumerate(vals): ws.cell(row=r, column=1 + i, value=v)
    style_range(ws, r, r, 1, 11, wrap=True); ws.cell(row=r, column=4).number_format = INT; ws.cell(row=r, column=5).number_format = INT; ws.cell(row=r, column=6).number_format = PCT; r += 1
ws.conditional_formatting.add(f'F5:F{r-1}', ColorScaleRule(start_type='num', start_value=0, start_color='FFFFFF', end_type='num', end_value=1, end_color='C00000'))
for c, wd in zip('ABCDEFGHIJK', [7, 38, 13, 10, 8, 9, 9, 30, 52, 40, 50]): ws.column_dimensions[c].width = wd
ws.freeze_panes = 'C5'; ws.auto_filter.ref = f'A4:K{r-1}'

# ---- Coverage
ws = wb.create_sheet('Coverage', 3)
title(ws, 'PM coverage of equipment in the asset register', 'Static, from the asset register joined to active/paused schedules. Direct = attached to an active schedule; Via parent = an ancestor location is; Paused only; None.')
eq = a[a['Record Level'] == 'Equipment']
g = eq.groupby(['Category', 'PM coverage']).size().unstack(fill_value=0).reindex(columns=['Direct', 'Via parent', 'Paused only', 'None'], fill_value=0)
g['Total'] = g.sum(axis=1); g['% with active coverage'] = (g['Direct'] + g['Via parent']) / g['Total']; g = g.sort_values('Total', ascending=False)
header(ws, 4, ['Category', 'Direct', 'Via parent', 'Paused only', 'None', 'Total', '% with active coverage'])
r = 5
for cat, row in g.iterrows():
    ws.cell(row=r, column=1, value=cat)
    for j, v in enumerate(row): ws.cell(row=r, column=2 + j, value=float(v) if j == 5 else int(v))
    style_range(ws, r, r, 1, 7, numfmt=INT); ws.cell(row=r, column=7).number_format = PCT; r += 1
ws.conditional_formatting.add(f'G5:G{r-1}', ColorScaleRule(start_type='num', start_value=0, start_color='C00000', mid_type='num', mid_value=0.5, mid_color='F4B183', end_type='num', end_value=1, end_color='FFFFFF'))
r += 1; ws.cell(row=r, column=1, value='By site').font = f(True); r += 1
gs = eq.groupby(['Site', 'PM coverage']).size().unstack(fill_value=0).reindex(columns=['Direct', 'Via parent', 'Paused only', 'None'], fill_value=0); gs['Total'] = gs.sum(axis=1); gs['% with active coverage'] = (gs['Direct'] + gs['Via parent']) / gs['Total']
header(ws, r, ['Site', 'Direct', 'Via parent', 'Paused only', 'None', 'Total', '% with active coverage']); r += 1
for s_, row in gs.iterrows():
    ws.cell(row=r, column=1, value=s_)
    for j, v in enumerate(row): ws.cell(row=r, column=2 + j, value=float(v) if j == 5 else int(v))
    style_range(ws, r, r, 1, 7, numfmt=INT); ws.cell(row=r, column=7).number_format = PCT; r += 1
r += 1; ws.cell(row=r, column=1, value='By criticality group (Childress-HPC only has criticality assigned)').font = f(True); r += 1
gc_ = eq.assign(cg=eq['Asset Criticality'].replace('', '(none)')).groupby(['cg', 'PM coverage']).size().unstack(fill_value=0).reindex(columns=['Direct', 'Via parent', 'Paused only', 'None'], fill_value=0); gc_['Total'] = gc_.sum(axis=1)
header(ws, r, ['Criticality', 'Direct', 'Via parent', 'Paused only', 'None', 'Total']); r += 1
for s_, row in gc_.iterrows():
    ws.cell(row=r, column=1, value=s_)
    for j, v in enumerate(row): ws.cell(row=r, column=2 + j, value=int(v))
    style_range(ws, r, r, 1, 6, numfmt=INT); r += 1
r += 1; ws.cell(row=r, column=1, value='Group A / B equipment without active coverage').font = f(True); r += 1
header(ws, r, ['Code', 'Name', 'Category', 'Site', 'Criticality', 'PM coverage']); r += 1
for _, row in eq[eq['Asset Criticality'].isin(['A', 'B']) & eq['PM coverage'].isin(['Paused only', 'None'])].sort_values(['Asset Criticality', 'Category']).iterrows():
    for j, v in enumerate([row['Code'], row['Name'], row['Category'], row['Site'], row['Asset Criticality'], row['PM coverage']]):
        ws.cell(row=r, column=1 + j, value=v)
    style_range(ws, r, r, 1, 6); r += 1
for c, wd in zip('ABCDEFG', [34, 40, 22, 22, 12, 12, 16]): ws.column_dimensions[c].width = wd
ws.freeze_panes = 'A5'

# ---- heat maps
def heatmap(name, seg_field, segments, ttl):
    ws = wb.create_sheet(name); title(ws, ttl, 'Fail rate = fails / applicable templates in the segment; second block shows counts.')
    segr = rng(seg_field); r = 4; header(ws, r, ['Rule', 'Description', 'All'] + segments, height=48); r += 1
    ws.cell(row=r, column=1, value='Templates'); ws.cell(row=r, column=3, value=f'=COUNTA({rng("Code")})')
    for j, s_ in enumerate(segments): ws.cell(row=r, column=4 + j, value=f'=COUNTIF({segr},"{s_}")' if s_ != '(blank)' else f'=COUNTBLANK({segr})')
    style_range(ws, r, r, 1, 3 + len(segments), numfmt=INT, bold=True); r += 1; top = r
    for rc in rule_cols:
        ws.cell(row=r, column=1, value=rc[:3]); ws.cell(row=r, column=2, value=RULES[rc[:3]][0]); ws.cell(row=r, column=3, value=f'=IFERROR(COUNTIF({rng(rc)},1)/COUNT({rng(rc)}),"")')
        for j, s_ in enumerate(segments):
            crit = '""' if s_ == '(blank)' else f'"{s_}"'
            ws.cell(row=r, column=4 + j, value=f'=IFERROR(COUNTIFS({rng(rc)},1,{segr},{crit})/COUNTIFS({rng(rc)},">=0",{segr},{crit}),"")')
        style_range(ws, r, r, 1, 3 + len(segments), numfmt=PCT); r += 1
    ws.conditional_formatting.add(f'C{top}:{L(3+len(segments))}{r-1}', ColorScaleRule(start_type='num', start_value=0, start_color='FFFFFF', mid_type='num', mid_value=0.5, mid_color='F4B183', end_type='num', end_value=1, end_color='C00000'))
    r += 1; header(ws, r, ['Rule', 'Description', 'All'] + segments, height=48); r += 1
    for rc in rule_cols:
        ws.cell(row=r, column=1, value=rc[:3]); ws.cell(row=r, column=2, value=RULES[rc[:3]][0] + ' — fails'); ws.cell(row=r, column=3, value=f'=COUNTIF({rng(rc)},1)')
        for j, s_ in enumerate(segments):
            crit = '""' if s_ == '(blank)' else f'"{s_}"'
            ws.cell(row=r, column=4 + j, value=f'=COUNTIFS({rng(rc)},1,{segr},{crit})')
        style_range(ws, r, r, 1, 3 + len(segments), numfmt=INT); r += 1
    ws.column_dimensions['A'].width = 7; ws.column_dimensions['B'].width = 40; ws.column_dimensions['C'].width = 9
    for j in range(len(segments)): ws.column_dimensions[L(4 + j)].width = 12
    ws.freeze_panes = 'D6'
heatmap('By Site', 'Site (short)', sites, 'Fail rate by site')
heatmap('By Type', 'Type', types, 'Fail rate by schedule type')

# ---- Asset references
ws = wb.create_sheet('Asset References')
title(ws, 'Schedule × asset references', 'One row per asset attached to a schedule. Filter Overlap = Yes to see assets covered by more than one active schedule of the same type.')
ex2 = ex.copy(); ex2['Overlap (active, same type)'] = ['Yes' if (r_.Asset, r_.Type, r_.When) in ov_pairs and r_.Active == '1' else '' for r_ in ex2.itertuples()]
ex2['Not-active WO closures'] = np.where(ex2.Asset.isin(na), 'Yes', '')
ex2 = ex2.rename(columns={'SM': 'Schedule', 'level': 'Record Level', 'crit': 'Criticality', 'cat': 'Category'})
ex2['Site'] = ex2['Site'].str.replace(r'\s*\(.*\)$', '', regex=True)
cols = ['Schedule', 'Active', 'Type', 'Priority', 'When', 'Site', 'Asset', 'Category', 'Record Level', 'Criticality', 'Overlap (active, same type)', 'Not-active WO closures']
header(ws, 4, cols)
for i, row in enumerate(ex2[cols].itertuples(index=False), 5):
    for j, v in enumerate(row): ws.cell(row=i, column=1 + j, value=None if (isinstance(v, float) and np.isnan(v)) else v)
for c, wd in zip('ABCDEFGHIJKL', [18, 8, 12, 22, 26, 16, 34, 24, 16, 10, 12, 12]): ws.column_dimensions[c].width = wd
ws.freeze_panes = 'A5'; ws.auto_filter.ref = f'A4:L{len(ex2)+4}'
for c in range(1, len(cols) + 1): ws.cell(row=4, column=c).font = f(True, 'FFFFFF')

wb._sheets = [wb[s] for s in ['Read Me', 'Summary', 'Gap Rules', 'Coverage', 'By Site', 'By Type', 'Asset References', 'Data']]
out = str(OUTPUTS['pm']); wb.save(out); print('saved', out)
print(pd.DataFrame({'applicable': F.notna().sum(), 'fails': F.eq(1).sum(), 'pct': (F.eq(1).sum() / F.notna().sum() * 100).round(1)}).to_string())
