"""Step 7 — build the work order workbook (closed history, backlog, permissions, failure codes).
Inputs : work/wo_scored.pkl, work/bad_actors.pkl, work/assets_scored.pkl, work/group_counts.csv, work/user_facts.json,
         OpenWorkOrderList.csv, UserGroups.csv, FailureCodeProblemsCausesActions.csv
Outputs: outputs/WorkOrder_Gap_Analysis.xlsx
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from config import *
# run-rate window: the last RUN_RATE_WEEKS complete ISO weeks (Mon–Sun) before the snapshot
_last = SNAPSHOT - pd.Timedelta(days=1)                                  # last day before the snapshot
RR_END = _last - pd.Timedelta(days=_last.weekday())                     # Monday of the last complete ISO week + 1 week boundary
RR_START = RR_END - pd.Timedelta(weeks=RUN_RATE_WEEKS)
import pandas as pd, numpy as np, re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import ColorScaleRule

df = pd.read_pickle(WORK['wo_scored'])
# Four source notes are literally "#NAME?" (a spreadsheet error pasted as text). Annotate so they are not read as Excel errors.
df['Completion Notes'] = df['Completion Notes'].replace('#NAME?', '(source note was the spreadsheet error text "#NAME")')
df['Note Placeholder'] = df['Note Placeholder'].replace('#name?', '#name (spreadsheet error text)')
n = len(df)
FIRST, LAST = 2, n + 1  # data rows on Data sheet

FONT = 'Arial'
def f(bold=False, color='000000', size=10, italic=False):
    return Font(name=FONT, bold=bold, color=color, size=size, italic=italic)
HDR_FILL = PatternFill('solid', fgColor='1F3864')
SUB_FILL = PatternFill('solid', fgColor='D9E1F2')
INPUT_FILL = PatternFill('solid', fgColor='FFFF00')
thin = Side(style='thin', color='BFBFBF')
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

def header(ws, row, values, col=1, height=30):
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=col + i, value=v)
        c.font = f(True, 'FFFFFF'); c.fill = HDR_FILL; c.border = BORDER
        c.alignment = Alignment(wrap_text=True, vertical='center')
    ws.row_dimensions[row].height = height

def style_range(ws, r1, r2, c1, c2, numfmt=None, bold=False, color='000000', wrap=False):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = f(bold, color); cell.border = BORDER
            if numfmt: cell.number_format = numfmt
            if wrap: cell.alignment = Alignment(wrap_text=True, vertical='top')

def title(ws, text, sub=None):
    ws['A1'] = text; ws['A1'].font = f(True, size=14)
    if sub:
        ws['A2'] = sub; ws['A2'].font = f(italic=True, color='595959')

PCT = '0.0%'
INT = '#,##0'

wb = Workbook()

# ================================================================== Data sheet
data_cols = list(df.columns)
wsD = wb.active; wsD.title = 'Data'
wsD.append(data_cols)
for row in df.itertuples(index=False):
    wsD.append([None if (isinstance(v, float) and np.isnan(v)) else v for v in row])
# Source notes literally containing "#NAME?" are text, not Excel errors — force string type
for ci in [data_cols.index('Completion Notes') + 1, data_cols.index('Note Placeholder') + 1]:
    for cell in wsD.iter_rows(min_row=2, min_col=ci, max_col=ci):
        if cell[0].data_type == 'e':
            cell[0].data_type = 's'
for c in range(1, len(data_cols) + 1):
    cell = wsD.cell(row=1, column=c); cell.font = f(True, 'FFFFFF'); cell.fill = HDR_FILL
    cell.alignment = Alignment(wrap_text=True, vertical='center')
wsD.freeze_panes = 'J2'
wsD.auto_filter.ref = f'A1:{L(len(data_cols))}{LAST}'
wsD.row_dimensions[1].height = 60
col = {name: L(i + 1) for i, name in enumerate(data_cols)}
rule_cols = [c for c in data_cols if (c[:1] in 'GVRH' and c[1:3].isdigit())]
for c, name in enumerate(data_cols, 1):
    wsD.column_dimensions[L(c)].width = 8 if name in rule_cols else 14
for name in ['Asset Name', 'Description', 'Completion Notes', 'Notes', 'Cause', 'Solution', 'Gaps']:
    wsD.column_dimensions[col[name]].width = 30
def rng(name): return f"Data!${col[name]}${FIRST}:${col[name]}${LAST}"

# Segments
def ordered(series):
    s = series.replace('', '(blank)')
    return list(s.value_counts().index)
sites = ordered(df['Site Name'])
types = ordered(df['Maintenance Type'])
years = sorted(df['Year Completed'].unique())
statuses = ordered(df['Status'])

# ================================================================== Rule metadata
RULES = {
 'G01': ('Priority missing', 'Completeness', 'All closed WOs', 'High',
         'Priority drives scheduling and backlog reporting; a blank is invisible to planners.',
         'Field not mandatory on the WO creation form.',
         'Make Priority mandatory; default it by Maintenance Type for PM-generated WOs.'),
 'G02': ('Maintenance Type missing', 'Completeness', 'All closed WOs', 'High',
         'Maintenance Type is the primary cut for cost, PM-compliance and reliability reporting.',
         'Field not mandatory; manually created WOs skip it.',
         'Make Maintenance Type mandatory at creation and at close.'),
 'G03': ('Site missing', 'Completeness', 'All closed WOs', 'High',
         'WO cannot be charged or reported by site.', 'Asset record has no site assigned.',
         'Fix the affected asset records; add Site to the asset creation standard.'),
 'G04': ('Assigned Users missing', 'Completeness', 'All closed WOs', 'Medium',
         'No accountable crew or person; distorts workload and scheduling reports.',
         'Assignment step skipped, concentrated at Mackenzie, Sweetwater and Canal Flats.',
         'Require assignment before status can move past Open.'),
 'G05': ('Completed By missing', 'Completeness', 'All closed WOs', 'Low',
         'Completion attribution lost.', 'System or bulk closure without a user.', 'Investigate the affected records.'),
 'G06': ('Account Code missing', 'Completeness', 'All closed WOs', 'High',
         'Cost cannot be posted to the right R&M account; finance reporting by asset class is understated.',
         'Account code inherited from asset class; Other/Project/Improvement/Safety WOs have no default.',
         'Add default account codes for non-PM maintenance types; make Acc Code mandatory at close.'),
 'G07': ('Charge Dept missing', 'Completeness', 'All closed WOs', 'Medium',
         'Cost cannot be allocated to a site cost centre.', 'Asset or site record missing charge department.',
         'Populate charge dept at site level so every WO inherits it.'),
 'G08': ('Asset Description missing', 'Completeness', 'All closed WOs', 'Low',
         'Technicians cannot confirm they are on the right equipment; mostly Road and location records.',
         'Asset master gap (not a WO gap) — surfaces here because WOs inherit it.',
         'Carry into the asset-register gap analysis.'),
 'G09': ('Target (Suggested Completion) missing', 'Completeness', 'All closed WOs', 'Medium',
         'No due date means schedule compliance cannot be measured for these WOs.',
         'Manually created WOs do not populate a target date.',
         'Derive target date from Priority automatically on create.'),
 'G10': ('Completion Notes missing', 'Completeness', 'Closed, Completed', 'High',
         'No record of what was found or done; history is useless for reliability analysis and audits.',
         'Notes not mandatory at close; heavy at Canal Flats (98%) and Mackenzie (72%); improving strongly year on year.',
         'Make Completion Notes mandatory to close (min length); for inspections allow a structured "no defects found" checkbox instead.'),
 'G11': ('Completion Notes are a placeholder', 'Validity', 'Closed, Completed', 'Medium',
         '"Completed", "done", "n/a" etc. look populated but carry no information. Notes beginning "overlapping" (~260, all Childress inspections) mean PMs are being closed as Completed when they were not performed, inflating PM compliance.',
         'Mandatory field satisfied with minimum text; overlapping PM schedules on the same asset.',
         'Add a minimum length / disallowed-values check; fix the overlapping PM schedules; close skipped PMs as Incomplete or Rejected, never Completed.'),
 'G12': ('Actual Hours missing', 'Completeness', 'Closed, Completed', 'High',
         'Labour cost is zero on every WO (Total Labour Costs 100% blank); no basis for crew sizing, PM optimisation or cost per asset.',
         'Hours not mandatory at close; no labour rates configured. Mackenzie 91% and Canal Flats 94% blank vs Childress 33%.',
         'Make Actual Hours mandatory to close; load labour rates so labour cost calculates.'),
 'G13': ('Actual Hours zero', 'Validity', 'Closed, Completed', 'Low',
         'Zero hours on a completed job is a placeholder.', 'Field forced but not valued.', 'Validate Actual Hours > 0 on completed WOs.'),
 'G14': ('Est Hours missing', 'Completeness', 'Closed, Completed', 'Medium',
         'Without estimates there is no planning baseline and no estimate-vs-actual feedback loop.',
         'PM templates carry no duration; corrective WOs are not planned.',
         'Add estimated duration to every PM/inspection template; require an estimate for planned corrective work.'),
 'G15': ('Scheduled Maintenance ID missing on Preventive/Inspection', 'Consistency', 'Preventive + Inspection', 'High',
         'PM-type WOs not generated from a schedule are either ad-hoc work mis-typed as PM or PMs being created manually — either way PM compliance is overstated.',
         'Manual WO creation with type Preventive/Inspection.',
         'Restrict Preventive/Inspection types to schedule-generated WOs, or add a "manual PM" type.'),
 'G16': ('Cause missing on failure WO', 'Completeness', 'Corrective + Breakdown + Damage + Warranty', 'High',
         'No failure coding means no bad-actor analysis, no root-cause trending, no basis for adjusting PM frequencies. Problem field is 100% blank across the whole export.',
         'Code lists ARE configured (35 problem, 63 cause, 53 action codes) but are optional at close and have never been used; the lists themselves need cleaning first — duplicates in every list, miner-repair codes mixed with facilities codes, causes/mechanisms/symptoms in one field (see Failure Codes sheet).',
         'Clean the code lists (de-duplicate, separate miner codes, split mechanism from cause, one Unknown per list), link codes to asset categories if the CMMS allows, then make Problem/Cause/Action mandatory to close failure WOs.'),
 'G17': ('Solution missing on failure WO', 'Completeness', 'Corrective + Breakdown + Damage + Warranty', 'High',
         'What fixed it is not captured in a reportable field.', 'As G16.', 'As G16.'),
 'G18': ('Requested By missing or "Guest" on failure WO', 'Completeness', 'Corrective + Breakdown + Damage + Warranty', 'Medium',
         'Cannot close the loop with the person who found the defect; "Guest" requests cannot be traced.',
         'Request portal allows anonymous submission; technicians create WOs directly without a requester.',
         'Require a named requester; disable Guest or force name/email capture.'),
 'V01': ('Completed more than 1 day before Created', 'Validity', 'All closed WOs', 'Medium',
         'WOs are being created after the fact (backfilled), so Days Open and response-time metrics are meaningless for these records.',
         'Work done then logged later; Date Completed manually back-dated.',
         'Acceptable for emergency work if flagged; add a "created retrospectively" indicator or block completion dates earlier than creation.'),
 'V02': ('Target date earlier than Created', 'Validity', 'WOs with a target date', 'Low',
         'Impossible due date; schedule compliance miscalculated.', 'Manual target entry / back-dating.', 'Validate Suggested Completion >= DateCreated.'),
 'V03': ('Actual Hours > 24 on one WO (review)', 'Validity', 'WOs with hours', 'Medium',
         'Very large hour totals are usually multi-day, multi-person jobs booked as a lump sum or a typo (max 792 h).',
         'No labour transactions — a single hours field per WO.',
         'Capture hours per person per day (labour transactions) or at least validate against crew size x days open.'),
 'V04': ('Charge Dept does not match Site', 'Consistency', 'All closed WOs', 'Low',
         'Cost lands in the wrong site cost centre.', 'Asset moved between sites without updating charge dept; dummy dept CA-DUM00.',
         'Derive Charge Dept from Site; review the mismatches.'),
 'V05': ('Rejected WO has hours booked', 'Validity', 'Rejected', 'Low',
         'Labour recorded against work that was not done.', 'Status changed to Rejected after time was booked.',
         'Block Rejected when hours > 0, or move hours to the WO that was actually worked.'),
 'V06': ('Status not a closed status in a closed-WO export', 'Consistency', 'All rows', 'Low',
         'Open/Draft/WIP records appearing in a closed report suggest a report filter or status-workflow issue.',
         'Report definition or status configuration.', 'Review the report filter and the status workflow.'),
 'V07': ('Scheduled Maintenance ID on non-PM type', 'Consistency', 'Non-PM maintenance types', 'Medium',
         'Schedule-generated WOs typed as Safety/Other/Meter reading are excluded from PM compliance and PM cost.',
         'PM schedule templates set with the wrong Maintenance Type (mostly Safety).',
         'Decide whether Safety routines count as PM; set the type on the schedule template accordingly.'),
 'V08': ('Days Open > 90', 'Validity', 'All closed WOs', 'Low',
         'Very old WOs closed in bulk hide real backlog age.', 'Backlog clean-ups; WOs left open after work done.',
         'Weekly backlog review; close WOs within a set window after completion.'),
 'V09': ('Completed after target date (late)', 'Performance', 'Closed, Completed with a target', 'Info',
         'Schedule compliance indicator, not a data gap. One third of completed WOs with a target were closed late.',
         'Capacity, priority inflation (69% of WOs are Highest), target dates auto-set too tight.',
         'Recalibrate priority-to-target mapping; report PM compliance with a grace window.'),
 'V10': ('Closed Incomplete / Rejected because asset not active', 'Consistency', 'Closed, Incomplete + Rejected', 'Medium',
         'PMs are being generated for assets that are not in service (not installed, under construction, shut down) and then manually closed — pure waste and it depresses PM compliance.',
         'Asset status (Active / Not in service) not maintained, or PM schedules not tied to asset status.',
         'Maintain asset operating status; suppress PM generation for assets not in service. Feed into the asset-register gap analysis.'),
 'V11': ('Highest priority on non-Breakdown work', 'Consistency', 'Priority = Highest', 'High',
         'AM-001 5.5 reserves Highest for Breakdown/Emergency, T4 conditions and Group A asset down. 94% of Highest PMs come from schedule templates, so the priority is meaningless for scheduling and hides real urgent work.',
         'PM schedule templates default to Highest; no priority-to-category validation.',
         'Reset PM template priorities per the 5.5 matrix (typically Medium/Low by frequency); validate Highest against category/triage class.'),
 'V12': ('Safety / Warranty / Damage used as a category', 'Consistency', 'All closed WOs', 'Medium',
         'AM-001 G-5.10: safety is a flag, warranty a cost-responsibility flag, damage a failure-cause code — not categories. Using them as types removes the underlying PM/CM/Breakdown category from KPIs and cost analysis.',
         'Current Maintenance Type list predates AM-001 taxonomy.',
         'Re-map the Maintenance Type list to the 13 AM-001 categories; add Safety and Warranty flags and a Damage cause code; migrate history.'),
 'V13': ('Category = Other', 'Consistency', 'All closed WOs', 'Medium',
         'AM-001 G-5.11 caps Other at 2% of work orders and requires monthly review; currently 5.8%, almost all Childress general/housekeeping work.',
         'No Housekeeping/Administrative sub-category; daily crew WOs logged as Other.',
         'Reclassify recurring Other work into defined categories (Asset Monitoring, Inspection, CM); monthly category review.'),
 'R01': ('Fails AM-001 reporting standard (composite)', 'Reporting standard', 'Closed, Completed', 'High',
         'AM-001 KPI "Work order quality": closed WOs meeting the reporting standard ÷ closed, target ≥ 95%. Composite of G-5.14 (asset, category, priority, technician, labour hours) and G-5.15/16 (meaningful description; cause and corrective action on failures). Photographs, parts and attachments cannot be assessed from this export.',
         'Closure checklist (G-5.18) not enforced in the CMMS.',
         'Implement the G-5.18 closure checklist as mandatory close-out fields; supervisor/planner review before close.'),
 'R02': ('Scheduled PM not completed by target date (PM compliance proxy)', 'Performance', 'Schedule-generated PM/Inspection with a target', 'Info',
         'AM-001 KPI "PM compliance" target ≥ 95% within the 10% window. The export has no due date or frequency code, so Suggested Completion is used as the proxy for the window. Trend is improving (43% in 2024, 73% in 2026).',
         'Capacity; overlapping schedules; targets set by priority rather than by frequency window.',
         'Add PM frequency code and due date to the WO; report compliance against the 10% window of G-7.8.'),
 'H01': ('Failure WO written against a location/area record, not equipment', 'Hierarchy', 'Corrective + Breakdown + Damage + Warranty', 'High',
         'Strategy Section 8 sets Level 3 Equipment as the level at which maintenance is managed. 42% of failure WOs are charged to buildings, roads, laydown areas or catch-all records ("Leased Kubotas", "Data Centers"), so failure history never reaches the equipment and the AM-005 bad-actor trigger (2 or 3 corrective WOs in 12 months) cannot fire on a real asset.',
         'Equipment not registered at Level 3 (or not findable), so technicians pick the nearest location; no validation of asset level at WO creation.',
         'Register the missing Level 3 equipment (vehicles, lighting circuits, doors/gates, tanks); require an Equipment-level asset on Corrective/Breakdown WOs; keep location records for Other/housekeeping only.'),
}

# ================================================================== Read Me sheet
ws = wb.create_sheet('Read Me', 0)
title(ws, 'CMMS Gap Analysis — Work Orders (closed WO export)',
      f'Source: ClosedWorkOrders.csv, {n:,} records, completed {df["Year Completed"].min()}–{df["Year Completed"].max()}. Prepared ' + PREPARED_LABEL + '. Draft for review.')
lines = [
 ('Purpose', 'Measure how complete, valid and consistent closed work order records are against a draft data standard, and locate where the gaps cluster (site, maintenance type, year) so remediation can target process and configuration rather than one-off clean-up.'),
 ('How to read it', 'Summary gives the headline numbers. Gap Rules lists every rule with the count and % of records failing it, plus impact, likely root cause and recommendation. By Site / By Type / By Year are heat maps of failure rate per rule. Field Completeness is the raw blank rate for every column in the export. Placeholder Notes lists the low-information completion notes found. Data holds every record with a 1 (fails) / 0 (passes) / blank (rule not applicable) flag per rule — filter on any flag column to get the record list for that gap.'),
 ('Governing documents', 'Rules are anchored to the Data Center Facilities Asset Management Strategy Rev 3.0 (third draft dated 08/21/2026 — Section 7 criticality groups A–D, Section 8 hierarchy/naming/materiality, G-13.2 mandatory asset data, AM-005 defect trigger), to IREN AM-001 Maintenance Workflow & Change Management Rev 3.0 (first draft dated 09/09/2026) — in particular 5.4 Work Order Categorization, 5.5 Prioritization, 5.6 Work Order Reporting and Closure Standard (G-5.14 to G-5.18), 7.3 PM windows and Section 10 KPIs — and to AM-002 Facilities Materials Management Rev 1.0 (first draft 09/23/2026) for parts consumption (G-8.12 to G-8.14) and CMMS reporting (G-12.5). All three documents are drafts pending approval and post-date the work orders analysed (Nov 2023 – Sep 2026). This workbook therefore measures the CURRENT-STATE BASELINE against the TARGET STATE the framework defines; it is not a compliance finding against a standard that was in force. Note that the WO export carries no Asset Criticality Group, so no criticality-dependent rule (Group A vs B/C/D thresholds, response times) can be applied until the asset register is joined in. Each rule carries its clause reference on the Gap Rules sheet; the AM-001 Traceability sheet lists the framework requirements that cannot be measured from this export at all.'),
 ('Assumptions', 'Only "Closed, Completed" WOs are expected to carry Completion Notes, Actual Hours and Est Hours. Preventive and Inspection WOs are expected to originate from a schedule (Scheduled Maintenance ID). Corrective, Breakdown, Damage and Warranty WOs are treated as failure work and expected to carry Cause and Solution. A completion note is treated as a placeholder if it matches a list of low-information phrases or is shorter than 12 characters. Dates are dd/mm/yyyy in the export.'),
 ('Not assessed', 'Parts, misc costs and labour cost fields are 100% blank in the export (parts and labour rates are evidently not configured or not exported) and are reported in Field Completeness only. Asset-master gaps (serial numbers, manufacturer, model) will be assessed from the asset and inventory exports when available.'),
 ('Legend', 'Blue text on Summary = hardcoded thresholds you can edit. Yellow cells = inputs to review or fill in. Heat maps: darker red = higher failure rate.'),
]
r = 4
for k, v in lines:
    ws.cell(row=r, column=1, value=k).font = f(True)
    c = ws.cell(row=r, column=2, value=v); c.font = f(); c.alignment = Alignment(wrap_text=True, vertical='top')
    ws.row_dimensions[r].height = max(30, 15 * (len(v) // 110 + 1))
    r += 1
ws.column_dimensions['A'].width = 18; ws.column_dimensions['B'].width = 120

r += 1
ws.cell(row=r, column=1, value='Work order data standard applied (derived from AM-001 5.4–5.6, 7.3, §10 and AM-002 8.4, 12)').font = f(True, size=12); r += 1
header(ws, r, ['Field', 'Requirement (per AM-001 unless noted)', 'Applies to', 'Rule ID']); r += 1
std = [
 ('Priority', 'Required', 'All WOs', 'G01'), ('Maintenance Type', 'Required', 'All WOs', 'G02'), ('Site Name', 'Required', 'All WOs', 'G03'),
 ('Assigned Users', 'Required', 'All WOs', 'G04'), ('Completed by', 'Required', 'All WOs', 'G05'), ('Acc Code', 'Required', 'All WOs', 'G06'),
 ('Charge Dept Code', 'Required; must match Site', 'All WOs', 'G07, V04'), ('Asset Description', 'Required (asset master)', 'All WOs', 'G08'),
 ('Suggested Completion', 'Required; >= DateCreated', 'All WOs', 'G09, V02'), ('Completion Notes', 'Required; meaningful text (>=12 chars, not a placeholder)', 'Closed, Completed', 'G10, G11'),
 ('Actual Hours', 'Required; > 0; reviewed if > 24', 'Closed, Completed', 'G12, G13, V03'), ('Est Hours', 'Recommended', 'Closed, Completed', 'G14'),
 ('Scheduled Maintenance ID', 'Required on Preventive/Inspection; not expected on other types', 'By Maintenance Type', 'G15, V07'),
 ('Problem / Cause / Solution', 'Required (coded) on failure work', 'Corrective, Breakdown, Damage, Warranty', 'G16, G17'),
 ('Requested By', 'Required, named (not Guest) on failure work', 'Corrective, Breakdown, Damage, Warranty', 'G18'),
 ('Date Completed', '>= DateCreated (within 1 day tolerance)', 'All WOs', 'V01'), ('Status', 'Rejected WOs carry no hours; closed export contains only closed statuses', 'All WOs', 'V05, V06'),
 ('Days Open', 'Reviewed if > 90', 'All WOs', 'V08'),
]
for row in std:
    for i, v in enumerate(row):
        ws.cell(row=r, column=1 + i, value=v)
    style_range(ws, r, r, 1, 4, wrap=True); r += 1
ws.column_dimensions['C'].width = 34; ws.column_dimensions['D'].width = 14

# ================================================================== Summary sheet
ws = wb.create_sheet('Summary', 1)
title(ws, 'Summary — Work Order Data Quality', 'All figures are formulas over the Data sheet.')
ws.column_dimensions['A'].width = 52; ws.column_dimensions['B'].width = 14; ws.column_dimensions['C'].width = 14; ws.column_dimensions['D'].width = 60
r = 4
header(ws, r, ['Headline', 'Count', '% of base', 'Note']); r += 1
gc = rng('Gap Count'); st = rng('Status'); mtr = rng('Maintenance Type')
kpis = [
 ('Work orders in export', f'=COUNTA({rng("Work Order Code")})', None, 'All rows'),
 ('  of which Closed, Completed', f'=COUNTIF({st},"Closed, Completed")', f'=B6/B5', ''),
 ('  of which Closed, Incomplete', f'=COUNTIF({st},"Closed, Incomplete")', f'=B7/B5', ''),
 ('  of which Rejected', f'=COUNTIF({st},"Rejected")', f'=B8/B5', ''),
 ('WOs with at least one completeness gap (G-rules)', f'=COUNTIF({gc},">0")', f'=B9/B5', 'A record fails at least one applicable G rule'),
 ('WOs with 3 or more completeness gaps', f'=COUNTIF({gc},">=3")', f'=B10/B5', ''),
 ('Average completeness gaps per WO', f'=AVERAGE({gc})', None, ''),
 ('Completed WOs with no Completion Notes', f'=COUNTIF({rng("G10 Completion Notes missing")},1)', f'=B12/B6', 'Base: Closed, Completed'),
 ('Completed WOs with no Actual Hours', f'=COUNTIF({rng("G12 Actual Hours missing")},1)', f'=B13/B6', 'Base: Closed, Completed'),
 ('Failure WOs with no Cause recorded', f'=COUNTIF({rng("G16 Cause missing (failure WO)")},1)', f'=B14/COUNT({rng("G16 Cause missing (failure WO)")})', 'Base: Corrective, Breakdown, Damage, Warranty'),
 ('PM/Inspection WOs not linked to a schedule', f'=COUNTIF({rng("G15 Sched Maint ID missing (PM/Insp)")},1)', f'=B15/COUNT({rng("G15 Sched Maint ID missing (PM/Insp)")})', 'Base: Preventive + Inspection'),
 ('WOs with no Account Code', f'=COUNTIF({rng("G06 Account Code missing")},1)', f'=B16/B5', ''),
 ('WOs created after the work was completed (>1 day)', f'=COUNTIF({rng("V01 Completed before Created (>1 day)")},1)', f'=B17/B5', 'Backfilled records'),
 ('Completed WOs closed after target date', f'=COUNTIF({rng("V09 Completed after Target (late)")},1)', f'=B18/COUNT({rng("V09 Completed after Target (late)")})', 'Base: completed WOs with a target date'),
 ('WOs with Priority = Highest - This Week', f'=COUNTIF({rng("Priority")},"Highest - This Week")', f'=B19/B5', 'Priority inflation indicator'),
 ('Closed as Completed with note "overlapping"', f'=COUNTIF({rng("Note Placeholder")},"overlapping")', f'=B20/B6', 'PMs closed Completed without being performed'),
 ('PMs closed Incomplete/Rejected because asset not active', '=COUNTIF(' + rng([c for c in data_cols if c.startswith('V10')][0]) + ',1)', f'=B21/B5', 'PMs generated for out-of-service assets'),
]
for label, cnt, pct, note in kpis:
    ws.cell(row=r, column=1, value=label); ws.cell(row=r, column=2, value=cnt)
    if pct: ws.cell(row=r, column=3, value=pct)
    ws.cell(row=r, column=4, value=note)
    style_range(ws, r, r, 1, 4); ws.cell(row=r, column=2).number_format = INT if 'Average' not in label else '0.00'
    ws.cell(row=r, column=3).number_format = PCT
    r += 1
r += 1
ws.cell(row=r, column=1, value='AM-001 Section 10 KPIs — current baseline vs target').font = f(True); r += 1
header(ws, r, ['Indicator (AM-001 §10)', 'Current', 'Target', 'How measured here / caveat']); r += 1
r01 = rng([c for c in data_cols if c.startswith('R01')][0]); r02 = rng([c for c in data_cols if c.startswith('R02')][0])
v11 = rng([c for c in data_cols if c.startswith('V11')][0])
kpi2 = [
 ('Work order quality (closed WOs meeting reporting standard ÷ closed)', f'=1-COUNTIF({r01},1)/COUNT({r01})', 0.95, 'Composite of G-5.14/15/16 fields present and meaningful; photos, parts, permits not in export'),
 ('PM compliance (proxy: scheduled PMs completed by target date)', f'=1-COUNTIF({r02},1)/COUNT({r02})', 0.95, 'Proxy — export has no PM frequency code or due date; Suggested Completion used as the window'),
 ('Reactive-to-total ratio (Breakdown + Damage ÷ total)', f'=(COUNTIF({mtr},"Breakdown")+COUNTIF({mtr},"Damage"))/B5', 0.20, 'Lower bound: Corrective cannot be split into planned vs unplanned in this export'),
 ('Reactive-to-total ratio incl. all Corrective (upper bound)', f'=(COUNTIF({mtr},"Breakdown")+COUNTIF({mtr},"Damage")+COUNTIF({mtr},"Corrective"))/B5', 0.20, 'Upper bound'),
 ('Other category share (G-5.11)', f'=COUNTIF({mtr},"Other")/B5', 0.02, 'Target is a maximum. By hours the share is larger: CMMS "Hours by Maintenance Type" report, 29 Sep 2025 – 29 Sep 2026: Other 8,897 of 43,007 h (21%) on 428 WOs, vs Preventive 33%, Inspection 17%, Corrective 6%. This export reconciles to that report within 1–2% by completion date'),
 ('Highest priority used outside Breakdown/Damage (5.5)', f'=COUNTIF({v11},1)/COUNT({v11})', 0.0, 'Share of Highest-priority WOs that are not Breakdown/Damage; target is a maximum'),
 ('Schedule compliance', None, 0.90, 'Not measurable — requires frozen-schedule week vs executed week (G-7.10)'),
 ('Triage response', None, 0.95, 'Not measurable — no triage class or triage timestamp in the CMMS export (G-5.7)'),
 ('Backlog (weeks of open ready work at booked pace)', "='Open Backlog'!B26", '4–6 wks', 'Ready estimated hours ÷ booked-hours run rate (see Open Backlog sheet). Below the band; understated because 53% of open WOs carry no estimate and hours are unrecorded at three sites'),
]
for label, cur, tgt, note in kpi2:
    ws.cell(row=r, column=1, value=label); ws.cell(row=r, column=2, value=cur if cur else 'n/a'); ws.cell(row=r, column=3, value=tgt); ws.cell(row=r, column=4, value=note)
    style_range(ws, r, r, 1, 4, wrap=True)
    ws.cell(row=r, column=2).number_format = PCT; ws.cell(row=r, column=3).number_format = PCT
    ws.cell(row=r, column=3).font = f(color='0000FF')
    r += 1

r += 1
ws.cell(row=r, column=1, value='Completeness gaps per WO (distribution)').font = f(True); r += 1
header(ws, r, ['Gaps on record', 'WOs', '% of WOs']); r += 1
for g in range(0, int(df['Gap Count'].max()) + 1):
    ws.cell(row=r, column=1, value=g); ws.cell(row=r, column=2, value=f'=COUNTIF({gc},A{r})'); ws.cell(row=r, column=3, value=f'=B{r}/$B$5')
    style_range(ws, r, r, 1, 3); ws.cell(row=r, column=2).number_format = INT; ws.cell(row=r, column=3).number_format = PCT; r += 1
r += 1
ws.cell(row=r, column=1, value='WOs by status').font = f(True); r += 1
header(ws, r, ['Status', 'WOs', '% of WOs']); r += 1
for s in statuses:
    ws.cell(row=r, column=1, value=s); ws.cell(row=r, column=2, value=f'=COUNTIF({st},A{r})'); ws.cell(row=r, column=3, value=f'=B{r}/$B$5')
    style_range(ws, r, r, 1, 3); ws.cell(row=r, column=2).number_format = INT; ws.cell(row=r, column=3).number_format = PCT; r += 1
r += 1
ws.cell(row=r, column=1, value='WOs by maintenance type').font = f(True); r += 1
header(ws, r, ['Maintenance Type', 'WOs', '% of WOs']); r += 1
for t in types:
    ws.cell(row=r, column=1, value=t)
    ws.cell(row=r, column=2, value=f'=COUNTIF({mtr},A{r})' if t != '(blank)' else f'=COUNTBLANK({mtr})')
    ws.cell(row=r, column=3, value=f'=B{r}/$B$5')
    style_range(ws, r, r, 1, 3); ws.cell(row=r, column=2).number_format = INT; ws.cell(row=r, column=3).number_format = PCT; r += 1
ws.freeze_panes = 'A5'

# ================================================================== Gap Rules sheet
ws = wb.create_sheet('Gap Rules', 2)
title(ws, 'Gap rules — fails, rate, impact and recommendation', 'Counts are formulas over the flag columns on the Data sheet. Severity, impact, root cause and recommendation are analyst judgement for review.')
cols = ['Rule', 'Description', 'Category', 'Population', 'Applicable records', 'Fails', 'Fail rate', 'Severity', 'AM-001 / AM-002 clause', 'Why it matters', 'Likely root cause', 'Recommendation']
CLAUSES = {'G01': 'G-5.3, G-5.14, 5.5', 'G02': 'G-5.3, G-5.14, 5.4', 'G03': 'G-5.14 (correct asset ID)', 'G04': 'G-5.14 (assigned technician)', 'G05': 'G-5.18 (supervisor/planner review)', 'G06': 'AM-002 G-12.5 (consumption by category/campus); Strategy cost analysis', 'G07': 'AM-002 G-12.5', 'G08': 'G-6.1 (AROE master data)', 'G09': '5.5 completion window; G-5.12', 'G10': 'G-5.15, G-5.18', 'G11': 'G-5.15 ("Fixed/OK/Done not acceptable")', 'G12': 'G-5.14 (labour hours), G-5.18', 'G13': 'G-5.14', 'G14': '5.3 step 7 (job plan, resources)', 'G15': '5.4 PM definition (generated from PM schedules); G-7.8', 'G16': 'G-5.16 (failure description: symptoms, cause, components)', 'G17': 'G-5.16 (corrective action)', 'G18': '5.3 step 4 (work request by technician/operations)', 'V01': 'G-5.14 (dates started and completed)', 'V02': '5.5 completion window', 'V03': 'G-5.14 (labour hours: technician, contractors, vendors)', 'V04': 'AM-002 G-12.5', 'V05': 'G-5.14', 'V06': 'G-5.18', 'V07': '5.4 Inspection/PM definitions; G-5.10', 'V08': 'G-5.12 (deferred >30 days flagged; backlog 4–6 weeks)', 'V09': '5.5 completion windows; §10 schedule compliance', 'V10': 'G-7.5 (asset registry basis for planning); AM-009 lifecycle status', 'V11': '5.5 priority matrix', 'V12': 'G-5.10', 'V13': 'G-5.11', 'R01': '§10 Work order quality ≥95%; G-5.14–5.18', 'R02': '§10 PM compliance ≥95%; G-7.8', 'H01': 'Strategy 8 (Level 3), G-8.1; AM-005 G-AM-005.1'}
r = 4; header(ws, r, cols, height=36); r += 1
for rc in rule_cols:
    rid = rc[:3]; meta = RULES[rid]
    ws.cell(row=r, column=1, value=rid); ws.cell(row=r, column=2, value=meta[0]); ws.cell(row=r, column=3, value=meta[1]); ws.cell(row=r, column=4, value=meta[2])
    ws.cell(row=r, column=5, value=f'=COUNT({rng(rc)})'); ws.cell(row=r, column=6, value=f'=COUNTIF({rng(rc)},1)')
    ws.cell(row=r, column=7, value=f'=IF(E{r}=0,0,F{r}/E{r})'); ws.cell(row=r, column=8, value=meta[3])
    ws.cell(row=r, column=9, value=CLAUSES.get(rid, '')); ws.cell(row=r, column=10, value=meta[4]); ws.cell(row=r, column=11, value=meta[5]); ws.cell(row=r, column=12, value=meta[6])
    style_range(ws, r, r, 1, 12, wrap=True)
    ws.cell(row=r, column=5).number_format = INT; ws.cell(row=r, column=6).number_format = INT; ws.cell(row=r, column=7).number_format = PCT
    r += 1
ws.conditional_formatting.add(f'G5:G{r-1}', ColorScaleRule(start_type='num', start_value=0, start_color='FFFFFF', end_type='num', end_value=1, end_color='C00000'))
for c, w in zip('ABCDEFGHIJKL', [7, 34, 13, 24, 11, 9, 9, 9, 22, 48, 44, 48]): ws.column_dimensions[c].width = w
ws.freeze_panes = 'C5'
ws.auto_filter.ref = f'A4:L{r-1}'

# ================================================================== Heat maps
def heatmap(name, seg_field, segments, blank_label='(blank)', title_text=''):
    ws = wb.create_sheet(name)
    title(ws, title_text, 'Fail rate = fails / applicable records in the segment. Second block shows fail counts. Formulas over the Data sheet.')
    segr = rng(seg_field)
    r = 4
    header(ws, r, ['Rule', 'Description', 'All'] + [str(s) for s in segments], height=48); r += 1
    ws.cell(row=r, column=1, value='Records'); ws.cell(row=r, column=2, value='WOs in segment')
    ws.cell(row=r, column=3, value=f'=COUNTA({rng("Work Order Code")})')
    for j, s in enumerate(segments):
        crit = f'"{s}"' if s != blank_label else '""'
        if isinstance(s, (int, np.integer)): crit = str(s)
        ws.cell(row=r, column=4 + j, value=f'=COUNTIF({segr},{crit})' if s != blank_label else f'=COUNTBLANK({segr})')
    style_range(ws, r, r, 1, 3 + len(segments), numfmt=INT, bold=True); r += 1
    top = r
    for rc in rule_cols:
        rid = rc[:3]
        ws.cell(row=r, column=1, value=rid); ws.cell(row=r, column=2, value=RULES[rid][0])
        ws.cell(row=r, column=3, value=f'=IFERROR(COUNTIF({rng(rc)},1)/COUNT({rng(rc)}),"")')
        for j, s in enumerate(segments):
            if s == blank_label:
                seg_crit = '""'
            elif isinstance(s, (int, np.integer)):
                seg_crit = str(s)
            else:
                seg_crit = f'"{s}"'
            ws.cell(row=r, column=4 + j, value=f'=IFERROR(COUNTIFS({rng(rc)},1,{segr},{seg_crit})/COUNTIFS({rng(rc)},">=0",{segr},{seg_crit}),"")')
        style_range(ws, r, r, 1, 3 + len(segments), numfmt=PCT); r += 1
    ws.conditional_formatting.add(f'C{top}:{L(3+len(segments))}{r-1}',
        ColorScaleRule(start_type='num', start_value=0, start_color='FFFFFF', mid_type='num', mid_value=0.5, mid_color='F4B183', end_type='num', end_value=1, end_color='C00000'))
    r += 1
    header(ws, r, ['Rule', 'Description', 'All'] + [str(s) for s in segments], height=48); r += 1
    top2 = r
    for rc in rule_cols:
        rid = rc[:3]
        ws.cell(row=r, column=1, value=rid); ws.cell(row=r, column=2, value=RULES[rid][0] + ' — fails')
        ws.cell(row=r, column=3, value=f'=COUNTIF({rng(rc)},1)')
        for j, s in enumerate(segments):
            seg_crit = '""' if s == blank_label else (str(s) if isinstance(s, (int, np.integer)) else f'"{s}"')
            ws.cell(row=r, column=4 + j, value=f'=COUNTIFS({rng(rc)},1,{segr},{seg_crit})')
        style_range(ws, r, r, 1, 3 + len(segments), numfmt=INT); r += 1
    ws.column_dimensions['A'].width = 7; ws.column_dimensions['B'].width = 40; ws.column_dimensions['C'].width = 9
    for j in range(len(segments)): ws.column_dimensions[L(4 + j)].width = 12
    ws.freeze_panes = 'D6'

heatmap('By Site', 'Site Name', sites, title_text='Fail rate by site')
heatmap('By Type', 'Maintenance Type', types, title_text='Fail rate by maintenance type')
heatmap('By Year', 'Year Completed', years, title_text='Fail rate by year completed (trend)')

# ================================================================== Field completeness
ws = wb.create_sheet('Field Completeness')
title(ws, 'Raw blank rate for every field in the export', 'COUNTBLANK over the Data sheet. Requirement column is the draft standard (edit as needed).')
req = {'Priority': 'Required', 'Asset Name': 'Required', 'Asset Code': 'Required', 'Asset Description': 'Required', 'Asset Category': 'Required', 'Site Name': 'Required',
       'Description': 'Required', 'Project': 'Optional', 'Work Order Code': 'Required', 'Maintenance Type': 'Required', 'Status': 'Required', 'DateCreated': 'Required',
       'Suggested Completion': 'Required', 'Date Completed': 'Required', 'Days Open': 'Derived', 'Requested By': 'Conditional (failure work)', 'Assigned Users': 'Required',
       'Scheduled Maintenance ID': 'Conditional (PM/Inspection)', 'Notes': 'Optional', 'Problem': 'Conditional (failure work)', 'Cause': 'Conditional (failure work)',
       'Solution': 'Conditional (failure work)', 'Completion Notes': 'Conditional (Completed)', 'Completed by': 'Required', 'Est Hours': 'Recommended', 'Actual Hours': 'Conditional (Completed)',
       'Total Labour Costs': 'Derived (needs labour rates)', 'Misc Costs': 'Optional', 'Total Misc Costs': 'Derived', 'Parts Used': 'Conditional (parts consumed)', 'Parts Codes': 'Conditional (parts consumed)',
       'Total Parts Costs': 'Derived', 'Total All Costs': 'Derived', 'Acc Code': 'Required', 'Acc Description': 'Derived', 'Charge Dept Code': 'Required', 'Charge Dept Description': 'Derived'}
r = 4; header(ws, r, ['Field', 'Draft requirement', 'Blank', 'Blank %', 'Distinct values']); r += 1
src_cols = [c for c in data_cols if c in req]
for c in src_cols:
    ws.cell(row=r, column=1, value=c); ws.cell(row=r, column=2, value=req[c]); ws.cell(row=r, column=2).fill = INPUT_FILL
    ws.cell(row=r, column=3, value=f'=COUNTBLANK({rng(c)})'); ws.cell(row=r, column=4, value=f'=C{r}/{n}')
    ws.cell(row=r, column=5, value=int(df[c].nunique()))
    style_range(ws, r, r, 1, 5); ws.cell(row=r, column=3).number_format = INT; ws.cell(row=r, column=4).number_format = PCT; ws.cell(row=r, column=5).number_format = INT
    r += 1
ws.conditional_formatting.add(f'D5:D{r-1}', ColorScaleRule(start_type='num', start_value=0, start_color='FFFFFF', end_type='num', end_value=1, end_color='C00000'))
ws.cell(row=r + 1, column=1, value=f'Distinct values are a static count from the source file ({n:,} records). Yellow cells are editable.').font = f(italic=True, color='595959')
for c, w in zip('ABCDE', [26, 30, 10, 10, 14]): ws.column_dimensions[c].width = w
ws.freeze_panes = 'A5'

# ================================================================== Placeholder notes
ws = wb.create_sheet('Placeholder Notes')
title(ws, 'Low-information completion notes', 'Normalised (lower-case) note text matched by the placeholder rule (G11), with status split. Counts are formulas over the Data sheet.')
ph = df.loc[df['Note Placeholder'] != '', 'Note Placeholder'].value_counts()
r = 4; header(ws, r, ['Note text (normalised)', 'WOs', 'Closed, Completed', 'Closed, Incomplete', 'Rejected', 'Reading']); r += 1
reading = {'not active': 'Asset not in service — PM should not have generated (asset status gap)',
           'overlapping': 'PM closed as Completed because another PM covered it — PM schedule overlap; inflates compliance',
           'completed': 'Placeholder — no information about what was done', 'complete': 'Placeholder', 'completed.': 'Placeholder', 'job complete': 'Placeholder',
           'n/a': 'Placeholder', 'duplicate wo': 'Duplicate WO — should be Rejected, not Completed', 'duplicate': 'Duplicate WO', '#name (spreadsheet error text)': 'Spreadsheet formula error pasted as a note',
           'pm completed': 'Placeholder', 'not active.': 'Asset not in service', 'completed -': 'Placeholder', 'completed by': 'Placeholder'}
npr = rng('Note Placeholder')
for text, cnt in ph.items():
    ws.cell(row=r, column=1, value=text)
    t = text.replace('"', '""')
    ws.cell(row=r, column=2, value=f'=COUNTIF({npr},"{t}")')
    ws.cell(row=r, column=3, value=f'=COUNTIFS({npr},"{t}",{st},"Closed, Completed")')
    ws.cell(row=r, column=4, value=f'=COUNTIFS({npr},"{t}",{st},"Closed, Incomplete")')
    ws.cell(row=r, column=5, value=f'=COUNTIFS({npr},"{t}",{st},"Rejected")')
    ws.cell(row=r, column=6, value=reading.get(text, 'Placeholder / very short note'))
    style_range(ws, r, r, 1, 6, numfmt=INT); r += 1
for c, w in zip('ABCDEF', [30, 8, 14, 14, 10, 70]): ws.column_dimensions[c].width = w
ws.freeze_panes = 'A5'

# ================================================================== AM-001 Traceability
ws = wb.create_sheet('AM-001 Traceability')
title(ws, 'Strategy / AM-001 / AM-002 requirements vs the closed work order export', 'What the framework requires on a work order, whether the export can show it, the current state, and the assessment. Status: Met / Partial / Gap / Not measurable / Configuration gap.')
header(ws, 4, ['Clause', 'Requirement (paraphrased)', 'Measurable in export?', 'Current state (from Data)', 'Status', 'Rule(s)'], height=36)
tr = [
 ('AM-001 G-5.14', 'Every WO records WO number', 'Yes', 'Present on 100% of records', 'Met', ''),
 ('AM-001 G-5.14', 'Correct asset ID', 'Partly — presence only; correctness needs asset register', 'Asset Code present on 100%; 47% of WOs are multi-asset routes carrying one code', 'Partial', 'G03'),
 ('AM-001 G-5.14', 'Work order category (one of 13 in 5.4)', 'Yes', 'Current type list has 11 values; Safety/Warranty/Damage used as types (G-5.10 says flags); no PdM, Emergency, Capital, Modification, Asset Monitoring', 'Configuration gap', 'G02, V12, V13'),
 ('AM-001 G-5.14', 'Priority (5.5 matrix)', 'Yes', 'Priority values already match the 5.5 names; 69% Highest, driven by PM templates', 'Partial', 'G01, V11'),
 ('AM-001 G-5.14', 'Triage class where applicable', 'No', 'No triage class field in the export / CMMS', 'Not measurable', ''),
 ('AM-001 G-5.14', 'Assigned technician', 'Yes', '95% populated; 4.9% blank, concentrated at Mackenzie/Sweetwater/Canal Flats', 'Partial', 'G04'),
 ('AM-001 G-5.14', 'Dates started and completed', 'Partly', 'Date Completed present; NO Date Started field exists; 8% completed >1 day before created', 'Gap', 'V01'),
 ('AM-001 G-5.14', 'Labour hours (technician, contractors, vendors)', 'Partly', 'Single Actual Hours field; blank on 49% of completed WOs; no split by technician/contractor/vendor; labour cost 100% blank', 'Gap', 'G12, G13, V03'),
 ('AM-001 G-5.14', 'Attachments where applicable', 'No', 'Attachment data not exportable through the front end', 'Not measurable', ''),
 ('AM-001 G-5.15', 'Work description answers found / done / tested / result; "Fixed/OK/Done" not acceptable', 'Partly — presence and placeholder detection only', '29% of completed WOs have no notes; 2% placeholders; four-question structure not enforced', 'Gap', 'G10, G11'),
 ('AM-001 G-5.16', 'Failure description (symptoms, cause, components) and corrective action on failures', 'Yes', 'Problem 100% blank; Cause blank on 94% and Solution on 88% of failure WOs. Code lists exist (35 / 63 / 53) but are not enforced — usage gap, not configuration', 'Gap (enforcement)', 'G16, G17'),
 ('AM-001 G-5.17 / AM-002 G-8.14', 'Parts consumed recorded (part no., description, qty, serials for RIK); WO not closed with unrecorded consumption', 'Partly', 'Parts Used / Parts Codes / Parts Costs blank on 100% of 18,384 WOs — inventory module not linked to WOs', 'Gap', '(Field Completeness)'),
 ('AM-001 G-5.17', 'Photographs (min 4) and recommendations', 'No', 'Attachment data cannot be exported through the front end (CMMS Administrator, 28 Sep 2026); to be sampled manually in the quarterly WO audit (G-14.1)', 'Not measurable', ''),
 ('AM-001 G-5.18', 'Closure checklist incl. supervisor/planner review', 'No', 'No checklist or review fields in export; "Completed by" is 61% one person at Childress, suggesting bulk closure by a planner', 'Not measurable', 'G05'),
 ('AM-001 5.4', 'PM = generated from CMMS PM schedules', 'Yes', '11% of Preventive/Inspection WOs have no schedule ID', 'Partial', 'G15'),
 ('AM-001 5.4', 'Inspections that are steps of a PM recorded under that PM', 'Partly', '~260 inspections closed Completed as "overlapping" another PM', 'Gap', 'G11'),
 ('AM-001 G-5.10', 'Safety flag; Warranty cost flag; Damage as failure-cause code', 'Yes', 'All three exist as Maintenance Types instead', 'Configuration gap', 'V12'),
 ('AM-001 G-5.11', 'Other ≤ 2% of WOs, reviewed monthly', 'Yes', '5.8% overall; 9.8% at Childress', 'Gap', 'V13'),
 ('AM-001 5.5', 'Highest = Breakdown/Emergency/T4/Group A down', 'Yes', '91% of Highest WOs are PM/Inspection', 'Gap', 'V11'),
 ('AM-001 G-5.12', 'Corrective deferred >30 days flagged; backlog 4–6 weeks', 'Yes (Open Backlog sheet)', '435 open WOs at 28 Sep 2026: 45 of 95 open corrective/breakdown WOs are >30 days old; 85 Drafts with median age 122 days; 36 WOs >1 year; 123 of 310 with a due date are overdue', 'Gap', 'V08, Open Backlog'),
 ('AM-001 G-7.8 / §10', 'PM WOs initiated within 10% window; PM compliance ≥95%', 'Proxy only', 'No frequency code or due date on WO; 58% of scheduled PMs completed by Suggested Completion (73% in 2026)', 'Gap', 'R02'),
 ('AM-001 §10', 'Work order quality ≥ 95%', 'Composite proxy', 'See Summary KPI block', 'Gap', 'R01'),
 ('AM-001 §10', 'Reactive-to-total ≤ 20%', 'Bounded', '0.7% (Breakdown+Damage) to 9.4% (incl. all Corrective) — within target either way, but Corrective planned/unplanned split not captured', 'Met (with caveat)', ''),
 ('AM-001 §10', 'Schedule compliance ≥ 90%; Triage response ≥ 95%', 'No', 'Frozen-schedule week and triage timestamps not in CMMS', 'Not measurable', ''),
 ('AM-001 G-6.1 / G-7.5', 'AROE asset registry as basis for planning; lifecycle status', 'Indirect', '~240 PMs closed Incomplete/Rejected as "not active" / not installed / under construction', 'Gap (asset master)', 'V10'),
 ('AM-001 5.3 step 4', 'Work request raised with problem description, triage class, requested priority', 'Partly', 'Requested By blank on 96%; "Guest" on 132; no triage class', 'Gap', 'G18'),
 ('Strategy G-7.1 / G-7.4', 'Every asset carries a criticality group; governance intensity keyed to Group A–D', 'No', 'No criticality field on the WO export; cannot distinguish Group A response/trigger thresholds', 'Not measurable (join asset register)', ''),
 ('Strategy 8 / G-8.1', 'Maintenance managed at Level 3 Equipment; assets tagged with level and parent', 'Indirect', '42% of failure WOs written against location/area records; 47% of all WOs are multi-asset routes carrying one Asset Code', 'Gap', 'H01'),
 ('Strategy G-8.4 / G-8.6', 'Assets named per campus naming convention; field tag = CMMS name', 'Partly', 'Asset Codes follow a SITE-SYSTEM-EQUIP-SEQ pattern; catch-all codes exist (CHI-DATACENTERS, CHI-MOBILE-OTHER-KUBOTA); full check needs the naming convention document and asset export', 'Partial', ''),
 ('Strategy G-13.2', 'Assets registered with ID, hierarchy, location, criticality, manufacturer, model, serial, install date, warranty, docs, lifecycle status', 'No (asset export)', 'Asset Description blank on 409 WOs (Road, Locations); remaining attributes not in WO export — next phase', 'Next phase', 'G08'),
 ('Strategy G-13.3 / AM-006 G-AM-006.4', 'CMMS captures data for every KPI without manual re-work', 'Yes', 'Of 8 AM-001 §10 execution KPIs, 3 measurable directly, 2 by proxy, 3 not at all from this export', 'Gap', 'R01, R02'),
 ('AM-005 G-AM-005.1', 'Defect trigger: ≥2 corrective WOs/12 months (Group A), ≥3 (B/C/D)', 'Yes (without criticality)', '133 asset records meet the ≥3 trigger, 287 meet ≥2; only 63 of the 133 are equipment-level records; Cause recorded on 6% of failure WOs so RCA has no data', 'Gap', 'H01, G16 — see Bad Actor Candidates'),
 ('AM-007 G-AM-007.3', 'CMMS master data conforms to AROE data standards incl. mandatory WO fields', 'Partly', 'This workbook is the baseline for the mandatory-WO-field standard', 'Gap', 'all'),
 ('AM-001 G-6.3', 'Permissions assigned through eight buckets (View-Only, Technician, Planner, Planning Manager, Inventory, Administrator, AROE, Owner)', 'Yes (UserGroups.csv)', '53 groups vs 8 buckets: 24 are capability add-ons ("Group - Can …") stacked on users, 12 are crew/dispatch groups, 7 are placeholders, test or Fiix-migration artefacts, 1 duplicate name, plus Guests. See Permissions Mapping for the proposed collapse.', 'Configuration gap', 'Permissions Mapping'),
 ('AM-001 G-6.4', 'One bucket per user; access requested by line manager, approved by AROE; quarterly review; no shared logins', 'Yes (UsersbySite.csv)', '65% of 142 active users hold more than one group (avg 2.3, max 12); 42 active users have never logged in; 34 inactive accounts retain groups; Guest active at 7 sites; one active test account with 6 groups', 'Gap', 'Permissions Mapping'),
 ('AM-001 G-6.5', 'Administrator and Owner not the same person; no execution-team member in Administrator, AROE or Owner', 'Partly', 'Administrators holds 4 active people (incl. GM Engineering and IT) plus 10 system/API accounts; no active member of an AROE-equivalent group; no identifiable Owner; 6 planning users hold master-data import rights outside buckets 6–8', 'Gap', 'Permissions Mapping'),
 ('AM-002 G-12.5', 'CMMS reports consumption by asset, WO category and campus without re-work', 'Yes', 'Account Code blank on 13% (60% of Other/Project/Improvement); Charge Dept mismatch on 27', 'Partial', 'G06, G07, V04'),
]
r = 5
for row in tr:
    for i, v in enumerate(row): ws.cell(row=r, column=1 + i, value=v)
    style_range(ws, r, r, 1, 6, wrap=True); r += 1
for c, w in zip('ABCDEF', [22, 50, 22, 60, 16, 14]): ws.column_dimensions[c].width = w
ws.freeze_panes = 'A5'; ws.auto_filter.ref = f'A4:F{r-1}'

# ================================================================== Bad Actor Candidates
ba = pd.read_pickle(WORK['bad_actors'])
ws = wb.create_sheet('Bad Actor Candidates')
title(ws, 'AM-005 defect-trigger candidates from work order history', 'Static table from the source data: Corrective, Breakdown, Damage and Warranty WOs (excluding Rejected) per Asset Code, and the maximum count in any rolling 12-month window. Criticality group is not in the export, so both trigger thresholds (G-AM-005.1) are shown. Location/Area records are buckets, not assets — see rule H01.')
cols = list(ba.columns)
header(ws, 4, cols, height=40)
r = 5
for row in ba.itertuples(index=False):
    for i, v in enumerate(row):
        ws.cell(row=r, column=1 + i, value=v)
    style_range(ws, r, r, 1, len(cols)); ws.cell(row=r, column=8).number_format = 'yyyy-mm-dd'; ws.cell(row=r, column=9).number_format = 'yyyy-mm-dd'; r += 1
for c, w in zip('ABCDEFGHIJKL', [30, 40, 26, 14, 16, 10, 12, 11, 11, 12, 12, 12]): ws.column_dimensions[c].width = w
ws.freeze_panes = 'C5'; ws.auto_filter.ref = f'A4:{L(len(cols))}{r-1}'
ws.conditional_formatting.add(f'G5:G{r-1}', ColorScaleRule(start_type='num', start_value=1, start_color='FFFFFF', end_type='num', end_value=10, end_color='C00000'))
ws.cell(row=r + 1, column=1, value=f'{len(ba):,} asset records carry failure work; {(ba["Max failure WOs in any 12 months"] >= 3).sum()} meet the >=3 trigger and {(ba["Max failure WOs in any 12 months"] >= 2).sum()} the >=2 trigger. Filter Record Level = Equipment for genuine candidates.').font = f(italic=True, color='595959')

# ================================================================== Next phase: asset & inventory standard
ws = wb.create_sheet('Next Phase Standard')
title(ws, 'Asset register and inventory data standard for the next exports', 'Fields the governing documents require. Use as the export column list and as the rule set for the asset and inventory gap analyses. Yellow = confirm with AROE.')
header(ws, 4, ['Object', 'Field', 'Requirement', 'Source clause', 'Proposed check'], height=30)
std2 = [
 ('Asset', 'Asset ID / name per campus naming convention', 'Required', 'Strategy G-8.4, G-13.2', 'Pattern match to naming convention; duplicates; field tag = CMMS name'),
 ('Asset', 'Hierarchy level (0–4) and parent asset', 'Required', 'Strategy 8, G-8.1', 'Level present; parent present for L1–L4; L3 exists under every L2'),
 ('Asset', 'Location / site', 'Required', 'Strategy G-13.2', 'Present; matches parent chain'),
 ('Asset', 'Criticality group (A–D) and six-dimension score', 'Required before handover', 'Strategy G-7.1, G-7.2', 'Group present; score total consistent with group band'),
 ('Asset', 'Manufacturer', 'Required', 'Strategy G-13.2', 'Present; normalised vendor list (no free-text variants)'),
 ('Asset', 'Model', 'Required', 'Strategy G-13.2', 'Present'),
 ('Asset', 'Serial number', 'Required (Level 3/4 equipment); n/a for locations/structures', 'Strategy G-13.2; AM-002 G-8.14 (RIK serials)', 'Present on equipment; not a placeholder; unique per manufacturer'),
 ('Asset', 'Installation date', 'Required', 'Strategy G-13.2', 'Present; not in future; <= first WO date'),
 ('Asset', 'Warranty (expiry / terms)', 'Required', 'Strategy G-13.2; AM-001 G-8.9', 'Present where install date < 5 yrs'),
 ('Asset', 'Linked documentation (OEM manuals, drawings)', 'Required', 'Strategy G-13.2, G-9.4', 'At least one linked document for Group A/B'),
 ('Asset', 'Lifecycle status (Active / Not in service / Retired)', 'Required', 'Strategy G-8.5, G-13.2; AM-009', 'Present; PMs not generating on non-Active (see V10)'),
 ('Asset', 'Obsolescence status (Supported / Declining / Constrained / Obsolete)', 'Required', 'AM-009 G-AM-009.1', 'Present; review date within 6/12 months by group'),
 ('Asset', 'Maintenance strategy / tactic per asset', 'Required', 'Strategy G-9.7', 'PM schedule exists for Group A/B/C; RTF only Group D with approval'),
 ('Asset', 'Value vs materiality threshold ($500)', 'Registered if > $500 or Group A/B/safety/data-bearing', 'Strategy G-8.7, G-8.8', 'Items > $500 present as assets; below threshold flagged consumable'),
 ('Inventory', 'Item code and description', 'Required', 'AM-002 G-12.1', 'Present; description not a placeholder'),
 ('Inventory', 'Manufacturer part number', 'Required', 'AM-002 G-12.1', 'Present'),
 ('Inventory', 'Manufacturer name', 'Required (implied by part number)', 'AM-002 G-12.1', 'Present; normalised'),
 ('Inventory', 'Approved equivalents', 'Where applicable', 'AM-002 G-12.1', 'Present for Group A/B'),
 ('Inventory', 'Supplier(s) and lead time', 'Required', 'AM-002 G-12.1, Table 6-1', 'Present; lead-time category 1–5 assigned'),
 ('Inventory', 'Unit cost', 'Required', 'AM-002 G-12.1', 'Present; > 0'),
 ('Inventory', 'Spare-parts group (A–D) and adjusted score', 'Required before stocking', 'AM-002 G-6.1, G-6.4', 'Present; never lower than parent asset group'),
 ('Inventory', 'Min / Max', 'Required for every stocked item', 'AM-002 G-7.3', 'Both present; Min < Max; Group A Min >= 1'),
 ('Inventory', 'Bin location', 'Required', 'AM-002 G-12.1', 'Present'),
 ('Inventory', 'Serial numbers', 'Where applicable (serialised spares, RIK)', 'AM-002 G-12.1, G-8.12', 'Present on serialised items'),
 ('Inventory', 'Hazardous-material flag and SDS link', 'Required where hazardous', 'AM-002 G-12.1', 'Flag set; SDS link present'),
 ('Inventory', 'Shelf-life / preservation requirement', 'Where applicable', 'AM-002 G-12.1', 'Present'),
 ('Inventory', 'Linked parent asset(s)', 'Required', 'AM-002 G-12.1, G-12.4', 'At least one linked asset; asset exists and is Active'),
 ('Inventory', 'Usage in last 12 months', 'Derived', 'AM-002 G-7.5, G-8.16', 'Zero-usage items listed for AM-009 review (excl. Group A insurance spares)'),
]
r = 5
for row in std2:
    for i, v in enumerate(row): ws.cell(row=r, column=1 + i, value=v)
    style_range(ws, r, r, 1, 5, wrap=True); ws.cell(row=r, column=3).fill = INPUT_FILL; r += 1
for c, w in zip('ABCDE', [11, 46, 34, 30, 52]): ws.column_dimensions[c].width = w
ws.freeze_panes = 'A5'

# ================================================================== Open Backlog (from OpenWorkOrderList.csv, snapshot 28 Sep 2026)
ob = pd.read_csv(INPUTS['open_wo'], dtype=str, keep_default_na=False); ob.columns = [c.strip() for c in ob.columns]
for c in ob.columns: ob[c] = ob[c].str.strip()
ob = ob[ob['WO Code'] != '']
_reg = pd.read_pickle(WORK['assets_scored']).drop_duplicates('Code').set_index('Code')
ob['Asset Code'] = ob['Asset'].str.extract(r'\(([^()]+)\)\s*$')[0].fillna('')
ob['Site'] = ob['Asset Code'].map(_reg['Site']).fillna('(unmatched)')
ob['Record Level'] = ob['Asset Code'].map(_reg['Record Level']).fillna('')
_today = SNAPSHOT
ob['Age (days)'] = (_today - pd.to_datetime(ob['Date Created'], format=FMT_OPEN_WO)).dt.days
ob['Age band'] = pd.cut(ob['Age (days)'], [-1, 7, 30, 90, 180, 365, 10000], labels=['0-7', '8-30', '31-90', '91-180', '181-365', '>365']).astype(str)
_dl = pd.to_numeric(ob['Days Late'], errors='coerce')
ob['Overdue'] = np.where(_dl > 0, 'Yes', np.where(_dl.notna(), 'No', 'No due date'))
ob['Corrective >30 days (G-5.12)'] = np.where(ob['Type'].isin(['Corrective', 'Breakdown', 'Damage']) & (ob['Age (days)'] > 30), 'Yes', '')
ob['Ready (Open/Assigned)'] = np.where(ob['Status'].isin(['Open', 'Assigned']), 'Yes', '')
ob['Est Hours (num)'] = pd.to_numeric(ob['Est Hours'], errors='coerce')
ob_cols = list(ob.columns)
wsO = wb.create_sheet('Open Backlog Data')
wsO.append(ob_cols)
for row in ob.itertuples(index=False):
    wsO.append([None if (isinstance(v, float) and np.isnan(v)) else v for v in row])
for c in range(1, len(ob_cols) + 1):
    cell = wsO.cell(row=1, column=c); cell.font = f(True, 'FFFFFF'); cell.fill = HDR_FILL; cell.alignment = Alignment(wrap_text=True, vertical='center'); wsO.column_dimensions[L(c)].width = 14
wsO.column_dimensions['B'].width = 50; wsO.column_dimensions['D'].width = 40
wsO.freeze_panes = 'B2'; wsO.auto_filter.ref = f'A1:{L(len(ob_cols))}{len(ob)+1}'; wsO.row_dimensions[1].height = 45
ocol = {nm: L(i + 1) for i, nm in enumerate(ob_cols)}
def orng(nm): return f"'Open Backlog Data'!${ocol[nm]}$2:${ocol[nm]}${len(ob)+1}"

ws = wb.create_sheet('Open Backlog')
title(ws, 'Open work order backlog — snapshot ' + SNAPSHOT_LABEL, f'{len(ob)} open work orders (OpenWorkOrderList.csv). Age measured from Date Created to {SNAPSHOT_LABEL}. Formulas over the Open Backlog Data sheet. Site is looked up from the asset register.')
for c, w in zip('ABCD', [56, 14, 14, 60]): ws.column_dimensions[c].width = w
r = 4; header(ws, r, ['Measure', 'Value', '% of open', 'Note / clause']); r += 1
N = f"COUNTA({orng('WO Code')})"
_wo = pd.read_pickle(WORK['wo_scored']); _wo['dc'] = pd.to_datetime(_wo['Date Completed'], format=FMT_CLOSED_WO); _wo['h'] = pd.to_numeric(_wo['Actual Hours'], errors='coerce').fillna(0)
_w = _wo[(_wo.dc >= RR_START) & (_wo.dc < RR_END)]
_rate = _w.groupby('Site Name')['h'].sum() / RUN_RATE_WEEKS; _hb = _w.groupby('Site Name').apply(lambda d: (d['Actual Hours'] == '').mean())
RUN_RATE_TOTAL = int(round(_rate.sum()))
_rr_label = f"ISO weeks {RR_START.isocalendar()[1]}–{(RR_END - pd.Timedelta(days=1)).isocalendar()[1]} {RR_END.year}"
items = [
 ('Open work orders', f'={N}', None, 'All non-closed statuses'),
 ('  in Draft (never released)', f'=COUNTIF({orng("Status")},"Draft")', f'=B6/B5', 'Median age 122 days; oldest 915 days; includes a "test workorder"'),
 ('  Work Completed but not closed', f'=COUNTIF({orng("Status")},"Work Completed")', f'=B7/B5', 'Closure lag — G-5.18 supervisor/planner review pending'),
 ('  Waiting for parts', f'=COUNTIF({orng("Status")},"Waiting for parts")', f'=B8/B5', 'Cannot be tracked against stock: inventory module not transacting (AM-002)'),
 ('  On Hold', f'=COUNTIF({orng("Status")},"On Hold")', f'=B9/B5', ''),
 ('Older than 30 days', f'=COUNTIF({orng("Age (days)")},">30")', f'=B10/B5', ''),
 ('Older than 1 year', f'=COUNTIF({orng("Age (days)")},">365")', f'=B11/B5', 'Mostly Canal Flats drafts from 2024'),
 ('Corrective/Breakdown open > 30 days (G-5.12 management review flag)', f'=COUNTIF({orng("Corrective >30 days (G-5.12)")},"Yes")', f'=B12/(COUNTIF({orng("Type")},"Corrective")+COUNTIF({orng("Type")},"Breakdown")+COUNTIF({orng("Type")},"Damage"))', 'Base: open corrective/breakdown/damage'),
 ('Overdue (past due date)', f'=COUNTIF({orng("Overdue")},"Yes")', f'=B13/(COUNTIF({orng("Overdue")},"Yes")+COUNTIF({orng("Overdue")},"No"))', 'Base: WOs with a due date'),
 ('No due date', f'=COUNTIF({orng("Overdue")},"No due date")', f'=B14/B5', 'Schedule compliance unmeasurable for these'),
 ('Open PM / Inspection WOs', f'=COUNTIF({orng("Type")},"Preventive")+COUNTIF({orng("Type")},"Inspection")', f'=B15/B5', ''),
 ('  of which overdue', f'=COUNTIFS({orng("Type")},"Preventive",{orng("Overdue")},"Yes")+COUNTIFS({orng("Type")},"Inspection",{orng("Overdue")},"Yes")', f'=B16/B15', 'Feeds PM compliance (target >= 95%)'),
 ('Priority blank', f'=COUNTBLANK({orng("Priority")})', f'=B17/B5', 'Concentrated in Draft / Requested — pre-triage (5.3 step 6)'),
 ('Assigned Users blank', f'=COUNTBLANK({orng("Assigned Users")})', f'=B18/B5', ''),
 ('Est Hours blank', f'=COUNTBLANK({orng("Est Hours")})', f'=B19/B5', 'Backlog hours below are understated by this share'),
 ('On a location/area record rather than equipment', f'=COUNTIF({orng("Record Level")},"Location / Area")', f'=B20/B5', 'Strategy 8 Level 3'),
 ('Highest priority', f'=COUNTIF({orng("Priority")},"Highest - This Week")', f'=B21/B5', '5.5 reserves Highest for Breakdown/Emergency/T4'),
 ('Type = Warranty', f'=COUNTIF({orng("Type")},"Warranty")', f'=B22/B5', 'Against 3 warranty records in the register'),
 ('Estimated hours in backlog (all open, where estimated)', f'=SUM({orng("Est Hours (num)")})', None, 'Hours'),
 ('Estimated hours of READY work (Open/Assigned, where estimated)', f'=SUMIF({orng("Ready (Open/Assigned)")},"Yes",{orng("Est Hours (num)")})', None, 'Hours — numerator of the AM-001 backlog KPI'),
 (f'Weekly booked technician hours — {RUN_RATE_WEEKS}-week run rate ({_rr_label}, from closed WO Actual Hours)', RUN_RATE_TOTAL, None, 'Demonstrated throughput, not rostered capacity. Cross-checked against the CMMS "Total Technician Hours Logged per Site per Work Week" report (twelve months to the snapshot): same pattern by site and week; weekly levels differ (e.g. Childress wk 51-2025 c.990 h in the report vs 1,293 h from WO Actual Hours by completion date), so the report likely draws hours from a different field or date than WO completion — source to be confirmed. Edit if a rostered-capacity figure becomes available'),
 ('Backlog in weeks of ready work at booked pace (AM-001 §10 band 4–6 weeks)', '=B24/B25', None, 'Below the band. Understated: 53% of open WOs have no estimate, and booked hours are blank on 58% of Mackenzie, 69% of Sweetwater and 100% of Canal Flats closures in the period'),
]
for label, val, pct, note in items:
    ws.cell(row=r, column=1, value=label); ws.cell(row=r, column=2, value=val)
    if pct: ws.cell(row=r, column=3, value=pct)
    ws.cell(row=r, column=4, value=note); style_range(ws, r, r, 1, 4, wrap=True)
    ws.cell(row=r, column=2).number_format = '0.0' if 'weeks' in label else INT; ws.cell(row=r, column=3).number_format = PCT
    if 'INPUT' in label: ws.cell(row=r, column=2).font = f(color='0000FF'); ws.cell(row=r, column=2).fill = INPUT_FILL
    r += 1
r += 1
for seg_name, seg_field in [('Status', 'Status'), ('Age band', 'Age band'), ('Site', 'Site'), ('Type', 'Type')]:
    ws.cell(row=r, column=1, value=f'Open WOs by {seg_name}').font = f(True); r += 1
    header(ws, r, [seg_name, 'WOs', '% of open', 'Median age (days)']); r += 1
    vals = list(ob[seg_field].replace('', '(blank)').value_counts().index)
    if seg_field == 'Age band': vals = [v for v in ['0-7', '8-30', '31-90', '91-180', '181-365', '>365'] if v in vals]
    for v in vals:
        ws.cell(row=r, column=1, value=v)
        crit = '""' if v == '(blank)' else f'"{v}"'
        ws.cell(row=r, column=2, value=f'=COUNTIF({orng(seg_field)},{crit})' if v != '(blank)' else f'=COUNTBLANK({orng(seg_field)})')
        ws.cell(row=r, column=3, value=f'=B{r}/$B$5')
        ws.cell(row=r, column=4, value=int(ob.loc[ob[seg_field].replace('', '(blank)') == v, 'Age (days)'].median()))
        style_range(ws, r, r, 1, 4); ws.cell(row=r, column=2).number_format = INT; ws.cell(row=r, column=3).number_format = PCT; r += 1
    r += 1
ws.cell(row=r, column=1, value=f'Backlog at booked pace by site (static; {RUN_RATE_WEEKS}-week run rate {_rr_label})').font = f(True); r += 1
header(ws, r, ['Site', 'Booked hours / week', 'Hours blank on closures (%)', 'Ready WOs', 'Ready hours (estimated)', 'Ready WOs without estimate (%)', 'Weeks to clear at booked pace']); r += 1
_wo = pd.read_pickle(WORK['wo_scored']); _wo['dc'] = pd.to_datetime(_wo['Date Completed'], format='%d/%m/%Y %H.%M.%S'); _wo['h'] = pd.to_numeric(_wo['Actual Hours'], errors='coerce').fillna(0)
_w = _wo[(_wo.dc >= RR_START) & (_wo.dc < RR_END)]
_rate = _w.groupby('Site Name')['h'].sum() / RUN_RATE_WEEKS; _hb = _w.groupby('Site Name').apply(lambda d: (d['Actual Hours'] == '').mean())
_rd = ob[ob['Ready (Open/Assigned)'] == 'Yes']; _r = _rd.groupby('Site').agg(n=('WO Code', 'size'), hrs=('Est Hours (num)', 'sum'), un=('Est Hours (num)', lambda s_: s_.isna().mean()))
for site in sorted(set(_rate.index) | set(_r.index)):
    br = float(_rate.get(site, 0)); rh = float(_r['hrs'].get(site, 0)) if site in _r.index else 0.0
    vals = [site, round(br, 1), round(float(_hb.get(site, np.nan)) * 100, 0) if site in _hb.index else None, int(_r['n'].get(site, 0)) if site in _r.index else 0, round(rh, 1), round(float(_r['un'].get(site, np.nan)) * 100, 0) if site in _r.index else None, round(rh / br, 1) if br > 0 else 'n/a (no hours booked)']
    for i, v in enumerate(vals): ws.cell(row=r, column=1 + i, value=v)
    style_range(ws, r, r, 1, 7); r += 1
ws.cell(row=r, column=1, value='Weeks to clear is a throughput measure (ready estimated hours ÷ hours actually booked per week). It understates the true backlog wherever estimates or booked hours are missing — Canal Flats and Sweetwater cannot be measured at all.').font = f(italic=True, color='595959')
for c, wd in zip('ABCDEFG', [56, 14, 14, 10, 14, 14, 16]): ws.column_dimensions[c].width = wd
ws.freeze_panes = 'A5'

# ================================================================== Permissions mapping template
ws = wb.create_sheet('Permissions Mapping')
title(ws, 'Permission groups → AM-001 §6 buckets (first-pass mapping)', 'The 53 configured groups from UserGroups.csv with a proposed target bucket and rationale. Yellow = review/complete (user counts, config rights, shared-login flag, and the proposed bucket itself). Summary counts below recalculate.')
BUCKETS = ['1 General / View-Only', '2 Technician', '3 Maintenance Planner', '4 Maintenance Planning Manager', '5 Inventory Team', '6 CMMS Administrator', '7 AROE', '8 Owner', 'None — grants rights outside the standard', 'Retire — unused / duplicate']
header(ws, 4, ['#', 'Configured group name', 'Active users in group (UsersBySite.csv, ' + SNAPSHOT_LABEL + ')', 'Can edit config / master data?', 'Shared login?', 'Target AM-001 bucket (proposed)', 'Notes'], height=36)
gcnt = pd.read_csv(WORK['group_counts'], index_col=0)
uf = __import__('json').load(open(WORK['user_facts']))
ug = pd.read_csv(INPUTS['user_groups'], dtype=str, keep_default_na=False)['Full Name'].str.strip()
ug = ug[ug != '']
def propose(g):
    gl = g.lower()
    if gl.startswith('group - can') or gl.startswith('group - able'):
        return ('None — grants rights outside the standard', 'Capability add-on group. Stacking these on users breaks G-6.4 "one bucket only"; fold the right into the bucket that needs it, then delete.')
    if g == 'Administrators': return ('6 CMMS Administrator', 'G-6.5: must not be the same person as the Owner; no execution staff.')
    if g == 'CMMS Management': return ('7 AROE', 'Clarify whether this is AROE (governance) or the Owner (VP Engineering). Only one named Owner is allowed.')
    if g in ('Corporate - Management', 'Corporate - Support', 'Site - Management'): return ('1 General / View-Only', 'Site / Facilities Managers approve MOPs and site scheduling (3.3, 7.5) but AM-001 gives them no CMMS edit bucket — confirm with author whether 4 Planning Manager is intended for site management.')
    if g in ('GPU - Data Technicians', 'Labourers - Critical Environment', 'Technicians') or gl.startswith('maintenance - ') and g not in ('Maintenance - Planning', 'Maintenance - Leaders'):
        return ('2 Technician', 'Crew group used as a WO assignment target (Assigned Users). Keep for dispatch; permissions should come from bucket 2, not from the crew group.')
    if g == 'Maintenance - Planning': return ('3 Maintenance Planner', '')
    if g == 'Maintenance - Leaders': return ('3 Maintenance Planner', 'Supervisors: AM-001 has no supervisor bucket although G-5.18 requires supervisor review at close — raise with author (Technician + close authority, or Planner).')
    if g == 'Logistics': return ('5 Inventory Team', '')
    if g == 'Purchasing': return ('5 Inventory Team', 'AM-002 G-5.3 segregation: the approver of a purchase shall not receive it — Purchasing must not hold receipt rights. May need to be View-Only plus PO.')
    if g == 'Technical Support': return ('6 CMMS Administrator', 'If this is vendor support, rights must be time-bound and logged (G-6.4).')
    if g == 'Guests': return ('Retire — unused / duplicate', '"Guest" is the requester on 132 WOs. G-6.4 prohibits shared/anonymous logins; replace with named requesters or a request portal that captures identity.')
    if gl.startswith('new group') or gl.startswith('new user group') or 'test' in gl or 'do not change' in gl or 'migrated' in gl:
        return ('Retire — unused / duplicate', 'Placeholder, test or migration artefact (Fiix legacy).')
    return ('', '')
seen = set()
for i, g in enumerate(ug, 1):
    r = 4 + i
    b, note = propose(g)
    if g in seen: note = ('DUPLICATE group name. ' + note).strip()
    seen.add(g)
    ws.cell(row=r, column=1, value=i); ws.cell(row=r, column=2, value=g); ws.cell(row=r, column=6, value=b); ws.cell(row=r, column=7, value=note)
    gk = g.replace('Group - Can Import, Add and Modify Parts', 'Group - Can Import Add and Modify Parts')
    ws.cell(row=r, column=3, value=int(gcnt.loc[gk, 'active_human']) if gk in gcnt.index else 0)
    if g.lower().startswith('group - can') or g.lower().startswith('group - able'): ws.cell(row=r, column=4, value='Yes' if ('Import' in g or 'Modify' in g or 'Edit' in g or 'Create' in g or 'Add' in g) else 'No')
    if g in ('Guests',): ws.cell(row=r, column=5, value='Yes')
    for c in (4, 5): ws.cell(row=r, column=c).fill = INPUT_FILL
    ws.cell(row=r, column=6).fill = INPUT_FILL
    style_range(ws, r, r, 1, 7, wrap=True)
for i in range(len(ug) + 1, 61):
    r = 4 + i; ws.cell(row=r, column=1, value=i)
    for c in range(2, 8): ws.cell(row=r, column=c).fill = INPUT_FILL
    style_range(ws, r, r, 1, 7)
from openpyxl.worksheet.datavalidation import DataValidation
dv = DataValidation(type='list', formula1='"' + ','.join(BUCKETS) + '"', allow_blank=True); ws.add_data_validation(dv); dv.add('F5:F64')
dv2 = DataValidation(type='list', formula1='"Yes,No"', allow_blank=True); ws.add_data_validation(dv2); dv2.add('D5:D64'); dv2.add('E5:E64')
r = 67
ws.cell(row=r, column=1, value='Summary by target bucket').font = f(True); r += 1
header(ws, r, ['Target bucket', 'Groups mapped', 'Active memberships (a user in several groups is counted in each)', 'G-6 test']); r += 1
tests = {'6 CMMS Administrator': 'G-6.5: not the same person as Owner; no execution staff', '7 AROE': 'G-6.5: no execution staff', '8 Owner': 'G-6.5: one named individual', 'None — grants rights outside the standard': 'Remove or redesign', 'Retire — unused / duplicate': 'Delete after access review (G-6.4)'}
for b in BUCKETS:
    ws.cell(row=r, column=1, value=b); ws.cell(row=r, column=2, value=f'=COUNTIF($F$5:$F$64,A{r})'); ws.cell(row=r, column=3, value=f'=SUMIF($F$5:$F$64,A{r},$C$5:$C$64)'); ws.cell(row=r, column=4, value=tests.get(b, ''))
    style_range(ws, r, r, 1, 4, numfmt=INT); r += 1
ws.cell(row=r, column=1, value='Groups entered'); ws.cell(row=r, column=2, value='=COUNTA($B$5:$B$64)'); style_range(ws, r, r, 1, 2, numfmt=INT, bold=True); r += 1
ws.cell(row=r, column=1, value='Groups flagged as shared logins'); ws.cell(row=r, column=2, value='=COUNTIF($E$5:$E$64,"Yes")'); style_range(ws, r, r, 1, 2, numfmt=INT, bold=True); r += 1
ws.cell(row=r, column=1, value='Groups with config/master-data rights'); ws.cell(row=r, column=2, value='=COUNTIF($D$5:$D$64,"Yes")'); ws.cell(row=r, column=4, value='Standard allows this only in buckets 6, 7, 8 (G-6.3)'); style_range(ws, r, r, 1, 4, numfmt=INT, bold=True)
r += 2
ws.cell(row=r, column=1, value='User population (aggregate, from UsersbySite.csv — no individuals listed)').font = f(True); r += 1
header(ws, r, ['Measure', 'Value', 'AM-001 test', 'Reading']); r += 1
urows = [
 ('Accounts', uf['accounts'], '', f"{uf['human']} human, {uf['system']} system/API, Guest at {uf['guest_sites']} sites, {uf['test']} test/placeholder"),
 ('Active human users', uf['human_active'], '', f"{uf['human_inactive']} inactive accounts still hold group memberships"),
 ('Active users holding more than one group', uf['multi_group'], 'G-6.4 one bucket per user', f"{uf['multi_group_pct']}% of active users; average {uf['avg_groups']} groups; maximum {uf['max_groups']}"),
 ('Active users carrying capability add-on groups', uf['addon_users'], 'G-6.3', 'Rights granted by stacking "Group – Can …" groups'),
 ('Active users with import / add / modify (master-data) rights', uf['masterdata_rights'], 'G-6.3: master data only in buckets 6, 7, 8', 'Planners and asset managers holding import rights outside Administrator/AROE'),
 ('Members of Administrators', uf['admins_total'], 'G-6.5', f"{uf['admins_active_human']} active people, {uf['admins_system']} system/API accounts, 1 inactive vendor account"),
 ('Active members of CMMS Management (AROE-equivalent)', uf['cmms_mgmt_active'], 'G-6.3 bucket 7', 'No active AROE-level group membership exists'),
 ('Active users in Maintenance – Leaders', uf['leaders_active'], '—', 'Catch-all: managers, supervisors, planners, coordinators, engineers, technicians, a cleaner, a director'),
 ('Active users in Technicians', uf['technicians_active'], '—', f"includes {uf['technicians_mgmt']} with manager / head / director titles"),
 ('Active users who have never logged in', uf['never_login'], 'G-6.4 quarterly access review', 'Provisioned but unused accounts'),
 ('Active users with no login in 90+ days', uf['login_gt90'], 'G-6.4', f"{uf['login_gt365']} over a year"),
 ('Active test account with live group memberships', uf['test_active'], 'Master data control', 'Logged in September 2026'),
 ('Active vendor personnel with WO rights', uf['vendor_active'], 'G-6.4 time-bound elevated rights', 'Contractor accounts in Leaders/Technicians'),
 ('Accounts with an hourly rate', 0, 'AM-001 G-5.14 labour cost', 'Rate field blank or zero on every account — labour cost cannot calculate'),
 ('Users assigned to more than one site', uf['multi_site_users'], '—', 'Site assignment is many-to-many; not a gap in itself'),
]
for row in urows:
    for i, v in enumerate(row): ws.cell(row=r, column=1 + i, value=v)
    style_range(ws, r, r, 1, 4, wrap=True); ws.cell(row=r, column=2).number_format = INT; r += 1
for c, wd in zip('ABCDEFG', [6, 40, 14, 22, 14, 40, 40]): ws.column_dimensions[c].width = wd
ws.freeze_panes = 'A5'

# ================================================================== Failure code lists
fc = pd.read_csv(INPUTS['failure_codes'], dtype=str, keep_default_na=False)
fc.columns = [c.strip() for c in fc.columns]
P = fc[['Problem Code', 'Problem Description']].rename(columns={'Problem Code': 'Code', 'Problem Description': 'Description'}); P = P[P.Code.str.strip() != '']
C = fc[['Cause Code', 'Cause Description']].rename(columns={'Cause Code': 'Code', 'Cause Description': 'Description'}); C = C[C.Code.str.strip() != '']
A = fc[['Action Code', 'Action Description']].rename(columns={'Action Code': 'Code', 'Action Description': 'Description'}); A = A[A.Code.str.strip() != '']
MINER_P = {'P30', 'P31', 'P32', 'P33', 'P34', 'P35', 'P36', 'P37', 'P38', 'P39'}
MINER_C = {'C60', 'C61', 'C62', 'C63', 'C64', 'C65', 'C66', 'C67', 'C68', 'C69', 'C71', 'C74', 'C75', 'C76', 'C77', 'C78'}
MINER_A = {'A60', 'A61', 'A62', 'A63', 'A64', 'A65', 'A66', 'A67', 'A68', 'A69', 'A70', 'A71', 'A72', 'A73', 'A74', 'A75', 'A76', 'A77', 'A78', 'A79', 'A80'}
NOTES_P = {
 'P007': 'Duplicate of P12 "Doesn\'t turn ON"; 3-digit code breaks the Pnn pattern', 'P008': '3-digit code breaks the Pnn pattern; "Breaker tripped" is arguably a cause/mechanism (see C23)',
 'P12': 'Duplicate of P007', 'P02': 'Overlaps P03 Damaged and P36 Physically Damaged Chasis', 'P03': 'Overlaps P02 / P36',
 'P33': 'Cause, not a symptom (control board / firmware) — belongs in the Cause list (C67/C68/C74)', 'P39': 'Cause, not a symptom (see C65 Dead PSU)',
 'P36': 'Miner-specific; typo "Chasis"; overlaps P02/P03', 'P37': '"Temperamental Machine" is not a failure symptom — remove', 'P40': 'Catch-all "Other / Unkown" (typo); make it require free text or remove',
 'P21': 'Duplicated as a cause (C43 Fluid leak) — a leak is a symptom', 'P38': 'Miner-specific; overlaps P35 Fans not Spinning',
}
NOTES_C = {
 'C28': 'Exact duplicate of C74 "Control Board Malfunction"', 'C74': 'Exact duplicate of C28', 'C05': 'Catch-all "Unknown Cause" — keep one catch-all only, require text', 'C73': 'Catch-all "TBD" — remove (duplicates C05)',
 'C16': 'Overlaps C27 Programming Issue and C67 Improper Firmware', 'C27': 'Overlaps C16 / C67', 'C67': 'Overlaps C16 / C27',
 'C43': 'Symptom, not a cause (duplicates Problem P21 Fluid Leak)', 'C31': 'Symptom, not a cause (loss of power is what happened, not why)',
 'C20': 'Overlaps C25 Loose wire and C22 Electrical fault — decide the split', 'C25': 'Overlaps C20', 'C22': 'Overlaps C20 / C23',
 'C76': 'Miner firmware error message, not a cause', 'C78': 'Miner firmware error message, not a cause', 'C77': 'Miner error message; overlaps C24 Excessive Voltage / C75', 'C75': 'Overlaps C77',
 'C01': 'Root cause (ISO 14224 "cause" class: design/fabrication/installation/operation/maintenance) — different level from mechanisms like C12 Wear, C42 Corrosion',
 'C02': 'Root cause (maintenance) — overlaps C14 Lack of Maintenance', 'C14': 'Overlaps C02', 'C03': 'Root cause (installation)', 'C04': 'Root cause (operation)',
 'C12': 'Failure mechanism (ISO 14224) — mixed in the same list as root causes', 'C42': 'Failure mechanism', 'C13': 'Failure mechanism; overlaps C46 Lack of cooling',
 'C09': 'Component damage (what failed), not why', 'C60': 'Component damage (what failed), not why', 'C65': 'Component failure; duplicates Problem P39', 'C70': '"End of Life Cycle" — obsolescence, belongs to AM-009 status rather than a WO cause',
}
NOTES_A = {
 'A63': 'Duplicate of A75 "Serial cable replaced"', 'A75': 'Duplicate of A63', 'A66': 'Duplicate of A79 "Power cable replaced" (and of A25 Cable Replaced)', 'A79': 'Duplicate of A66 / A25',
 'A67': 'Duplicate of A78 "Firmware recovered/updated"', 'A78': 'Duplicate of A67', 'A64': 'Near-duplicate of A76 Buss bar screws tightened and A28 Bolt/s Tightened', 'A76': 'Near-duplicate of A64 / A28',
 'A65': 'Triplicate: A65 Replace PSU = A73 Replace Power Supply = A18 Power Supply Replaced', 'A73': 'Triplicate with A65 / A18', 'A18': 'Triplicate with A65 / A73',
 'A62': 'Overlaps A71 Replace Damaged Components and A17 Part Replaced', 'A71': 'Overlaps A62 / A17', 'A70': 'Duplicate of A13 Fan Replaced', 'A13': 'Duplicate of A70',
 'A68': 'Overlaps A11 Electronic Board Replaced', 'A77': 'Overlaps A23 Asset De-Commissioned', 'A74': 'Catch-all "TBD" — remove', 'A72': 'Troubleshooting step, not a corrective action',
 'A60': 'Imperative tense ("Replace X") — the A01–A33 list uses past tense ("X Replaced"): two lists merged', 'A69': 'Imperative tense; miner-specific',
}
def tense(desc):
    d = desc.strip().lower()
    return 'Imperative' if re.match(r'^(replace|tighten|recover|reconfigure|try|remove)\b', d) else 'Past / noun'
def build_list(df_, miner, notes, kind):
    out = df_.copy(); out['Code'] = out['Code'].str.strip(); out['Description'] = out['Description'].str.strip()
    dup = out['Description'].str.lower().str.replace(r"[^a-z0-9 ]", '', regex=True).str.replace(r'\s+', ' ', regex=True).duplicated(keep=False)
    out['Exact duplicate description'] = np.where(dup, 'Yes', '')
    out['Miner-specific'] = np.where(out['Code'].isin(miner), 'Yes', '')
    out['Catch-all'] = np.where(out['Description'].str.contains(r'\b(other|unknown|unkown|tbd)\b', case=False, regex=True), 'Yes', '')
    out['Code format'] = np.where(out['Code'].str.match(r'^[PCA]\d{2}$'), '', 'Non-standard')
    if kind == 'A': out['Tense'] = out['Description'].map(tense)
    out['Review note'] = out['Code'].map(notes).fillna('')
    out['Issue'] = np.where((out['Exact duplicate description'] == 'Yes') | (out['Catch-all'] == 'Yes') | (out['Review note'] != '') | (out['Code format'] != ''), 'Yes', '')
    return out
PL = build_list(P, MINER_P, NOTES_P, 'P'); CL = build_list(C, MINER_C, NOTES_C, 'C'); AL = build_list(A, MINER_A, NOTES_A, 'A')

ws = wb.create_sheet('Failure Codes')
title(ws, 'Problem / Cause / Action code lists — review against AM-001 G-5.16 and AM-005', f'{len(PL)} problem, {len(CL)} cause, {len(AL)} action codes configured (FailureCodeProblemsCausesActions.csv, ' + SNAPSHOT_LABEL + '). Lists are flat — no problem→cause→action linkage and no filtering by asset category is visible.')
r = 4
summary_rows = [
 ('Problem codes', len(PL), int((PL['Issue'] == 'Yes').sum()), int((PL['Miner-specific'] == 'Yes').sum()), int((PL['Exact duplicate description'] == 'Yes').sum() // 2), int((PL['Catch-all'] == 'Yes').sum())),
 ('Cause codes', len(CL), int((CL['Issue'] == 'Yes').sum()), int((CL['Miner-specific'] == 'Yes').sum()), int((CL['Exact duplicate description'] == 'Yes').sum() // 2), int((CL['Catch-all'] == 'Yes').sum())),
 ('Action codes', len(AL), int((AL['Issue'] == 'Yes').sum()), int((AL['Miner-specific'] == 'Yes').sum()), int((AL['Exact duplicate description'] == 'Yes').sum() // 2), int((AL['Catch-all'] == 'Yes').sum())),
]
header(ws, r, ['List', 'Codes', 'Codes with a review note', 'Miner-specific', 'Exact duplicate pairs', 'Catch-all codes']); r += 1
for row in summary_rows:
    for i, v in enumerate(row): ws.cell(row=r, column=1 + i, value=v)
    style_range(ws, r, r, 1, 6, numfmt=INT); r += 1
r += 1
findings = [
 'Two lists have been merged: a facilities list (P01–P24, C01–C55, A01–A33, past-tense actions) and a miner-repair list (P30–P40, C60–C78, A60–A80, imperative actions such as "Replace Hash Board", "Get Bitmain Power Status Failed"). The Strategy scope is Facilities assets; miner codes should be a separate set or filtered by asset category so that a technician closing a transformer WO does not scroll past "Replace LDO".',
 'Exact duplicates exist in every list (e.g. P007/P12 "Doesn\'t turn ON"; C28/C74 "Control Board Malfunction"; A63/A75, A66/A79, A67/A78, A13/A70; A65/A73/A18 power supply replaced three ways). Duplicates split the statistics AM-005 bad-actor analysis depends on.',
 'The Cause list mixes three levels that ISO 14224 keeps apart: root causes (C01 Improper material, C03 Installation error, C04 Operating error, C14 Lack of maintenance), failure mechanisms (C12 Wear and tear, C42 Corrosion, C13 Overheating) and component failures / symptoms (C60 Burnt hash board, C65 Dead PSU, C43 Fluid leak, C31 Loss of power). A single pick-list cannot trend any of them cleanly. Recommend splitting into Mechanism and Cause, or at minimum tagging each code with its level.',
 'Several codes sit in the wrong list: P33 Control Board / Firmware Issues and P39 Power Supply Issue are causes; C43 Fluid leak and C31 Loss of Power are symptoms; A72 "Try another Port/PDU/Section" is a troubleshooting step.',
 'Catch-alls: P40 "Other / Unkown", C05 "Unknown Cause", C73 "TBD", A74 "TBD". Keep one Unknown per list, require free text when it is chosen, and report its usage rate (target declining).',
 'Code numbering has gaps (P07, P18, P25–P29; C17–C19, C32–C39, C56–C59; A12, A34–A59) and two non-standard codes (P007, P008). Gaps usually mean codes were deleted — check whether historical WOs reference them.',
 'Usage: despite 151 configured codes, Problem is blank on 100% of the 18,384 closed WOs and Cause on 94% of failure WOs, so the lists have never been exercised. Clean the lists first, then make Problem/Cause/Action mandatory at close for Corrective, Breakdown, Damage and Warranty (G-5.16); enforcing an un-cleaned 63-item cause list will drive "Unknown" usage.',
 'No problem→cause→action dependency or asset-category filtering is visible in the export. If the CMMS supports linking codes to asset categories, that is the single biggest usability lever for adoption.',
]
ws.cell(row=r, column=1, value='Findings').font = f(True); r += 1
for t in findings:
    c = ws.cell(row=r, column=1, value=t); c.font = f(); c.alignment = Alignment(wrap_text=True, vertical='top'); ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6); ws.row_dimensions[r].height = 15 * (len(t) // 130 + 1) + 6; r += 1
r += 1
for name, tbl in [('Problem codes', PL), ('Cause codes', CL), ('Action codes', AL)]:
    ws.cell(row=r, column=1, value=name).font = f(True, size=12); r += 1
    cols = list(tbl.columns); header(ws, r, cols, height=30); r += 1
    for row in tbl.itertuples(index=False):
        for i, v in enumerate(row): ws.cell(row=r, column=1 + i, value=v)
        style_range(ws, r, r, 1, len(cols), wrap=True); r += 1
    r += 1
for c, wd in zip('ABCDEFGHI', [10, 44, 14, 14, 12, 14, 14, 70, 8]): ws.column_dimensions[c].width = wd

# order sheets
order = ['Read Me', 'Summary', 'Gap Rules', 'AM-001 Traceability', 'Bad Actor Candidates', 'Open Backlog', 'Permissions Mapping', 'Failure Codes', 'By Site', 'By Type', 'By Year', 'Field Completeness', 'Placeholder Notes', 'Next Phase Standard', 'Data', 'Open Backlog Data']
wb._sheets = [wb[s] for s in order]
out = str(OUTPUTS['wo'])
wb.save(out); print('saved', out)
