"""Step 2 — score the asset register against Strategy G-13.2, Section 7 and Section 8; cross-reference WO history.
Inputs : AllAssets.csv, AssetWarranty.csv, work/wo_scored.pkl, work/bad_actors.pkl
Outputs: outputs/Asset_Register_Gap_Analysis.xlsx, work/assets_scored.pkl
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

# ------------------------------------------------------------------ load
df = pd.read_csv(INPUTS['assets'], dtype=str, keep_default_na=False)
df.columns = [c.strip() for c in df.columns]
for c in df.columns: df[c] = df[c].str.strip()
df = df.drop(columns=['Created Date', 'Last Price Currency'])
n = len(df)

wo = pd.read_pickle(WORK['wo_scored'])
ba = pd.read_pickle(WORK['bad_actors'])

# ------------------------------------------------------------------ derived
df['Parent Code'] = df['Asset Location'].str.extract(r'\(([^()]+)\)\s*$')[0].fillna('')
codes = set(df['Code']); parent = dict(zip(df['Code'], df['Parent Code']))
def depth(c):
    d = 0; seen = set()
    while parent.get(c, '') != '' and c not in seen:
        seen.add(c); c = parent[c]; d += 1
        if c not in parent: return d
    return d
df['Hierarchy depth'] = df['Code'].map(depth)
_cd = pd.to_datetime(df['Created'], format=FMT_ASSETS)
_day = _cd.dt.date.astype(str)
_big = _day.value_counts(); _big = set(_big[_big >= 500].index)
df['Creation batch'] = np.where(_day.isin(_big), 'Bulk load ' + _day, 'Manual / small batch')
df['Created quarter'] = _cd.dt.to_period('Q').astype(str)
is_site = df['Is Site'].eq('on')
LOC = LOCATION_CATEGORY_PREFIXES
is_loc = df['Category'].str.startswith(LOC) | is_site
df['Record Level'] = np.where(is_loc, 'Location / Area', 'Equipment')
is_eq = ~is_loc

PH = {'N/A', 'NA', 'N.A.', 'UNKNOWN', 'TBD', 'TBC', '-', '0', 'NONE', 'NIL', 'NOT AVAILABLE', 'SEE TAG', 'X', '?', 'UNK'}
def is_ph(s): return s.str.upper().str.strip().isin(PH)
mk, md, sn = df['Make'], df['Model'], df['Serial Number']

# Make normalisation
def norm_make(s):
    s = s.upper().replace('.', '').replace(',', '')
    s = re.sub(r'\b(INC|LTD|LLC|CORP|CORPORATION|CO|COMPANY|ELECTRIC|ELECTRONICS|ELECTRICAL|SOLUTIONS|INSTALLATIONS|GMBH|SA|LIMITED)\b', '', s)
    s = re.sub(r'[^A-Z0-9]+', ' ', s).strip()
    return s.split(' ')[0] if s else ''
df['Make (normalised)'] = mk.map(norm_make)
variants = df[mk.ne('') & ~is_ph(mk)].groupby('Make (normalised)')['Make'].nunique()
df['Make variants for this manufacturer'] = df['Make (normalised)'].map(variants).fillna(0).astype(int)

sci = sn.str.match(r'^\d+(\.\d+)?E\+\d+$', case=False)
sn_valid = sn.ne('') & ~is_ph(sn) & ~sci
dup_serial = sn_valid & df.assign(k=mk.str.upper() + '|' + sn.str.upper()).duplicated('k', keep=False)
dup_code = df['Code'].duplicated(keep=False)

SITECODE = SITE_CODES
prefix = df['Code'].str.split('-').str[0]
expected_prefix = df['Site'].map(SITECODE)
prefix_ok = (prefix == expected_prefix) | is_site

# WO cross-reference
wo_codes = set(wo['Asset Code'])
wo_names = set(n_.strip() for s in wo['Asset Name'] for n_ in s.split(','))
df['WO history (by code)'] = np.where(df['Code'].isin(wo_codes), 'Yes', 'No')
df['WO history (by code or name)'] = np.where(df['Code'].isin(wo_codes) | df['Name'].isin(wo_names), 'Yes', 'No')
v10col = [c for c in wo.columns if c.startswith('V10')][0]
not_active_codes = set(wo.loc[wo[v10col] == 1, 'Asset Code'])
df['WOs closed as asset not active'] = np.where(df['Code'].isin(not_active_codes), 'Yes', '')
ba_codes = dict(zip(ba['Asset Code'], ba['Max failure WOs in any 12 months']))
df['Max failure WOs in 12 months'] = df['Code'].map(ba_codes).fillna(0).astype(int)

def flag(cond, applicable=None):
    out = cond.astype(int)
    if applicable is not None: out = out.where(applicable, other=np.nan)
    return out

flags = {}
flags['A01 Criticality group missing']                 = flag(df['Asset Criticality'].eq(''), ~is_site)
flags['A02 Make missing (equipment)']                  = flag(mk.eq(''), is_eq)
flags['A03 Make is a placeholder']                     = flag(is_ph(mk), mk.ne(''))
flags['A04 Make spelled more than one way']            = flag(df['Make variants for this manufacturer'] > 1, mk.ne('') & ~is_ph(mk))
flags['A05 Model missing (equipment)']                 = flag(md.eq(''), is_eq)
flags['A06 Model is a placeholder']                    = flag(is_ph(md), md.ne(''))
flags['A07 Serial number missing (equipment)']         = flag(sn.eq(''), is_eq)
flags['A08 Serial number is a placeholder']            = flag(is_ph(sn), sn.ne(''))
flags['A09 Serial number corrupted (scientific notation)'] = flag(sci, sn.ne(''))
flags['A10 Serial number duplicated (same make)']      = flag(dup_serial, sn_valid)
flags['A11 Manufacturer part number missing (equipment)'] = flag(df['Manufacturer Part Number'].eq(''), is_eq)
flags['A12 No parent in hierarchy (non-site)']         = flag(df['Parent Code'].eq(''), ~is_site)
flags['A13 Parent code not in register (orphan)']      = flag(df['Parent Code'].ne('') & ~df['Parent Code'].isin(codes))
flags['A14 Duplicate asset code']                      = flag(dup_code)
flags['A15 Code prefix does not match site code']      = flag(~prefix_ok & df['Site'].ne('(No Site)'))
flags['A16 Site missing']                              = flag(df['Site'].eq('(No Site)'))
flags['A17 Description missing']                       = flag(df['Description'].eq(''))
flags['A18 Status missing']                            = flag(df['Asset Status'].eq(''))
flags['A19 Status "on" but PMs closed as asset not active'] = flag(df['Asset Status'].eq('on') & df['Code'].isin(not_active_codes))
flags['A20 Equipment with no work order history']      = flag(df['WO history (by code or name)'].eq('No'), is_eq & ~df['Category'].str.startswith('HPC'))
F = pd.DataFrame(flags)
df = pd.concat([df, F], axis=1)
core = [c for c in F.columns if c[:3] not in ('A19', 'A20')]
df['Gap Count'] = F[core].fillna(0).sum(axis=1).astype(int)
FIRST, LAST = 2, n + 1
data_cols = list(df.columns); rule_cols = list(F.columns)

# ------------------------------------------------------------------ styles
FONT = 'Arial'
def f(bold=False, color='000000', size=10, italic=False): return Font(name=FONT, bold=bold, color=color, size=size, italic=italic)
HDR_FILL = PatternFill('solid', fgColor='1F3864'); INPUT_FILL = PatternFill('solid', fgColor='FFFF00')
thin = Side(style='thin', color='BFBFBF'); BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
PCT = '0.0%'; INT = '#,##0'
def header(ws, row, values, col=1, height=30):
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=col + i, value=v); c.font = f(True, 'FFFFFF'); c.fill = HDR_FILL; c.border = BORDER
        c.alignment = Alignment(wrap_text=True, vertical='center')
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
# ------------------------------------------------------------------ Data
wsD = wb.active; wsD.title = 'Data'
wsD.append(data_cols)
for row in df.itertuples(index=False):
    wsD.append([None if (isinstance(v, float) and np.isnan(v)) else v for v in row])
for ci in [data_cols.index('Serial Number') + 1, data_cols.index('Model') + 1, data_cols.index('Name') + 1, data_cols.index('Description') + 1, data_cols.index('Notes') + 1]:
    for cell in wsD.iter_rows(min_row=2, min_col=ci, max_col=ci):
        if cell[0].data_type in ('e', 'f'): cell[0].data_type = 's'
for c in range(1, len(data_cols) + 1):
    cell = wsD.cell(row=1, column=c); cell.font = f(True, 'FFFFFF'); cell.fill = HDR_FILL; cell.alignment = Alignment(wrap_text=True, vertical='center')
    wsD.column_dimensions[L(c)].width = 8 if data_cols[c - 1] in rule_cols else 14
for nm in ['Name', 'Code', 'Asset Location', 'Description']: wsD.column_dimensions[L(data_cols.index(nm) + 1)].width = 32
wsD.freeze_panes = 'C2'; wsD.auto_filter.ref = f'A1:{L(len(data_cols))}{LAST}'; wsD.row_dimensions[1].height = 60
col = {nm: L(i + 1) for i, nm in enumerate(data_cols)}
def rng(nm): return f"Data!${col[nm]}${FIRST}:${col[nm]}${LAST}"

sites = list(df['Site'].value_counts().index)
cats = list(df.loc[is_eq, 'Category'].value_counts().head(30).index)

RULES = {
 'A01': ('Criticality group missing', 'Completeness', 'All assets except site roots', 'High', 'Strategy G-7.1, G-7.2, G-13.2',
         'Criticality drives every governance intensity in the framework (tactics, change control, spares group, defect trigger, audit). Only Childress-HPC has assigned it (1,525 assets); 94% of the register and 99% of work orders have no group.',
         'Criticality Map exercise has been done for one site only; field not mandatory at onboarding.',
         'Run the Table 7-1 scoring for Level 3 equipment at every site, Group A candidates first (transformers, switchgear, gensets, UPS, cooling); make the field mandatory at onboarding (G-8.2).'),
 'A02': ('Make missing on equipment', 'Completeness', 'Equipment records', 'High', 'Strategy G-13.2', 'Warranty, recall and spares matching impossible without the manufacturer.', 'Not mandatory at onboarding; bulk loads without OEM data.', 'Mandatory at onboarding; backfill from nameplates during PM rounds.'),
 'A03': ('Make is a placeholder (N/A etc.)', 'Validity', 'Assets with a Make', 'Medium', 'Strategy G-13.2', 'Looks populated but carries nothing; hides the gap from completeness reports.', 'Mandatory field satisfied with N/A.', 'Disallow placeholder values; use an explicit "Unknown — to verify" status.'),
 'A04': ('Manufacturer spelled more than one way', 'Consistency', 'Assets with a real Make', 'Medium', 'Strategy G-8.5 (master data)', 'DELL/Dell, TRANE/Trane, two ITI variants: spend, warranty and spares analysis by manufacturer fragments.', 'Free-text Make field.', 'Controlled manufacturer list; merge variants.'),
 'A05': ('Model missing on equipment', 'Completeness', 'Equipment records', 'High', 'Strategy G-13.2', 'Cannot identify the correct spare, manual or OEM bulletin.', 'As A02.', 'As A02.'),
 'A06': ('Model is a placeholder', 'Validity', 'Assets with a Model', 'Low', 'Strategy G-13.2', 'As A03.', 'As A03.', 'As A03.'),
 'A07': ('Serial number missing on equipment', 'Completeness', 'Equipment records', 'High', 'Strategy G-13.2; AM-002 G-8.14 (RIK serials); AM-001 G-8.9 (vendor reports)',
         'Three quarters of equipment has no serial. Without it there is no warranty claim, no recall match, and no way to follow a unit that is swapped between positions.',
         'Not captured at receiving/commissioning; bulk loads; no field walkdown.',
         'Capture at receiving (AM-002 8.1) and commissioning; barcode/QR walkdown for Group A/B equipment; mandatory for Level 3/4 equipment, n/a for locations.'),
 'A08': ('Serial number is a placeholder', 'Validity', 'Assets with a Serial', 'Medium', 'Strategy G-13.2', '423 assets carry "N/A" as a serial.', 'As A03.', 'As A03.'),
 'A09': ('Serial number corrupted (scientific notation)', 'Validity', 'Assets with a Serial', 'High', 'Strategy G-13.2; AM-007 G-AM-007.3',
         'Values like "8.8224E+11" are long numeric serials that were opened in Excel and re-imported — the real serial is lost and the same corrupted value now appears on many assets.',
         'Bulk import via spreadsheet without text formatting.', 'Re-capture from nameplates; import serials as text; add a format validation on the field.'),
 'A10': ('Serial number duplicated on assets of the same make', 'Validity', 'Assets with a valid Serial', 'High', 'Strategy G-13.2',
         'A serial identifies one physical unit. The same serial on 10 PDUs means either copy-paste or that one position record was cloned; either way the physical unit cannot be traced.',
         'Cloning asset records without clearing the serial.', 'Uniqueness check on Make + Serial; walkdown of duplicates.'),
 'A11': ('Manufacturer part number missing', 'Completeness', 'Equipment records', 'Medium', 'Strategy G-13.2; AM-002 G-12.1 (link to spares)', 'Field exists but is blank on 100%; it is the natural key to the inventory part master.', 'Field never used.', 'Populate for Level 4 components and RIK items; link to inventory item codes.'),
 'A12': ('No parent in the hierarchy', 'Hierarchy', 'All assets except site roots', 'High', 'Strategy G-8.1 (parent-child tagged)', '2,616 non-site records float outside the hierarchy, so they roll up to nothing and inherit no location or criticality.', 'Assets created without selecting a location; legacy imports.', 'Assign a parent to every record; block creation without a parent.'),
 'A13': ('Parent code not in the register (orphan)', 'Hierarchy', 'Assets with a parent', 'Low', 'Strategy G-8.1', 'Parent was deleted or renamed after children were created.', 'No referential control on rename/delete.', 'Re-point the children; retire via master data process.'),
 'A14': ('Duplicate asset code', 'Validity', 'All', 'High', 'Strategy G-8.4, G-8.5 (naming is the controlled key)', '113 codes appear twice (PDUs, PLC panels). Work orders and history split between two records.', 'Code not enforced unique in the CMMS.', 'Enforce uniqueness; merge the pairs.'),
 'A15': ('Code prefix does not match site code', 'Consistency', 'All with a site', 'Medium', 'Strategy G-8.4 (campus naming convention)', 'Prince George HPC assets use PG11B…PG32B and SRV prefixes instead of PG-…; 18 (No Site) assets carry CHI/CF codes.', 'HPC assets loaded under a different naming scheme.', 'Confirm the campus naming convention document; re-code or document the HPC exception.'),
 'A16': ('Site missing', 'Completeness', 'All', 'Medium', 'Strategy 8 (Level 0)', '31 assets in (No Site).', 'Created outside a site.', 'Assign site.'),
 'A17': ('Description missing', 'Completeness', 'All', 'Low', 'Strategy G-13.2', 'Mostly locations; 1,193 equipment records.', 'Not mandatory.', 'Mandatory at onboarding.'),
 'A18': ('Status missing', 'Completeness', 'All', 'Low', 'Strategy G-8.5 (lifecycle status)', 'Only on/off exists; 19 blank.', 'Field not set.', 'Set; replace on/off with the lifecycle statuses of G-8.5 / AM-009.'),
 'A19': ('Status "on" but PMs closed as asset not active', 'Consistency', 'All', 'High', 'Strategy G-8.5 lifecycle status; AM-001 V10 cross-reference',
         'Technicians are closing PMs as "not active / not installed / under construction" on assets whose register status is "on", so the CMMS keeps generating PMs for them.',
         'Lifecycle status is binary and not maintained; PM generation not tied to status.',
         'Introduce lifecycle statuses (Planned, Commissioning, Active, Not in service, Retired); suppress PM generation except for Active.'),
 'A20': ('Equipment with no work order history (review list)', 'Coverage', 'Equipment excl. HPC', 'Info', 'Strategy G-9.1 (PM minimum for Groups A–C), G-9.7 (tactic recorded per asset)',
         'Equipment that has never appeared on a WO by its own code or by name in a route. Caveat: work recorded against a parent section (e.g. "Section 141 - PDUs") is not attributed to the child PDUs, and Childress-HPC is still commissioning, so this overstates the true gap; use it as a review list, not a count.',
         'Tactics not assigned; assets registered but not in PM routes; assets not yet commissioned.',
         'Join with criticality once assigned; every Group A/B/C asset must be on at least one PM schedule.'),
}
NOT_MEASURABLE = [
 ('Strategy G-13.2', 'Installation date', 'No field. Created date exists but is a system load date: 86% of equipment was created on 10 bulk-load days (see By Creation Batch). Notes carry manufacturing dates on ~20 records.'),
 ('Strategy G-13.2', 'Warranty', 'Asset Warranty report holds 3 records out of 27,193 assets (1 Dell HPC server, 2 fan records with no site, make, model or serial). 13,848 equipment records were created within the last two years and would ordinarily still be under OEM warranty. Inventory warranty report is empty. See Warranty sheet.'),
 ('Strategy G-13.2 / G-9.4', 'Linked documentation (OEM manuals, drawings)', 'Separate table; not in this export (same 2,000-row limit).'),
 ('Strategy G-8.5 / G-13.2', 'Lifecycle status', 'System offers Active / Inactive only (exported as on/off). A location record "Assets Removed from Service" (ASSETS-REMOVED-FROM-SERVICE, no site) exists as a decommissioning mechanism but holds no assets. Interim convention proposed: set Inactive + Notes (date, reason), leave in hierarchy position.'),
 ('AM-009 G-AM-009.1', 'Obsolescence status (Supported / Declining / Constrained / Obsolete)', 'No field'),
 ('Strategy G-8.1', 'Hierarchy level tag (0–4)', 'No field; depth derived from parent chain runs 1–7 levels'),
 ('Strategy G-8.7', 'Value vs $500 materiality threshold', 'No cost field (Last Price Currency populated on 94 rows only)'),
 ('Strategy G-9.7', 'Maintenance strategy / tactic per asset', 'Not in export; approximated by WO history (A20)'),
 ('Strategy G-7.2', 'Six-dimension criticality score', 'Group only, on 1,525 assets; no score'),
 ('Strategy G-13.3 / AM-006 G-AM-006.4', 'CMMS produces KPI data without manual re-work', 'Reporting is capped at 2,000 rows per extract; the 27,193-record asset list took ~5 hours to pull. This is a system capability gap, not a data gap.'),
]

# ------------------------------------------------------------------ Read Me
ws = wb.create_sheet('Read Me', 0)
title(ws, 'CMMS Gap Analysis — Asset Register (all sites)', f'Source: AllAssets.csv, {n:,} records, {df["Code"].nunique():,} distinct codes, {df["Site"].nunique()} sites, created Nov 2022 – Sep 2026. Prepared ' + PREPARED_LABEL + '. Draft for review.')
lines = [
 ('Purpose', 'Measure the asset register against the Strategy Rev 3.0 requirements for asset master data (G-13.2), criticality (Section 7), hierarchy, naming and lifecycle status (Section 8) and the AM-009 obsolescence status, and cross-reference it to the work order history so that data gaps can be weighted by what the asset actually does.'),
 ('Headline', 'The register is large and well populated at the location level but thin at the equipment level: 94% of assets have no criticality group (only Childress-HPC has been scored), 75% of equipment has no serial number, 17% no manufacturer and 23% no model, and the Manufacturer Part Number field is empty on every record. 2,616 records sit outside the hierarchy and 113 codes are duplicated. Serial data that does exist has integrity problems: 423 "N/A" placeholders, 49 values corrupted to scientific notation by a spreadsheet round-trip, and 787 assets sharing a serial with another asset of the same make. Of the lifecycle fields the Strategy requires, install date and obsolescence do not exist, lifecycle status is Active/Inactive only, and the warranty table holds 3 records for 27,193 assets. The register was built by bulk import — 86% of equipment was created on ten days — and the serial/make/model gaps follow the import batches (By Creation Batch sheet), so the primary fix is the load template, not field discipline.'),
 ('Equipment vs location', 'Records whose category is a location, building, road, structure, piping run or general-infrastructure bucket (5,093) are treated as Location / Area; the other 22,100 as Equipment. Make, Model, Serial and Part Number rules apply to Equipment only, so a road without a serial number is not counted as a gap. The split is by category and can be refined.'),
 ('How to read it', 'Summary gives headline numbers and the cross-reference to work orders. Gap Rules lists each rule with counts, clause, impact and recommendation. Strategy Traceability lists every G-13.2 attribute and whether the export can show it. By Site and By Category are heat maps of fail rate. Field Completeness is the raw blank rate per column. Data carries every asset with derived columns (parent code, depth, record level, WO history, failure count) and a 1/0/blank flag per rule.'),
 ('Assumptions', 'Parent is parsed from the code in parentheses at the end of Asset Location. Hierarchy depth is the number of parent hops to a root. Placeholders: N/A, NA, Unknown, TBD, -, 0, None, See Tag and similar. Make variants are detected by normalising case, punctuation and corporate suffixes and taking the first word. Serial duplicates are counted within the same Make. WO history matches on Asset Code, or on the asset Name appearing in a multi-asset work order route.'),
 ('Governing documents', 'Data Center Facilities Asset Management Strategy Rev 3.0 (third draft 08/21/2026); AM-001 Rev 3.0; AM-002 Rev 1.0. All are drafts pending approval and post-date the data, so this is a current-state baseline against the target state.'),
]
r = 4
for k, v in lines:
    ws.cell(row=r, column=1, value=k).font = f(True); c = ws.cell(row=r, column=2, value=v); c.font = f(); c.alignment = Alignment(wrap_text=True, vertical='top')
    ws.row_dimensions[r].height = max(30, 15 * (len(v) // 110 + 1)); r += 1
ws.column_dimensions['A'].width = 20; ws.column_dimensions['B'].width = 120

# ------------------------------------------------------------------ Summary
ws = wb.create_sheet('Summary', 1)
title(ws, 'Summary — Asset Register Data Quality', 'Formulas over the Data sheet.')
for c, w in zip('ABCD', [60, 14, 14, 60]): ws.column_dimensions[c].width = w
gc = rng('Gap Count'); lvl = rng('Record Level'); crit = rng('Asset Criticality')
r = 4; header(ws, r, ['Headline', 'Count', '% of base', 'Note']); r += 1
kp = [
 ('Asset records', f'=COUNTA({rng("Code")})', None, ''),
 ('  Equipment records', f'=COUNTIF({lvl},"Equipment")', '=B6/B5', 'By category'),
 ('  Location / Area records', f'=COUNTIF({lvl},"Location / Area")', '=B7/B5', ''),
 ('Distinct asset codes', df['Code'].nunique(), None, 'Static'),
 ('Assets with a criticality group', f'=COUNTIF({crit},"?*")', '=B9/B5', 'All at Childress-HPC'),
 ('  Group A', f'=COUNTIF({crit},"A")', None, ''), ('  Group B', f'=COUNTIF({crit},"B")', None, ''), ('  Group C', f'=COUNTIF({crit},"C")', None, ''), ('  Group D', f'=COUNTIF({crit},"D")', None, ''),
 ('Equipment with no serial number', f'=COUNTIF({rng(rule_cols[6])},1)', '=B14/B6', 'Base: equipment'),
 ('Equipment with no make', f'=COUNTIF({rng(rule_cols[1])},1)', '=B15/B6', 'Base: equipment'),
 ('Equipment with no model', f'=COUNTIF({rng(rule_cols[4])},1)', '=B16/B6', 'Base: equipment'),
 ('Serials that are placeholders / corrupted / duplicated', f'=COUNTIF({rng(rule_cols[7])},1)+COUNTIF({rng(rule_cols[8])},1)+COUNTIF({rng(rule_cols[9])},1)', f'=B17/COUNTIF({rng("Serial Number")},"?*")', 'Base: assets with a serial'),
 ('Records outside the hierarchy (no parent, not a site)', f'=COUNTIF({rng(rule_cols[11])},1)', '=B18/B5', ''),
 ('Duplicate asset codes (records)', f'=COUNTIF({rng(rule_cols[13])},1)', '=B19/B5', ''),
 ('Records with at least one master-data gap (A01–A18)', f'=COUNTIF({gc},">0")', '=B20/B5', ''),
 ('Average gaps per record', f'=AVERAGE({gc})', None, ''),
 ('Assets "on" whose PMs were closed as not active', f'=COUNTIF({rng(rule_cols[18])},1)', None, 'Cross-reference to WO rule V10'),
 ('Equipment (excl. HPC) with no work order history', f'=COUNTIF({rng(rule_cols[19])},1)', f'=B23/COUNT({rng(rule_cols[19])})', 'By code or by name in a route'),
 ('AM-005 bad-actor candidates (>=3 failure WOs/12 mo) with a criticality group', int(((df['Max failure WOs in 12 months'] >= 3) & df['Asset Criticality'].ne('')).sum()), None, f'of {int((df["Max failure WOs in 12 months"] >= 3).sum())} candidates — static'),
 ('Work orders whose asset has a criticality group', int((wo['Asset Code'].map(dict(zip(df['Code'], df['Asset Criticality']))).fillna('') != '').sum()), f'=B25/{len(wo)}', f'of {len(wo):,} closed WOs — static'),
]
for label, cnt, pct, note in kp:
    ws.cell(row=r, column=1, value=label); ws.cell(row=r, column=2, value=cnt)
    if pct: ws.cell(row=r, column=3, value=pct)
    ws.cell(row=r, column=4, value=note); style_range(ws, r, r, 1, 4)
    ws.cell(row=r, column=2).number_format = '0.00' if 'Average' in label else INT; ws.cell(row=r, column=3).number_format = PCT; r += 1
r += 1
ws.cell(row=r, column=1, value='Hierarchy depth (parent hops to root)').font = f(True); r += 1
header(ws, r, ['Depth', 'Records', '% of records', 'Note']); r += 1
for d in sorted(df['Hierarchy depth'].unique()):
    ws.cell(row=r, column=1, value=int(d)); ws.cell(row=r, column=2, value=f'=COUNTIF({rng("Hierarchy depth")},A{r})'); ws.cell(row=r, column=3, value=f'=B{r}/$B$5')
    ws.cell(row=r, column=4, value='Site roots + records with no parent' if d == 0 else ('Strategy taxonomy stops at Level 4 (component)' if d >= 5 else ''))
    style_range(ws, r, r, 1, 4); ws.cell(row=r, column=2).number_format = INT; ws.cell(row=r, column=3).number_format = PCT; r += 1
ws.freeze_panes = 'A5'

# ------------------------------------------------------------------ Gap Rules
ws = wb.create_sheet('Gap Rules', 2)
title(ws, 'Gap rules — fails, rate, clause, impact and recommendation', 'Counts are formulas over the flag columns on the Data sheet.')
header(ws, 4, ['Rule', 'Description', 'Category', 'Population', 'Applicable', 'Fails', 'Fail rate', 'Severity', 'Strategy / AM clause', 'Why it matters', 'Likely root cause', 'Recommendation'], height=36)
r = 5
for rc in rule_cols:
    rid = rc[:3]; m = RULES[rid]
    vals = [rid, m[0], m[1], m[2], f'=COUNT({rng(rc)})', f'=COUNTIF({rng(rc)},1)', f'=IF(E{r}=0,0,F{r}/E{r})', m[3], m[4], m[5], m[6], m[7]]
    for i, v in enumerate(vals): ws.cell(row=r, column=1 + i, value=v)
    style_range(ws, r, r, 1, 12, wrap=True); ws.cell(row=r, column=5).number_format = INT; ws.cell(row=r, column=6).number_format = INT; ws.cell(row=r, column=7).number_format = PCT; r += 1
ws.conditional_formatting.add(f'G5:G{r-1}', ColorScaleRule(start_type='num', start_value=0, start_color='FFFFFF', end_type='num', end_value=1, end_color='C00000'))
for c, w in zip('ABCDEFGHIJKL', [7, 36, 13, 22, 10, 9, 9, 9, 28, 50, 40, 50]): ws.column_dimensions[c].width = w
ws.freeze_panes = 'C5'; ws.auto_filter.ref = f'A4:L{r-1}'

# ------------------------------------------------------------------ Traceability
ws = wb.create_sheet('Strategy Traceability', 3)
title(ws, 'Strategy G-13.2 / Section 7 / Section 8 requirements vs the asset register export', 'Status: Met / Partial / Gap / Not measurable (no field — confirm whether it exists in the CMMS).')
header(ws, 4, ['Clause', 'Requirement', 'In export?', 'Current state', 'Status', 'Rule'], height=30)
tr = [
 ('Strategy G-13.2', 'Asset ID', 'Yes', '27,193 records; 113 duplicate codes', 'Partial', 'A14'),
 ('Strategy G-8.1 / G-13.2', 'Hierarchy position (level, parent)', 'Parent only', 'Parent derivable for 90%; 2,616 non-site records have none; no level tag; depth runs 1–7', 'Partial', 'A12, A13'),
 ('Strategy G-13.2', 'Location', 'Yes', 'Site on all but 31; Asset Location on 90%', 'Partial', 'A16'),
 ('Strategy G-7.1 / G-13.2', 'Criticality group', 'Yes', '5.6% populated, one site only', 'Gap', 'A01'),
 ('Strategy G-13.2', 'Manufacturer', 'Yes', '17% of equipment blank; 155 "N/A"; spelling variants', 'Partial', 'A02–A04'),
 ('Strategy G-13.2', 'Model', 'Yes', '23% of equipment blank; 218 "N/A"', 'Partial', 'A05, A06'),
 ('Strategy G-13.2', 'Serial number', 'Yes', '75% of equipment blank; 423 "N/A"; 49 corrupted; 787 duplicated', 'Gap', 'A07–A10'),
 ('Strategy G-13.2 / AM-002 G-12.1', 'Manufacturer part number', 'Yes (empty)', 'Blank on 100%', 'Gap', 'A11'),
 ('Strategy G-8.4 / G-8.6', 'Named per campus naming convention', 'Partly', 'Consistent SITE-… pattern except PG HPC (PG11B…, SRV) and 18 no-site records; convention document not yet reviewed', 'Partial', 'A15'),
 ('Strategy G-8.5', 'Lifecycle status', 'on/off only', '27,174 on, 19 blank, 0 off — yet PMs are being closed as "asset not active" on 60 of these', 'Gap', 'A18, A19'),
 ('Strategy G-8.2', 'Criticality assigned before operational handover', 'Indirect', 'Assets created through Sep 2026 without criticality', 'Gap', 'A01'),
 ('Strategy G-9.1 / G-9.7', 'Tactic recorded per asset; PM minimum for Groups A–C', 'Indirect (WO history)', 'Equipment never seen on a WO listed under A20; cannot be judged without criticality', 'Not measurable', 'A20'),
 ('AM-005 G-AM-005.1', 'Defect trigger by criticality group', 'Join', '1 of 133 bad-actor candidates has a criticality group', 'Gap', 'A01'),
] + [(c, req, 'No', state, 'Not measurable', '') for c, req, state in NOT_MEASURABLE]
r = 5
for row in tr:
    for i, v in enumerate(row): ws.cell(row=r, column=1 + i, value=v)
    style_range(ws, r, r, 1, 6, wrap=True); r += 1
for c, w in zip('ABCDEF', [26, 46, 16, 64, 16, 12]): ws.column_dimensions[c].width = w
ws.freeze_panes = 'A5'; ws.auto_filter.ref = f'A4:F{r-1}'

# ------------------------------------------------------------------ heat maps
def heatmap(name, seg_field, segments, ttl):
    ws = wb.create_sheet(name)
    title(ws, ttl, 'Fail rate = fails / applicable records in the segment; second block shows fail counts. Formulas over the Data sheet.')
    segr = rng(seg_field); r = 4
    header(ws, r, ['Rule', 'Description', 'All'] + segments, height=60); r += 1
    ws.cell(row=r, column=1, value='Records'); ws.cell(row=r, column=2, value='Assets in segment'); ws.cell(row=r, column=3, value=f'=COUNTA({rng("Code")})')
    for j, s in enumerate(segments): ws.cell(row=r, column=4 + j, value=f'=COUNTIF({segr},"{s}")')
    style_range(ws, r, r, 1, 3 + len(segments), numfmt=INT, bold=True); r += 1; top = r
    for rc in rule_cols:
        ws.cell(row=r, column=1, value=rc[:3]); ws.cell(row=r, column=2, value=RULES[rc[:3]][0])
        ws.cell(row=r, column=3, value=f'=IFERROR(COUNTIF({rng(rc)},1)/COUNT({rng(rc)}),"")')
        for j, s in enumerate(segments):
            ws.cell(row=r, column=4 + j, value=f'=IFERROR(COUNTIFS({rng(rc)},1,{segr},"{s}")/COUNTIFS({rng(rc)},">=0",{segr},"{s}"),"")')
        style_range(ws, r, r, 1, 3 + len(segments), numfmt=PCT); r += 1
    ws.conditional_formatting.add(f'C{top}:{L(3+len(segments))}{r-1}', ColorScaleRule(start_type='num', start_value=0, start_color='FFFFFF', mid_type='num', mid_value=0.5, mid_color='F4B183', end_type='num', end_value=1, end_color='C00000'))
    r += 1; header(ws, r, ['Rule', 'Description', 'All'] + segments, height=60); r += 1
    for rc in rule_cols:
        ws.cell(row=r, column=1, value=rc[:3]); ws.cell(row=r, column=2, value=RULES[rc[:3]][0] + ' — fails'); ws.cell(row=r, column=3, value=f'=COUNTIF({rng(rc)},1)')
        for j, s in enumerate(segments): ws.cell(row=r, column=4 + j, value=f'=COUNTIFS({rng(rc)},1,{segr},"{s}")')
        style_range(ws, r, r, 1, 3 + len(segments), numfmt=INT); r += 1
    ws.column_dimensions['A'].width = 7; ws.column_dimensions['B'].width = 40; ws.column_dimensions['C'].width = 9
    for j in range(len(segments)): ws.column_dimensions[L(4 + j)].width = 11
    ws.freeze_panes = 'D6'
heatmap('By Site', 'Site', sites, 'Fail rate by site')
heatmap('By Category', 'Category', cats, 'Fail rate by equipment category (top 30 by count)')
batches = list(df['Creation batch'].value_counts().index)
heatmap('By Creation Batch', 'Creation batch', batches, 'Fail rate by creation batch — bulk loads (>=500 assets created in one day) vs hand-entered records')

# ------------------------------------------------------------------ Field completeness
ws = wb.create_sheet('Field Completeness')
title(ws, 'Raw blank rate for every field in the export', 'COUNTBLANK over the Data sheet.')
req = {'Name': 'Required', 'Code': 'Required (unique)', 'Asset Status': 'Required (lifecycle)', 'Notes': 'Optional', 'Created': 'System', 'Site': 'Required', 'Is Site': 'System',
       'Make': 'Required (equipment)', 'Model': 'Required (equipment)', 'Category': 'Required', 'Asset Location': 'Required (parent)', 'Description': 'Required',
       'Serial Number': 'Required (equipment)', 'Manufacturer Part Number': 'Required (components / RIK)', 'Asset Criticality': 'Required'}
header(ws, 4, ['Field', 'Strategy requirement', 'Blank', 'Blank %', 'Distinct values']); r = 5
for c in [c for c in data_cols if c in req]:
    ws.cell(row=r, column=1, value=c); ws.cell(row=r, column=2, value=req[c]); ws.cell(row=r, column=2).fill = INPUT_FILL
    ws.cell(row=r, column=3, value=f'=COUNTBLANK({rng(c)})'); ws.cell(row=r, column=4, value=f'=C{r}/{n}'); ws.cell(row=r, column=5, value=int(df[c].nunique()))
    style_range(ws, r, r, 1, 5); ws.cell(row=r, column=3).number_format = INT; ws.cell(row=r, column=4).number_format = PCT; r += 1
ws.conditional_formatting.add(f'D5:D{r-1}', ColorScaleRule(start_type='num', start_value=0, start_color='FFFFFF', end_type='num', end_value=1, end_color='C00000'))
r += 1; ws.cell(row=r, column=1, value='Attributes the Strategy requires that do not appear in the export:').font = f(True); r += 1
for c, req_, state in NOT_MEASURABLE:
    ws.cell(row=r, column=1, value=req_); ws.cell(row=r, column=2, value=c); ws.cell(row=r, column=3, value=state); style_range(ws, r, r, 1, 3, wrap=True); r += 1
for c, w in zip('ABCDE', [44, 30, 40, 10, 14]): ws.column_dimensions[c].width = w
ws.freeze_panes = 'A5'

# ------------------------------------------------------------------ Warranty sheet
w = pd.read_csv(INPUTS['warranty'], dtype=str, keep_default_na=False)
ws = wb.create_sheet('Warranty')
title(ws, 'Asset Warranty report — all records', 'Reproduced in full from AssetWarranty.csv (report run ' + SNAPSHOT_LABEL + '). The Inventory Warranty report returned no records.')
header(ws, 4, list(w.columns) + ['Site (register)', 'Category (register)', 'Make (register)', 'Created (register)'])
reg = df.drop_duplicates('Code').set_index('Code')
r = 5
for row in w.itertuples(index=False):
    vals = list(row); code = row[0]
    vals += [reg.at[code, 'Site'], reg.at[code, 'Category'], reg.at[code, 'Make'], reg.at[code, 'Created']] if code in reg.index else ['not in register', '', '', '']
    for i, v in enumerate(vals): ws.cell(row=r, column=1 + i, value=v)
    style_range(ws, r, r, 1, len(vals)); r += 1
r += 1
_cd2 = pd.to_datetime(df['Created'], format=FMT_ASSETS); _eq2 = df['Record Level'] == 'Equipment'
notes = [
 f'{len(w)} warranty records exist against {n:,} assets ({(df["Record Level"] == "Equipment").sum():,} equipment).',
 f'{int((_eq2 & (_cd2 >= SNAPSHOT - pd.DateOffset(months=24))).sum()):,} equipment records were created in the last 24 months and {int((_eq2 & (_cd2 >= SNAPSHOT - pd.DateOffset(months=12))).sum()):,} in the last 12 — most OEM warranties on transformers, switchgear, drives, fans, servers and cooling run 12–36 months, so the majority of these are probably still in warranty with no record of it.',
 'Consequence (Strategy G-13.2; AM-001 G-8.9 "warranty terms shall be recorded"; AM-001 5.4 Warranty flag): failures on in-warranty equipment are being repaired at IREN cost. 163 work orders were typed Warranty in the WO export, so warranty work is happening without the register knowing which assets are covered.',
 'Two of the three records are fan assets with no site, make, model or serial — a warranty cannot be claimed on them as recorded.',
 'The report status is computed as IF(Warranty 1 - Expiry Date < NOW(), "Expired", "Valid for N days"), so expired warranties would appear as rows labelled Expired. None do: these three are the only warranties ever recorded on the register, not just the unexpired ones. ("Warranty 1" in the expression is the report alias for the joined warranty table.) Warranty is a child table of the asset, so an asset can hold multiple warranty records (e.g. parts / labour / extended) — the capability exists; it has been used three times.',
 'Recommendation: capture warranty start/end and terms at receiving (AM-002 8.1) and commissioning, as warranty records on the asset; backfill for equipment created in the last 24 months from purchase orders, starting with Group A/B candidate categories.',
]
for t in notes:
    c = ws.cell(row=r, column=1, value=t); c.font = f(); c.alignment = Alignment(wrap_text=True, vertical='top'); ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9); ws.row_dimensions[r].height = 32; r += 1
for c, wd in zip('ABCDEFGHI', [22, 36, 14, 16, 18, 20, 18, 14, 16]): ws.column_dimensions[c].width = wd

# ------------------------------------------------------------------ Make variants sheet (static)
ws = wb.create_sheet('Make Variants')
title(ws, 'Manufacturer names spelled more than one way', 'Static list from the source data, grouped by normalised name. Use to build the controlled manufacturer list.')
header(ws, 4, ['Normalised', 'Variant as entered', 'Assets'])
vt = df[mk.ne('') & ~is_ph(mk)].groupby(['Make (normalised)', 'Make']).size().reset_index(name='n')
vt = vt[vt['Make (normalised)'].map(variants) > 1].sort_values(['Make (normalised)', 'n'], ascending=[True, False])
r = 5
for row in vt.itertuples(index=False):
    for i, v in enumerate(row): ws.cell(row=r, column=1 + i, value=v)
    style_range(ws, r, r, 1, 3); ws.cell(row=r, column=3).number_format = INT; r += 1
for c, w in zip('ABC', [22, 44, 10]): ws.column_dimensions[c].width = w
ws.freeze_panes = 'A5'

wb._sheets = [wb[s] for s in ['Read Me', 'Summary', 'Gap Rules', 'Strategy Traceability', 'By Site', 'By Category', 'By Creation Batch', 'Field Completeness', 'Warranty', 'Make Variants', 'Data']]
out = str(OUTPUTS['assets'])
wb.save(out); print('saved', out)
print(pd.DataFrame({'applicable': F.notna().sum(), 'fails': F.eq(1).sum(), 'pct': (F.eq(1).sum() / F.notna().sum() * 100).round(1)}).to_string())
print('gap>0', (df['Gap Count'] > 0).sum(), 'avg', df['Gap Count'].mean().round(2), ' A19', F.iloc[:, 18].eq(1).sum(), ' sci', sci.sum())
df.to_pickle(WORK['assets_scored'])
