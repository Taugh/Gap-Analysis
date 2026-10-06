"""Step 6 — score the stock list and parts usage against AM-002.
Inputs : StockList.csv, PartsUsage.csv
Outputs: outputs/Inventory_Gap_Analysis.xlsx
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

# ------------------------------------------------------------------ load & score
df = pd.read_csv(INPUTS['stock'], dtype=str, keep_default_na=False)
df.columns = [c.strip() for c in df.columns]
for c in df.columns: df[c] = df[c].str.strip()
df = df.drop(columns=['Part URL'])
usage = pd.read_csv(INPUTS['parts_usage'], dtype=str, keep_default_na=False)
usage.columns = [c.strip() for c in usage.columns]
usage = usage[usage['Part Code'].str.strip() != '']

n = len(df)
mn = pd.to_numeric(df['Min. Qty'], errors='coerce')
oh = pd.to_numeric(df['Qty on Hand'], errors='coerce')
lp = pd.to_numeric(df['Last Price'], errors='coerce')
name = df['Stock Item']
words = name.str.split().str.len()
has_ref = name.str.contains(r'Supplier Ref|Part No|P/N|Model', case=False, regex=True)
has_token = name.str.contains(r'\b[A-Z0-9]*\d[A-Z0-9\-/]{4,}\b', regex=True)
codes_per_name = df.drop_duplicates('Part Code').groupby('Stock Item')['Part Code'].nunique()
shared_name = name.map(codes_per_name).fillna(1) > 1
rows_per_code = df.groupby('Part Code').size()
multi_loc = df['Part Code'].map(rows_per_code) > 1
data_bearing = name.str.contains(r'hard drive|ssd|nvme|memory|rdimm|processor|controller card|transceiver|adapter|adaptor', case=False, regex=True) | df['Category'].eq('PRT - HPC parts')
used_codes = set(usage['Part Code'])

df['Words in name'] = words
df['Priced'] = np.where(lp > 0, 'Yes', 'No')
df['Data-bearing / IT component'] = np.where(data_bearing, 'Yes', '')
df['Locations for code'] = df['Part Code'].map(rows_per_code)
df['Issued via CMMS 2019-2026'] = np.where(df['Part Code'].isin(used_codes), 'Yes', 'No')

def flag(cond, applicable=None):
    out = cond.astype(int)
    if applicable is not None: out = out.where(applicable, other=np.nan)
    return out

flags = {}
flags['I01 Description generic (<=2 words)']            = flag(words <= 2)
flags['I02 No manufacturer part number in record']       = flag(~(has_ref | has_token))
flags['I03 Same description used for different codes']  = flag(shared_name)
flags['I04 Unit cost (Last Price) missing or zero']      = flag(~(lp > 0))
flags['I05 Min quantity not set']                        = flag(mn.isna())
flags['I06 Min set but zero']                            = flag(mn.eq(0), mn.notna())
flags['I07 On hand below Min']                           = flag(oh < mn, mn.gt(0))
flags['I08 Zero on hand']                                = flag(oh.eq(0))
flags['I09 Not linked to a parent asset (BOM Groups)']   = flag(df['BOM Groups'].eq(''))
flags['I10 Bin location blank (Aisle/Row/Bin)']          = flag(df['Aisle'].eq('') & df['Row'].eq('') & df['Bin'].eq(''))
flags['I11 Account code blank']                          = flag(df['Account Code'].eq(''))
flags['I12 Charge dept / campus blank']                  = flag(df['Charge Dept Code'].eq(''))
flags['I13 Inventory / UNSPSC code blank']               = flag(df['Inventory Code'].eq('') & df['UNSPC Code'].eq(''))
flags['I14 Uncategorised (Other / Parts And Supplies)']  = flag(df['Category'].isin(['PRT - Other', 'Parts And Supplies']))
flags['I15 Test / dummy record']                         = flag(name.str.contains(r'\bdummy\b|\btest\b', case=False, regex=True) | df['Part Code'].eq('PRT-00000000'))
flags['I16 Data-bearing component without unit cost']    = flag(~(lp > 0), data_bearing)
flags['I17 No CMMS issue transaction 2019-2026']         = flag(~df['Part Code'].isin(used_codes))
F = pd.DataFrame(flags)
df = pd.concat([df, F], axis=1)
core = [c for c in F.columns if c[:3] not in ('I06', 'I07', 'I08', 'I17')]
df['Gap Count'] = F[core].fillna(0).sum(axis=1).astype(int)
FIRST, LAST = 2, n + 1
data_cols = list(df.columns)
rule_cols = list(F.columns)

# ------------------------------------------------------------------ styles
FONT = 'Arial'
def f(bold=False, color='000000', size=10, italic=False): return Font(name=FONT, bold=bold, color=color, size=size, italic=italic)
HDR_FILL = PatternFill('solid', fgColor='1F3864'); INPUT_FILL = PatternFill('solid', fgColor='FFFF00')
thin = Side(style='thin', color='BFBFBF'); BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
PCT = '0.0%'; INT = '#,##0'; CUR = '$#,##0.00'
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
    out = []
    for c, v in zip(data_cols, row):
        if isinstance(v, float) and np.isnan(v): out.append(None)
        elif c in ('Min. Qty', 'Qty on Hand', 'Last Price', 'Total Value') and v != '': out.append(float(v))
        else: out.append(v)
    wsD.append(out)
for c in range(1, len(data_cols) + 1):
    cell = wsD.cell(row=1, column=c); cell.font = f(True, 'FFFFFF'); cell.fill = HDR_FILL; cell.alignment = Alignment(wrap_text=True, vertical='center')
    wsD.column_dimensions[L(c)].width = 8 if data_cols[c - 1] in rule_cols else 14
wsD.column_dimensions['A'].width = 40; wsD.column_dimensions[L(data_cols.index('BOM Groups') + 1)].width = 30
wsD.freeze_panes = 'C2'; wsD.auto_filter.ref = f'A1:{L(len(data_cols))}{LAST}'; wsD.row_dimensions[1].height = 60
col = {nm: L(i + 1) for i, nm in enumerate(data_cols)}
def rng(nm): return f"Data!${col[nm]}${FIRST}:${col[nm]}${LAST}"

# ------------------------------------------------------------------ Usage sheet
wsU = wb.create_sheet('Parts Usage')
title(wsU, 'Parts usage export (PartsUsage.csv) — 2019 to date', 'Every issue transaction recorded in the CMMS inventory module over the period. Reproduced in full.')
header(wsU, 4, list(usage.columns))
for i, row in enumerate(usage.itertuples(index=False), 5):
    for j, v in enumerate(row): wsU.cell(row=i, column=1 + j, value=v)
    style_range(wsU, i, i, 1, len(usage.columns))
wsU.cell(row=7, column=1, value=f'{len(usage)} transaction(s) in total. Compare: 18,384 closed work orders in the same CMMS over Nov 2023 – Sep 2026 with Parts Used blank on 100%.').font = f(italic=True, color='595959')
for c, w in zip('ABCDEFG', [18, 10, 50, 12, 12, 14, 14]): wsU.column_dimensions[c].width = w

# ------------------------------------------------------------------ Rules metadata
RULES = {
 'I01': ('Description generic (2 words or fewer)', 'Completeness', 'High', 'AM-002 G-12.1 (description)', '"Feeder" ×18, "Unprotected Reducer", "Flng" — a technician cannot pick the right part, and duplicates cannot be detected.', 'Busway components loaded from a vendor BOM without specification text.', 'Adopt a noun–modifier description standard (type, rating, size, OEM ref); rewrite the busway items from the vendor drawing.'),
 'I02': ('No manufacturer part number in the record', 'Completeness', 'High', 'AM-002 G-12.1 (manufacturer part number, approved equivalents)', 'The export has no Manufacturer or Part Number field; only 11 items carry a "Supplier Ref" inside the description text. Re-ordering and equivalence checks are impossible from the CMMS.', 'Fields not configured on the part record, or not included in the report.', 'Add Manufacturer, Manufacturer Part No. and Approved Equivalent fields; make mandatory before an item can be stocked; migrate Supplier Refs out of the description.'),
 'I03': ('Same description on more than one part code', 'Consistency', 'Medium', 'AM-002 G-12.1; Strategy G-8.5 (master data)', '81 codes share a name with another code. Either the descriptions are incomplete or the codes are duplicates.', 'Generic descriptions (I01).', 'Resolve with I01; merge true duplicates through the master data process.'),
 'I04': ('Unit cost missing or zero', 'Completeness', 'High', 'AM-002 G-12.1 (unit cost); G-5.5 (financial reconciliation); Strategy G-8.7 (materiality)', '89% of lines have no price, so total stock value (reported $181k) is understated and the $500 materiality test cannot be applied.', 'Receipts not processed through the CMMS (no PO/receipt linkage); prices typed manually on a few items.', 'Receive stock against POs in the CMMS so Last Price populates; one-off backfill of unit cost from procurement for Group A/B items.'),
 'I05': ('Min quantity not set', 'Completeness', 'High', 'AM-002 G-7.3 (Min/Max configured for every stocked item)', 'No reorder point, so the CMMS cannot generate replenishment requests.', 'Min/Max never configured; no Max field exists in the export at all.', 'Set Min/Max for every stocked item, starting with Group A/B; confirm a Max field exists in the system.'),
 'I06': ('Min set but zero', 'Validity', 'Low', 'AM-002 G-7.1', 'A zero Min is equivalent to no reorder point.', 'Default value left in place.', 'Treat as unset; include in I05 remediation.'),
 'I07': ('On hand below Min', 'Performance', 'Medium', 'AM-002 G-7.3, G-8.17 (replenishment)', 'Items already below their reorder point with no automatic replenishment.', 'No replenishment trigger configured.', 'Raise requisitions now; enable automatic replenishment.'),
 'I08': ('Zero on hand', 'Performance', 'Info', 'AM-002 §13 Group A stockouts (target zero)', 'Includes the two 250/500 HP motors valued at $35k–$63k with zero on hand. Without spare-parts group the Group A stockout KPI cannot be computed.', 'Unknown — may be legitimate on-demand items.', 'Classify items (I18) then assess stockouts against Group A/B.'),
 'I09': ('Not linked to a parent asset', 'Completeness', 'High', 'AM-002 G-12.1, G-12.4 (linked parent assets)', 'Spares cannot be flagged for reallocation or disposal when the asset lifecycle changes, and PM parts lists cannot be linked (G-12.2).', 'BOM Groups populated only for busway, switchgear, PDU and HPC parts.', 'Link every stocked item to at least one Level 3 asset or asset class.'),
 'I10': ('Bin location blank', 'Completeness', 'Medium', 'AM-002 G-12.1 (bin location); 8.3 storage controls', 'Only a storage-area name is held (20 areas); no aisle/row/bin on any item, so cycle counts and issue verification depend on memory.', 'Fields exist but were never populated.', 'Assign bin locations during the first physical count.'),
 'I11': ('Account code blank', 'Completeness', 'Low', 'AM-002 G-12.5 (consumption reporting)', 'Consumption cannot post to an R&M account.', 'Default not set on category.', 'Default account code by part category.'),
 'I12': ('Charge dept / campus blank', 'Completeness', 'High', 'AM-002 G-12.5 (consumption by campus); Strategy 8 (Level 0 site)', 'Blank on 100% — the stock list does not say which campus holds the item; storage-area names are not site-qualified.', 'Inventory not tied to the site hierarchy.', 'Add site to every storage location and default charge dept from it.'),
 'I13': ('Inventory / UNSPSC classification blank', 'Completeness', 'Low', 'AM-002 G-12.1', 'Both classification fields are unused on 100% of items.', 'Fields not adopted.', 'Decide whether UNSPSC is wanted; otherwise remove the field to avoid false expectations.'),
 'I14': ('Uncategorised item', 'Consistency', 'Low', 'AM-002 §6', 'Items in Other / Parts And Supplies escape category-based defaults.', 'Catch-all category used at creation.', 'Reclassify.'),
 'I15': ('Test / dummy record in live stock', 'Validity', 'Low', 'AM-002 G-12.6 (master data control)', 'A test record with 2 on hand sits in the live list and inflates counts.', 'Not cleaned up after configuration.', 'Retire through the master data process.'),
 'I16': ('Data-bearing / IT component without unit cost', 'Completeness', 'High', 'Strategy G-8.8 (data-bearing items controlled regardless of value); AM-007 G-AM-007.4', 'Drives, memory, processors and transceivers (1,000+ transceivers) are controlled items irrespective of value and need cost for the financial register (G-5.5).', 'HPC spares loaded without cost.', 'Backfill cost from purchase records; apply the Group A/B issue controls of 6.1.'),
 'I17': ('No CMMS issue transaction 2019–2026', 'Process', 'High', 'AM-002 G-8.12, G-8.14 (every issue recorded; WO not closed with unrecorded consumption); §13 Unrecorded consumption = zero', 'One issue transaction exists for 314 stock lines over seven years, while 18,384 work orders were closed. Either parts are not being consumed through the store, or consumption is entirely unrecorded — on-hand quantities cannot be trusted.', 'Inventory module not used for issues; WO parts tab not used; no stores discipline.', 'Make the WO parts tab the only route for consumption; enforce G-8.11 (no release without WO); run a wall-to-wall count to re-baseline on-hand.'),
}
NOT_MEASURABLE = [
 ('AM-002 G-12.1', 'Manufacturer and manufacturer part number', 'No field in export (Supplier Ref in free text on 11 items)'),
 ('AM-002 G-12.1', 'Approved equivalents', 'No field'),
 ('AM-002 G-12.1 / Table 6-1', 'Supplier(s) and lead time; lead-time consequence score', 'No field'),
 ('AM-002 G-6.1 / G-6.4', 'Spare-parts group (A–D) and adjusted score', 'No field — the entire classification scheme of Section 6 is not represented'),
 ('AM-002 G-7.3', 'Max quantity', 'No field (Min only)'),
 ('AM-002 G-12.1 / G-8.12', 'Serial numbers on serialised spares', 'No field'),
 ('AM-002 G-12.1 / 8.2', 'Hazardous-material flag and SDS link', 'No field'),
 ('AM-002 G-12.1', 'Shelf-life / preservation requirement', 'No field'),
 ('AM-002 G-8.1–8.4', 'Receipts against PO; receiving discrepancy reports', 'No receipt transactions in export'),
 ('AM-002 §11', 'Cycle count records, count accuracy', 'No count data'),
 ('AM-002 G-12.3', 'PO status against item', 'No field'),
]

# ------------------------------------------------------------------ Read Me
ws = wb.create_sheet('Read Me', 0)
title(ws, 'CMMS Gap Analysis — Inventory (stock list and parts usage)', f'Source: StockList.csv ({n} stock lines, {df["Part Code"].nunique()} part codes, {df["Location Name"].nunique()} storage areas) and PartsUsage.csv ({len(usage)} transaction, 2019–2026). Prepared ' + PREPARED_LABEL + '. Draft for review.')
lines = [
 ('Purpose', 'Measure the inventory master data and transaction history against AM-002 Facilities Materials Management Rev 1.0 (first draft 09/23/2026) and the Asset Management Strategy Rev 3.0 (Section 8.3 materiality, G-8.8 controlled items, G-13.2). As with the work order analysis, the documents post-date the data, so this is a current-state baseline against the target state, not a compliance finding.'),
 ('Headline', 'The stock list is a location-keyed parts catalogue rather than a controlled inventory: 89% of lines carry no unit cost, 55% no reorder point, there is no manufacturer, part number, supplier, lead time, spare-parts group, Max, serial, hazmat or bin data at all, and one issue transaction was recorded in seven years against 18,384 work orders. The first remediation step is therefore process (receive and issue through the CMMS), not data clean-up.'),
 ('How to read it', 'Summary gives the headline numbers and the AM-002 Section 13 KPIs that can be baselined. Gap Rules lists each rule with counts, clause, impact and recommendation. AM-002 Traceability lists every G-12.1 field and whether the export can show it. Field Completeness is the raw blank rate per column. Data carries every stock line with a 1/0/blank flag per rule. Parts Usage reproduces the usage export in full.'),
 ('Assumptions', 'A stock line is one part code at one storage area (21 codes appear at more than one area). "Generic description" means two words or fewer. A part number is deemed present if the description contains "Supplier Ref" or a token of 5+ characters mixing letters and digits. Data-bearing / IT components are identified by keyword and the HPC parts category. Gap Count excludes the performance rules I06–I08 and I17.'),
 ('Legend', 'Yellow cells are inputs to review. Heat colours: darker red = higher failure rate.'),
]
r = 4
for k, v in lines:
    ws.cell(row=r, column=1, value=k).font = f(True); c = ws.cell(row=r, column=2, value=v); c.font = f(); c.alignment = Alignment(wrap_text=True, vertical='top')
    ws.row_dimensions[r].height = max(30, 15 * (len(v) // 110 + 1)); r += 1
ws.column_dimensions['A'].width = 18; ws.column_dimensions['B'].width = 120

# ------------------------------------------------------------------ Summary
ws = wb.create_sheet('Summary', 1)
title(ws, 'Summary — Inventory Data Quality', 'Formulas over the Data sheet.')
for c, w in zip('ABCD', [56, 14, 14, 62]): ws.column_dimensions[c].width = w
gc = rng('Gap Count'); pc = rng('Part Code')
r = 4; header(ws, r, ['Headline', 'Count', '% of base', 'Note']); r += 1
kpis = [
 ('Stock lines (part code × storage area)', f'=COUNTA({pc})', None, ''),
 ('Distinct part codes', df['Part Code'].nunique(), None, 'Static count'),
 ('Storage areas', df['Location Name'].nunique(), None, 'Static count; none is site-qualified'),
 ('Total on hand (units)', f'=SUM({rng("Qty on Hand")})', None, '1,000+ are transceivers and 1,000 air filters'),
 ('Stock value as reported', f'=SUM({rng("Total Value")})', None, 'Only priced items contribute — understated'),
 ('Lines with a unit cost', f'=COUNTIF({rng("Priced")},"Yes")', f'=B10/B5', ''),
 ('Lines with at least one master-data gap (I01–I05, I09–I16)', f'=COUNTIF({gc},">0")', f'=B11/B5', ''),
 ('Lines with 4 or more gaps', f'=COUNTIF({gc},">=4")', f'=B12/B5', ''),
 ('Average gaps per line', f'=AVERAGE({gc})', None, ''),
 ('Lines with no reorder point (Min blank)', f'=COUNTIF({rng(rule_cols[4])},1)', f'=B14/B5', ''),
 ('Lines not linked to an asset', f'=COUNTIF({rng(rule_cols[8])},1)', f'=B15/B5', ''),
 ('Lines with generic description', f'=COUNTIF({rng(rule_cols[0])},1)', f'=B16/B5', ''),
 ('Lines with no manufacturer part number', f'=COUNTIF({rng(rule_cols[1])},1)', f'=B17/B5', ''),
 ('Lines below Min', f'=COUNTIF({rng(rule_cols[6])},1)', f'=B18/COUNT({rng(rule_cols[6])})', 'Base: lines with Min > 0'),
 ('Lines with zero on hand', f'=COUNTIF({rng(rule_cols[7])},1)', f'=B19/B5', ''),
 ('Part codes issued through the CMMS, 2019–2026', len(used_codes), f'=B20/B6', 'From PartsUsage.csv'),
]
for label, cnt, pct, note in kpis:
    ws.cell(row=r, column=1, value=label); ws.cell(row=r, column=2, value=cnt)
    if pct: ws.cell(row=r, column=3, value=pct)
    ws.cell(row=r, column=4, value=note); style_range(ws, r, r, 1, 4)
    ws.cell(row=r, column=2).number_format = CUR if 'value' in label.lower() else ('0.00' if 'Average' in label else INT)
    ws.cell(row=r, column=3).number_format = PCT; r += 1
r += 1
ws.cell(row=r, column=1, value='AM-002 Section 13 KPIs — current baseline vs target').font = f(True); r += 1
header(ws, r, ['Indicator (AM-002 §13)', 'Current', 'Target', 'How measured here / caveat']); r += 1
kpi2 = [
 ('Unrecorded consumption (closed WOs with parts used but not recorded)', 'n/a', 0, 'Cannot be measured directly, but 1 issue in 7 years vs 18,384 WOs implies consumption is essentially unrecorded'),
 ('Min/Max review currency (items with Min/Max reviewed in last quarter)', f'=1-COUNTIF({rng(rule_cols[4])},1)/B5', 1.0, 'Upper bound: share of items with any Min at all; no Max field, no review date'),
 ('Group A stockouts', 'n/a', 0, 'No spare-parts group on any item'),
 ('Spare parts availability at scheduled start', 'n/a', 0.98, 'No parts reservations on WOs'),
 ('Inventory accuracy (count / value)', 'n/a', '97% / 99%', 'No count records; 89% of lines unpriced'),
 ('Zero-usage stock value (12 months, excl. insurance spares)', f'=SUMIFS({rng("Total Value")},{rng("Issued via CMMS 2019-2026")},"No")', 'Declining', 'By the CMMS record, every priced item is zero-usage'),
 ('Cycle count compliance; Replenishment cycle time; Receiving discrepancy rate; Long-lead register; Hazmat inspections; Turnover', 'n/a', 'various', 'No transactions, counts, POs, lead times or hazmat flags in the CMMS export'),
]
for label, cur, tgt, note in kpi2:
    ws.cell(row=r, column=1, value=label); ws.cell(row=r, column=2, value=cur); ws.cell(row=r, column=3, value=tgt); ws.cell(row=r, column=4, value=note)
    style_range(ws, r, r, 1, 4, wrap=True); ws.cell(row=r, column=3).font = f(color='0000FF')
    ws.cell(row=r, column=2).number_format = CUR if 'value' in label.lower() else PCT; ws.cell(row=r, column=3).number_format = PCT; r += 1
r += 1
ws.cell(row=r, column=1, value='Stock lines by storage area').font = f(True); r += 1
header(ws, r, ['Storage area', 'Lines', 'Unpriced lines', 'On hand (units)']); r += 1
for loc in df['Location Name'].value_counts().index:
    ws.cell(row=r, column=1, value=loc); ws.cell(row=r, column=2, value=f'=COUNTIF({rng("Location Name")},A{r})')
    ws.cell(row=r, column=3, value=f'=COUNTIFS({rng("Location Name")},A{r},{rng("Priced")},"No")'); ws.cell(row=r, column=4, value=f'=SUMIF({rng("Location Name")},A{r},{rng("Qty on Hand")})')
    style_range(ws, r, r, 1, 4, numfmt=INT); r += 1
ws.freeze_panes = 'A5'

# ------------------------------------------------------------------ Gap Rules
ws = wb.create_sheet('Gap Rules', 2)
title(ws, 'Gap rules — fails, rate, clause, impact and recommendation', 'Counts are formulas over the flag columns on the Data sheet.')
header(ws, 4, ['Rule', 'Description', 'Category', 'Applicable lines', 'Fails', 'Fail rate', 'Severity', 'AM-002 / Strategy clause', 'Why it matters', 'Likely root cause', 'Recommendation'], height=36)
r = 5
for rc in rule_cols:
    rid = rc[:3]; m = RULES[rid]
    vals = [rid, m[0], m[1], f'=COUNT({rng(rc)})', f'=COUNTIF({rng(rc)},1)', f'=IF(D{r}=0,0,E{r}/D{r})', m[2], m[3], m[4], m[5], m[6]]
    for i, v in enumerate(vals): ws.cell(row=r, column=1 + i, value=v)
    style_range(ws, r, r, 1, 11, wrap=True); ws.cell(row=r, column=4).number_format = INT; ws.cell(row=r, column=5).number_format = INT; ws.cell(row=r, column=6).number_format = PCT; r += 1
ws.conditional_formatting.add(f'F5:F{r-1}', ColorScaleRule(start_type='num', start_value=0, start_color='FFFFFF', end_type='num', end_value=1, end_color='C00000'))
for c, w in zip('ABCDEFGHIJK', [7, 36, 13, 10, 8, 9, 9, 30, 50, 40, 50]): ws.column_dimensions[c].width = w
ws.freeze_panes = 'C5'; ws.auto_filter.ref = f'A4:K{r-1}'

# ------------------------------------------------------------------ Traceability
ws = wb.create_sheet('AM-002 Traceability', 3)
title(ws, 'AM-002 G-12.1 mandatory inventory fields and related requirements vs the stock list export', 'Status: Met / Partial / Gap / Not measurable (no field in export — confirm whether the field exists in the CMMS).')
header(ws, 4, ['Clause', 'Requirement', 'In export?', 'Current state', 'Status', 'Rule'], height=30)
tr = [
 ('AM-002 G-12.1', 'Item code', 'Yes', '290 codes, all PRT-nnnn except one test record', 'Met', 'I15'),
 ('AM-002 G-12.1', 'Description', 'Yes', '24% generic (2 words or fewer); 81 codes share a description', 'Partial', 'I01, I03'),
 ('AM-002 G-12.1', 'Unit cost', 'Yes', 'Zero or blank on 89%', 'Gap', 'I04, I16'),
 ('AM-002 G-7.3', 'Min / Max for every stocked item', 'Min only', 'Min blank on 55%, zero on 9% more; no Max field', 'Gap', 'I05, I06'),
 ('AM-002 G-12.1', 'Bin location', 'Yes (empty)', 'Aisle/Row/Bin blank on 100%; 20 storage-area names only', 'Gap', 'I10'),
 ('AM-002 G-12.1 / G-12.4', 'Linked parent assets', 'Yes (BOM Groups)', 'Blank on 36%; links are to asset classes, not Level 3 assets', 'Partial', 'I09'),
 ('AM-002 G-12.5', 'Consumption by asset, category and campus', 'Partly', 'Charge dept blank 100%; account code blank 7%; no consumption to report', 'Gap', 'I11, I12, I17'),
 ('AM-002 G-8.12 / G-8.14', 'Every issue recorded; WO not closed with unrecorded consumption', 'Yes (usage export)', '1 transaction 2019–2026; WO Parts Used blank on 100% of 18,384 WOs', 'Gap', 'I17'),
 ('Strategy G-8.7', 'Items > $500 registered and controlled', 'Partly', '27 priced lines exceed $500; the other 279 lines cannot be tested', 'Not measurable', 'I04'),
 ('Strategy G-8.8 / AM-007 G-AM-007.4', 'Data-bearing components controlled regardless of value; sanitised before disposal', 'Partly', 'Drives, memory, processors, 1,000+ transceivers held; 5 of 13 HPC lines unpriced; no serial tracking', 'Gap', 'I16'),
] + [(c, req, 'No', state, 'Not measurable', '') for c, req, state in NOT_MEASURABLE]
r = 5
for row in tr:
    for i, v in enumerate(row): ws.cell(row=r, column=1 + i, value=v)
    style_range(ws, r, r, 1, 6, wrap=True); r += 1
for c, w in zip('ABCDEF', [24, 48, 16, 60, 16, 12]): ws.column_dimensions[c].width = w
ws.freeze_panes = 'A5'; ws.auto_filter.ref = f'A4:F{r-1}'

# ------------------------------------------------------------------ Field completeness
ws = wb.create_sheet('Field Completeness', 4)
title(ws, 'Raw blank rate for every field in the stock list export', 'COUNTBLANK over the Data sheet. Requirement column is the AM-002 standard (edit as needed).')
req = {'Stock Item': 'Required', 'Part Code': 'Required', 'Description': 'Required', 'Category': 'Required', 'Inventory Code': 'Optional', 'UNSPC Code': 'Optional',
       'Location Name': 'Required (site-qualified)', 'Account Code': 'Required', 'Account Description': 'Derived', 'Charge Dept Code': 'Required', 'Charge Dept Description': 'Derived',
       'Aisle': 'Required', 'Row': 'Required', 'Bin': 'Required', 'Min. Qty': 'Required (G-7.3)', 'Qty on Hand': 'Required', 'Last Price': 'Required (unit cost)', 'Total Value': 'Derived', 'BOM Groups': 'Required (linked assets)'}
header(ws, 4, ['Field', 'AM-002 requirement', 'Blank', 'Blank %', 'Distinct values'])
r = 5
for c in [c for c in data_cols if c in req]:
    ws.cell(row=r, column=1, value=c); ws.cell(row=r, column=2, value=req[c]); ws.cell(row=r, column=2).fill = INPUT_FILL
    ws.cell(row=r, column=3, value=f'=COUNTBLANK({rng(c)})'); ws.cell(row=r, column=4, value=f'=C{r}/{n}'); ws.cell(row=r, column=5, value=int(df[c].nunique()))
    style_range(ws, r, r, 1, 5); ws.cell(row=r, column=3).number_format = INT; ws.cell(row=r, column=4).number_format = PCT; r += 1
ws.conditional_formatting.add(f'D5:D{r-1}', ColorScaleRule(start_type='num', start_value=0, start_color='FFFFFF', end_type='num', end_value=1, end_color='C00000'))
r += 1
ws.cell(row=r, column=1, value='Fields AM-002 G-12.1 requires that do not appear in the export at all:').font = f(True); r += 1
for c, req_, state in NOT_MEASURABLE:
    ws.cell(row=r, column=1, value=req_); ws.cell(row=r, column=2, value=c); ws.cell(row=r, column=3, value=state); style_range(ws, r, r, 1, 3, wrap=True); r += 1
for c, w in zip('ABCDE', [40, 30, 34, 10, 14]): ws.column_dimensions[c].width = w
ws.freeze_panes = 'A5'

wb._sheets = [wb[s] for s in ['Read Me', 'Summary', 'Gap Rules', 'AM-002 Traceability', 'Field Completeness', 'Parts Usage', 'Data']]
out = str(OUTPUTS['inv'])
wb.save(out); print('saved', out)
print(pd.DataFrame({'applicable': F.notna().sum(), 'fails': F.eq(1).sum(), 'pct': (F.eq(1).sum() / F.notna().sum() * 100).round(1)}).to_string())
print('gap>0', (df['Gap Count'] > 0).sum(), 'gap>=4', (df['Gap Count'] >= 4).sum(), 'avg', df['Gap Count'].mean().round(2))
