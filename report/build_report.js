const fs = require('fs');
const path = require('path');
const ROOT = path.resolve(__dirname, '..');
const CH = path.join(ROOT, 'outputs', 'charts');
const REV = process.env.REPORT_REVISION || '0.1';
const OUT = path.join(ROOT, 'outputs', `CMMS_Gap_Analysis_Report_Rev${REV}.docx`);
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, AlignmentType,
  ShadingType, BorderStyle, ImageRun, PageBreak, TableOfContents, Header, Footer, PageNumber, LevelFormat,
  PageOrientation, VerticalAlign,
} = require('docx');

// ------------------------------------------------------------------ helpers
const FONT = 'Arial';
const NAVY = '1F3864', GREY = '595959', LIGHT = 'D9E1F2', ZEBRA = 'F5F7FA', BORDER = 'BFBFBF';
const P = (text, opts = {}) => new Paragraph({
  spacing: { after: opts.after ?? 140, before: opts.before ?? 0, line: 276 },
  alignment: opts.align, keepNext: opts.keepNext,
  children: Array.isArray(text) ? text : [new TextRun({ text, font: FONT, size: opts.size ?? 21, bold: opts.bold, italics: opts.italics, color: opts.color })],
});
const R = (text, o = {}) => new TextRun({ text, font: FONT, size: o.size ?? 21, bold: o.bold, italics: o.italics, color: o.color });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 160 }, children: [R(t, { size: 30, bold: true, color: NAVY })] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 280, after: 120 }, keepNext: true, children: [R(t, { size: 25, bold: true, color: NAVY })] });
const H3 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_3, spacing: { before: 200, after: 80 }, keepNext: true, children: [R(t, { size: 22, bold: true, color: GREY })] });
const bullet = (text, ref = 'bullets') => new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 80, line: 264 }, children: Array.isArray(text) ? text : [R(text)] });
const num = (text, ref = 'numbers') => new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 80, line: 264 }, children: Array.isArray(text) ? text : [R(text)] });
const num2 = (text) => num(text, 'numbers2');
const caption = (t) => P(t, { italics: true, size: 18, color: GREY, after: 200 });
const brk = () => new Paragraph({ children: [new PageBreak()] });
const img = (file, w, h) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: 60 }, keepNext: true, children: [new ImageRun({ type: 'png', data: fs.readFileSync(file), transformation: { width: w, height: h } })] });

const cellBorders = { top: { style: BorderStyle.SINGLE, size: 4, color: BORDER }, bottom: { style: BorderStyle.SINGLE, size: 4, color: BORDER }, left: { style: BorderStyle.SINGLE, size: 4, color: BORDER }, right: { style: BorderStyle.SINGLE, size: 4, color: BORDER } };
function table(headers, rows, widths, opts = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  const mk = (t, w, hdr, i, zebra) => new TableCell({
    width: { size: w, type: WidthType.DXA }, borders: cellBorders, verticalAlign: VerticalAlign.TOP,
    shading: hdr ? { type: ShadingType.CLEAR, fill: NAVY, color: 'auto' } : (zebra ? { type: ShadingType.CLEAR, fill: ZEBRA, color: 'auto' } : undefined),
    margins: { top: 60, bottom: 60, left: 90, right: 90 },
    children: [new Paragraph({ spacing: { after: 0, line: 252 }, alignment: (opts.numCols || []).includes(i) && !hdr ? AlignmentType.RIGHT : AlignmentType.LEFT,
      children: [R(String(t), { size: opts.size ?? 18, bold: hdr || ((opts.boldCols || []).includes(i)), color: hdr ? 'FFFFFF' : undefined })] })],
  });
  return new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    rows: [new TableRow({ tableHeader: true, cantSplit: true, children: headers.map((h, i) => mk(h, widths[i], true, i, false)) }),
      ...rows.map((r, ri) => new TableRow({ cantSplit: true, children: r.map((c, i) => mk(c, widths[i], false, i, ri % 2 === 1)) }))],
  });
}
const spacer = () => P('', { after: 120 });

// ------------------------------------------------------------------ content
const children = [];

// Cover
children.push(P('IREN', { bold: true, size: 28, color: NAVY, after: 60 }));
children.push(P('Data Center Facilities — Asset Management Governance System', { size: 22, color: GREY, after: 900 }));
children.push(P('CMMS Gap Analysis', { bold: true, size: 48, color: NAVY, after: 120 }));
children.push(P('Current-state baseline of CMMS data, configuration and PM program against the Asset Management Strategy, AM-001 and AM-002', { size: 26, color: GREY, after: 600 }));
children.push(P('Report to VP Engineering', { size: 22, after: 60 }));
children.push(P('Revision 0.1 — Draft for review — 28 September 2026', { size: 22, color: GREY, after: 600 }));
children.push(table(['Item', 'Detail'], [
  ['Document type', 'Assessment report (baseline)'],
  ['Prepared by', 'Troy Brannon, CMMS Administrator'],
  ['Reviewed by', '[Name / Role]'],
  ['Approved by', '[Name / Role]'],
  ['Distribution', 'VP Engineering; Senior Asset Manager (AROE); Maintenance Planning Manager; Site / Facilities Managers'],
  ['Supporting workbooks', 'WorkOrder_Gap_Analysis.xlsx · PM_Schedule_Gap_Analysis.xlsx · Asset_Register_Gap_Analysis.xlsx · Inventory_Gap_Analysis.xlsx'],
  ['Status of governing documents', 'Strategy Rev 3.0 (third draft, 21 Aug 2026); AM-001 Rev 3.0 (first draft, 9 Sep 2026); AM-002 Rev 1.0 (first draft, 23 Sep 2026) — all pending approval'],
], [2600, 6760]));
children.push(spacer());
children.push(P('All outputs in this report were generated from CMMS exports taken between 25 and 28 September 2026 and should be reviewed before decisions are taken on them. Figures are reproducible from the supporting workbooks, where every count is a live formula over the source records.', { italics: true, size: 18, color: GREY }));
children.push(brk());

// TOC
children.push(P('Contents', { bold: true, size: 28, color: NAVY, after: 200 }));
children.push(new TableOfContents('Contents', { hyperlink: true, headingStyleRange: '1-2' }));
children.push(P('Right-click and choose Update Field to refresh the table of contents after editing.', { italics: true, size: 16, color: GREY, before: 120 }));
children.push(brk());

// ------------------------------------------------------------------ Executive summary
children.push(H1('Executive Summary'));
children.push(P('IREN has issued three draft governance documents for Facilities asset management — the Data Center Facilities Asset Management Strategy (Rev 3.0), AM-001 Maintenance Workflow & Change Management (Rev 3.0) and AM-002 Facilities Materials Management (Rev 1.0). Together they define what the CMMS must hold and do: a criticality-driven asset register, a controlled hierarchy, mandatory work order fields, a PM program built from schedules, a transacting inventory, an eight-bucket permission model and a set of KPIs the system must produce without manual re-work.'));
children.push(P('This report measures where the CMMS stands today against that target. It draws on the complete closed work order history (18,384 records, November 2023 to September 2026), the open backlog (435), every PM schedule (3,005), the full asset register (27,193 records across nine sites), the stock list and seven years of parts transactions, the permission groups and the failure code lists. Because all three governing documents post-date the data, the results are a current-state baseline against the target state, not a compliance finding against a standard that was in force.'));
children.push(H2('Headline results'));
children.push(P('Against the AM-001 Section 10 indicators that can be measured or approximated from the exports, the baseline is well short of target in every case but one. Work order closure quality is about 40 percent against a 95 percent target; PM compliance is about 58 percent by the only proxy available; the "Other" category share is 5.8 percent against a 2 percent cap; and 99 percent of Highest-priority work orders are routine PMs rather than the breakdowns and emergencies the priority is reserved for. Only the reactive-work ratio is inside target. Backlog in weeks of ready work measures about two weeks at booked pace, below the 4–6 week band, though estimates and booked hours are missing widely enough that the figure is understated. Schedule compliance and triage response cannot yet be measured because the CMMS does not hold the data they need.'));
children.push(P('The trend, however, is strongly positive. Closure quality has risen from 12 percent of work orders in 2024 to 37 percent in 2025 and 67 percent so far in 2026; completion notes and labour hours have gone from largely absent to largely present over the same period. The field discipline is improving on its own. What has not moved is the structural layer underneath it: the register, the hierarchy, the PM templates, the code lists, the inventory module and the permission model, none of which have been aligned to the framework and most of which were built by bulk import or accreted over time.'));
children.push(H2('Five root causes explain most of the gaps'));
children.push(num([R('Criticality has not been assigned. ', { bold: true }), R('Only 1,525 of 27,193 assets (5.6 percent), all at Childress-HPC, carry a criticality group. Every governance intensity in the framework — maintenance tactic, change control, spares classification, defect trigger, audit frequency — is keyed to criticality, so almost none of it can be applied today. This is the prerequisite for the rest.')]));
children.push(num([R('Maintenance is recorded above the equipment level. ', { bold: true }), R('42 percent of failure work orders are written against buildings, sections or catch-all records rather than Level 3 equipment, and 2,616 assets sit outside the hierarchy altogether. Failure history therefore never reaches the asset, and the AM-005 bad-actor trigger cannot fire on a real machine.')]));
children.push(num([R('PM templates carry wrong defaults. ', { bold: true }), R('67 percent of the 3,005 schedules generate Highest-priority work, 73 percent carry no estimated hours, and due dates are set independently of priority. These three template settings are the direct source of the priority inflation, the missing estimates and the arbitrary lateness figures seen on every work order. They are one-time edits.')]));
children.push(num([R('The inventory module is not transacting. ', { bold: true }), R('One parts issue has been recorded in seven years against 18,384 work orders; parts fields are blank on every work order; 89 percent of stock lines have no unit cost and 55 percent no reorder point. On-hand quantities cannot be trusted and AM-002 describes a process that does not yet exist.')]));
children.push(num([R('The configuration predates the framework. ', { bold: true }), R('Maintenance types include Safety, Warranty and Damage (which AM-001 makes flags) and omit five of its thirteen categories; asset status is Active/Inactive only; 53 permission groups exist against eight buckets and two thirds of active users sit in more than one; 151 failure codes are configured but have never been used and contain duplicates and a merged miner-repair list; and reporting is capped at 2,000 rows per extract, so the KPIs the Strategy requires "without manual re-work" cannot be produced.')]));
children.push(H2('What is being asked of leadership'));
children.push(P('The remediation plan in Section 5 is sequenced so that the cheapest, highest-leverage changes come first. Roughly a third of the findings are configuration changes that can be made within thirty days without any data entry — pausing 84 schedules on assets already reported as not in service, resetting template priorities and estimates, making completion notes and hours mandatory, retiring test records and the anonymous Guest account. The second phase needs decisions and modest resourcing: a criticality-scoring programme for the estate starting with Group A candidates, a taxonomy re-map, a permissions redesign and a return to receiving and issuing parts through the CMMS. The third phase is the field work — a serial, make and model walkdown for Group A and B equipment, hierarchy repair, and warranty backfill for the 13,848 equipment records created in the last two years, of which three currently have a warranty recorded.'));
children.push(P('Six decisions are needed from leadership and six clarifications from the framework author; both lists are in Section 6. The exposure indicators in Section 7 give the order of magnitude of what the gaps cost today: warranty work being done at IREN cost with no coverage record, labour hours unrecorded on 8,290 completed work orders, 78 Group A UPS units at Childress-HPC covered only by paused schedules, and a Sweetwater HV transformer with thirteen failure work orders in twelve months and no recorded cause.'));
children.push(brk());

// ------------------------------------------------------------------ 1 Introduction
children.push(H1('1. Introduction and Scope'));
children.push(H2('1.1 Purpose'));
children.push(P('The purpose of this assessment is to establish, with evidence, how far the CMMS and the data in it currently are from the requirements of the Asset Management Governance Framework, to locate where the gaps cluster (by site, by data area, by creation batch), to separate configuration and process causes from data-entry causes, and to propose a sequenced remediation plan with owners drawn from the framework\'s own accountabilities.'));
children.push(H2('1.2 Governing documents'));
children.push(P('The requirements measured against are those of the three artifacts issued to date. All are drafts pending approval and all post-date the data assessed; where this report refers to a clause it uses the numbering of the revision cited.'));
children.push(table(['Document', 'Revision / date', 'Requirements used in this assessment'], [
  ['Data Center Facilities Asset Management Strategy', 'Rev 3.0, third draft, 21 Aug 2026', 'Section 7 criticality groups A–D and Table 7-4 governance intensity; Section 8 hierarchy (Level 3 equipment), naming (G-8.4), master data (G-8.5), materiality $500 (G-8.7, G-8.8); Section 9 tactics (G-9.1, G-9.7); G-13.2 mandatory asset attributes; G-13.3 KPIs without re-work; AM-005 defect trigger; AM-007 data standards; AM-009 obsolescence status'],
  ['AM-001 Maintenance Workflow & Change Management', 'Rev 3.0, first draft, 9 Sep 2026', '5.4 work order categories and G-5.10/G-5.11; 5.5 priority matrix; 5.6 reporting and closure standard (G-5.14 to G-5.18); Section 6 permission buckets (G-6.3 to G-6.5); 7.3 frequency codes and G-7.8; G-8.9 vendor reports and warranty; Section 10 KPIs'],
  ['AM-002 Facilities Materials Management', 'Rev 1.0, first draft, 23 Sep 2026', 'Section 5 inventory governance; Section 6 spare-parts groups; Section 7 Min/Max (G-7.3); 8.4 issuance (G-8.12 to G-8.14); Section 12 CMMS requirements (G-12.1, G-12.4, G-12.5); Section 13 KPIs'],
], [2600, 2000, 4760]));
children.push(spacer());
children.push(H2('1.3 Data assessed'));
children.push(table(['Export', 'Snapshot', 'Records', 'Used for'], [
  ['Closed work orders', '25 Sep 2026', '18,384 (Nov 2023 – Sep 2026)', 'Reporting standard, categories, priority, failure coding, PM linkage, backfilling, lateness, bad-actor candidates'],
  ['Open work orders', '28 Sep 2026', '435', 'Backlog age, overdue, G-5.12 deferral flag, readiness'],
  ['Scheduled maintenance list', '28 Sep 2026', '3,005 schedules; 10,338 asset references', 'Template priority, estimates, triggers, frequency codes, overlap, coverage'],
  ['Asset register (all sites)', '25 Sep 2026', '27,193 (21,901 equipment; 5,292 locations)', 'G-13.2 attributes, criticality, hierarchy, naming, duplicates, lifecycle status, creation batches'],
  ['Asset warranty report', '28 Sep 2026', '3', 'G-13.2 warranty'],
  ['Stock list and parts usage', '25 Sep 2026', '314 stock lines; 1 issue transaction (2019–2026)', 'AM-002 master data, Min/Max, linkage, consumption'],
  ['User groups', '28 Sep 2026', '53 groups', 'AM-001 Section 6 permission model'],
  ['Failure code lists', '28 Sep 2026', '35 problem, 63 cause, 53 action codes', 'G-5.16 failure coding; AM-005 readiness'],
], [2300, 1300, 2300, 3460]));
children.push(spacer());
children.push(P('Not assessed, because the data cannot be exported through the CMMS front end: work order attachments and photographs (G-5.17), asset linked documents (G-13.2), and API-level permissions. Physical accuracy of the register (assets that exist but are not registered, or registered but no longer exist) cannot be assessed from any export and requires a field sample; see Section 8.'));
children.push(H2('1.4 Method'));
children.push(P('Each data area was scored record by record against a set of rules derived from the governing documents. A rule states which records it applies to (for example, Completion Notes are required only on work orders closed as Completed; serial numbers only on equipment, not on locations), what constitutes a failure, and which clause it traces to. Every record in the supporting workbooks carries a flag per rule, and every count in this report is a formula over those flags, so any rule can be adjusted and the results recomputed. Requirements that the exports cannot show are listed as "not measurable" on each workbook\'s traceability sheet rather than being silently omitted.'));
children.push(P('Where the framework depends on criticality group — response times, defect thresholds, spares classification — the assessment shows both thresholds or notes that the rule cannot be applied, because 94 percent of assets have no group. Rates are reported against the applicable population, not the whole file.'));
children.push(H2('1.5 Limitations'));
children.push(bullet('The governing documents are drafts. Where a requirement seemed internally inconsistent (for example the Group A threshold of "above 25" in Strategy Table 7-2 against "above 21" in AM-002 Table 6-2) the discrepancy is recorded in Section 6 rather than resolved.'));
children.push(bullet('The CMMS caps report extracts at 2,000 rows. Assembling the asset register took approximately five hours of repeated pulls. This is itself recorded as a capability gap against Strategy G-13.3 and AM-006 G-AM-006.4.'));
children.push(bullet('PM compliance is a proxy: work orders carry no PM frequency code or due date, so the Suggested Completion date stands in for the 10 percent window of AM-001 G-7.8.'));
children.push(bullet('Backlog in weeks uses booked technician hours (12-week run rate) as its denominator because rostered capacity is not held in the CMMS; 53 percent of open work orders carry no estimate.'));
children.push(bullet('The equipment/location split in the register is by asset category and can be refined; a road without a serial number is not counted as a gap.'));
children.push(brk());

// ------------------------------------------------------------------ 2 KPI baseline
children.push(H1('2. Baseline against the Framework KPIs'));
children.push(P('AM-001 Section 10 defines the execution indicators; AM-002 Section 13 the materials indicators. The table shows the current value where the exports allow one, the target, and the caveat that applies.'));
children.push(table(['Indicator', 'Current', 'Target', 'How measured / caveat'], [
  ['Work order quality (closed WOs meeting reporting standard)', '40%', '≥ 95%', 'Composite of G-5.14/15/16 fields present and meaningful, on 17,033 completed WOs; photos and parts not assessable. 12% (2024) → 37% (2025) → 67% (2026 YTD)'],
  ['PM compliance', '58%', '≥ 95%', 'Proxy: schedule-generated PM/Inspection WOs completed by Suggested Completion date. 73% in 2026 YTD'],
  ['Reactive-to-total work order ratio', '0.7% – 9.4%', '≤ 20%', 'Lower bound Breakdown + Damage; upper bound includes all Corrective (planned/unplanned not distinguishable)'],
  ['"Other" category share', '5.8%', '≤ 2%', 'G-5.11; 9.8% at Childress'],
  ['Highest priority outside Breakdown/Damage', '99%', '0%', 'Share of 12,696 Highest-priority WOs that are not breakdown work; 94% originate from PM templates'],
  ['Backlog (weeks of ready work at booked pace)', '≈ 2.1', '4 – 6 weeks', '2,009 estimated hours of Open/Assigned work ÷ 951 booked hours/week (12-week run rate). Below the band; understated because 53% of open WOs carry no estimate and hours are unrecorded at three sites. Mackenzie ≈ 8.8 weeks; Childress ≈ 0.3'],
  ['Schedule compliance', 'n/a', '≥ 90%', 'Frozen-schedule week not held in CMMS'],
  ['Triage response', 'n/a', '≥ 95%', 'No triage class or timestamp in CMMS'],
  ['Unrecorded consumption (AM-002)', 'n/a', '0', '1 issue transaction in 7 years vs 18,384 WOs — consumption is essentially unrecorded'],
  ['Group A stockouts; spares availability; inventory accuracy (AM-002)', 'n/a', '0 / ≥ 98% / ≥ 97%', 'No spare-parts group, reservations or count records in the CMMS'],
], [2900, 1150, 1100, 4210], { numCols: [] }));
children.push(spacer());
children.push(P('Of the ten indicators, three can be measured directly, three by proxy or bound, and four not at all from the CMMS as configured. Strategy G-13.3 requires the system to capture the data for every KPI without manual re-work; the gap between what the framework measures and what the system can report is therefore itself a finding, and it is addressed in Phase 2 of the roadmap.'));
children.push(img(path.join(CH, 'fig_wo_trend.png'), 520, 252));
children.push(caption('Figure 1. Work order closure quality by year. Field discipline has improved sharply; the remaining gap is structural.'));
// Findings
children.push(H1('3. Findings by Area'));
children.push(H2('3.1 Work orders — closed history'));
children.push(P('The closed work order history is the richest dataset and the one that has improved most. Of 18,384 closed work orders, 80 percent fail at least one completeness rule derived from AM-001 G-5.14 to G-5.16, averaging two gaps each, but the failures are concentrated in three fields and two sites rather than spread evenly.'));
children.push(table(['Rule (AM-001 clause)', 'Applicable', 'Fails', 'Rate'], [
  ['Completion notes missing on completed WO (G-5.15)', '17,033', '4,923', '28.9%'],
  ['Actual hours missing on completed WO (G-5.14)', '17,033', '8,290', '48.7%'],
  ['Estimated hours missing (5.3 step 7)', '17,033', '12,548', '73.7%'],
  ['Cause missing on failure WO (G-5.16)', '1,885', '1,778', '94.3%'],
  ['Solution missing on failure WO (G-5.16)', '1,885', '1,656', '87.9%'],
  ['PM/Inspection WO not generated from a schedule (5.4)', '13,948', '1,534', '11.0%'],
  ['Account code missing', '18,384', '2,374', '12.9%'],
  ['Failure WO written against a location, not equipment (Strategy 8)', '1,885', '785', '41.6%'],
  ['Completed more than a day before created (backfilled)', '18,384', '1,488', '8.1%'],
  ['Completed after target date', '16,273', '5,367', '33.0%'],
  ['Highest priority on non-breakdown work (5.5)', '12,696', '12,612', '99.3%'],
  ['Fails composite reporting standard (G-5.14–5.16)', '17,033', '10,193', '59.8%'],
], [5200, 1400, 1400, 1360], { numCols: [1, 2, 3] }));
children.push(spacer());
children.push(P('Three observations shape the remediation. First, labour and failure coding are the largest gaps and both are enforcement gaps: hours and notes are not mandatory at close, and the failure-code lists, although configured, have never been used — the Problem field is empty on all 18,384 records. Second, the notes and hours gaps are site-driven: Canal Flats has no completion notes on 98 percent of completed work orders and Mackenzie on 72 percent, against 10 percent at Childress. Third, several findings are process signals rather than data gaps: about 260 Childress inspections were closed as Completed with the note "overlapping" (that is, not performed), about 240 PMs were closed Incomplete or Rejected because the asset was not in service, and 8 percent of work orders were created after the work was done, which makes response-time metrics unreliable for those records.'));
children.push(P('Hours tell a sharper story than counts. In the twelve months to 29 September 2026 the CMMS recorded 43,007 technician hours on 8,263 closed work orders (its own Hours by Maintenance Type report, which this export reconciles to within one to two percent by completion date). Of those hours, 21 percent — 8,897 — were booked to the "Other" category on just 428 work orders, more than Inspection (17 percent) and more than three times Corrective (6 percent); Preventive took 33 percent. The G-5.11 cap of two percent is therefore being exceeded by a factor of ten in effort terms, and a fifth of the workforce\'s recorded time is invisible to category-based cost and KPI analysis. Parts, misc-cost and labour-cost fields are blank on every one of the 18,384 records, which ties the work order history directly to the inventory finding in Section 3.4. Twelve test work orders and a differently-formatted "WO-prefixed" numbering series sit in the production history.'));
children.push(H3('Bad-actor candidates'));
children.push(P('Applying the AM-005 defect trigger (two or more failure work orders in twelve months for Group A, three for other groups) across the history identifies 133 asset records at the three-or-more threshold and 287 at two-or-more. Only 63 of the 133 are equipment records — the rest are buckets such as "Leased Kubotas" (54 corrective work orders in a year), roads and data-centre structures — and only one of the 133 has a criticality group. Among genuine equipment, the Sweetwater HV transformer T12 (13 failure work orders in twelve months) and autotransformer T1, the Block 1 and Block 6 gensets at Childress, MV transformer T111 and a group of exhaust fans and lighting circuits stand out and warrant review regardless of the data programme. With Cause recorded on 6 percent of failure work orders, root-cause analysis on any of them would start from almost nothing.'));
children.push(H2('3.2 Work orders — open backlog'));
children.push(P('At 28 September 2026 there were 435 open work orders. 264 are Open, 85 are in Draft and have never been released (median age 122 days, oldest 915 days, including a test work order), 13 are Work Completed but not closed, 14 are On Hold and 10 are Waiting for Parts. 187 are older than thirty days and 36 older than a year. Of the 95 open corrective and breakdown work orders, 45 exceed the 30-day deferral flag of AM-001 G-5.12. Of the 310 with a due date, 123 are overdue; 125 have no due date at all. 204 are PM or inspection work orders, 92 of them overdue. Half sit on location records rather than equipment, and 62 are typed Warranty against a register that holds three warranty records.'));
children.push(P('Backlog in weeks of ready work — the AM-001 indicator — has been measured at booked pace rather than against rostered capacity, which the CMMS does not hold. Over the twelve weeks to 20 September 2026 technicians booked about 951 hours a week across the estate (Childress 399, Prince George 203, Childress-HPC 197, Mackenzie 145, Sweetwater 7, Canal Flats none); the CMMS\'s own "Technician Hours Logged per Site per Work Week" report shows the same pattern. Against roughly 2,009 estimated hours of Open and Assigned work, the ready backlog is about 2.1 weeks, below the 4–6 week band — but the figure is understated on both sides. 53 percent of open work orders carry no estimate (64 percent at Childress, 76 percent at Prince George), and booked hours are missing on 58 percent of Mackenzie, 69 percent of Sweetwater and 100 percent of Canal Flats closures in the period, so two sites cannot be measured at all. Mackenzie shows about 8.8 weeks because a handful of large estimated work orders sit against a low booked rate. Priority is blank on 11 percent of open work orders but almost entirely in Draft and Requested status, which is the pre-triage stage where AM-001 step 6 assigns it; this is the workflow operating as intended.'));
children.push(H2('3.3 PM program — scheduled maintenance'));
children.push(P('The schedule master (3,005 schedules; 1,665 active, 1,340 paused; 10,338 asset references) is the root of several work order findings and a finding in its own right.'));
children.push(table(['Finding (clause)', 'Value'], [
  ['Templates set to Highest priority (AM-001 5.5)', '2,006 of 3,005 (67%) — including 1,799 of 2,423 preventive and 188 of 479 inspection templates, and all 10 meter-reading templates'],
  ['Templates with no estimated hours (5.3 step 7)', '2,206 (73%)'],
  ['Templates whose due-date setting exceeds the priority window (5.5)', '740 of 2,965 (25%) — Highest templates allow 1 to 36 days'],
  ['Active templates with no recurrence visible (G-9.7)', '300 — 299 are meter-based greasing routes at Mackenzie and Prince George; verify in system'],
  ['Active templates on assets reported as not in service', '84 templates covering 60 assets'],
  ['Overlap candidates (same asset, type and frequency in more than one active template)', '466 templates / 372 combinations — mostly un-bundled tasks on the same asset (motor and shaft greasing), not straight duplicates'],
  ['Intervals outside the AM-001 7.3 code set', '183 at 26 weeks (7.3 defines semi-annual as 24 weeks); 44 calendar-date schedules'],
  ['Safety / Other / Improvement typed schedules (G-5.10)', '93'],
  ['Test schedules live; legacy "SMnnn" codes', '19; 30'],
], [4300, 5060]));
children.push(spacer());
children.push(P('Joining the schedules to the register and crediting coverage through parent locations (many routes are built on section records), 13.7 percent of equipment is directly attached to an active schedule, 45.9 percent is covered through a parent, 2.3 percent only by paused schedules and 38.1 percent (8,352 records) by nothing. HPC servers and racks account for a large share of the uncovered population and are probably vendor-maintained. Among the assets that do have a criticality group, however, 78 Group A UPS units at Childress-HPC are covered only by paused schedules, which Strategy G-9.1 does not permit for Group A. Mackenzie has 1,426 equipment records with no coverage; neither Prince George HPC site has any active schedule.'));
children.push(img(path.join(CH, 'fig_pm_coverage.png'), 520, 352));
children.push(caption('Figure 2. PM schedule coverage of equipment by category. "Via parent" credits routes built on section or building records.'));
children.push(H2('3.4 Asset register'));
children.push(P('The register is large and well populated at the location level but thin at the equipment level, and its lifecycle attributes largely do not exist.'));
children.push(table(['Finding (clause)', 'Value'], [
  ['Assets with a criticality group (Strategy G-7.1)', '1,525 of 27,193 (5.6%) — all at Childress-HPC; 1.2% of work orders touch an asset with a group'],
  ['Equipment with no serial number (G-13.2)', '16,332 of 21,901 (74.6%)'],
  ['Serials that are placeholders, corrupted or duplicated', '1,259 of 5,608 (22%): 423 "N/A"; 49 converted to scientific notation by a spreadsheet round-trip; 787 shared with another asset of the same make'],
  ['Equipment with no make / no model', '16.5% / 22.3%; 155 and 218 "N/A"; manufacturers spelled multiple ways on 1,800 assets'],
  ['Manufacturer part number', 'Blank on 100%'],
  ['Records outside the hierarchy (no parent) (G-8.1)', '2,616'],
  ['Duplicate asset codes (G-8.4)', '113 codes on 223 records'],
  ['Codes not following the SITE- prefix convention', '2,010, mostly Prince George HPC (PG11B…, SRV)'],
  ['Lifecycle status (G-8.5; AM-009)', 'Active/Inactive only; 0 Inactive; 60 "Active" assets whose PMs were closed as not in service; "Assets Removed from Service" location exists but is unused; no obsolescence status'],
  ['Installation date; linked documents', 'No field; not exportable'],
  ['Warranty (G-13.2; AM-001 G-8.9)', '3 records for 27,193 assets; 13,848 equipment records created in the last 24 months; 163 closed and 62 open WOs typed Warranty'],
], [4300, 5060]));
children.push(spacer());
children.push(P('The register was built by bulk import: 86 percent of equipment was created on ten calendar days, and the serial, make and model gaps track those import batches almost exactly (Figure 3). The January 2025 load is 92 percent serial-blank, the August 2026 load 100 percent, while hand-entered records are 49 percent blank — but carry nearly half the duplicated serials, the cloning pattern. The primary fix is therefore the import template and the onboarding standard (Strategy G-8.2), not field discipline.'));
children.push(img(path.join(CH, 'fig_serial_batch.png'), 520, 302));
children.push(caption('Figure 3. Serial-number completeness by asset creation batch.'));
children.push(img(path.join(CH, 'fig_crit_site.png'), 520, 252));
children.push(caption('Figure 4. Criticality assignment by site. Childress-HPC has scored 41 percent of its equipment; no other site has started.'));
children.push(H2('3.5 Materials and inventory'));
children.push(P('The stock list is a location-keyed parts catalogue rather than a controlled inventory, and the parts-usage history shows the inventory module is not transacting. Across 314 stock lines (290 part codes in 20 storage areas), 89 percent have no unit cost, so the reported value of $181,218 is the value of 35 priced lines and Strategy G-8.7\'s $500 materiality test cannot be applied to the rest; 55 percent have no reorder point and there is no Max field, so AM-002 G-7.3 replenishment has nothing to work from; 36 percent are not linked to any asset and the links that exist point at asset classes rather than Level 3 equipment; bin location, charge department and classification codes are empty on every line, and storage-area names are not site-qualified, so the list does not say which campus holds an item. There is no manufacturer or part-number field in the export; eleven items carry a supplier reference in the description text and 24 percent of descriptions are two words or fewer ("Feeder" appears eighteen times under different codes).'));
children.push(P('One parts issue was recorded between 2019 and 2026 — a single MV disconnect switch in May 2025 — while 18,384 work orders were closed and the Parts Used field remained blank on every one. AM-002 G-8.12 and G-8.14 (every issue recorded; no work order closed with unrecorded consumption) therefore describe a process that does not yet exist rather than one being done badly, and on-hand quantities cannot be trusted until stock is re-baselined. Nine of the G-12.1 mandatory fields — manufacturer, part number, approved equivalents, supplier and lead time, spare-parts group, Max, serials, hazmat/SDS, shelf life — do not appear in the export and need confirmation as to whether they exist in the system. The HPC storage area holds drives, memory, processors and over a thousand transceivers, which Strategy G-8.8 treats as controlled items regardless of value; five of those thirteen lines are unpriced and none are serial-tracked.'));
children.push(H2('3.6 CMMS configuration and governance'));
children.push(P('Several findings are properties of how the CMMS is configured rather than of the data in it. They are listed here because they are owned by the CMMS Administrator under AROE approval (AM-001 G-6.2) and because most can be changed quickly. The user population (193 accounts: 176 people, 10 system or API accounts, a Guest login at seven sites and five test or placeholder accounts) was analysed in aggregate; no individual is identified in this report or the workbooks.'));
children.push(table(['Area', 'Current configuration', 'Framework requirement', 'Gap type'], [
  ['Work order categories', '11 Maintenance Types incl. Safety, Warranty, Damage, Other, Meter reading', 'AM-001 5.4: 13 categories; Safety and Warranty are flags, Damage a cause code (G-5.10); PdM, Emergency, Capital, Modification, Asset Monitoring absent', 'Configuration'],
  ['Priority', 'Names already match the 5.5 matrix', 'Highest reserved for Breakdown/Emergency/T4', 'Template defaults (Section 3.3)'],
  ['Failure coding', '35 problem, 63 cause, 53 action codes configured; optional; never used; 61 codes with duplicates, wrong-list or level-mixing issues; a miner-repair list merged into the facilities list; four catch-alls', 'G-5.16 failure description and corrective action; AM-005 trending', 'Enforcement, after list clean-up'],
  ['Asset status', 'Active / Inactive; "Assets Removed from Service" location unused', 'G-8.5 lifecycle status; AM-009 four obsolescence states', 'Configuration'],
  ['Asset categories', '127 categories incl. typo ("Dry Coole Fan"), singular/plural pairs, generic "Equipment" and "HPC"', 'Strategy G-8.2: new categories require AROE approval', 'Master data control'],
  ['Permissions', '53 groups: 24 capability add-ons ("Group – Can …"), 12 crew/dispatch groups, 7 placeholders or migration artefacts, Guests, 1 duplicate. 142 active users: 65% hold more than one group (average 2.3, maximum 12); 42 have never logged in; Administrators holds 4 people plus 10 system/API accounts; no active member of an AROE-level group; 6 planning users hold master-data import rights; one test account is active with six groups; Guest is active at 7 sites', 'AM-001 G-6.3: eight buckets; G-6.4 one bucket per user, quarterly review, no shared logins; G-6.5 separation of Administrator and Owner, no execution staff in buckets 6–8', 'Configuration / redesign'],
  ['Reporting', 'Extracts capped at 2,000 rows; asset list took ~5 hours to assemble', 'Strategy G-13.3 and AM-006 G-AM-006.4: KPI data without manual re-work', 'System capability'],
  ['Test records in production', 'Dummy part; 12 test WOs closed; 1 test WO open; 19 test schedules; "New Group" placeholders', 'AM-002 G-12.6; Strategy G-8.5 master data control', 'Housekeeping'],
], [1600, 2900, 3160, 1700]));
// Root causes
children.push(H1('4. Cross-cutting Root Causes'));
children.push(P('The gap analysis produced more than sixty individual rules across five data areas. Read together they resolve into five causes, each of which explains findings in several areas at once. Fixing them in this order gives the largest return for the least effort.'));
children.push(H3('4.1 Criticality has not been assigned'));
children.push(P('The Strategy makes criticality the primary input to governance intensity (Table 7-4): tactic selection, condition monitoring, change control, spares group, defect trigger, obsolescence review and audit frequency all key off Group A to D. With 94 percent of assets unscored and one site having done the exercise, the framework is effectively inoperable outside Childress-HPC. It also means the assessment could not weight any gap by consequence — a missing serial on a Group A transformer and on a Group D fan count the same today. Owner: AROE (G-7.1), with Maintenance Planning; effort concentrated in Group A candidate categories first (transformers, switchgear, gensets, UPS, cooling, HPC).'));
children.push(H3('4.2 Maintenance is recorded above the equipment level'));
children.push(P('Strategy Section 8 sets Level 3 Equipment as the level at which maintenance is managed. In practice 42 percent of failure work orders, 49 percent of open work orders, 24 percent of active schedule references and most of the inventory asset links point at buildings, sections or areas. History accumulates on the bucket, the bad-actor trigger fires on "Roads", and spares cannot follow an asset. The hierarchy also has 2,616 records with no parent and 113 duplicated codes. Owner: AROE (hierarchy, G-8.1), CMMS Administrator (validation at WO creation), Maintenance Planning (route rebuild).'));
children.push(H3('4.3 PM templates carry wrong defaults'));
children.push(P('Priority, estimated hours and days-to-complete are set once on the template and inherited by every work order it generates. With 67 percent of templates at Highest, 73 percent without hours and due dates decoupled from priority, the work order history shows exactly those defects at scale. No amount of field discipline corrects them; three template edits do. Owner: Maintenance Planning Manager (G-7.1), executed by the CMMS Administrator.'));
children.push(H3('4.4 The inventory module is not transacting'));
children.push(P('Receipts, issues and returns are not being processed through the CMMS, so unit costs, on-hand balances, reorder triggers and consumption history are all absent, and the work order side has no parts data to report. This is an operating-model gap as much as a system one: stores discipline (AM-002 G-8.11, no release without a work order) has to exist before the master data is worth cleaning. Owner: Inventory Team with Maintenance Planning; Procurement for PO-based receiving.'));
children.push(H3('4.5 The configuration predates the framework'));
children.push(P('Categories, statuses, code lists, permission groups and reporting limits were set up before the framework was written and have accreted since — including a Fiix migration and a miner-repair code set. Aligning them is mostly configuration work under G-6.2 (AROE approval, logged), and several items block other remediation: failure codes must be cleaned before they are enforced, lifecycle statuses must exist before PM generation can be tied to them, and permission buckets must exist before access can be reviewed. Owner: CMMS Administrator under AROE approval.'));
children.push(brk());

// ------------------------------------------------------------------ 5 Roadmap
children.push(H1('5. Remediation Roadmap'));
children.push(P('Actions are grouped into three phases by dependency and effort, with owners taken from the accountabilities the framework itself assigns. Phase 0 requires no data entry and can proceed immediately under existing authority; Phase 1 needs the decisions in Section 6; Phase 2 is field and backfill work that depends on Phases 0 and 1.'));
children.push(H2('5.1 Phase 0 — configuration and quick wins (0 to 30 days)'));
children.push(table(['#', 'Action', 'Clause', 'Owner', 'Effect'], [
  ['0.1', 'Pause the 84 active schedules covering the 60 assets already reported as not in service; set those assets Inactive with reason in Notes', 'G-8.5; V10/A19/S14', 'Maintenance Planning Mgr; CMMS Admin', 'Stops ~240 wasted PM closures per year immediately'],
  ['0.2', 'Reset PM template priorities per the 5.5 matrix (routine PM/inspection to Medium or Low; Highest only for statutory or Group A tasks); derive Suggested Completion from priority', 'AM-001 5.5; S01, S06', 'Maintenance Planning Mgr', 'Removes 99% of false Highest work orders; makes lateness measurable'],
  ['0.3', 'Add estimated hours to every active template', 'AM-001 5.3 step 7; S03', 'Maintenance Planning', 'Fixes 74% blank estimates on all future WOs; enables backlog KPI'],
  ['0.4', 'Make Completion Notes (minimum length, disallowed placeholder values) and Actual Hours (> 0) mandatory to close a Completed WO', 'G-5.15, G-5.14, G-5.18', 'CMMS Admin (AROE approval)', 'Closes the two largest completeness gaps at source'],
  ['0.5', 'Retire test and placeholder records: dummy part, 13 test WOs, 19 test schedules, four "New Group" placeholders, "Original Fiix" and migrated groups', 'AM-002 G-12.6; Strategy G-8.5', 'CMMS Admin', 'Master-data hygiene'],
  ['0.6', 'Disable Guest requester; require a named requester on work requests', 'AM-001 G-6.4; 5.3 step 4', 'CMMS Admin', 'Closes anonymous access; enables closing the loop with requesters'],
  ['0.7', 'Reclassify recurring "Other" work at Childress into defined categories; institute monthly category review', 'G-5.11', 'Maintenance Planning; AROE', 'Brings Other from 5.8% toward the 2% cap'],
  ['0.8', 'Load labour rates so labour cost calculates from Actual Hours', 'G-5.14', 'CMMS Admin; Finance', 'Enables cost per asset'],
  ['0.9', 'Obtain rostered productive hours per site and publish backlog-in-weeks against both booked pace and capacity; the gap between the two is the hours-recording gap', 'AM-001 §10', 'Maintenance Planning Mgr', 'Backlog KPI with a capacity denominator'],
], [500, 3900, 1500, 1700, 1760]));
children.push(spacer());
children.push(H2('5.2 Phase 1 — decisions, taxonomy and controls (30 to 90 days)'));
children.push(table(['#', 'Action', 'Clause', 'Owner', 'Effect'], [
  ['1.1', 'Criticality scoring programme: apply Table 7-1 to Level 3 equipment at every site, Group A candidate categories first; make the field mandatory at onboarding', 'Strategy G-7.1, G-7.2, G-8.2', 'AROE (owner); Maintenance Planning; Site Managers', 'Unlocks every criticality-dependent control; lets all other gaps be weighted by consequence'],
  ['1.2', 'Re-map Maintenance Types to the 13 AM-001 categories; add Safety and Warranty flags and a Damage cause code; migrate history', 'AM-001 5.4, G-5.10', 'CMMS Admin (AROE approval)', 'KPI-accurate categories; warranty charge-back possible'],
  ['1.3', 'Clean the failure code lists (de-duplicate, separate miner codes, split mechanism from cause, one Unknown per list with required text, fix numbering); link codes to asset categories if supported; then make Problem/Cause/Action mandatory to close failure WOs', 'G-5.16; AM-005', 'CMMS Admin; Maintenance Planning; discipline leads', 'Enables root-cause trending and the AM-005 bad-actor process'],
  ['1.4', 'Redesign permissions to the eight buckets: fold capability add-ons into buckets, keep crew groups for dispatch only, retire artefacts; assign every user one bucket; confirm Administrator ≠ Owner', 'AM-001 G-6.3 to G-6.5', 'CMMS Admin; AROE; Owner', 'Auditable access model; quarterly review becomes possible; 42 never-used and 34 inactive accounts cleared'],
  ['1.5', 'Introduce lifecycle statuses (Planned, Commissioning, Active, Not in service, Retired) and AM-009 obsolescence status; suppress PM generation except for Active', 'Strategy G-8.5; AM-009', 'CMMS Admin (AROE approval)', 'Ends PMs on assets that are not in service'],
  ['1.6', 'Inventory operating model: receive against PO and issue against WO in the CMMS only (G-8.11); WO parts tab as the sole consumption route; wall-to-wall count to re-baseline on hand; set Min/Max for stocked items starting with Group A/B', 'AM-002 5, 7, 8.4, G-12.2', 'Inventory Team; Maintenance Planning; Procurement', 'Trustworthy stock; parts on work orders; unit costs populate'],
  ['1.7', 'Asset import template: serial, make, model, criticality, parent and site mandatory; serials imported as text', 'Strategy G-8.2, G-13.2', 'AROE; CMMS Admin', 'Prevents the next site repeating the batch pattern'],
  ['1.8', 'Controlled lists for manufacturer and asset category; merge variants and typos', 'Strategy G-8.5', 'AROE', 'Consistent master data'],
  ['1.9', 'Establish an extraction route that removes the 2,000-row limit (vendor report configuration or API) and configure the Section 10 / Section 13 KPIs as system reports', 'Strategy G-13.3; AM-006 G-AM-006.4', 'CMMS Admin; vendor', 'KPIs without manual re-work; repeatable re-baseline'],
], [500, 3900, 1500, 1700, 1760]));
children.push(spacer());
children.push(H2('5.3 Phase 2 — field verification and backfill (90 to 180 days)'));
children.push(table(['#', 'Action', 'Clause', 'Owner', 'Effect'], [
  ['2.1', 'Serial, make and model walkdown for Group A and B equipment using barcode/QR capture; resolve the 787 duplicated and 49 corrupted serials', 'Strategy G-13.2', 'Site Managers; AROE', 'Warranty, recall and unit traceability for critical assets'],
  ['2.2', 'Hierarchy repair: assign parents to the 2,616 orphaned records, merge the 113 duplicate codes, register Level 3 equipment behind the location-charged failure work (vehicles, lighting circuits, gates, tanks); require an equipment-level asset on Corrective/Breakdown WOs', 'Strategy 8, G-8.1', 'AROE; CMMS Admin', 'Failure history reaches the asset; bad-actor trigger works'],
  ['2.3', 'Warranty backfill from purchase orders for equipment created in the last 24 months, Group A/B first; capture warranty at receiving and commissioning going forward', 'Strategy G-13.2; AM-001 G-8.9', 'Procurement; AROE', 'Warranty recovery on in-scope failures'],
  ['2.4', 'PM coverage: every Group A/B/C asset on at least one active schedule; reactivate or retire the 1,340 paused schedules by site; bundle same-frequency tasks into one template per asset; add PM frequency code and due date to generated WOs', 'Strategy G-9.1, 9.4; AM-001 G-7.8', 'Maintenance Planning Mgr', 'True PM compliance measurement against the 10% window'],
  ['2.5', 'Backfill unit cost, spare-parts group and lead time for Group A/B spares; classify HPC data-bearing components and apply Section 6 issue controls', 'AM-002 6, G-12.1; Strategy G-8.8', 'Inventory Team; AROE', 'Materiality and stockout KPIs become measurable'],
  ['2.6', 'Field verification sample: about fifty assets per site checked nameplate-to-record', 'Strategy G-8.6; AM-010', 'AROE; Site Managers', 'Register accuracy figure to sit alongside completeness'],
  ['2.7', 'Re-run this baseline from the same exports; report against Section 10 / Section 13 targets', 'AM-006; AM-010', 'CMMS Admin; AROE', 'Progress evidence for the quarterly asset management review'],
], [500, 3900, 1500, 1700, 1760]));
children.push(spacer());
children.push(P('Ongoing controls once the phases are complete: monthly category and priority review (G-5.11), quarterly work order audit against the reporting standard (G-14.1), quarterly bad-actor analysis (AM-005 G-AM-005.5), quarterly access review (G-6.4), and an annual re-baseline using the workbooks delivered with this report.'));
children.push(brk());

// ------------------------------------------------------------------ 6 Decisions
children.push(H1('6. Decisions Required and Questions for the Framework Author'));
children.push(H2('6.1 Decisions for leadership'));
children.push(num2([R('Approve the criticality scoring programme (1.1) as the prerequisite workstream, with AROE as owner and site input resourced. ', { bold: true }), R('Without it the framework cannot be applied and no gap can be prioritised by consequence.')]));
children.push(num2([R('Authorise the Phase 0 and Phase 1 configuration changes under AM-001 G-6.2 ', { bold: true }), R('(taxonomy, statuses, mandatory close-out fields, failure codes, permissions), logged and approved by AROE.')]));
children.push(num2([R('Adopt the inventory operating model in 1.6 ', { bold: true }), R('— no release without a work order, receipts and issues through the CMMS — and resource the wall-to-wall count.')]));
children.push(num2([R('Resource the Phase 2 field work ', { bold: true }), R('(walkdown, hierarchy repair, warranty backfill), scoped to Group A and B once criticality exists.')]));
children.push(num2([R('Engage the CMMS vendor on report extraction and API access ', { bold: true }), R('so that Strategy G-13.3 can be met and the baseline re-run without five hours of manual pulls.')]));
children.push(num2([R('Confirm the scope of miners in the Facilities CMMS. ', { bold: true }), R('The Strategy scopes Facilities assets and the register describes data centres as "excluding miners", yet miner repair codes, a Miners site and an R&M – Miners account sit in the same system. If miners remain, they need their own category, code set and permission profile.')]));
children.push(H2('6.2 Clarifications for the framework author'));
children.push(bullet('Group A threshold: Strategy Table 7-2 defines Group A as a score above 25 and Group B as 18–24, leaving 25 unassigned, while AM-002 Table 6-2 uses "above 21" for spare-parts Group A on the same base score. The two documents draw the A/B line in different places.'));
children.push(bullet('Semi-annual frequency: AM-001 7.3 defines S as 168 days / 24 weeks; 183 existing schedules run at 26 weeks. Confirm whether 26 weeks is accepted (and amend 7.3) or the schedules migrate.'));
children.push(bullet('Supervisor role: G-5.18 requires supervisor/planner review at close, but the Section 6 permission model has no supervisor bucket. Maintenance – Leaders currently maps to Planner by default.'));
children.push(bullet('Site / Facilities Manager role: 3.3 and 7.5 give site management MOP and scheduling approvals, but no edit bucket exists in Section 6; the default mapping is View-Only.'));
children.push(bullet('Safety routines: 64 schedules and 640 closed work orders are typed Safety. Under G-5.10 these become PM or Inspection with a Safety flag; confirm they then count toward PM compliance.'));
children.push(bullet('Procurement segregation: AM-002 G-5.3 says the approver of a purchase must not receive it, so Purchasing cannot simply map to the Inventory Team bucket; a narrower profile is implied.'));
children.push(bullet('Created date versus installation date: the register has no installation date and G-13.2 requires one. Confirm whether the CMMS created date may stand in for assets created at commissioning, or whether a separate field is to be added.'));
// Exposure
children.push(H1('7. Risk and Exposure Indicators'));
children.push(P('The assessment does not cost the gaps — the CMMS holds no labour rates, unit costs or warranty values to cost them with. The following indicators give the order of magnitude and are the places where a defensible number can be built once the Phase 0 and Phase 1 data starts to flow.'));
children.push(table(['Indicator', 'Evidence', 'Framework reference'], [
  ['Warranty leakage', '163 closed and 62 open work orders typed Warranty; 3 warranty records in the register; 13,848 equipment records created within the typical warranty period', 'Strategy G-13.2; AM-001 G-8.9, 5.4 Warranty flag'],
  ['Unrecorded labour', '8,290 completed work orders with no hours; labour cost zero on all 18,384; no labour rates configured', 'AM-001 G-5.14; §10 backlog'],
  ['Unrecorded consumption', 'One parts issue in seven years; parts blank on every work order; on-hand balances untrustworthy; 10 open WOs waiting for parts that cannot be tracked', 'AM-002 G-8.12, G-8.14, §13'],
  ['Critical assets without active PM', '78 Group A UPS units covered only by paused schedules; 8,352 equipment records with no schedule (criticality unknown for most)', 'Strategy G-9.1'],
  ['Chronic failures without root cause', '133 asset records at the AM-005 trigger; Sweetwater HV transformer T12 with 13 failure WOs in 12 months; Cause recorded on 6% of failure WOs', 'AM-005 G-AM-005.1; AM-001 G-5.16'],
  ['PM effort on out-of-service assets', '~240 PMs closed as not in service; 84 schedules still active on those assets', 'Strategy G-8.5; AM-001 G-7.5'],
  ['Untraceable physical units', '74.6% of equipment without a serial; 22% of existing serials unreliable; corrupted serials cannot be recovered from the system', 'Strategy G-13.2; AM-002 G-8.14'],
  ['Access control', '53 permission groups; 65% of active users in more than one group; 42 active accounts never used; anonymous Guest active at 7 sites; 10 system/API accounts with Administrator rights; no identifiable AROE or Owner membership; hourly rate blank on every account so labour cost cannot calculate', 'AM-001 G-6.3 to G-6.5; G-5.14'],
], [2200, 4700, 2460]));
// Next steps
children.push(H1('8. Next Steps and Assurance'));
children.push(P('The immediate next steps are to circulate this report and the four workbooks to AROE and the Maintenance Planning Manager for review of the rule set and the proposed owners, to take the six decisions in Section 6.1 to the VP Engineering, and to begin Phase 0 under existing authority. All data items requested for the assessment have now been received; the backlog KPI uses booked hours as its denominator until rostered productive hours per site are available, at which point both figures can be shown side by side.'));
children.push(P('The baseline should be re-run at the end of Phase 1 and again at the end of Phase 2 from the same exports, and thereafter annually as part of the AM-010 assurance cycle. Because every count in the workbooks is a formula over the record-level flags, a re-run requires only replacing the source extracts; the rule set, clause references and traceability sheets carry forward, and any rule that AROE chooses to change is changed in one place.'));
children.push(P('Finally, the register\'s completeness has been measured but its accuracy has not. A stratified field sample of roughly fifty assets per site, checked nameplate-to-record (action 2.6), is the only way to establish how many registered assets no longer exist and how many installed assets were never registered, and it is likely to reshape the walkdown priorities in Phase 2.'));
// ------------------------------------------------------------------ Appendices
children.push(H1('Appendix A — Requirement Traceability Summary'));
children.push(P('Each supporting workbook contains a traceability sheet listing the framework requirements relevant to that data area and whether the export can show them. The counts below summarise those sheets.'));
children.push(table(['Area', 'Requirements traced', 'Met', 'Partial', 'Gap', 'Not measurable / configuration gap'], [
  ['Work orders (AM-001 5.4–5.6, §6, §10; AM-002 G-12.5; Strategy 7, 8, 13; AM-005, AM-007)', '36', '2', '6', '17', '11'],
  ['PM schedules (AM-001 5.4, 5.5, 7.3; Strategy 9) — rule-based, no traceability sheet', '15 rules', '—', '—', '14', '1 (meter triggers to verify)'],
  ['Asset register (Strategy G-13.2, 7, 8; AM-009; AM-005)', '23', '0', '6', '6', '11'],
  ['Inventory (AM-002 G-12.1, 7.3, 8.4, 12.5; Strategy 8.3)', '21', '1', '2', '6', '12'],
], [3800, 1300, 800, 900, 800, 1760], { numCols: [1, 2, 3, 4] }));
children.push(spacer());
children.push(P('Counts are taken from the traceability sheets as of 28 September 2026. "Met" includes one "met with caveat"; "Gap" includes enforcement and asset-master gaps; the work order row excludes one "next phase" item.', { italics: true, size: 18, color: GREY }));
children.push(H1('Appendix B — Supporting Workbooks'));
children.push(table(['Workbook', 'Contents'], [
  ['WorkOrder_Gap_Analysis.xlsx', 'Read Me (rule set and standard applied); Summary with §10 KPI block; Gap Rules (35 rules with clause, impact, root cause, recommendation); AM-001 Traceability; Bad Actor Candidates; Open Backlog and backlog data; Permissions Mapping (53 groups → 8 buckets); Failure Codes (151 codes annotated); heat maps by site, type and year; Field Completeness; Placeholder Notes; Next Phase Standard; Data (18,384 records with flags)'],
  ['PM_Schedule_Gap_Analysis.xlsx', 'Summary; Gap Rules (15); Coverage by category, site and criticality group with Group A/B exceptions; heat maps by site and type; Asset References (10,338 rows with overlap flags); Data (3,005 schedules with flags)'],
  ['Asset_Register_Gap_Analysis.xlsx', 'Summary; Gap Rules (20); Strategy Traceability; heat maps by site, category and creation batch; Field Completeness; Warranty; Make Variants; Data (27,193 records with derived hierarchy, coverage and WO-history columns and flags)'],
  ['Inventory_Gap_Analysis.xlsx', 'Summary with AM-002 §13 KPI block; Gap Rules (17); AM-002 Traceability (G-12.1 field by field); Field Completeness; Parts Usage (full export); Data (314 lines with flags)'],
], [2800, 6560]));
children.push(spacer());
children.push(H1('Appendix C — Glossary'));
children.push(table(['Term', 'Meaning in this report'], [
  ['AROE', 'Asset Management and Operational Excellence team; owner of the CMMS, asset registry and data standards under the Strategy'],
  ['Bad-actor candidate', 'An asset record meeting the AM-005 defect trigger: two or more failure work orders in twelve months (Group A) or three or more (Groups B–D)'],
  ['Bulk load / creation batch', 'A calendar day on which 500 or more asset records were created; used to trace data quality to import events'],
  ['Equipment / Location record', 'Register records classified by asset category: physical equipment (Level 3/4) versus sites, buildings, areas, roads, structures and piping (Levels 0–2)'],
  ['Failure work order', 'Work orders typed Corrective, Breakdown, Damage or Warranty'],
  ['Level 3', 'The equipment level of the Strategy Section 8 hierarchy, at which maintenance is managed in the CMMS'],
  ['Not measurable', 'A framework requirement the CMMS export cannot show, either because no field exists or because the data cannot be exported'],
  ['Placeholder', 'A populated value that carries no information (N/A, TBD, Unknown, Done, Completed)'],
  ['Reporting standard', 'AM-001 G-5.14 to G-5.18: the fields and content a closed work order must carry'],
  ['Via parent (coverage)', 'Equipment credited as PM-covered because an ancestor location in the hierarchy is attached to an active schedule'],
], [2400, 6960]));

// ------------------------------------------------------------------ document
const doc = new Document({
  creator: 'Troy Brannon', title: 'CMMS Gap Analysis — Current-State Baseline', description: 'Assessment of CMMS data, configuration and PM program against the IREN Asset Management Governance Framework',
  styles: { default: { document: { run: { font: FONT, size: 21 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 30, bold: true, color: NAVY }, paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 25, bold: true, color: NAVY }, paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 1 } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 22, bold: true, color: GREY }, paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2 } },
    ] },
  numbering: { config: [
    { reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] },
    { reference: 'numbers', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 360 } } } }] },
    { reference: 'numbers2', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 360 } } } }] },
  ] },
  features: { updateFields: true },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, bottom: 1300, left: 1440, right: 1440 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [R('IREN — CMMS Gap Analysis — Rev 0.1 DRAFT for review', { size: 16, color: GREY })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [R('Page ', { size: 16, color: GREY }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: GREY }), R(' of ', { size: 16, color: GREY }), new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 16, color: GREY })] })] }) },
    children,
  }],
});
Packer.toBuffer(doc).then(buf => { fs.writeFileSync(OUT, buf); console.log('written', OUT); });
