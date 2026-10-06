# Rule catalogue (as scored at the current snapshot)

Rule text, impact, root cause and recommendation are on each workbook's Gap Rules sheet; this is the index.

## Work orders (WorkOrder_Gap_Analysis.xlsx)

| Rule | Description | Applicable | Fails | Rate |
|---|---|---|---|---|
| G01 | Priority missing | 18,384 | 247 | 1.3% |
| G02 | Maintenance Type missing | 18,384 | 97 | 0.5% |
| G03 | Site missing | 18,384 | 2 | 0.0% |
| G04 | Assigned Users missing | 18,384 | 903 | 4.9% |
| G05 | Completed By missing | 18,384 | 12 | 0.1% |
| G06 | Account Code missing | 18,384 | 2,374 | 12.9% |
| G07 | Charge Dept missing | 18,384 | 32 | 0.2% |
| G08 | Asset Description missing | 18,384 | 409 | 2.2% |
| G09 | Target (Suggested Completion) missing | 18,384 | 883 | 4.8% |
| G10 | Completion Notes missing | 17,033 | 4,923 | 28.9% |
| G11 | Completion Notes are a placeholder | 17,033 | 363 | 2.1% |
| G12 | Actual Hours missing | 17,033 | 8,290 | 48.7% |
| G13 | Actual Hours zero | 17,033 | 30 | 0.2% |
| G14 | Est Hours missing | 17,033 | 12,548 | 73.7% |
| G15 | Scheduled Maintenance ID missing on Preventive/Inspection | 13,948 | 1,534 | 11.0% |
| G16 | Cause missing on failure WO | 1,885 | 1,778 | 94.3% |
| G17 | Solution missing on failure WO | 1,885 | 1,656 | 87.9% |
| G18 | Requested By missing or "Guest" on failure WO | 1,885 | 1,618 | 85.8% |
| V01 | Completed more than 1 day before Created | 18,384 | 1,488 | 8.1% |
| V02 | Target date earlier than Created | 17,501 | 659 | 3.8% |
| V03 | Actual Hours > 24 on one WO (review) | 9,150 | 742 | 8.1% |
| V04 | Charge Dept does not match Site | 18,384 | 27 | 0.1% |
| V05 | Rejected WO has hours booked | 881 | 87 | 9.9% |
| V06 | Status not a closed status in a closed-WO export | 18,384 | 12 | 0.1% |
| V07 | Scheduled Maintenance ID on non-PM type | 4,339 | 443 | 10.2% |
| V08 | Days Open > 90 | 18,384 | 292 | 1.6% |
| V09 | Completed after target date (late) | 16,273 | 5,367 | 33.0% |
| V10 | Closed Incomplete / Rejected because asset not active | 1,339 | 255 | 19.0% |
| V11 | Highest priority on non-Breakdown work | 12,696 | 12,612 | 99.3% |
| V12 | Safety / Warranty / Damage used as a category | 18,384 | 841 | 4.6% |
| V13 | Category = Other | 18,384 | 1,063 | 5.8% |
| R01 | Fails AM-001 reporting standard (composite) | 17,033 | 10,193 | 59.8% |
| R02 | Scheduled PM not completed by target date (PM compliance proxy) | 12,391 | 5,216 | 42.1% |
| H01 | Failure WO written against a location/area record, not equipment | 1,885 | 785 | 41.6% |

## PM schedules (PM_Schedule_Gap_Analysis.xlsx)

| Rule | Description | Applicable | Fails | Rate |
|---|---|---|---|---|
| S01 | Priority = Highest on a routine template | 3,005 | 2,006 | 66.8% |
| S02 | Priority blank | 3,005 | 35 | 1.2% |
| S03 | Estimated hours blank | 3,005 | 2,206 | 73.4% |
| S04 | Active with no trigger visible | 1,665 | 300 | 18.0% |
| S05 | Interval not an AM-001 7.3 frequency code | 1,540 | 283 | 18.4% |
| S06 | Days-to-complete exceeds the priority window | 2,965 | 740 | 25.0% |
| S07 | Type is not PM / Inspection / Meter reading | 3,005 | 93 | 3.1% |
| S08 | Code outside SITE-TYPE-nnnn convention | 3,005 | 30 | 1.0% |
| S09 | No assets attached | 3,005 | 5 | 0.2% |
| S10 | Assigned Users blank | 3,005 | 230 | 7.7% |
| S11 | Test / dummy schedule | 3,005 | 19 | 0.6% |
| S12 | References location-level records | 3,000 | 371 | 12.4% |
| S13 | Overlap candidate (same asset, type & frequency) | 1,665 | 466 | 28.0% |
| S14 | Active on an asset whose PMs were closed as not active | 1,665 | 84 | 5.0% |
| S15 | Paused | 3,005 | 1,340 | 44.6% |

## Asset register (Asset_Register_Gap_Analysis.xlsx)

| Rule | Description | Applicable | Fails | Rate |
|---|---|---|---|---|
| A01 | Criticality group missing | 27,184 | 25,659 | 94.4% |
| A02 | Make missing on equipment | 21,901 | 3,619 | 16.5% |
| A03 | Make is a placeholder (N/A etc.) | 18,511 | 155 | 0.8% |
| A04 | Manufacturer spelled more than one way | 18,356 | 1,800 | 9.8% |
| A05 | Model missing on equipment | 21,901 | 4,880 | 22.3% |
| A06 | Model is a placeholder | 17,099 | 218 | 1.3% |
| A07 | Serial number missing on equipment | 21,901 | 16,332 | 74.6% |
| A08 | Serial number is a placeholder | 5,608 | 423 | 7.5% |
| A09 | Serial number corrupted (scientific notation) | 5,608 | 49 | 0.9% |
| A10 | Serial number duplicated on assets of the same make | 5,136 | 787 | 15.3% |
| A11 | Manufacturer part number missing | 21,901 | 21,901 |  |
| A12 | No parent in the hierarchy | 27,184 | 2,616 | 9.6% |
| A13 | Parent code not in the register (orphan) | 27,193 | 4 | 0.0% |
| A14 | Duplicate asset code | 27,193 | 223 | 0.8% |
| A15 | Code prefix does not match site code | 27,193 | 2,010 | 7.4% |
| A16 | Site missing | 27,193 | 31 | 0.1% |
| A17 | Description missing | 27,193 | 1,244 | 4.6% |
| A18 | Status missing | 27,193 | 19 | 0.1% |
| A19 | Status "on" but PMs closed as asset not active | 27,193 | 60 | 0.2% |
| A20 | Equipment with no work order history (review list) | 19,443 | 16,173 | 83.2% |

## Inventory (Inventory_Gap_Analysis.xlsx)

| Rule | Description | Applicable | Fails | Rate |
|---|---|---|---|---|
| I01 | Description generic (2 words or fewer) | 314 | 76 | 24.2% |
| I02 | No manufacturer part number in the record | 314 | 259 | 82.5% |
| I03 | Same description on more than one part code | 314 | 81 | 25.8% |
| I04 | Unit cost missing or zero | 314 | 279 | 88.9% |
| I05 | Min quantity not set | 314 | 174 | 55.4% |
| I06 | Min set but zero | 140 | 28 | 20.0% |
| I07 | On hand below Min | 112 | 11 | 9.8% |
| I08 | Zero on hand | 314 | 13 | 4.1% |
| I09 | Not linked to a parent asset | 314 | 113 | 36.0% |
| I10 | Bin location blank | 314 | 314 |  |
| I11 | Account code blank | 314 | 23 | 7.3% |
| I12 | Charge dept / campus blank | 314 | 314 |  |
| I13 | Inventory / UNSPSC classification blank | 314 | 314 |  |
| I14 | Uncategorised item | 314 | 3 | 1.0% |
| I15 | Test / dummy record in live stock | 314 | 2 | 0.6% |
| I16 | Data-bearing / IT component without unit cost | 14 | 6 | 42.9% |
| I17 | No CMMS issue transaction 2019–2026 | 314 | 313 | 99.7% |
