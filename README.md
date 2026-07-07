Excel VBA Multi-Report Auditor (Cross-Workbook Data Sync Suite)

A robust, two-phase VBA automation system designed to handle complex weekly report rollovers and dynamically audit/synchronize daily guest records from raw external workbooks against target corporate account arrays.

This toolkit turns what was once a tedious, manual, 45-minute daily cross-referencing task into a single-button, 3-second operational workflow, reducing human error to 0%.

🚀 Repository Description (GitHub About Section)

A multi-file VBA automation suite that manages weekly report transitions and dynamically audits/syncs daily raw arrivals against targeted corporate account arrays, eliminating manual cross-referencing and reducing processing time by 90%.

📂 Repository Structure

This repository is structured to allow recruiters and developers to test the automation from a clean slate (before-vba) and compare it with the successful operational outputs (after-vba):

excel-vba-multi-report-auditor/
│
├── before-vba/                           # Sandbox testing environment (clean state)
│   ├── 2102.xlsx                         # Raw Daily Arrival report (Day 1)
│   ├── 2202.xlsx                         # Raw Daily Arrival report (Day 2)
│   ├── 2302.xlsx                         # Raw Daily Arrival report (Day 3)
│   ├── 2402.xlsx                         # Raw Daily Arrival report (Day 4)
│   ├── 2502.xlsx                         # Raw Daily Arrival report (Day 5)
│   ├── 2602.xlsx                         # Raw Daily Arrival report (Day 6 - Blank State)
│   ├── 2702.xlsx                         # Raw Daily Arrival report (Day 7 - Blank State)
│   └── Double Points Strategy Arrival List.xlsm  # Master Tracker Template
│
├── after-vba/                            # Output reference showing final automated state
│   ├── 2102.xlsx                         
│   ├── ...                               # (Daily files updated with filtering flags)
│   └── Double Points Strategy Arrival List.xlsm  # Fully audited & populated Master Tracker
│
├── images/                               # Screenshots visualizing execution flow
│   ├── 1. Phase 1 - Report Before.png
│   ├── 2. Phase 1 - Report Rollover Complete.png
│   ├── 3. Phase 2 - All 7 Days Data Open.png
│   ├── 4. Phase 2 - Complete Message.png
│   ├── 5. Phase 2 - Complete Each Day.png
│   ├── 6. Phase 2 - Report - Example Day With Target.png
│   └── 7. Phase 2 - Report - Example Day With None.png
│
├── src/                                  # Modular VBA Source Code (.bas)
│   ├── Phase1_WeeklyReportRollover.bas   # Structural setup & cleanup module
│   └── Phase2_ProcessDailyReports.bas    # Relational cross-workbook processing module
│
└── README.md                             # Documentation


🛠️ The Two-Phase Workflow

This automation is split into two operational phases to handle database structural management (Phase 1) and cross-workbook querying (Phase 2).

Phase 1: Automated Weekly Rollover

Before tracking begins for a new week, this macro clones the master tracker, updates its internal metadata, and completely resets sheet layouts for the upcoming week.

Dynamic Date Range Tracking: Reads the date name of the first tab (e.g., 28 Mar 2026), adds 7 days to calculate the new tracking boundaries, and dynamically generates a cleanly-formatted output filename (e.g., Double Points Strategy Arrival List - 0404 - 1004).

Safe Duplication (SaveAs): Converts the workbook into a fresh Macro-Enabled Workbook (.xlsm) while preserving the historical file.

Tab & Metadata Re-calibration: Iterates through all 7 daily sheets, re-dating the tabs (e.g., 04 Apr 2026) and updating inner sheet title headers in Cell B2 simultaneously.

Layout-Preserving Scrub: Locates the boundaries of the existing tracking grid, flushes out stale data rows, and strips color highlights while leaving cell borders, grids, and formatting intact.

Before Phase 1 Rollover:


After Phase 1 Rollover Complete:


Phase 2: Relational Cross-Report Auditing

This engine acts as an in-memory database query that cross-references newly exported raw daily arrivals with target accounts.

Fuzzy File-Match Extraction: Automatically formats target date parameters into search hashes (e.g., "04 Apr 2026" -> "0404" / "0221"), loops through all open system sheets, and pairs them cleanly with the correct daily tab.

Fuzzy Account Match Filtering: Iterates through every record in the raw sheets, comparing the company column (Column S) against a defined array of target corporate accounts. Using InStr combined with vbTextCompare, it handles spacing/punctuation variations safely.

Automated Column Strip & Pivot: Injects a dynamic MatchHelper flag column in Column Z, filters the raw sheet, creates an in-memory temporary sheet to wipe unneeded columns, re-orders structural metrics, and extracts clean values.

Graceful Empty-State Flags: If a target day has zero corporate arrivals, the macro bypasses standard copy-paste values, injecting a stylized warning label ("There is none for [Date]" highlighted in operational yellow) so audits are visually explicit.

Preparing Data Workspace (All 7 Days Open):


Macro Processing and Completion Dialog:


Audit Output - Active Day (With target corporate arrivals):


Audit Output - Blank Day (Zero corporate arrivals flagged in yellow):


🌟 Advanced VBA Code Design Highlights

This suite showcases advanced programming patterns beyond basic recordable macros:

Dual Naming Fallbacks: Supports both calendar dates (e.g., 0404) and portfolio testing arrays (e.g., 0221 fallback) to accommodate both live data and static mock validation.

Target Array Hardening: Features fuzzy root-matching keywords (e.g., "Hewlett Packard", "HPE", "Aust High") designed to safely filter spelling variations output by legacy property management software.

Aggressive Runtime Optimization: Disables screen updating, forces manual calculation rules, and suppresses alert dialogs during execution—preventing performance chokepoints.

Defensive Error Handling: Uses On Error Resume Next and IsError safeguards around array loops to ensure missing tabs, corrupted text, or system error values (e.g., #N/A or #REF!) never crash the main macro execution.

💻 How to Use the Sandbox Assets

To test this locally on your machine, follow these steps using the assets in before-vba:

Download the Testing Workspace:

Open before-vba/Double Points Strategy Arrival List.xlsm.

Open all 7 raw files (2102.xlsx through 2702.xlsx).

Import code modules:

Press Alt + F11 to open the VBA editor.

Right-click the project hierarchy and import both .bas files from the src/ folder.

Execute:

Run Phase1_WeeklyReportRollover to generate a newly-dated tracker file.

Run Phase2_ProcessDailyReports to instantly audit the raw sheets and watch the matching corporate arrivals copy over automatically.
