# Excel VBA Multi-Report Auditor (Cross-Workbook Data Sync Suite)

A robust, two-phase VBA automation system designed to handle complex weekly report rollovers and dynamically audit/synchronize daily guest records from raw external workbooks against target corporate account arrays.

This toolkit turns what was once a tedious, manual, 45-minute daily cross-referencing task into a single-button, 3-second operational workflow, reducing human error to 0%.

---

## 🚀 Repository Description (GitHub About Section)

A multi-file VBA automation suite that manages weekly report transitions and dynamically audits/syncs daily raw arrivals against targeted corporate account arrays, eliminating manual cross-referencing and reducing processing time by 90%.

---

## 📂 Repository Structure

This repository is structured to allow recruiters and developers to test the automation from a clean slate (`before-vba`) and compare it with the successful operational outputs (`after-vba`):

    excel-vba-multi-report-auditor/
    │
    ├── before-vba/                               # Sandbox testing environment (clean state)
    │   ├── 2102.xlsx                             # Raw Daily Arrival report (Day 1)
    │   ├── 2202.xlsx                             # Raw Daily Arrival report (Day 2)
    │   ├── 2302.xlsx                             # Raw Daily Arrival report (Day 3)
    │   ├── 2402.xlsx                             # Raw Daily Arrival report (Day 4)
    │   ├── 2502.xlsx                             # Raw Daily Arrival report (Day 5)
    │   ├── 2602.xlsx                             # Raw Daily Arrival report (Day 6 - Blank State)
    │   ├── 2702.xlsx                             # Raw Daily Arrival report (Day 7 - Blank State)
    │   └── Double Points Strategy Arrival List.xlsm # Master Tracker Template
    │
    ├── after-vba/                                # Output reference showing final automated state
    │   ├── 2102.xlsx                             # (Daily files updated with filtering flags)
    │   └── Double Points Strategy Arrival List.xlsm # Fully audited & populated Master Tracker
    │
    ├── images/                                   # Screenshots visualizing execution flow
    │   ├── 1. Phase 1 - Report Before.png
    │   ├── 2. Phase 1 - Report Rollover Complete.png
    │   ├── 3. Phase 2 - All 7 Days Data Opened.png
    │   ├── 4. Phase 2 - Complete Message.png
    │   ├── 5. Phase 2 - Complete Each Day Data Sheet.png
    │   ├── 6. Phase 2 - Report - Example day with targets.png
    │   └── 7. Phase 2 - Report - Example day without targets.png
    │
    ├── src/                                      # Modular VBA Source Code (.bas)
    │   ├── Phase1_WeeklyReportRollover.bas       # Structural setup & cleanup module
    │   └── Phase2_ProcessDailyReports.bas        # Relational cross-workbook processing module
    │
    └── README.md                                 # Documentation

---

## 🛠️ The Two-Phase Workflow

This automation is split into two operational phases to handle database structural management (Phase 1) and cross-workbook querying (Phase 2).

### Phase 1: Automated Weekly Rollover
Before tracking begins for a new week, this macro clones the master tracker, updates its internal metadata, and completely resets sheet layouts for the upcoming week.
* **Dynamic Date Range Tracking:** Reads the date name of the first tab (e.g., `28 Mar 2026`), adds 7 days to calculate the new tracking boundaries, and dynamically generates a cleanly formatted output filename (e.g., `Double Points Strategy Arrival List - 0404 - 1004.xlsm`).
* **Environment Scrubbing:** Loops through all 7 internal daily tabs, updates the visual header banners to reflect the upcoming tracking dates, and cleanly wipes structural layout ranges (Row 5 downwards) of transactional contents and manual formatting flags while strictly preserving table borders and background formulas.

#### Execution Visuals (Phase 1)
| Before Rollover (Stale Data) | After Rollover (Clean Workspace & Updated Dates) |
| :---: | :---: |
| <img src="./images/1. Phase 1 - Report Before.png" width="100%"> | <img src="./images/2. Phase 1 - Report Rollover Complete.png" width="100%"> |

---

### Phase 2: Relational Cross-Workbook Auditing
Once the clean weekly folder environment is staged, Phase 2 processes all raw transactional spreadsheets simultaneously.
* **Dynamic Title & Fallback Binding:** Dynamically reads the alphanumeric `CDate` format of each tracker tab, transforms it to an expected four-character target string (`mmdd`), and scans memory for matching open workbooks. Includes a sequential numerical fallback (`0221` through `0227`) for swift portfolio deployment.
* **Array-Based Root Match Engine:** Compiles an internal key array of target corporate clients (e.g., *WAYNE ENTERPRISES*, *STARK INDUSTRIES*). It injects a specialized in-memory `MatchHelper` tracker field into Column Z of the raw files, mapping partial strings using `InStr` combined with `vbTextCompare` to bypass inconsistent system spaces or regional entity suffixes.
* **Visibility Inversion & Safe Copying:** Activates an `AutoFilter` pass to catch positive rows instantly, aggregates visible elements with `SpecialCells(xlCellTypeVisible)`, and cuts away unnecessary formatting tracks before deploying raw unformatted text blocks down into the Master sheet structure using `.PasteSpecial Paste:=xlPasteValues`.
* **Zero-Condition Exception Handling:** If zero targeted corporate matches occur within a raw document (e.g., weekend slow periods), the engine catches the empty row collection safely, aborts paste loops to preserve sheet structures, writes an administrative message (`"There is none for [Date]"`), and flags the block with a light visual highlight for quick operational audit reviews.

#### Execution Visuals (Phase 2)
| Multi-File Environment State | Processing Success Catch |
| :---: | :---: |
| <img src="./images/3. Phase 2 - All 7 Days Data Opened.png" width="100%"> | <img src="./images/4. Phase 2 - Complete Message.png" width="100%"> |

#### Master Tracker Outputs After Matching
| Example: Active Tracking Day with Targets Matched | Example: Weekend Day with Zero Matches Handled |
| :---: | :---: |
| <img src="./images/6. Phase 2 - Report - Example day with targets.png" width="100%"> | <img src="./images/7. Phase 2 - Report - Example day without targets.png" width="100%"> |

---

## 💻 How To Run

1. Clone or download the `before-vba` sandbox folder directory to your local machine.
2. Open all 7 raw day data files (`2102.xlsx` to `2702.xlsx`) along with the master `Double Points Strategy Arrival List.xlsm` template.
3. Access the VBA Editor window (**Alt + F11**).
4. Run `Phase1_WeeklyReportRollover` to stage the structural parameters of your sheet.
5. Run `Phase2_ProcessDailyReports` to execute the automated relational lookup loops across all active workbooks simultaneously.
