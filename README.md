# Consolidation Engine
### Multi-Source Excel VBA Data Consolidation & Quality Control System

> A portfolio-grade Excel/VBA automation system that consolidates operational workbooks, standardizes source structures, validates records, detects duplicates, maintains a master dataset, logs exceptions, and produces management-ready reporting.

**Project status:** v1.0 — Portfolio Master  
**Platform:** Microsoft Excel  
**Automation:** VBA  
**Data:** Synthetic / simulated portfolio data

---

## 1. Business Problem

Operational reporting workflows often begin with multiple Excel workbooks produced by different sources, stations, teams, or reporting processes.

When those files are consolidated manually, common risks include:

- inconsistent source structures and field names
- invalid or incomplete records entering the reporting dataset
- duplicate records across files or repeated runs
- weak traceability of what was imported and rejected
- time spent repeatedly preparing the same reporting dataset

The goal of this project was to build a reusable Excel-based operational consolidation workflow rather than a standalone dashboard.

---

## 2. Solution

The Consolidation Engine automates the workflow from source-file ingestion to reporting:

```text
Source Workbooks
      ↓
Import
      ↓
Source Profile Matching
      ↓
Field Mapping
      ↓
Validation
      ↓
Deduplication
      ↓
Master Data Consolidation
      ↓
Audit / Quality Logging
      ↓
Dashboard & Operational Reporting
```

The operator interacts with the system through a dedicated **Control Panel**, while configuration and helper layers remain separated from the user-facing workflow.

---

## 3. System Architecture

### Input Layer
Multiple `.xlsx` source files are discovered from the input workflow and processed through VBA.

### VBA Processing Layer
The main processing sequence is coordinated by `RunFullImport()` and `ProcessOneFile()`.

### Configuration Layer
Source profiles and validation rules are stored in dedicated configuration tables so that processing logic is not hard-coded entirely into the workflow.

### Data Quality Layer
Records are split into clean and flagged outcomes. Duplicate detection is applied before consolidation, including protection against duplicates introduced by rerunning previously processed data.

### Output Layer
Clean records are written to `Master_Data`, while operational activity and exceptions are captured in `Import_Log` and `Data_Quality_Log`.

### Reporting Layer
A dashboard presents KPIs and operational charts using the consolidated data and helper calculations.

---

## 4. VBA Architecture

The workbook is organized into focused VBA modules:

| Module | Responsibility |
|---|---|
| `mod_Main` | Main orchestration and `RunFullImport` entry point |
| `mod_Import` | Workbook import and file processing |
| `mod_Mapping` | Source-profile matching and row remapping |
| `mod_Validate` | Validation and clean/flagged separation |
| `mod_Dedupe` | Duplicate detection across files and reruns |
| `mod_Consolidate` | Append validated records to the master dataset |
| `mod_Utilities` | Logging and shared helper routines |
| `mod_Config` | Configuration access |
| `mod_Protection` | Worksheet protection and controlled VBA writing |

This separation keeps the automation easier to test, troubleshoot, and extend.

---

## 5. Workbook Structure

### User-facing sheets

- `Control_Panel` — operator entry point and system status
- `Dashboard` — KPI and operational reporting view
- `Master_Data` — standardized consolidated dataset
- `Import_Log` — batch-level processing history
- `Data_Quality_Log` — validation exceptions and rejected rows
- `Instructions` — operating guide

### Hidden / backend sheets

- `Dashboard_Data`
- `Config_SourceProfiles`
- `Config_ValidationRules`

---

## 6. Key Features

### Multi-source ingestion
Processes multiple Excel source files through one controlled workflow.

### Configurable source mapping
Source profiles identify incoming structures and map fields into a standardized schema.

### Validation and exception handling
Invalid records are separated from clean records and written to an auditable quality log instead of being silently ignored.

### Cross-file and cross-run deduplication
Business keys are used to prevent repeated records from being written into the consolidated master dataset, including protection against rerunning previously processed files.

### Audit trail
Each run records a batch ID, timestamp, source file, processing status, row counts, duplicate counts, and processing messages.

### Operator Control Panel
A central interface provides the main run action, workbook navigation, and system status indicators.

### Automated reporting
Processed data feeds five KPI cards and four operational dashboard visuals.

### Worksheet protection
Operational data sheets are protected while retaining controlled VBA write capability through the protection workflow.

---

## 7. Dashboard

The final dashboard contains five KPI cards:

| KPI | Result |
|---|---:|
| Total Flights Handled | **1,392** |
| Fleet-wide On-Time % | **76.9%** |
| Total Passengers Handled | **148,253** |
| Avg. Turnaround Time | **33.6 min** |
| Data Quality Rate | **94.0%** |

It also includes four visuals:

1. On-Time Performance by Station
2. Data Quality Mix
3. Daily Flights Handled by Station
4. Flights Handled by Station

---

## 8. Final Test Results

The final live test of the protected workbook completed successfully.

| Test | Result |
|---|---|
| VBA compile | ✅ Passed |
| Workbook opens on Control Panel | ✅ Passed |
| Protected sheets before run | ✅ Passed |
| Source files processed | **4** |
| Files failed | **0** |
| Master records after processing | **47** |
| Protected sheets after run | ✅ Passed |
| Dashboard integrity | ✅ Passed |
| Navigation | ✅ Passed |

The workbook also recorded **2 validation exceptions** and **48 duplicates** in the final control-panel status after the latest rerun. These are tracked separately from the 47 retained master records.

---

## 9. Data Quality Approach

The system deliberately treats data quality as part of the processing workflow rather than as a post-processing reporting task.

A record can be:

- accepted as clean
- flagged for validation review
- rejected as a duplicate

The quality rate shown on the dashboard is based on the workbook's processed clean, flagged, and duplicate populations.

---

## 10. Business Value Demonstrated

This project demonstrates how Excel/VBA can be used to build a lightweight operational automation application for recurring workflows such as:

- recurring report consolidation
- operational data preparation
- validation and exception handling
- reconciliation support
- audit logging
- management reporting

The project is intended as a portfolio simulation and does **not** represent proprietary company data or a live production system.

---

## 11. Project Screenshots

Recommended portfolio gallery:

```text
screenshots/
├── 01-control-panel.png
├── 02-dashboard.png
├── 03-master-data.png
├── 04-import-log.png
├── 05-data-quality-log.png
└── 06-system-architecture.png
```

The architecture image should illustrate:

```text
Input Files
   ↓
RunFullImport
   ↓
ProcessOneFile
   ↓
Import → Mapping → Validation → Deduplication
   ↓
AppendToMaster
   ↓
Import Log + Data Quality Log
   ↓
Dashboard
```

---

## 12. Technical Lessons

Building this system reinforced several practical lessons:

- dashboards are only as reliable as the process that feeds them
- validation should happen before consolidation
- deduplication needs to consider repeated runs, not only duplicates inside one file
- configuration tables can make automation more adaptable than hard-coded mappings
- audit logs make automated workflows easier to investigate
- protection needs to be designed alongside VBA write operations rather than added at the end

---

## 13. Future Enhancements

Potential next versions could introduce:

- user-selectable input folders
- configurable business-key definitions
- stronger error recovery and retry handling
- automated email/report distribution
- Power Query or Power BI integration
- database-backed master storage
- scheduled execution through an external orchestration layer

---

## 14. Author

**Lukman Oyewo**  
Aeronautical & Astronautical Engineer | Data Analytics | Excel VBA | Power BI | SQL | Python  
Lagos, Nigeria

---
