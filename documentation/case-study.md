# Consolidation Engine — Portfolio Case Study

## Overview

I built the Consolidation Engine as a practical simulation of an operational data-processing problem: multiple Excel workbooks need to be converted into one standardized, validated, traceable reporting dataset.

Rather than stopping at a dashboard, the project focuses on the process behind the dashboard — ingestion, mapping, validation, deduplication, consolidation, audit logging, and reporting.

**Technology:** Excel VBA, Excel Tables, PivotTables, worksheet protection  
**Data:** Synthetic / simulated portfolio data

---

## Business Problem

A recurring operational reporting process may rely on several workbooks that do not always arrive in exactly the same structure.

Manual consolidation creates opportunities for:

- inconsistent field mapping
- invalid records
- duplicate records
- incomplete audit trails
- repeated manual preparation before reporting

The business requirement simulated here was simple: **turn multiple source workbooks into a controlled operational dataset that can be reviewed and reported with confidence.**

---

## Solution

The solution is a VBA-driven workflow controlled from a single Control Panel.

The process is:

**Import → Map → Validate → Deduplicate → Consolidate → Log → Report**

The operator launches one process while the VBA engine handles the repeatable processing steps behind the interface.

---

## Architecture

The system is separated into practical layers:

### 1. Input
Multiple operational `.xlsx` workbooks.

### 2. Automation Engine
`RunFullImport` coordinates the run and `ProcessOneFile` handles each source workbook.

### 3. Mapping and Validation
Source profiles standardize incoming fields, while validation rules identify data-quality problems.

### 4. Deduplication and Consolidation
Clean rows are checked against existing keys and then appended to the master dataset.

### 5. Audit Layer
Batch processing and validation activity are written to dedicated logs.

### 6. Reporting
The consolidated dataset supports the dashboard and KPI calculations.

---

## Key Features

### Multi-source processing
The workflow is designed to process multiple source workbooks in one run.

### Configurable mapping
Source-specific structures can be translated into a standardized destination schema.

### Validation and exception handling
Records that fail business or data-quality rules are logged for investigation.

### Cross-run deduplication
The engine seeds existing master keys before deduplication so that rerunning previously processed files does not simply append the same records again.

### Auditability
Each run receives a batch ID and records file-level processing outcomes and row counts.

### Operational interface
The Control Panel provides run control, navigation, and status visibility.

### Protected workbook design
Operational sheets are protected after processing while the VBA workflow is allowed controlled write access during execution.

---

## Data Quality

The latest completed run produced:

- **47** retained master records
- **2** validation exceptions
- **48** duplicate records identified in the run history
- **94.0%** dashboard data-quality rate

The exception and duplicate populations remain traceable rather than being silently discarded.

---

## Dashboard Results

The finished dashboard summarizes the processed operational dataset with:

- **1,392** flights handled
- **76.9%** fleet-wide on-time performance
- **148,253** passengers handled
- **33.6 minutes** average turnaround time
- **94.0%** data-quality rate

It includes four analytical visuals covering station performance, quality mix, daily volumes, and station totals.

---

## Final System Test

The final live test confirmed:

- VBA compilation passed
- workbook startup protection passed
- four source files were processed
- zero files failed
- 47 master records remained after processing
- protected operational sheets were restored after the run
- dashboard and navigation remained functional

---

## Business Impact Demonstrated

The project demonstrates a repeatable pattern for reducing manual spreadsheet handling and improving traceability in recurring operational workflows.

Potential real-world applications include:

- daily or weekly operations reporting
- multi-branch consolidation
- recurring management reports
- quality-control workflows
- reconciliation preparation
- exception monitoring

This case study uses simulated data and is a portfolio demonstration, not a representation of a live enterprise deployment.

---

## Technical Lessons

One of the most important design decisions was treating the workbook as a small application rather than a collection of spreadsheets.

The project therefore combines:

**interface + business rules + data processing + audit trail + reporting + protection**

That structure makes the solution more relevant to real operational automation work than a dashboard-only exercise.
