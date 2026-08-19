# Microsoft Purview Activity Explorer Reporting Platform

## Overview

The solution uses native Activity Explorer exports combined with Excel automation and SharePoint publishing.

---

# Phase 1 - Reporting Framework

## Sensitivity Label Activity

Exports:

- LabelApplied
- LabelChanged
- LabelRemoved
- LabelRecommended
- LabelApplyFailed
- LabelRecommendedAndDismissed

Output:

SensitivityLabels.csv

---

## DLP Activity

Exports all DLP-related Activity Explorer events.

Output:

DLPActivities.csv

---

## File Activity

Exports:

- File Uploaded To Cloud
- File Printed
- File Copied To Clipboard
- File Copied To Removable Media
- File Read
- File Modified

Output:

FileActivities.csv

---

## AI Activity

Exports:

- Copilot Interactions
- AI Events

Output:

AIActivities.csv

---

# Phase 2 - Dashboard

Workbook:

PurviewExecutiveDashboard.xlsx

## Worksheets

Executive Summary

Labels Data

DLP Data

File Data

AI Data

Labels Pivot

DLP Pivot

File Pivot

AI Pivot

Trends

---

## Executive Summary Metrics

- Label Events
- DLP Events
- File Activities
- AI Interactions
- Last Refreshed Timestamp

---

# Phase 3 - Automation

The reporting suite is executed through:

Run-PurviewReportingSuite.ps1

Workflow:

Generate Reports

Refresh Dashboard

Update Trend Summary

Archive Snapshot

Publish Reports

---

# Phase 4 - Historical Trending

Trend data stored in:

PurviewTrendSummary.csv

Columns:

Snapshot Date

Period Start

Period End

Label Events

DLP Events

File Activities

AI Interactions

Cloud Uploads

USB Copies

Print Activities

Highly Confidential Label Events

Unique AI Users

Unique File Users

Unique DLP Users

Unique Label Users

---

# Phase 5 - Publishing

Published Location

Purview Admins - Purview Reports

Publishing mechanism:

OneDrive Sync

No PnP PowerShell required.

---

# Scheduling

Task Scheduler

Weekly

PowerShell 7

Action:

pwsh.exe -ExecutionPolicy Bypass -File Run-PurviewReportingSuite.ps1
