# Microsoft Purview Activity Explorer Executive Dashboard

A PowerShell-based reporting framework for Microsoft Purview Activity Explorer that automatically:

- Exports Activity Explorer data
- Builds executive-friendly reports
- Refreshes an Excel dashboard
- Tracks historical trends
- Archives weekly snapshots
- Publishes reports to SharePoint

## Features

### Reporting

- Sensitivity Label Activity
- DLP Activity
- File Activity
- AI Activity

### Dashboard

- Executive Summary
- KPI Cards
- Interactive Pivot Tables
- Trend Reporting
- Historical Archive

### Automation

- Scheduled reporting
- Automated dashboard refresh
- Automated SharePoint publishing
- Historical snapshot retention

---

## Solution Architecture

Activity Explorer

↓ 

Export Scripts

↓ 

CSV Reports

↓ 

Excel Dashboard Refresh

↓ 

Trend Summary Generation

↓ 

Archive Snapshot

↓ 

Publish to SharePoint

---

## Scripts

| Script | Purpose |
|----------|----------|
| Run-All-PurviewReports.ps1 | Generate all report exports |
| Export-SensitivityLabelActivities.ps1 | Export label activity |
| Export-DLPActivities.ps1 | Export DLP activity |
| Export-FileActivities.ps1 | Export file activity |
| Export-AIActivities.ps1 | Export AI activity |
| Refresh-PurviewDashboard.ps1 | Refresh workbook |
| Update-PurviewTrendSummary.ps1 | Create trend data |
| Archive-PurviewReports.ps1 | Archive weekly snapshot |
| Publish-PurviewReports.ps1 | Publish reports |
| Run-PurviewReportingSuite.ps1 | Master orchestration script |

---

## Output Files

PurviewExecutiveDashboard.xlsx

SensitivityLabels.csv

DLPActivities.csv

FileActivities.csv

AIActivities.csv

PurviewTrendSummary.csv

---

## Requirements

### PowerShell

- PowerShell 7+

### Modules

ExchangeOnlineManagement

### Permissions

Microsoft Purview Activity Explorer access

Compliance Administrator role

or

Compliance Data Administrator role

---

## Future Enhancements

- Risk Summary Dashboard
- Historical Trend Charts
- Monthly KPI Reporting
- Data Quality Improvements
- Executive Scorecards
