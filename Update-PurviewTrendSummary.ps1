#requires -Version 5.1

# ============================================================
# Update Purview Trend Summary
# ============================================================

[CmdletBinding()]
param(
    [string]$SourceFolder = "C:\PurviewExports",
    [string]$TrendFolder = "C:\PurviewExports\TrendData",
    [int]$DaysBack = 7
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$TrendFile = Join-Path `
    $TrendFolder `
    "PurviewTrendSummary.csv"

# ============================================================
# Create trend folder
# ============================================================

if (!(Test-Path $TrendFolder))
{
    New-Item `
        -ItemType Directory `
        -Path $TrendFolder |
        Out-Null
}

# ============================================================
# Source files
# ============================================================

$LabelsFile = Join-Path `
    $SourceFolder `
    "SensitivityLabels.csv"

$DlpFile = Join-Path `
    $SourceFolder `
    "DLPActivities.csv"

$FilesFile = Join-Path `
    $SourceFolder `
    "FileActivities.csv"

$AiFile = Join-Path `
    $SourceFolder `
    "AIActivities.csv"

$RequiredFiles = @(
    $LabelsFile,
    $DlpFile,
    $FilesFile,
    $AiFile
)

foreach ($RequiredFile in $RequiredFiles)
{
    if (!(Test-Path $RequiredFile))
    {
        throw "Required report not found: $RequiredFile"
    }
}

# ============================================================
# Import current reporting-period data
# ============================================================

$Labels = @(
    Import-Csv -Path $LabelsFile
)

$Dlp = @(
    Import-Csv -Path $DlpFile
)

$Files = @(
    Import-Csv -Path $FilesFile
)

$Ai = @(
    Import-Csv -Path $AiFile
)

# ============================================================
# Helper function
# ============================================================

function Get-UniqueNonBlankCount
{
    param(
        [object[]]$Records,
        [string]$PropertyName
    )

    if ($null -eq $Records -or $Records.Count -eq 0)
    {
        return 0
    }

    $Values = @(
        $Records |
        ForEach-Object {
            $_.$PropertyName
        } |
        Where-Object {
            -not [string]::IsNullOrWhiteSpace($_)
        } |
        Sort-Object -Unique
    )

    return $Values.Count
}

# ============================================================
# Reporting-period dates
# ============================================================

$SnapshotDate = Get-Date
$PeriodEnd = $SnapshotDate
$PeriodStart = $SnapshotDate.AddDays(-$DaysBack)

# ============================================================
# Calculate trend metrics
# ============================================================

$DlpRuleMatches = @(
    $Dlp |
    Where-Object {
        $_.'Activity ID' -eq "DLPRuleMatch"
    }
).Count

$DlpRuleUndos = @(
    $Dlp |
    Where-Object {
        $_.'Activity ID' -eq "DLPRuleUndo"
    }
).Count

$CloudUploads = @(
    $Files |
    Where-Object {
        $_.'Activity ID' -eq "FileUploadedToCloud"
    }
).Count

$UsbCopies = @(
    $Files |
    Where-Object {
        $_.'Activity ID' -eq "FileCopiedToRemovableMedia"
    }
).Count

$PrintActivities = @(
    $Files |
    Where-Object {
        $_.'Activity ID' -eq "FilePrinted"
    }
).Count

$HighlyConfidentialEvents = @(
    $Labels |
    Where-Object {
        $_.'Sensitivity Label' -like "Highly Confidential*"
    }
).Count

$AiWebSearchInteractions = @(
    $Ai |
    Where-Object {
        $_.'Web Search Used' -match "^(True|TRUE|true)$"
    }
).Count

# ============================================================
# Build this run's summary row
# ============================================================

$TrendRow = [PSCustomObject]@{
    "Snapshot Date"                    =
        $SnapshotDate.ToString("yyyy-MM-dd HH:mm:ss")

    "Period Start"                     =
        $PeriodStart.ToString("yyyy-MM-dd")

    "Period End"                       =
        $PeriodEnd.ToString("yyyy-MM-dd")

    "Label Events"                     =
        $Labels.Count

    "DLP Events"                       =
        $Dlp.Count

    "File Activities"                  =
        $Files.Count

    "AI Interactions"                  =
        $Ai.Count

    "DLP Rule Matches"                 =
        $DlpRuleMatches

    "DLP Rule Undos"                   =
        $DlpRuleUndos

    "Cloud Uploads"                    =
        $CloudUploads

    "USB Copies"                       =
        $UsbCopies

    "Print Activities"                 =
        $PrintActivities

    "Highly Confidential Label Events" =
        $HighlyConfidentialEvents

    "Web Search AI Interactions"        =
        $AiWebSearchInteractions

    "Unique Label Users"               =
        Get-UniqueNonBlankCount `
            -Records $Labels `
            -PropertyName "User"

    "Unique DLP Users"                 =
        Get-UniqueNonBlankCount `
            -Records $Dlp `
            -PropertyName "User"

    "Unique File Users"                =
        Get-UniqueNonBlankCount `
            -Records $Files `
            -PropertyName "User"

    "Unique AI Users"                  =
        Get-UniqueNonBlankCount `
            -Records $Ai `
            -PropertyName "User"
}

# ============================================================
# Prevent duplicate snapshots on the same calendar date
# ============================================================

$ExistingRows = @()

if (Test-Path $TrendFile)
{
    $ExistingRows = @(
        Import-Csv -Path $TrendFile
    )
}

$TodayPrefix = $SnapshotDate.ToString("yyyy-MM-dd")

$ExistingRows = @(
    $ExistingRows |
    Where-Object {
        $_.'Snapshot Date' -notlike "$TodayPrefix*"
    }
)

$UpdatedTrendData = @(
    $ExistingRows
    $TrendRow
)

# ============================================================
# Export summary history
# ============================================================

$UpdatedTrendData |
Sort-Object {
    [datetime]$_.'Snapshot Date'
} |
Export-Csv `
    -Path $TrendFile `
    -NoTypeInformation `
    -Encoding UTF8

Write-Host ""
Write-Host "==========================================" `
    -ForegroundColor Green

Write-Host "Trend Summary Updated" `
    -ForegroundColor Green

Write-Host "Trend File: $TrendFile" `
    -ForegroundColor Green

Write-Host "Snapshots Stored: $($UpdatedTrendData.Count)" `
    -ForegroundColor Green

Write-Host "==========================================" `
    -ForegroundColor Green