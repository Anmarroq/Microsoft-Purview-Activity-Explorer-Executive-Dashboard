#requires -Version 5.1

# ============================================================
# Run Purview Reporting Suite
# ============================================================

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# The folder containing this script and all child scripts.
$ScriptRoot = $PSScriptRoot

if ([string]::IsNullOrWhiteSpace($ScriptRoot))
{
    $ScriptRoot = "C:\Users\Angel"
}

# ============================================================
# Script paths
# ============================================================

$GenerateReportsScript = Join-Path `
    $ScriptRoot `
    "Run-All-PurviewReports.ps1"

$RefreshDashboardScript = Join-Path `
    $ScriptRoot `
    "Refresh-PurviewDashboard.ps1"

$UpdateTrendScript = Join-Path `
    $ScriptRoot `
    "Update-PurviewTrendSummary.ps1"

$ArchiveReportsScript = Join-Path `
    $ScriptRoot `
    "Archive-PurviewReports.ps1"

$PublishReportsScript = Join-Path `
    $ScriptRoot `
    "Publish-PurviewReports.ps1"

$RequiredScripts = @(
    $GenerateReportsScript,
    $RefreshDashboardScript,
    $UpdateTrendScript,
    $ArchiveReportsScript,
    $PublishReportsScript
)

foreach ($RequiredScript in $RequiredScripts)
{
    if (!(Test-Path $RequiredScript))
    {
        throw "Required script not found: $RequiredScript"
    }
}

# ============================================================
# Logging
# ============================================================

$LogFolder = "C:\PurviewExports\Logs"

if (!(Test-Path $LogFolder))
{
    New-Item `
        -ItemType Directory `
        -Path $LogFolder |
        Out-Null
}

$LogFileName = "PurviewReporting_{0}.log" -f `
    (Get-Date -Format "yyyyMMdd_HHmmss")

$LogFile = Join-Path `
    $LogFolder `
    $LogFileName

$TranscriptStarted = $false
$SuiteSucceeded = $false
$SuiteStartTime = Get-Date

try
{
    Start-Transcript `
        -Path $LogFile `
        -Force |
        Out-Null

    $TranscriptStarted = $true

    Write-Host ""
    Write-Host "==========================================" `
        -ForegroundColor Cyan

    Write-Host "Purview Reporting Suite Starting" `
        -ForegroundColor Cyan

    Write-Host "Started: $($SuiteStartTime.ToString('yyyy-MM-dd HH:mm:ss'))" `
        -ForegroundColor Cyan

    Write-Host "Script folder: $ScriptRoot" `
        -ForegroundColor Cyan

    Write-Host "==========================================" `
        -ForegroundColor Cyan

    # ========================================================
    # Step 1: Generate current seven-day reports
    # ========================================================

    Write-Host ""
    Write-Host "Step 1 of 5 - Generate Reports" `
        -ForegroundColor Yellow

    & $GenerateReportsScript

    if (-not $?)
    {
        throw "The report-generation script did not complete successfully."
    }

    # ========================================================
    # Step 2: Refresh the executive dashboard
    # ========================================================

    Write-Host ""
    Write-Host "Step 2 of 5 - Refresh Dashboard" `
        -ForegroundColor Yellow

    & $RefreshDashboardScript

    if (-not $?)
    {
        throw "The dashboard-refresh script did not complete successfully."
    }

    # ========================================================
    # Step 3: Add or replace today's trend-summary row
    # ========================================================

    Write-Host ""
    Write-Host "Step 3 of 5 - Update Trend Summary" `
        -ForegroundColor Yellow

    & $UpdateTrendScript

    if (-not $?)
    {
        throw "The trend-summary script did not complete successfully."
    }

    # ========================================================
    # Step 4: Archive the current snapshot
    # ========================================================

    Write-Host ""
    Write-Host "Step 4 of 5 - Archive Reports" `
        -ForegroundColor Yellow

    & $ArchiveReportsScript

    if (-not $?)
    {
        throw "The archive script did not complete successfully."
    }

    # ========================================================
    # Step 5: Publish current reports to synced SharePoint
    # ========================================================

    Write-Host ""
    Write-Host "Step 5 of 5 - Publish Reports" `
        -ForegroundColor Yellow

    & $PublishReportsScript

    if (-not $?)
    {
        throw "The publishing script did not complete successfully."
    }

    $SuiteSucceeded = $true
    $SuiteEndTime = Get-Date
    $Elapsed = $SuiteEndTime - $SuiteStartTime

    Write-Host ""
    Write-Host "==========================================" `
        -ForegroundColor Green

    Write-Host "Purview Reporting Suite Complete" `
        -ForegroundColor Green

    Write-Host "Completed: $($SuiteEndTime.ToString('yyyy-MM-dd HH:mm:ss'))" `
        -ForegroundColor Green

    Write-Host "Duration: $($Elapsed.ToString())" `
        -ForegroundColor Green

    Write-Host "Log: $LogFile" `
        -ForegroundColor Green

    Write-Host "==========================================" `
        -ForegroundColor Green
}
catch
{
    $SuiteEndTime = Get-Date

    Write-Host ""
    Write-Host "==========================================" `
        -ForegroundColor Red

    Write-Host "Purview Reporting Suite Failed" `
        -ForegroundColor Red

    Write-Host "Failed: $($SuiteEndTime.ToString('yyyy-MM-dd HH:mm:ss'))" `
        -ForegroundColor Red

    Write-Host "Error: $($_.Exception.Message)" `
        -ForegroundColor Red

    Write-Host "Script: $($_.InvocationInfo.ScriptName)" `
        -ForegroundColor Red

    Write-Host "Line: $($_.InvocationInfo.ScriptLineNumber)" `
        -ForegroundColor Red

    Write-Host "Log: $LogFile" `
        -ForegroundColor Red

    Write-Host "==========================================" `
        -ForegroundColor Red

    throw
}
finally
{
    if ($TranscriptStarted)
    {
        try
        {
            Stop-Transcript | Out-Null
        }
        catch
        {
            Write-Warning "The transcript could not be stopped cleanly."
        }
    }

    if (-not $SuiteSucceeded)
    {
        $Host.SetShouldExit(1)
    }
}