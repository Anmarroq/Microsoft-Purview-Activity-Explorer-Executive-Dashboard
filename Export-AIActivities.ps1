# ============================================================
# Purview Activity Explorer - AI Activities Report
# ============================================================

# Remove these later when using the Master Script

# ============================================================
# Settings
# ============================================================

$DaysBack = 7

$OutputFolder = "C:\PurviewExports"

if (!(Test-Path $OutputFolder))
{
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}

$OutputFile = Join-Path $OutputFolder "AIActivities.csv"

# ============================================================
# Activities
# ============================================================

$Activities = @(
    "CopilotInteraction"
)

# ============================================================
# Collect Data
# ============================================================

$AllData = @()

foreach ($Activity in $Activities)
{
    Write-Host "Exporting $Activity..." `
        -ForegroundColor Yellow

    $Export = Export-ActivityExplorerData `
        -StartTime (Get-Date).AddDays(-7) `
        -EndTime (Get-Date) `
        -OutputFormat Json `
        -Filter1 @("Activity",$Activity) `
        -PageSize 5000

    if ($Export.ResultData)
    {
        $ActivityData = $Export.ResultData | ConvertFrom-Json

foreach ($Record in @($ActivityData))
{
    $AllData += $Record
}

        Write-Host "Retrieved $($ActivityData.Count) records" `
            -ForegroundColor Green
    }
    else
    {
        Write-Host "No records found" `
            -ForegroundColor DarkYellow
    }
}

# ============================================================
# Build Friendly Report
# ============================================================

Write-Host "Building AI Report..." `
    -ForegroundColor Cyan

$Output = $AllData | ForEach-Object {

    $ThreadId = $null
    $AppHost = $null

    if ($_.CopilotEventData)
    {
        $ThreadId = $_.CopilotEventData.ThreadId
        $AppHost = $_.CopilotEventData.AppHost
    }

    [PSCustomObject]@{

        "Activity" = $_.Activity

        "Activity ID" = $_.ActivityId

        "User" = $_.User

        "Happened" = $_.Happened

        "AI App Name" = $_.PurviewAIAppName

        "AI App Location" = $_.PurviewAIAppLocation

        "App Host" = $AppHost

        "App Identity" = $_.AppIdentity

        "App Category" = $_.AppIdentityCategory

        "App Group" = $_.AppIdentityGroup

        "Workload" = $_.Workload

        "Data Platform" = $_.DataPlatform

        "Thread ID" = $ThreadId

        "Correlation ID" = $_.EnrichedCopilotThreadOrCorrelationId

        "Files Referenced" = $_.AreFilesReferenced

        "Sensitive Files Referenced" = $_.AreSensitiveFilesReferenced

        "Web Search Used" = $_.HasWebsearchQuery

        "Client IP" = $_.ClientIP

    }

}


# ============================================================
# Export CSV
# ============================================================

$Output |
Export-Csv `
    -Path $OutputFile `
    -NoTypeInformation `
    -Encoding UTF8

# ============================================================
# Complete
# ============================================================

Write-Host ""
Write-Host "==========================================" `
    -ForegroundColor Green

Write-Host "AI Activity Export Complete" `
    -ForegroundColor Green

Write-Host "Records Exported: $($Output.Count)" `
    -ForegroundColor Green

Write-Host "Output File:" `
    -ForegroundColor Green

Write-Host $OutputFile `
    -ForegroundColor Green

Write-Host "==========================================" `
    -ForegroundColor Green