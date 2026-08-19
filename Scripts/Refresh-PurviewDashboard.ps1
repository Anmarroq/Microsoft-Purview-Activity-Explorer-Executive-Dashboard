# ==========================================
# Refresh Purview Dashboard
# ==========================================

$WorkbookPath =
    "C:\PurviewExports\PurviewExecutiveDashboard.xlsx"

Write-Host "Opening dashboard..." `
    -ForegroundColor Cyan

$Excel = New-Object -ComObject Excel.Application

$Excel.Visible = $false
$Excel.DisplayAlerts = $false

$Workbook = $Excel.Workbooks.Open($WorkbookPath)

Write-Host "Waiting for workbook to load..." `
    -ForegroundColor Yellow

Start-Sleep -Seconds 10

Write-Host "Refreshing data..." `
    -ForegroundColor Yellow

$RefreshSuccess = $false

for ($i = 1; $i -le 5; $i++)
{
    try
    {
        $Workbook.RefreshAll()

        $RefreshSuccess = $true

        break
    }
    catch
    {
        Write-Warning "Refresh attempt $i failed."

        Start-Sleep -Seconds 5
    }
}

if (-not $RefreshSuccess)
{
    throw "Unable to refresh workbook."
}

Start-Sleep -Seconds 10

$Workbook.Worksheets("Executive Summary").Range("B3").Value2 =
    "Last Refreshed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

# Give Excel a moment to finish refreshes
Start-Sleep -Seconds 10

Write-Host "Saving workbook..." `
    -ForegroundColor Yellow

$Workbook.Save()

$Workbook.Close($true)

$Excel.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null

[GC]::Collect()
[GC]::WaitForPendingFinalizers()


Write-Host ""
Write-Host "==========================================" -ForegroundColor Green
Write-Host "Dashboard Refresh Complete" -ForegroundColor Green
Write-Host $WorkbookPath -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green