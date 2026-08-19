# ==========================================
# Publish Purview Reports
# ==========================================

$SourceFolder = "C:\PurviewExports"

$DestinationFolder = `
"C:\Users\Angel\Contoso\Purview Admins - Purview Reports"

Write-Host ""
Write-Host "==========================================" `
    -ForegroundColor Cyan

Write-Host "Purview Report Publishing Started" `
    -ForegroundColor Cyan

Write-Host "==========================================" `
    -ForegroundColor Cyan

Write-Host ""
Write-Host "Source Folder:" `
    -ForegroundColor Yellow

Write-Host $SourceFolder

Write-Host ""
Write-Host "Destination Folder:" `
    -ForegroundColor Yellow

Write-Host $DestinationFolder

Write-Host ""

if (!(Test-Path $SourceFolder))
{
    throw "Source folder not found: $SourceFolder"
}

if (!(Test-Path $DestinationFolder))
{
    throw "Destination folder not found: $DestinationFolder"
}

$Files = @(
    "PurviewExecutiveDashboard.xlsx",
    "SensitivityLabels.csv",
    "DLPActivities.csv",
    "FileActivities.csv",
    "AIActivities.csv"
)

# Add trend file separately
$TrendFile =
"C:\PurviewExports\TrendData\PurviewTrendSummary.csv"

$Published = 0

foreach ($File in $Files)
{
    $SourceFile = Join-Path `
        $SourceFolder `
        $File

    Write-Host ""
    Write-Host "Checking $File..." `
        -ForegroundColor Yellow

    if (Test-Path $SourceFile)
    {
        Write-Host "Found file" `
            -ForegroundColor Green

        Copy-Item `
            -Path $SourceFile `
            -Destination $DestinationFolder `
            -Force

        Write-Host `
            "Published $File" `
            -ForegroundColor Green

        $Published++
    }
    else
    {
        Write-Warning `
            "$File not found"
    }
}

# Trend Summary

Write-Host ""
Write-Host "Checking PurviewTrendSummary.csv..." `
    -ForegroundColor Yellow

if (Test-Path $TrendFile)
{
    Copy-Item `
        -Path $TrendFile `
        -Destination $DestinationFolder `
        -Force

    Write-Host `
        "Published PurviewTrendSummary.csv" `
        -ForegroundColor Green

    $Published++
}
else
{
    Write-Warning `
        "Trend file not found:"

    Write-Warning `
        $TrendFile
}

Write-Host ""
Write-Host "==========================================" `
    -ForegroundColor Green

Write-Host "Publishing Complete" `
    -ForegroundColor Green

Write-Host "Files Published: $Published" `
    -ForegroundColor Green

Write-Host "==========================================" `
    -ForegroundColor Green