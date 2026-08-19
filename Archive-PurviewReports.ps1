# ==========================================
# Archive Purview Reports
# ==========================================

$SourceFolder =
"C:\PurviewExports"

$ArchiveRoot =
"C:\PurviewExports\Archive"

$ArchiveFolder =
Join-Path `
    $ArchiveRoot `
    (Get-Date -Format "yyyy-MM-dd")

if (!(Test-Path $ArchiveFolder))
{
    New-Item `
        -ItemType Directory `
        -Path $ArchiveFolder | Out-Null
}

$Files = @(
    "PurviewExecutiveDashboard.xlsx",
    "SensitivityLabels.csv",
    "DLPActivities.csv",
    "FileActivities.csv",
    "AIActivities.csv"
)

foreach ($File in $Files)
{
    $SourceFile = Join-Path `
        $SourceFolder `
        $File

    if (Test-Path $SourceFile)
    {
        Copy-Item `
            $SourceFile `
            $ArchiveFolder `
            -Force
    }
}

Write-Host ""
Write-Host "Archived to:" `
    -ForegroundColor Green

Write-Host $ArchiveFolder `
    -ForegroundColor Green