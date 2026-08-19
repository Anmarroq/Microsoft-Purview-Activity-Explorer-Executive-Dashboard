#requires -Version 5.1

<#
.SYNOPSIS
Creates a refreshable Excel workbook shell for the Purview Activity Explorer CSV reports.

.DESCRIPTION
Creates PurviewExecutiveDashboard.xlsx with:
- Executive Summary
- Labels Data, DLP Data, File Data, AI Data
- Labels Pivot, DLP Pivot, File Pivot, AI Pivot placeholder sheets

Each Data sheet is connected to its corresponding CSV through an Excel QueryTable.
The queries refresh when the workbook opens and can also be refreshed with Data > Refresh All.

Prerequisites:
- Microsoft Excel desktop installed
- The four CSV files present in the configured source folder

Run this script once to create the workbook. Build PivotTables and charts once in Excel.
The workbook can then refresh against replacement CSV files that retain the same paths,
file names, and column headings.
#>

[CmdletBinding()]
param(
    [string]$SourceFolder = "C:\PurviewExports",
    [string]$WorkbookPath = "C:\PurviewExports\PurviewExecutiveDashboard.xlsx",
    [switch]$Overwrite
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$CsvDefinitions = @(
    [pscustomobject]@{
        Name       = "Labels"
        CsvName    = "SensitivityLabels.csv"
        SheetName  = "Labels Data"
        PivotSheet = "Labels Pivot"
        QueryName  = "qryLabels"
    },
    [pscustomobject]@{
        Name       = "DLP"
        CsvName    = "DLPActivities.csv"
        SheetName  = "DLP Data"
        PivotSheet = "DLP Pivot"
        QueryName  = "qryDLP"
    },
    [pscustomobject]@{
        Name       = "Files"
        CsvName    = "FileActivities.csv"
        SheetName  = "File Data"
        PivotSheet = "File Pivot"
        QueryName  = "qryFiles"
    },
    [pscustomobject]@{
        Name       = "AI"
        CsvName    = "AIActivities.csv"
        SheetName  = "AI Data"
        PivotSheet = "AI Pivot"
        QueryName  = "qryAI"
    }
)

function Release-ComObject {
    param([object]$ComObject)

    if ($null -ne $ComObject) {
        try {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
        }
        catch {
            # Cleanup should not hide the original error.
        }
    }
}

function Set-WorksheetTitle {
    param(
        [Parameter(Mandatory)] [object]$Worksheet,
        [Parameter(Mandatory)] [string]$Title,
        [string]$Subtitle
    )

    $Worksheet.Cells.Item(1, 1).Value2 = $Title
    $TitleRange = $Worksheet.Range("A1:H1")
    $TitleRange.Merge()
    $TitleRange.Font.Name = "Aptos Display"
    $TitleRange.Font.Size = 20
    $TitleRange.Font.Bold = $true
    $TitleRange.Font.Color = 0xFFFFFF
    $TitleRange.Interior.Color = 0xA05A00
    $TitleRange.HorizontalAlignment = -4131
    $TitleRange.VerticalAlignment = -4108
    $Worksheet.Rows.Item(1).RowHeight = 32

    if (-not [string]::IsNullOrWhiteSpace($Subtitle)) {
        $Worksheet.Cells.Item(2, 1).Value2 = $Subtitle
        $SubtitleRange = $Worksheet.Range("A2:H2")
        $SubtitleRange.Merge()
        $SubtitleRange.Font.Name = "Aptos"
        $SubtitleRange.Font.Size = 10
        $SubtitleRange.Font.Color = 0x666666
        $Worksheet.Rows.Item(2).RowHeight = 22
    }
}

function Add-CsvQueryTable {
    param(
        [Parameter(Mandatory)] [object]$Worksheet,
        [Parameter(Mandatory)] [string]$CsvPath,
        [Parameter(Mandatory)] [string]$QueryName
    )

    $Destination = $Worksheet.Range("A1")
    $QueryTable = $Worksheet.QueryTables.Add("TEXT;$CsvPath", $Destination)
    $QueryTable.Name = $QueryName
    $QueryTable.FieldNames = $true
    $QueryTable.RowNumbers = $false
    $QueryTable.FillAdjacentFormulas = $false
    $QueryTable.PreserveFormatting = $true
    $QueryTable.RefreshOnFileOpen = $true
    $QueryTable.RefreshStyle = 1
    $QueryTable.SavePassword = $false
    $QueryTable.SaveData = $true
    $QueryTable.AdjustColumnWidth = $true
    $QueryTable.RefreshPeriod = 0
    $QueryTable.TextFilePromptOnRefresh = $false
    $QueryTable.TextFilePlatform = 65001
    $QueryTable.TextFileStartRow = 1
    $QueryTable.TextFileParseType = 1
    $QueryTable.TextFileTextQualifier = 1
    $QueryTable.TextFileConsecutiveDelimiter = $false
    $QueryTable.TextFileTabDelimiter = $false
    $QueryTable.TextFileSemicolonDelimiter = $false
    $QueryTable.TextFileCommaDelimiter = $true
    $QueryTable.TextFileSpaceDelimiter = $false
    $QueryTable.TextFileTrailingMinusNumbers = $true
    $QueryTable.BackgroundQuery = $false
    [void]$QueryTable.Refresh($false)

    $ResultRange = $QueryTable.ResultRange
    if ($null -ne $ResultRange -and $ResultRange.Rows.Count -ge 1) {
        $HeaderRange = $Worksheet.Range(
            $Worksheet.Cells.Item(1, 1),
            $Worksheet.Cells.Item(1, $ResultRange.Columns.Count)
        )
        $HeaderRange.Font.Bold = $true
        $HeaderRange.Font.Color = 0xFFFFFF
        $HeaderRange.Interior.Color = 0xA05A00
        $HeaderRange.HorizontalAlignment = -4108
        $HeaderRange.WrapText = $true
        $Worksheet.Application.ActiveWindow.SplitRow = 1
        $Worksheet.Application.ActiveWindow.FreezePanes = $true
        $Worksheet.UsedRange.Columns.AutoFit() | Out-Null

        for ($Column = 1; $Column -le $ResultRange.Columns.Count; $Column++) {
            if ($Worksheet.Columns.Item($Column).ColumnWidth -gt 42) {
                $Worksheet.Columns.Item($Column).ColumnWidth = 42
            }
        }
    }

    return $QueryTable
}

function Add-KpiCard {
    param(
        [Parameter(Mandatory)] [object]$Worksheet,
        [Parameter(Mandatory)] [string]$RangeAddress,
        [Parameter(Mandatory)] [string]$Title,
        [Parameter(Mandatory)] [string]$Formula,
        [Parameter(Mandatory)] [int]$FillColor
    )

    $CardRange = $Worksheet.Range($RangeAddress)
    $CardRange.Merge()
    $CardRange.Interior.Color = $FillColor
    $CardRange.Borders.LineStyle = 1
    $CardRange.Borders.Color = 0xD9D9D9
    $CardRange.HorizontalAlignment = -4108
    $CardRange.VerticalAlignment = -4108

    $TopLeft = $CardRange.Cells.Item(1, 1)
    $TopLeft.Formula = $Formula
    $TopLeft.NumberFormat = "#,##0"
    $TopLeft.Font.Name = "Aptos Display"
    $TopLeft.Font.Size = 24
    $TopLeft.Font.Bold = $true
    $TopLeft.Font.Color = 0xFFFFFF

    $LabelRow = $CardRange.Row + $CardRange.Rows.Count
    $LabelRange = $Worksheet.Range(
        $Worksheet.Cells.Item($LabelRow, $CardRange.Column),
        $Worksheet.Cells.Item($LabelRow, $CardRange.Column + $CardRange.Columns.Count - 1)
    )
    $LabelRange.Merge()
    $LabelRange.Value2 = $Title
    $LabelRange.Font.Name = "Aptos"
    $LabelRange.Font.Size = 10
    $LabelRange.Font.Bold = $true
    $LabelRange.Font.Color = 0x404040
    $LabelRange.HorizontalAlignment = -4108
}

if (!(Test-Path $SourceFolder)) {
    New-Item -ItemType Directory -Path $SourceFolder | Out-Null
}

$MissingFiles = @()
foreach ($Definition in $CsvDefinitions) {
    $CsvPath = Join-Path $SourceFolder $Definition.CsvName
    if (!(Test-Path $CsvPath)) {
        $MissingFiles += $CsvPath
    }
}

if ($MissingFiles.Count -gt 0) {
    $MissingList = $MissingFiles -join [Environment]::NewLine
    throw "The following required CSV files were not found:`n$MissingList"
}

if (Test-Path $WorkbookPath) {
    if ($Overwrite) {
        Remove-Item $WorkbookPath -Force
    }
    else {
        throw "The workbook already exists: $WorkbookPath. Use -Overwrite to replace it."
    }
}

$Excel = $null
$Workbook = $null
$CreatedQueryTables = @()

try {
    Write-Host "Starting Excel..." -ForegroundColor Cyan
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $Excel.ScreenUpdating = $false

    $Workbook = $Excel.Workbooks.Add()

    while ($Workbook.Worksheets.Count -gt 1) {
        $Workbook.Worksheets.Item($Workbook.Worksheets.Count).Delete()
    }

    $ExecutiveSheet = $Workbook.Worksheets.Item(1)
    $ExecutiveSheet.Name = "Executive Summary"

    foreach ($Definition in $CsvDefinitions) {
        $DataSheet = $Workbook.Worksheets.Add()
        $DataSheet.Name = $Definition.SheetName

        $CsvPath = Join-Path $SourceFolder $Definition.CsvName
        Write-Host "Connecting $($Definition.CsvName)..." -ForegroundColor Yellow
        $QueryTable = Add-CsvQueryTable -Worksheet $DataSheet -CsvPath $CsvPath -QueryName $Definition.QueryName
        $CreatedQueryTables += $QueryTable

        $PivotSheet = $Workbook.Worksheets.Add()
        $PivotSheet.Name = $Definition.PivotSheet
        Set-WorksheetTitle -Worksheet $PivotSheet -Title "$($Definition.Name) Analysis" -Subtitle "Create PivotTables and charts on this sheet using '$($Definition.SheetName)' as the source."
        $PivotSheet.Cells.Item(4, 1).Value2 = "Suggested PivotTable"
        $PivotSheet.Cells.Item(4, 1).Font.Bold = $true
        $PivotSheet.Cells.Item(5, 1).Value2 = switch ($Definition.Name) {
            "Labels" { "Rows: Sensitivity Label | Values: Count of Activity | Filters: Activity, Location, User" }
            "DLP"    { "Rows: Policy or Sensitive Information Type | Values: Count of Activity | Filters: Rule, User, Device" }
            "Files"  { "Rows: Activity | Values: Count of Activity | Filters: User, Device, Application, Target Domain" }
            "AI"     { "Rows: AI App Name or User | Values: Count of Activity | Filters: AI App Location, App Group, Web Search Used" }
        }
        $PivotSheet.Range("A4:H5").WrapText = $true
        $PivotSheet.Columns.Item(1).ColumnWidth = 100
        $PivotSheet.Range("A7").Value2 = "After creating the PivotTable, add a PivotChart and optional slicers."
        $PivotSheet.Range("A7").Font.Italic = $true
        $PivotSheet.Range("A7").Font.Color = 0x666666
        $PivotSheet.Activate() | Out-Null
        $Excel.ActiveWindow.DisplayGridlines = $false
    }

    $ExecutiveSheet.Activate() | Out-Null
    $Excel.ActiveWindow.DisplayGridlines = $false
    Set-WorksheetTitle -Worksheet $ExecutiveSheet -Title "Microsoft Purview Executive Dashboard" -Subtitle "Weekly Activity Explorer reporting. Use Data > Refresh All after the CSV exports are replaced."

    $ExecutiveSheet.Columns.Item("A").ColumnWidth = 3
    foreach ($ColumnName in @("B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M")) {
        $ExecutiveSheet.Columns.Item($ColumnName).ColumnWidth = 12
    }

    Add-KpiCard -Worksheet $ExecutiveSheet -RangeAddress "B4:D7" -Title "Sensitivity Label Events" -Formula "=MAX(0,COUNTA('Labels Data'!A:A)-1)" -FillColor 0xD47800
    Add-KpiCard -Worksheet $ExecutiveSheet -RangeAddress "E4:G7" -Title "DLP Events" -Formula "=MAX(0,COUNTA('DLP Data'!A:A)-1)" -FillColor 0x0078D4
    Add-KpiCard -Worksheet $ExecutiveSheet -RangeAddress "H4:J7" -Title "File Activities" -Formula "=MAX(0,COUNTA('File Data'!A:A)-1)" -FillColor 0x008272
    Add-KpiCard -Worksheet $ExecutiveSheet -RangeAddress "K4:M7" -Title "AI Interactions" -Formula "=MAX(0,COUNTA('AI Data'!A:A)-1)" -FillColor 0x5C2D91

    $ExecutiveSheet.Range("B10:M10").Merge()
    $ExecutiveSheet.Range("B10").Value2 = "Dashboard Build Checklist"
    $ExecutiveSheet.Range("B10").Font.Bold = $true
    $ExecutiveSheet.Range("B10").Font.Size = 14
    $ExecutiveSheet.Range("B10").Font.Color = 0xFFFFFF
    $ExecutiveSheet.Range("B10:M10").Interior.Color = 0x404040

    $Checklist = @(
        "1. Open each Pivot sheet and insert the suggested PivotTable from its matching Data sheet.",
        "2. Add PivotCharts for top labels, policies, file activities, and AI apps.",
        "3. Add slicers for User, Activity, Workload/Location, and date where useful.",
        "4. Move or copy the most important charts to this Executive Summary sheet.",
        "5. Use Data > Refresh All after the weekly CSV exports are replaced.",
        "6. Save the workbook after refresh so SharePoint users see current values."
    )

    $Row = 11
    foreach ($Item in $Checklist) {
        $ExecutiveSheet.Range("B$Row:M$Row").Merge()
        $ExecutiveSheet.Range("B$Row").Value2 = $Item
        $ExecutiveSheet.Range("B$Row").WrapText = $true
        $ExecutiveSheet.Rows.Item($Row).RowHeight = 24
        $Row++
    }

    $ExecutiveSheet.Range("B18:M18").Merge()
    $ExecutiveSheet.Range("B18").Value2 = "Last workbook build: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $ExecutiveSheet.Range("B18").Font.Italic = $true
    $ExecutiveSheet.Range("B18").Font.Color = 0x666666

    $Workbook.Worksheets.Item("Executive Summary").Move($Workbook.Worksheets.Item(1))
    $Workbook.Worksheets.Item("Labels Data").Move($Workbook.Worksheets.Item(2))
    $Workbook.Worksheets.Item("DLP Data").Move($Workbook.Worksheets.Item(3))
    $Workbook.Worksheets.Item("File Data").Move($Workbook.Worksheets.Item(4))
    $Workbook.Worksheets.Item("AI Data").Move($Workbook.Worksheets.Item(5))

    $Workbook.Worksheets.Item("Executive Summary").Activate() | Out-Null
    $Excel.CalculateFull()

    $Workbook.SaveAs($WorkbookPath, 51)
    $Workbook.Save()

    Write-Host "" 
    Write-Host "Dashboard workbook created successfully." -ForegroundColor Green
    Write-Host "Workbook: $WorkbookPath" -ForegroundColor Green
    Write-Host "CSV source folder: $SourceFolder" -ForegroundColor Green
    Write-Host "Open the workbook and create the PivotTables once on the four Pivot sheets." -ForegroundColor Cyan
}
finally {
    foreach ($QueryTable in $CreatedQueryTables) {
        Release-ComObject $QueryTable
    }

    if ($null -ne $Workbook) {
        try { $Workbook.Close($true) } catch {}
        Release-ComObject $Workbook
    }

    if ($null -ne $Excel) {
        try { $Excel.Quit() } catch {}
        Release-ComObject $Excel
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
