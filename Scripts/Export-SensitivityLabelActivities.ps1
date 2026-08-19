# ==========================================
# Purview Activity Explorer - Label Report
# ==========================================

# ==========================================
# Settings
# ==========================================

$DaysBack = 7

$OutputFolder = "C:\PurviewExports"

if (!(Test-Path $OutputFolder))
{
    New-Item `
        -ItemType Directory `
        -Path $OutputFolder | Out-Null
}

# ==========================================
# Build Label Lookup Table
# ==========================================

Write-Host "Building label lookup..." `
    -ForegroundColor Cyan

$Labels = @(Get-Label)

$LabelLookup = @{}

foreach ($Label in $Labels)
{
    if (![string]::IsNullOrEmpty($Label.ParentId))
    {
        $Parent = $Labels |
            Where-Object {
                $_.ImmutableId.ToString() -eq `
                $Label.ParentId.ToString()
            }

        $FriendlyName =
            "$($Parent.DisplayName)/$($Label.DisplayName)"
    }
    else
    {
        $FriendlyName =
            $Label.DisplayName
    }

    $LabelLookup[
        $Label.ImmutableId.ToString()
    ] = $FriendlyName
}

# ==========================================
# Activities
# ==========================================

$Activities = @(
    "LabelApplied",
    "LabelApplyFailed",
    "LabelChanged",
    "LabelRemoved",
    "LabelRecommended",
    "LabelRecommendedAndDismissed"
)

# ==========================================
# Collect Activity Explorer Data
# ==========================================

$Data = @()

foreach ($Activity in $Activities)
{
    Write-Host `
        "Exporting $Activity..." `
        -ForegroundColor Yellow

    $Export =
        Export-ActivityExplorerData `
            -StartTime (Get-Date).AddDays(-$DaysBack) `
            -EndTime (Get-Date) `
            -OutputFormat Json `
            -Filter1 @(
                "Activity",
                $Activity
            ) `
            -PageSize 5000

    if ($Export.ResultData)
    {
        $ActivityData = @(
            $Export.ResultData |
            ConvertFrom-Json
        )

        $Data += $ActivityData

        Write-Host `
            "Retrieved $($ActivityData.Count) records" `
            -ForegroundColor Green
    }
    else
    {
        Write-Host `
            "No records found" `
            -ForegroundColor DarkYellow
    }
}

# ==========================================
# Create Friendly Output
# ==========================================

Write-Host `
    "Building report..." `
    -ForegroundColor Cyan

$Output = @(
    $Data |
    Select-Object `
    @{
        Name='Activity'
        Expression={$_.Activity}
    },
    @{
        Name='Activity ID'
        Expression={$_.ActivityId}
    },
    @{
        Name='File'
        Expression={$_.FilePath}
    },
    @{
        Name='Location'
        Expression={$_.Workload}
    },
    @{
        Name='User'
        Expression={$_.User}
    },
    @{
        Name='Happened'
        Expression={$_.Happened}
    },
    @{
        Name='Sensitivity Label'
        Expression={
            if ($_.SensitivityLabel)
            {
                $LabelLookup[
                    $_.SensitivityLabel.ToString()
                ]
            }
        }
    },
    @{
        Name='Old Sensitivity Label'
        Expression={
            if ($_.OldSensitivityLabel)
            {
                $LabelLookup[
                    $_.OldSensitivityLabel.ToString()
                ]
            }
        }
    },
    @{
        Name='Application'
        Expression={$_.Application}
    },
    @{
        Name='Device'
        Expression={$_.DeviceName}
    },
    @{
        Name='How Applied'
        Expression={$_.HowApplied}
    },
    @{
        Name='Label Event Type'
        Expression={$_.LabelEventType}
    },
    @{
        Name='Protection Event Type'
        Expression={$_.ProtectionEventType}
    },
    @{
        Name='Client IP'
        Expression={$_.ClientIP}
    },
    @{
        Name='File Extension'
        Expression={$_.FileExtension}
    }
)

# ==========================================
#V
# ==========================================

$OutputFile = Join-Path `
    $OutputFolder `
    "SensitivityLabels.csv"

$Output |
Export-Csv `
    -Path $OutputFile `
    -NoTypeInformation `
    -Encoding UTF8

# ==========================================
# Summary
# ==========================================

Write-Host ""
Write-Host `
    "=======================================" `
    -ForegroundColor Green

Write-Host `
    "Export Complete" `
    -ForegroundColor Green

Write-Host `
    "Records Exported: $($Output.Count)" `
    -ForegroundColor Green

Write-Host `
    "Output File:" `
    -ForegroundColor Green

Write-Host `
    $OutputFile `
    -ForegroundColor Green

Write-Host `
    "=======================================" `
    -ForegroundColor Green