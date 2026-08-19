# ==========================================
# Purview Activity Explorer - File Activities
# ==========================================

# ==========================================
# Settings
# ==========================================

$DaysBack = 7

$OutputFolder = "C:\PurviewExports"

if (!(Test-Path $OutputFolder))
{
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}

# ==========================================
# Build Label Lookup Table
# ==========================================

Write-Host "Building Label Lookup..." -ForegroundColor Cyan

$Labels = Get-Label

$LabelLookup = @{}

foreach ($Label in $Labels)
{
    if (![string]::IsNullOrEmpty($Label.ParentId))
    {
        $Parent = $Labels |
            Where-Object {
                $_.ImmutableId.ToString() -eq $Label.ParentId.ToString()
            }

        $FriendlyName = "$($Parent.DisplayName)/$($Label.DisplayName)"
    }
    else
    {
        $FriendlyName = $Label.DisplayName
    }

    $LabelLookup[$Label.ImmutableId.ToString()] = $FriendlyName
}

# ==========================================
# File Activities
# ==========================================

$Activities = @(
    "FileAccessedByUnallowedApp",
    "FileArchived",
    "FileCopiedToClipboard",
    "FileCopiedToNetworkShare",
    "FileCopiedToRemoteDesktopSession",
    "FileCopiedToRemovableMedia",
    "FileCreated",
    "FileCreatedOnNetworkShare",
    "FileCreatedOnRemovableMedia",
    "FileDeleted",
    "FileModified",
    "FilePrinted",
    "FileRead",
    "FileRenamed",
    "FileTransferredByBluetooth",
    "FileUploadedToCloud"
)

# ==========================================
# Collect Activity Explorer Data
# ==========================================

$Data = @()

foreach ($Activity in $Activities)
{
    Write-Host "Exporting $Activity..." -ForegroundColor Yellow

    $Export = Export-ActivityExplorerData `
        -StartTime (Get-Date).AddDays(-$DaysBack) `
        -EndTime (Get-Date) `
        -OutputFormat Json `
        -Filter1 @("Activity",$Activity) `
        -PageSize 5000

    if ($Export.ResultData)
    {
        $ActivityData = @(
    $Export.ResultData | ConvertFrom-Json
)

        $Data += $ActivityData

        Write-Host "Retrieved $($ActivityData.Count) records" `
            -ForegroundColor Green
    }
    else
    {
        Write-Host "No records found" `
            -ForegroundColor DarkYellow
    }
}

# ==========================================
# Build Friendly Output
# ==========================================

Write-Host "Building Report..." `
    -ForegroundColor Cyan

$Output = $Data |
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
        if($_.SensitivityLabel)
        {
            $LabelLookup[$_.SensitivityLabel.ToString()]
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
    Name='Platform'
    Expression={$_.Platform}
},
@{
    Name='Source Location'
    Expression={$_.SourceLocationType}
},
@{
    Name='Destination Location'
    Expression={$_.DestinationLocationType}
},
@{
    Name='Target Domain'
    Expression={$_.TargetDomain}
},
@{
    Name='Target URL'
    Expression={$_.TargetUrl}
},
@{
    Name='File Type'
    Expression={$_.FileType}
},
@{
    Name='File Extension'
    Expression={$_.FileExtension}
},
@{
    Name='File Size'
    Expression={$_.FileSize}
},
@{
    Name='Protected'
    Expression={$_.IsProtected}
},
@{
    Name='Protection Event'
    Expression={$_.ProtectionEventType}
},
@{
    Name='Data State'
    Expression={$_.DataState}
},
@{
    Name='Client IP'
    Expression={$_.ClientIP}
},
@{
    Name='Enforcement Mode'
    Expression={$_.EnforcementMode}
}

# ==========================================
# Export CSV
# ==========================================

$OutputFile = Join-Path `
    $OutputFolder `
    "FileActivities.csv"

$Output |
Export-Csv `
    -Path $OutputFile `
    -NoTypeInformation `
    -Encoding UTF8

# ==========================================
# Complete
# ==========================================

Write-Host ""
Write-Host "==========================================" -ForegroundColor Green
Write-Host "File Activity Export Complete" -ForegroundColor Green
Write-Host "Records Exported: $(@($Output).Count)" -ForegroundColor Green
Write-Host "Output File:" -ForegroundColor Green
Write-Host $OutputFile -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green