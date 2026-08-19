# ==========================================
# Purview Activity Explorer - DLP Report
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
# Build Sensitive Information Type Lookup
# ==========================================

Write-Host "Building Sensitive Information Type Lookup..." `
    -ForegroundColor Cyan

$SITLookup = @{}

Get-DlpSensitiveInformationType |
ForEach-Object {

    $SITLookup[$_.Id.ToString()] = $_.Name

}

# ==========================================
# Activities
# ==========================================

$Activities = @(
    "DLPRuleMatch",
    "DLPRuleUndo"
)

# ==========================================
# Collect Activity Explorer Data
# ==========================================

$Data = @()

foreach ($Activity in $Activities)
{
    Write-Host "Exporting $Activity..." `
        -ForegroundColor Yellow

    $Export = Export-ActivityExplorerData `
        -StartTime (Get-Date).AddDays(-$DaysBack) `
        -EndTime (Get-Date) `
        -OutputFormat Json `
        -Filter1 @("Activity",$Activity) `
        -PageSize 5000

    if ($Export.ResultData)
    {
        $ActivityData = $Export.ResultData | ConvertFrom-Json

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
    Name='Sensitive Information Type'
    Expression={

        if($_.SensitiveInfoTypeData)
        {
            $SITLookup[
                $_.SensitiveInfoTypeData[0].SensitiveInfoTypeId.ToString()
            ]
        }

    }
},
@{
    Name='Match Count'
    Expression={

        if($_.SensitiveInfoTypeData)
        {
            $_.SensitiveInfoTypeData[0].Count
        }

    }
},
@{
    Name='Confidence'
    Expression={

        if($_.SensitiveInfoTypeData)
        {
            $_.SensitiveInfoTypeData[0].Confidence
        }

    }
},
@{
    Name='Policy'
    Expression={
        $_.PolicyMatchInfo.PolicyName
    }
},
@{
    Name='Rule'
    Expression={
        $_.PolicyMatchInfo.RuleName
    }
},
@{
    Name='Policy Mode'
    Expression={
        $_.PolicyMatchInfo.PolicyMode
    }
},
@{
    Name='Enforcement Mode'
    Expression={
        $_.EnforcementMode
    }
},
@{
    Name='Endpoint Operation'
    Expression={
        $_.EndpointOperation
    }
},
@{
    Name='Application'
    Expression={
        $_.Application
    }
},
@{
    Name='Device'
    Expression={
        $_.DeviceName
    }
},
@{
    Name='Client IP'
    Expression={
        $_.ClientIP
    }
},
@{
    Name='Platform'
    Expression={
        $_.Platform
    }
},
@{
    Name='File Type'
    Expression={
        $_.FileType
    }
},
@{
    Name='File Extension'
    Expression={
        $_.FileExtension
    }
},
@{
    Name='File Size'
    Expression={
        $_.FileSize
    }
},
@{
    Name='Source Location'
    Expression={
        $_.SourceLocationType
    }
},
@{
    Name='Target Location'
    Expression={
        $_.DestinationLocationType
    }
}

# ==========================================
# Export CSV
# ==========================================

$OutputFile = Join-Path `
    $OutputFolder `
    "DLPActivities.csv"

$Output |
Export-Csv `
    -Path $OutputFile `
    -NoTypeInformation `
    -Encoding UTF8

# ==========