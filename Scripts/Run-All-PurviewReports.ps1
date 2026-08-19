Import-Module ExchangeOnlineManagement

if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue))
{
    Connect-IPPSSession
}

.\Export-SensitivityLabelActivities.ps1

.\Export-DLPActivities.ps1

.\Export-FileActivities.ps1

.\Export-AIActivities.ps1