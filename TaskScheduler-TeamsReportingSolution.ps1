param (
    [Parameter(Mandatory=$true)]
    [Array]$UpnFilter,
    [Parameter(Mandatory=$true)]
    [String]$ReportFolder,
    [Parameter(Mandatory=$true, ParameterSetName='FirstHalfOfMonth')]
    [switch]$FirstHalfOfMonth,
    [Parameter(Mandatory=$true, ParameterSetName='SecondHalfOfMonth')]
    [switch]$SecondHalfOfMonth,    
    [Parameter(Mandatory=$false, ParameterSetName='ManualDates')]
    [Parameter(Mandatory=$true, ParameterSetName='SecondHalfOfMonth')]
    [switch]$EmailAddressList,
    [Parameter(Mandatory=$true, ParameterSetName='ManualDates')]
    [switch]$ManualDates
)
$SharedParams = @{
    UpnFilter = $UpnFilter
    ReportFolder = $ReportFolder
    ExportFileName = "TeamsCallReport"
    PowerBiDatasetId = "5ffb107c-32b9-4b26-b01e-5833102409c1"
    AppId = '8e171bdb-1c73-4e55-9683-db76ebd770e5'
    TenantId = '04b9e073-f7cf-4c95-9f91-e6d55d5a3797'
    ClientSecretFile = "C:\Temp\TeamsPBI\ClientSecret.txt"
    AdminUserName = "jakegwynn@jakegwynndemo.com"
    AdminPwFile = "C:\Temp\TeamsPBI\AdminPassword.txt"
    LogFile = "C:\Temp\TeamsPBI\TeamsCallReportingLog.txt"
}

$EmailParams = @{
    EmailAddressList = $EmailAddressList
    SmtpUserName = "postmaster@jakegwynn.com"
    SmtpPwFile = "C:\Temp\TeamsPBI\SmtpPassword.txt"
    SmtpServer = "smtp.mailgun.org"
    SmtpPort = 587
    SendAsUser = "noreply@jakegwynn.com"
}

# FIRST HALF OF MONTH:
$FirstHalfScriptParams = $SharedParams

#SECOND HALF OF MONTH:
$SecondHalfScriptParams = $SharedParams + $EmailParams

# MANUAL DATES:
$ManualDatesScriptParams = $SecondHalfScriptParams + @{
    StartDay = 1
    StartMonthNumber = 2
    StartYear = 2022
    EndDay = 1
    EndMonthNumber = 3
    EndYear = 2022
}

$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition

cd $scriptPath

if ($FirstHalfOfMonth) {
    .\TeamsReportingSolution-Workflow.ps1 @FirstHalfScriptParams -FirstHalfOfMonth #-DebugMode
}
elseif ($SecondHalfOfMonth) {
    .\TeamsReportingSolution-Workflow.ps1 @SecondHalfScriptParams -SecondHalfOfMonth #-DebugMode
} 
else {
    .\TeamsReportingSolution.ps1 @ManualDatesScriptParams #-DebugMode
}

<#
    .\TeamsReportingSolution-Workflow.ps1 @FirstHalfScriptParams 
    .\TeamsReportingSolution-Workflow.ps1 @SecondHalfScriptParams
#>