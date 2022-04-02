<#
Use examples:
  FIRST HALF OF MONTH:
   $ScriptParams = @{
    UpnFilter = "user1","user2"
    ReportFolder = ""
    ExportFileName = "TestExportWithFunctions1"
    PowerBiDatasetId = ""
    AppId = ''
    TenantId = ''
    ClientSecretFile = ""
    AdminUserName = ""
    AdminPwFile = ""
    LogFile = "" 
    FirstHalfOfMonth = $true
    MonthNumber = 2
  }

  SECOND HALF OF MONTH:
   $ScriptParams = @{
    UpnFilter = "user1","user2"
    ReportFolder = ""
    ExportFileName = "TestExportWithFunctions1"
    PowerBiDatasetId = ''
    AppId = ''
    TenantId = ''
    ClientSecretFile = ""
    AdminUserName = ""
    AdminPwFile = ""
    LogFile = "" 
    SecondHalfOfMonth = $true
    MonthNumber = 2
    EmailAddressList = "",""
    SmtpUserName = ""
    SmtpPwFile = ""
    SmtpServer = ""
    SmtpPort = 587
  }

    MANUAL DATES:
   $ScriptParams = @{
    UpnFilter = "user1","user2"
    ReportFolder = ""
    ExportFileName = "TestExportWithFunctions1"
    PowerBiDatasetId = ''
    AppId = ''
    TenantId = ''
    ClientSecretFile = ""
    AdminUserName = ""
    AdminPwFile = ""
    LogFile = "C:\Temp\TeamsPBI\log.txt" 
    StartDay = 1
    StartMonthNumber = 2
    StartYear = 2022
    EndDay = 1
    EndMonthNumber = 3
    EndYear = 2022
    EmailAddressList = ""
    SmtpUserName = ""
    SmtpPwFile = ""
    SmtpServer = ""
    SmtpPort = 587
  }


  cd "C:\ScriptFolder"
  .\TeamsReportingSolution.ps1 @ScriptParams

Copyright 2022 Jake Gwynn

DISCLAIMER:
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>
param (
  [Parameter(Mandatory=$true)]
  [Array]$UpnFilter,

  [Parameter(Mandatory=$true)]
  [String]$ReportFolder,

  [Parameter(Mandatory=$true)]
  [String]$ExportFileName,

  [Parameter(Mandatory=$true)]
  [String]$PowerBiDatasetId,

  [Parameter(Mandatory=$true)]
  [String]$TenantId,

  [Parameter(Mandatory=$true)]
  [String]$AppId,

  [Parameter(Mandatory=$true)]
  [String]$ClientSecretFile,

  [Parameter(Mandatory=$false)]
  [switch]$PowerBiAppAuthentication,

  [Parameter(Mandatory=$false)]
  [String]$LogFile,

  [Parameter(Mandatory=$false)]
  [switch]$DebugMode,

  [Parameter(Mandatory=$false, ParameterSetName='ManualDates')]
  [Parameter(Mandatory=$true, ParameterSetName='FirstHalfOfMonth')]
  [Parameter(Mandatory=$true, ParameterSetName='SecondHalfOfMonth')]
  [String]$AdminUserName,

  [Parameter(Mandatory=$false, ParameterSetName='ManualDates')]
  [Parameter(Mandatory=$true, ParameterSetName='FirstHalfOfMonth')]
  [Parameter(Mandatory=$true, ParameterSetName='SecondHalfOfMonth')]
  [String]$AdminPwFile,

  [Parameter(Mandatory=$false, ParameterSetName='FirstHalfOfMonth')]
  [Parameter(Mandatory=$false, ParameterSetName='SecondHalfOfMonth')]
  [int]$MonthNumber,

  [Parameter(Mandatory=$false, ParameterSetName='FirstHalfOfMonth')]
  [Parameter(Mandatory=$false, ParameterSetName='SecondHalfOfMonth')]
  [int]$Year,

  [Parameter(Mandatory=$true, ParameterSetName='FirstHalfOfMonth')]
  [switch]$FirstHalfOfMonth,

  [Parameter(Mandatory=$true, ParameterSetName='SecondHalfOfMonth')]
  [switch]$SecondHalfOfMonth,

  [Parameter(Mandatory=$false, ParameterSetName='SecondHalfOfMonth')]
  [string]$FirstHalfOfMonthFile,

  [Parameter(Mandatory=$false, ParameterSetName='SecondHalfOfMonth')]
  [Parameter(Mandatory=$false, ParameterSetName='ManualDates')]
  [array]$EmailAddressList,

  [Parameter(Mandatory=$false, ParameterSetName='SecondHalfOfMonth')]
  [Parameter(Mandatory=$false, ParameterSetName='ManualDates')]
  [string]$SmtpUserName,

  [Parameter(Mandatory=$false, ParameterSetName='SecondHalfOfMonth')]
  [Parameter(Mandatory=$false, ParameterSetName='ManualDates')]
  [string]$SmtpPwFile,

  [Parameter(Mandatory=$false, ParameterSetName='SecondHalfOfMonth')]
  [Parameter(Mandatory=$false, ParameterSetName='ManualDates')]
  [string]$SmtpServer,

  [Parameter(Mandatory=$false, ParameterSetName='SecondHalfOfMonth')]
  [Parameter(Mandatory=$false, ParameterSetName='ManualDates')]
  [string]$SmtpPort,

  [Parameter(Mandatory=$false, ParameterSetName='SecondHalfOfMonth')]
  [Parameter(Mandatory=$false, ParameterSetName='ManualDates')]
  [switch]$UseAnonymousSmtpAuth,

  [Parameter(Mandatory=$true, ParameterSetName='ManualDates')]
  [int]$StartDay,

  [Parameter(Mandatory=$true, ParameterSetName='ManualDates')]
  [int]$StartMonthNumber,

  [Parameter(Mandatory=$true, ParameterSetName='ManualDates')]
  [int]$StartYear,

  [Parameter(Mandatory=$true, ParameterSetName='ManualDates')]
  [int]$EndDay,

  [Parameter(Mandatory=$true, ParameterSetName='ManualDates')]
  [int]$EndMonthNumber,

  [Parameter(Mandatory=$true, ParameterSetName='ManualDates')]
  [int]$EndYear
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$TodaysDate = Get-Date 
$DateString = $TodaysDate.ToString("MM-dd-yyyy_HH-mm")

if ($ReportFolder.Substring($ReportFolder.Length - 1, 1) -eq "\") {
  $script:ReportFolder = $ReportFolder.Substring(0, $ReportFolder.Length - 1)
}

If($LogFile){
  if ($LogFile.Substring($LogFile.Length - 4, 4) -eq ".txt") {
    $LogFile = $LogFile.Substring(0, $LogFile.Length - 4)
  }
  If($DebugMode) {
    $LogFile = "$($LogFile)_DEBUG_$($UpnFilter[0])_$DateString.txt"
  }
  else {
    $LogFile = "$($LogFile)_$($UpnFilter[0])_$DateString.txt"
  }
}

If($DebugMode) {
  If (-not $LogFile) {
    $LogFile = "$ReportFolder\TeamsCallReportingLog_DEBUG_$DateString.txt"
  }
  Start-Transcript -Path $LogFile -Force
}

function Log ($Text,$ForegroundColor = "Yellow",$BackgroundColor = "Black") {
  $LogDate = (Get-Date).tostring()
  $LogText = "$($LogDate): $Text"
  If($LogFile -and (-not $DebugMode)) {
    Add-Content -path $LogFile -Value $LogText
  }
  Write-Host $LogText -ForegroundColor $ForegroundColor -BackgroundColor $BackgroundColor
}

function Set-QueryDates {
  Log "Setting Dates for Queries"
  If($FirstHalfOfMonth) {
    If (-not $MonthNumber) {$script:MonthNumber = $TodaysDate.Month}
    If (-not $Year) {$script:Year = $TodaysDate.Year}
    $script:Year = $TodaysDate.Year
    $script:StartDay = 1
    $script:EndDay = 15
  } 
  If($SecondHalfOfMonth) {
    If (-not $MonthNumber) {
      If($TodaysDate.Day -lt 16) {
        $script:MonthNumber = $TodaysDate.AddMonths(-1).Month
      }
      else {
        $script:MonthNumber = $TodaysDate.Month
      }
    }
    If (-not $Year) {$script:Year = If($MonthNumber -ne 12) {$TodaysDate.Year} else {$TodaysDate.AddYears(-1).Year}}
    $script:StartDay = 16
    $StartOfMonth = get-date -Month $MonthNumber -Year $Year -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0
    $EndOfMonth = ($startOfMonth).AddMonths(1).AddTicks(-1)
    $script:EndDay = $EndOfMonth.Day
  }
  If($FirstHalfOfMonth -or $SecondHalfOfMonth) {
    $script:StartMonthNumber = $MonthNumber
    $script:StartYear = $Year
    $script:EndMonthNumber = $MonthNumber
    $script:EndYear = $Year
  }
  $script:StartMonthNumberString = if($StartMonthNumber.tostring().Length -eq 1) {"0$StartMonthNumber"} else {$StartMonthNumber}
  $script:StartDayString = if($StartDay.tostring().Length -eq 1) {"0$StartDay"} else {$StartDay}
  $script:EndMonthNumberString = if($EndMonthNumber.tostring().Length -eq 1) {"0$EndMonthNumber"} else {$EndMonthNumber}
  $script:EndDayString = if($EndDay.tostring().Length -eq 1) {"0$EndDay"} else {$EndDay}
}

function Set-ReportFiles {
  Log "Setting Report Files"
  $FileNameSplit = $ExportFileName -split '\\'
  $ExportFileName = $FileNameSplit[$FileNameSplit.Count - 1]
  if ($ExportFileName.Substring($ExportFileName.Length - 4, 4) -eq ".csv") {
    $script:ExportFileName = $ExportFileName.Substring(0, $ReportFolder.Length - 4)
  }
  if ($FirstHalfOfMonth) {
    $script:ExportFileName = "$($ExportFileName)_FirstHalfOfMonth_$MonthNumber-$Year.csv"
  }
  elseif ($SecondHalfOfMonth) {
    If(-not $FirstHalfOfMonthFile) {
      $script:FirstHalfOfMonthFile = "$ReportFolder\$($ExportFileName)_FirstHalfOfMonth_$MonthNumber-$Year.csv"
    }
    Log "Report Folder: $ReportFolder"
    Log "FirstHalfOfMonthFile: $FirstHalfOfMonthFile"
    [System.Collections.Generic.List[psobject]]$script:FirstHalfOfMonthRecords = Import-Csv -Path "$FirstHalfOfMonthFile"
    $script:ExportFileName = "$($ExportFileName)_$MonthNumber-$Year.csv"
  }
  else {
    $script:ExportFileName = "$($ExportFileName)_$StartMonth-$StartDay-$($StartYear)_to_$EndMonth-$EndDay-$EndYear.csv"
  }
}

function Get-TeamsUsersWithNoLineUri {
  # Select necessary properties of and filter Teams Users 
  Log "Collecting Teams Users" 
  [array]$TeamsUsers = Get-CsOnlineUser -Filter {LineUri -ne $null} | Select-object UserPrincipalName,LineUri,SipAddress 

  Log "Parsing Teams Users LineUri"
  Foreach ($TeamsUser in $TeamsUsers) {
    If($TeamsUser.LineUri.Substring(0,4) -eq "tel:") {
      $TeamsUser.LineUri = ($TeamsUser.LineUri -split ":")[1]
    } 
    If ($TeamsUser.LineUri.Substring(0,1) -ne "+") {
      $TeamsUser.LineUri = "+$($TeamsUser.LineURI)"
    } 
  }
  return $TeamsUsers
}
function Get-UpnFilterLineUris ($AllTeamsUsers) {
  [System.Collections.Generic.List[psobject]]$UpnFilterLineUris = @()
  foreach ($UPN in $UpnFilter) {
    [array]$MatchedTeamsUsersLineUris = @()
    $MatchedTeamsUsersLineUris = $AllTeamsUsers.Where({$_.UserPrincipalName -like "*$UPN*"}).LineUri
    foreach ($LineUri in $MatchedTeamsUsersLineUris) {
        $UpnFilterLineUris.Add($LineUri)
    }
  }
  return $UpnFilterLineUris
}

function Connect-GraphApiWithClientSecret {
  Log "Connecting to Graph API to collect Teams PSTN call records"
  $Body = @{    
      Grant_Type    = "client_credentials"
      Scope         = "https://graph.microsoft.com/.default"
      client_Id     = $AppId
      Client_Secret = $ClientSecret
      } 
  $ConnectGraph = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method POST -Body $Body
  $script:ClientSecret = $null
  return $ConnectGraph.access_token
}
function Get-TeamsPstnCallRecords {  
  [System.Collections.Generic.List[psobject]]$PSTNCallRecords = @()
  $CallType = If ($apiUrl -like "*getPstnCalls*") {"PSTN"} elseif ($apiUrl -like "*getDirectRoutingCalls*") {"Direct Routing"} 
  Log "Collecting Teams $CallType call records from Graph API"
  $LoopIteration = 0
  $CurrentBatchCount = 0
  $TotalProcessedCount = 0
  $TotalMatchedCount = 0
  $MatchedBatchCount = 0
  $GetMoreRecords = $true
  While ($GetMoreRecords -eq $true) {
    $MatchedBatchCount = 0
    $LoopIteration++
    Log $apiUrl
    $CallRecordsResponse = (Invoke-RestMethod -Headers @{Authorization = "Bearer $Token"} -Uri $apiUrl -Method Get)
    $CurrentBatchCount = $CallRecordsResponse.'@odata.count'
    $TotalProcessedCount += $CurrentBatchCount
    $apiUrl = $CallRecordsResponse.'@odata.nextLink'
    If ($null -eq $apiUrl) {
      $GetMoreRecords = $false
    } else {
      $GetMoreRecords = $true
    }
    foreach ($Call in $CallRecordsResponse.Value) {  
      If (($Call.callerNumber -in $UpnFilterLineUris) -or ($Call.calleeNumber -in $UpnFilterLineUris)) {
        $PSTNCallRecord = [PSCustomObject]@{}
        $CallerTeamsUser = $TeamsUsers.Where({$_.LineURI -eq $Call.callerNumber})
        $CalleeTeamsUser = $TeamsUsers.Where({$_.LineURI -eq $Call.calleeNumber})
        $CallDuration = [math]::Round(($Call.duration/60),0)
        $CallStartTime = try{Get-Date $Call.startDateTime} catch {}
        $CallEndTime = try{Get-Date $Call.endDateTime} Catch {}
        $PSTNCallRecord = [PSCustomObject]@{
          CallerNumber = $Call.callerNumber
          CalleeNumber = $Call.calleeNumber
          CallerUPN = $CallerTeamsUser.UserPrincipalName
          CalleeUPN = $CalleeTeamsUser.UserPrincipalName
          CallerSIP = (($CallerTeamsUser.SipAddress) -split ":")[1]
          CalleeSIP = (($CalleeTeamsUser.SipAddress) -split ":")[1]
          StartDateTime = $CallStartTime
          EndDateTime = $CallEndTime
          'Duration(min)' = "$CallDuration"
          CallType = $CallType
        }
        $PSTNCallRecords.Add($PSTNCallRecord)
        $MatchedBatchCount++
      }
    }
    $TotalMatchedCount += $MatchedBatchCount
    Log "  Batch $LoopIteration finished. Records Processed:" -ForegroundColor Gray
    Log "    This Batch ($MatchedBatchCount Matched / $CurrentBatchCount Total)" -BackgroundColor DarkGreen -ForegroundColor Black
    Log "    Total $CallType ($TotalMatchedCount Matched / $TotalProcessedCount Total)" -BackgroundColor Yellow -ForegroundColor Black
  }
  return $PSTNCallRecords
}

function Set-TeamsCqdPowerBiQueries {
  Log "Storing Power BI Dax Queries"
  [array]$RequestBodyArray = @()
  Foreach ($UPN in $UpnFilter){
    $FirstUpnFilter = @"
      FILTER(
        KEEPFILTERS(VALUES('CQD'[First UPN])),
        SEARCH(\"$UPN\", 'CQD'[First UPN], 1, 0) >= 1
      ) 
"@
    $SecondUpnFilter = @"
      FILTER(
        KEEPFILTERS(VALUES('CQD'[Second UPN])),
        SEARCH(\"$UPN\", 'CQD'[Second UPN], 1, 0) >= 1
      ) 
"@
  $QueryFilterArray = @()
    $QueryFilterArray = @($FirstUpnFilter,$SecondUpnFilter)
    Foreach ($QueryFilter in $QueryFilterArray) {
      if ($QueryFilter -match "First UPN") {$StreamDirection = "First-to-Second"} else {$StreamDirection = "Second-to-First"}
      $RequestBodyArray += [PSCustomObject]@{
        UPN = $UPN
        StreamDirection = $StreamDirection
        # Request body must have
        RequestBody = @"
{
  "queries":
  [
    {"query": "
    // DAX Query
    DEFINE
    VAR __DS0FilterTable = 
    $QueryFilter

    VAR __DS0FilterTable2 = 
    FILTER(
        KEEPFILTERS(VALUES('CQD'[PSTN Call Type])),
        ISBLANK('CQD'[PSTN Call Type])
    )

    VAR __DS0FilterTable3 = 
    FILTER(KEEPFILTERS(VALUES('CQD'[Media Type])), 'CQD'[Media Type] = \"Audio\")

    VAR __DS0FilterTable4 = 
    FILTER(
        KEEPFILTERS(VALUES('CQD'[Start Time])),
        AND(
            'CQD'[Start Time] >= (DATE($StartYear, $StartMonthNumber, $StartDay) + TIME(0, 0, 1)),
            'CQD'[Start Time] < DATE($EndYear, $EndMonthNumber, $EndDay)
        )
    )

  VAR __DS0Core = 
    SUMMARIZECOLUMNS(
      ROLLUPADDISSUBTOTAL(
        ROLLUPGROUP(
            'CQD'[First UPN],
            'CQD'[Second UPN],
            'CQD'[First Is Caller],
            'CQD'[Start Time],
            'CQD'[End Time],
            'CQD'[Stream Direction]
        ), \"IsGrandTotalRowTotal\"
      ),
      __DS0FilterTable,
      __DS0FilterTable2,
      __DS0FilterTable3,
      __DS0FilterTable4,
      \"SumTotal_Audio_Stream_Duration__Minutes_\", CALCULATE(SUM('CQD'[Total Audio Stream Duration (Minutes)]))
    )

  VAR __DS0PrimaryWindowed = 
    TOPN(
      502,
      __DS0Core,
      [IsGrandTotalRowTotal],
      0,
      'CQD'[First UPN],
      1,
      'CQD'[Second UPN],
      1,
      'CQD'[First Is Caller],
      1,
      'CQD'[Start Time],
      1,
      'CQD'[End Time],
      1,
      'CQD'[Stream Direction],
      1
    )

EVALUATE
  __DS0PrimaryWindowed

ORDER BY
  [IsGrandTotalRowTotal] DESC,
  'CQD'[First UPN],
  'CQD'[Second UPN],
  'CQD'[First Is Caller],
  'CQD'[Start Time],
  'CQD'[End Time],
  'CQD'[Stream Direction]
      "
    }
  ],
  "serializerSettings": {"includeNulls": false}
}
"@ 
      }
    }
  }
  return $RequestBodyArray
}

function Get-TeamsCallRecords {
  $requestUrl = "https://api.powerbi.com/v1.0/myorg/datasets/$PowerBiDatasetId/executeQueries"
  Log "Collecting Teams call records with Power BI PowerShell module"
  $i = 0
  $PbiToken = Get-PowerBIAccessToken
  foreach ($RequestBody in $RequestBodyArray) {
    $PbiRequestBody = ""
    $PbiRequestBody = $RequestBody.RequestBody
    Start-Job -Name "PbiRequest" {
    $i++
    $PbiRequestBody = $RequestBody.RequestBody
    $PbiHttpRequestRaw = Invoke-RestMethod -ContentType 'application/json' -Method POST -Uri $using:requestUrl -Body $using:PbiRequestBody -Headers $using:PbiToken
    $PbiHttpRequest = $PbiHttpRequestRaw.Substring(3) | ConvertFrom-Json
    [System.Collections.Generic.List[psobject]]$TeamsCallRecordsTable = $PbiHttpRequest.results[0].tables[0].rows
    #Log "Teams Call Records Table Count: $($TeamsCallRecordsTable.Count)"
    If($TeamsCallRecordsTable.Count -gt 1) {
        $TeamsCallRecordsTable.RemoveAt(0)
        foreach ($TeamsCall in $TeamsCallRecordsTable) {
            if (($RequestBody.StreamDirection -eq $TeamsCall.'CQD[Stream Direction]') -or ($null -eq $TeamsCall.'CQD[Stream Direction]')) {
                $CallerUPN = If($TeamsCall.'CQD[First Is Caller]' -eq $true) {$TeamsCall.'CQD[First UPN]'} else {$TeamsCall.'CQD[Second UPN]'}
                $CalleeUPN = If($TeamsCall.'CQD[First Is Caller]' -eq $true) {$TeamsCall.'CQD[Second UPN]'} else {$TeamsCall.'CQD[First UPN]'}
                $TeamsCallRecord = [PSCustomObject]@{}
                $TeamsCallRecord = [PSCustomObject]@{
                    CallerNumber = ""
                    CalleeNumber = ""
                    CallerUPN = $CallerUPN
                    CalleeUPN = $CalleeUPN
                    CallerSIP = (($TeamsUsers.Where({$_.UserPrincipalName -eq $CallerUPN}).SipAddress) -split ":")[1]
                    CalleeSIP = (($TeamsUsers.Where({$_.UserPrincipalName -eq $CalleeUPN}).SipAddress) -split ":")[1]            
                    StartDateTime = Get-Date $TeamsCall.'CQD[Start Time]'
                    EndDateTime = Get-Date $TeamsCall.'CQD[End Time]'
                    'Duration(min)' = "$($TeamsCall.'[SumTotal_Audio_Stream_Duration__Minutes_]')"
                    CallType = "Teams"
                }
            $TeamsCallRecord      
          }
        }
      }
      #Log "  Teams Call Records query $i/$($RequestBodyArray.Count) complete" -BackgroundColor Yellow -ForegroundColor Black
    }
  }
  get-job -Name "PbiRequest" | Wait-Job | Out-Null
  [System.Collections.Generic.List[psobject]]$TeamsCallRecords = Receive-Job -Name "PbiRequest"
  remove-job -Name "PbiRequest"
  <#foreach ($Job in $JobTeamsCallRecords) {
    $TeamsCallRecords.Add($Job)
  }#>
  return $TeamsCallRecords
}
# get-job | remove-job
function Send-ReportEmail {
    Log "Sending Report via Email"
    if ($UseAnonymousSmtpAuth) {
        $SmtpUserName = "anonymous"
        $SecureSmtpPassword = ConvertTo-SecureString -String "anonymous" -AsPlainText -Force
        $SmtpCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($SmtpUserName, $SecureSmtpPassword)
        $SmtpSsl = $false
        If(-not $SmtpPort) {
            $SmtpPort = 25
        }
    }
    else {
        $SmtpSsl = $true
        If(-not $SmtpPort) {
            $SmtpPort = 587
        }
        $SecureSmtpPassword = ConvertTo-SecureString (Get-Content $SmtpPwFile)
        $SmtpCred = New-Object System.Management.Automation.PSCredential ($SmtpUserName, $SecureSmtpPassword)         
    }
    $MailParams = @{
        SmtpServer                 = $SmtpServer
        Port                       = $SmtpPort 
        UseSSL                     = $SmtpSsl
        Credential                 = $SmtpCred
        From                       = $SmtpUserName
        To                         = $EmailAddressList
        Subject                    = "Teams Call Records Report - $StartMonthNumber/$StartYear"
        Attachments                = "$ReportFolder\$ExportFileName"
    }
    Send-MailMessage @MailParams
}

################################################### Begin Script Here ####################################################################

#### Vars to Keep
Set-QueryDates
Set-ReportFiles 
$ClientSecretFileContents = Get-Content $ClientSecretFile
$ClientSecretTemp = ConvertTo-SecureString $ClientSecretFileContents
$ClientSecretPointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecretTemp)
$ClientSecret = [Runtime.InteropServices.Marshal]::PtrToStringAuto($ClientSecretPointer)
[Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ClientSecretPointer)

If($AdminUserName -and $AdminPwFile) {
  $SecureAdminPassword = ConvertTo-SecureString (Get-Content $AdminPwFile)
  $AdminCred = New-Object System.Management.Automation.PSCredential ($AdminUserName, $SecureAdminPassword)   
}

#### Vars to Keep

Log "Importing MicrosoftTeams PowerShell Module"
Import-Module MicrosoftTeams
Log "Importing MicrosoftPowerBIMgmt PowerShell Module"
Import-Module MicrosoftPowerBIMgmt

Log "Connecting to MicrosoftPowerBIMgmt PowerShell Module"
if ($PowerBiAppAuthentication) {
  $PowerBiCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AppId,$ClientSecret
  Connect-PowerBIServiceAccount -Tenant $TenantId -ServicePrincipal -Credential $PowerBiCred
}
elseif ($AdminCred) {
  Connect-PowerBIServiceAccount -Credential $AdminCred
}
else {
  Connect-PowerBIServiceAccount
}

Log "Connecting to MicrosoftTeams PowerShell Module"
if ($AdminCred) {
  Connect-MicrosoftTeams -Credential $AdminCred
}
else {
  Connect-MicrosoftTeams
}

[array]$TeamsUsers = Get-TeamsUsersWithNoLineUri
[System.Collections.Generic.List[psobject]]$UpnFilterLineUris = Get-UpnFilterLineUris $TeamsUsers

$Token = Connect-GraphApiWithClientSecret
$ClientSecret = $null
$apiUrls = "https://graph.microsoft.com/v1.0/communications/callRecords/getDirectRoutingCalls(fromDateTime=$StartYear-$StartMonthNumberString-$StartDayString,toDateTime=$EndYear-$EndMonthNumberString-$EndDayString)",`
"https://graph.microsoft.com/v1.0/communications/callRecords/getPstnCalls(fromDateTime=$StartYear-$StartMonthNumberString-$StartDayString,toDateTime=$EndYear-$EndMonthNumberString-$EndDayString)"

[System.Collections.Generic.List[psobject]]$PSTNCallRecords = @()
foreach ($apiUrl in $apiUrls) {
  [System.Collections.Generic.List[psobject]]$TempCallRecords = @()
  $TempCallRecords = Get-TeamsPstnCallRecords
  if ($TempCallRecords) {
    #$TempCallRecords
    $PSTNCallRecords.AddRange($TempCallRecords)
  }
}

[array]$RequestBodyArray = @()
$RequestBodyArray = Set-TeamsCqdPowerBiQueries
[System.Collections.Generic.List[psobject]]$TeamsCallRecords = @()
$TeamsCallRecords = Get-TeamsCallRecords

[System.Collections.Generic.List[psobject]]$AllCallRecords = @()
Log "Adding Teams Records to final collection"
if ($TeamsCallRecords) {
  $AllCallRecords.AddRange($TeamsCallRecords)
}
Log "Adding PSTN Records to final collection"
if ($PSTNCallRecords) {
  $AllCallRecords.AddRange($PSTNCallRecords)
}

if ($FirstHalfOfMonthRecords) {
  Log "Adding First Half of Month Records to final collection"
  $AllCallRecords.AddRange($FirstHalfOfMonthRecords)
}

if ($AllCallRecords) {
  Log "Sorting all records and removing duplicates"
  $AllCallRecords = $AllCallRecords | Sort-Object -Property StartDateTime -Unique
  
  Log "Exporting final collection with all records to $ReportFolder\$ExportFileName"
  $AllCallRecords | export-csv $ReportFolder\$ExportFileName -NoTypeInformation
  
  if ($EmailAddressList) {
    Send-ReportEmail
  }
}

Disconnect-MicrosoftTeams
Disconnect-PowerBIServiceAccount

If($DebugMode) {
  Stop-Transcript
}