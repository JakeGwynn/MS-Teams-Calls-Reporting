function Connect-GraphApiWithClientSecret {
    param(
        [Parameter(Mandatory=$true)] [string] $TenantId,
        [Parameter(Mandatory=$true)] [string] $AppId,
        [Parameter(Mandatory=$true)] [string] $ClientSecretFile
    )
    Write-Host "Connecting to Graph API to collect Teams PSTN call records"

    $ClientSecretFileContents = Get-Content $ClientSecretFile
    $ClientSecretTemp = ConvertTo-SecureString $ClientSecretFileContents
    $ClientSecretPointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecretTemp)
    $ClientSecret = [Runtime.InteropServices.Marshal]::PtrToStringAuto($ClientSecretPointer)
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ClientSecretPointer)

    $Body = @{    
        Grant_Type    = "client_credentials"
        Scope         = "https://graph.microsoft.com/.default"
        client_Id     = $AppId
        Client_Secret = $ClientSecret
        } 

    $ConnectGraph = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method POST -Body $Body
    $ClientSecret = $null
    return $ConnectGraph.access_token
}

function Get-CallRecords {
    param(
        [Parameter(Mandatory=$true)] [string] $Token,
        [Parameter(Mandatory=$true)] [string] $Url,
        [Parameter(Mandatory=$true)] [string] $CsvExportPath
    )
    $GetMoreRecords = $true
    while ($GetMoreRecords -eq $true) {
        $response = Invoke-RestMethod -Headers @{Authorization = "Bearer $Token"} -Uri $Url -Method Get
        if ($response."@odata.nextLink") {
            $Url = $response."@odata.nextLink"
        } else {
            $GetMoreRecords = $false
        }
        $response.Value | Export-Csv -Path $CsvExportPath -NoTypeInformation -Append
        $response | fl
    }
}

$Token = Connect-GraphApiWithClientSecret -TenantId "04b9e073-f7cf-4c95-9f91-e6d55d5a3797" -AppId "54df59eb-521a-44b9-b93e-907815a23adb" -ClientSecretFile "C:\temp\ClientSecret.txt" 

$EndYear = Get-Date -Format "yyyy"
$EndMonthNumber = Get-Date -Format "MM"
$EndMonthNumberString = $EndMonthNumber.ToString()
$EndDay = Get-Date -Format "dd"
$EndDayString = $EndDay.ToString()

$StartYear = (Get-Date).AddDays(-28).ToString("yyyy")
$StartMonthNumber = (Get-Date).AddDays(-28).ToString("MM")
$StartMonthNumberString = $StartMonthNumber.ToString()
$StartDay = (Get-Date).AddDays(-28).ToString("dd")
$StartDayString = $StartDay.ToString()

$apiUrls = "https://graph.microsoft.com/v1.0/communications/callRecords/getDirectRoutingCalls(fromDateTime=$StartYear-$StartMonthNumberString-$StartDayString,toDateTime=$EndYear-$EndMonthNumberString-$EndDayString)",`
"https://graph.microsoft.com/v1.0/communications/callRecords/getPstnCalls(fromDateTime=$StartYear-$StartMonthNumberString-$StartDayString,toDateTime=$EndYear-$EndMonthNumberString-$EndDayString)"

foreach ($apiUrl in $apiUrls) {
    Get-CallRecords -Token $Token -Url $apiUrl -CsvExportPath "C:\temp\TeamsCallRecordsTest.csv"
}
