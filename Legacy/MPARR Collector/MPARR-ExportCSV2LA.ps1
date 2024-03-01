<#
.SYNOPSIS
    Exports CSV file to Log Analytics.

.DESCRIPTION
    Exports CSV file to Log Analytics. Files of size bigger than 100MB will generate high memory and CPU usage.
    
.PARAMETER CustomerID
    Log Analytics workspace ID.

.PARAMETER SharedKey
    Workspace key (secret).

.PARAMETER FileName
    Path to the file that will be exported.

.PARAMETER TableName
    Name of the table data will be exported to. "_CL" will be added to the table name.

.PARAMETER TimeGeneratedColumnName
    CSV file coulmn that holds time values that should be passed as TimeGenerated. If not specified, current time will be used for TimeGenerated.

.NOTES
    Version 1.0
    Date: 2022-10-17
	
.NOTES to execute
Run this command: .\ExportCSV2LA.ps1 -FileName '.\Support\Product names and service plan identifiers for licensing.csv' -TableName "MSProducts"

#>
<#
HISTORY
Script      : ExportCSV2LA.ps1
Author      : G. Berdzik
Version     : 1.0.0
Description : The script exports a CSV file as table on Logs Analytics
2022-10-12		S. Zamorano		- Added laconfig.json file for configuration and decryption function
#>


param (
    [Parameter(Mandatory=$true)]
        [string] $FileName,
    [Parameter(Mandatory=$true)]
        $TableName,
    $TimeGeneratedColumnName
)

# Function to decrypt shared key
function DecryptSharedKey 
{
    param(
        [string] $encryptedKey
    )

    try {
        $secureKey = $encryptedKey | ConvertTo-SecureString -ErrorAction Stop  
    }
    catch {
        Write-Error "Workspace key: $($_.Exception.Message)"
        exit(1)
    }
    $BSTR =  [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureKey)
    $plainKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    $plainKey
}

$CONFIGFILE = "$PSScriptRoot\laconfig.json"
$json = Get-Content -Raw -Path $CONFIGFILE
[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
$EncryptedKeys = $config.EncryptedKeys
$AppClientID = $config.AppClientID
$ClientSecretValue = $config.ClientSecretValue
$TenantGUID = $config.TenantGUID
$TenantDomain = $config.TenantDomain
$CustomerID = $config.LA_CustomerID
$SharedKey = $config.LA_SharedKey
$CertificateThumb = $config.CertificateThumb
$OnmicrosoftTenant = $config.OnmicrosoftURL
if ($EncryptedKeys -eq "True")
{
    $SharedKey = DecryptSharedKey $SharedKey
    $ClientSecretValue = DecryptSharedKey $ClientSecretValue
	$CertificateThumb = DecryptSharedKey $CertificateThumb
}

#region functions

# Function to create the authorization signature
function BuildSignature 
{
    param(
        $customerId, $sharedKey, $date, $contentLength, $method, $contentType, $resource
    )

    $xHeaders = "x-ms-date:" + $date
    $stringToHash = $method + "`n" + $contentLength + "`n" + $contentType + "`n" + $xHeaders + "`n" + $resource

    $bytesToHash = [Text.Encoding]::UTF8.GetBytes($stringToHash)
    $keyBytes = [Convert]::FromBase64String($sharedKey)

    $sha256 = New-Object System.Security.Cryptography.HMACSHA256
    $sha256.Key = $keyBytes
    $calculatedHash = $sha256.ComputeHash($bytesToHash)
    $encodedHash = [Convert]::ToBase64String($calculatedHash)
    $authorization = 'SharedKey {0}:{1}' -f $customerId,$encodedHash
    return $authorization
}


# Function to create and post the request
function PostLogAnalyticsData
{
    param(
        $customerId, $sharedKey, $json, $logType
    )

    $body = ([System.Text.Encoding]::UTF8.GetBytes($json))
    $method = "POST"
    $contentType = "application/json"
    $resource = "/api/logs"
    $rfc1123date = [DateTime]::UtcNow.ToString("r")
    $contentLength = $body.Length
    $bsParams = @{
        customerId = $customerId
        sharedKey = $sharedKey
        date = $rfc1123date
        contentLength = $contentLength
        method = $method
        contentType = $contentType
        resource = $resource
    }
    $signature = BuildSignature @bsParams
    $uri = "https://" + $customerId + ".ods.opinsights.azure.com" + $resource + "?api-version=2016-04-01"

    $headers = @{
        "Authorization" = $signature;
        "Log-Type" = $logType;
        "x-ms-date" = $rfc1123date;
    }
    if ($TimeGeneratedColumnName -ne $null)
    {
        $headers["time-generated-field"] = $TimeGeneratedColumnName
    }

    try
    {
        $response = Invoke-WebRequest -Uri $uri -Method $method -ContentType $contentType -Headers $headers -Body $body -UseBasicParsing -ErrorAction Stop
    }
    catch
    {
        $response = New-Object psobject
        $response | Add-Member -MemberType NoteProperty -Name "StatusCode" -Value 400
        $msg = $_.ErrorDetails.Message | ConvertFrom-Json
        $errString = $_.Exception.Message + "`n" + $msg.Error + ": " + $msg.Message
        $response | Add-Member -MemberType NoteProperty -Name "Exception" -Value $errString
    }
    $response

}

#endregion

# *** Main code

if (Test-Path $FileName)
{
    Write-Host "Importing CSV file..."
    $data = Import-Csv -Path $FileName
    Write-Host "Calculating batch size..."

    $maximumBatchSize = 15MB
    $max = 0
    foreach($item in $data)
    {
        $size = [System.Text.Encoding]::UTF8.GetByteCount(($item | ConvertTo-Json))
        $max = [math]::Max($size, $max)
    }
    $BatchSize = [math]::Round($maximumBatchSize / $max, 0, [System.MidpointRounding]::ToZero)
   
    Write-Host "Starting export to LA..."
    $add = 0
    for ($i = 0; $i -le [math]::Floor($data.count/$BatchSize); $i++)
    {
        Write-Host "." -NoNewline
        $eventJSON = $data[($i*$BatchSize+$add)..($i*$BatchSize+$BatchSize)] | ConvertTo-Json -Depth 100

        $result = PostLogAnalyticsData -customerId $CustomerID -sharedKey $SharedKey -json $eventJSON -logType $TableName 
        if ([int]$result.StatusCode -ne 200)
        {
            Write-Host "Error exporting file $FileName to the Log Analytics. Exception: $($result.Exception)" -ForegroundColor Red
        }
        $add = 1
    }
}
else 
{
    Write-Host "File $FileName not found. Exiting."
    return
}
Write-Host "`nExport finished."
