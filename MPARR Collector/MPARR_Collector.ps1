<#
.SYNOPSIS
    Exports Office 365 Compliance data to Log Analytics and / or file.
.DESCRIPTION
    Exports Office 365 Compliance data to Log Analytics. Optionaly data from files can be created. 
    
    Script uses configuration data file 'laconfig.json' to connect to Azure resources. Config file should be placed in the same directory as the script file.
    Secrets in the file can be encrypted with DPAPI mechanism. Check examples to learn how encrypt secrets.

    Syntax of the laconfig.json is as follows:

        {
            "EncryptedKeys":  "True",
            "AppClientID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
            "ClientSecretValue": "zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzz",
            "TenantGUID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
            "TenantDomain": "your.tenant.domain",
            "LA_CustomerID":  "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
            "LA_SharedKey":  "zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzz",
            "CertificateThumb": "",
	        "OnmicrosoftURL": "your_tenant.onmicrosoft.com",
	        "RMSLogs": "c:\\APILogs\\RMSLogs\\",
	        "OutPutLogs": "c:\\APILogs",
            "Cloud": "Commercial"
        }

    EncryptedKeys - possible values True/False. If 'True', 'ClientSecretValue' and 'LA_SharedKey' should be encrypted.
    AppClientID - client app ID
    ClientSecretValue - secret for the app
    TenantGUID - GUID of the tenant
    TenantDomain - tenant FQDN
    LA_CustomerID - Log Analytics workspace ID
    LA_SharedKey - Log Analytics workspace key
    Cloud - optional parameter to specify Microsoft cloud. If not specified, defaults to 'Commercial'. Possible values are:
        Commercial - Commercial Cloud
        GCC - Government Community Cloud
        GCCH - Government Community High Cloud
        DOD - Department of Defense Cloud

    
    List of the content types script is able to query (i.e. Audit.AzureActiveDirectory, Audit.Exchange, DLP.All, etc.) depends on 'schema.json' file. You can add
    new content types as these become available. 
    
    The same config file is responsible for the filter list. Filter parameters are created dynamically. Regex match is used as filtering engine.

.PARAMETER UseCustomParameters
    Switch to enable custom parameters regarding start time, end time and output file name.

.PARAMETER pStartTime
    Start time of data to be exported.

.PARAMETER pEndTime
    End time of data to be exported.

.PARAMETER pFilenameCode
    Name of the export file(s).

.PARAMETER ExportToFileOnly
    Switch to disable export to Log Analytics.

.PARAMETER ExportWithFile
    Switch to export to Log Analytics creating output files at the same time.

.EXAMPLE
    mparr_collector.ps1 -FilterAuditSharepoint "Accessed"

    Exports compliance data to LA with filtering enabled for Sharepoint data. Please note that list of the filters depends on the 'schema.json' content.

.EXAMPLE
    "your_secret" | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString

    Encrypts secret string. Resulting string should be pasted to the "laconfig.json" file.
    When enabling secret encryption, both secrets are required be encrypted - "ClientSecretValue" and "LA_SharedKey" from the "laconfig.json" file 
    (replace "your_secret" with these values and put results into the corresponding fields of the config file).
    Value of "EncryptedKeys" must be set to "True".

.NOTES
    Version 4.12
    Current version - 21.09.2023

     Original Version:        2.7
              Author:         Walid Elmorsy - Principal Program Manager - Compliance CAT team.
                              Brendon Lee - Senior Program Manager - Compliance CAT team
              Creation Date:  11/11/2021
              Purpose :       Collect Microsoft 365 Compliance Audit Log Activity Information via Office 365 Management API endpoints, and export to JSON files (For testing purposes only)
#> 

<#
HISTORY
  2022-04-01    S. Carstens  - make code more readable, structure, added parameter for start/end time
  ...
  2022-09-13    G.Berdzik   - Fixes related to timestamp.
  2022-09-16    G.Berdzik   - Fixes related to secret encryption.
  2022-09-21    G.Berdzik   - Fixed issue with encoding. Improved help. Added warning for logs older than 2 days.
  2022-09-22	S.Zamorano  - Fixed Azure AD Filter
  2022-09-27    G.Berdzik   - Change to Version 3. Added support for direct export to LA (no files required).
  2022-11-04    G.Berdzik   - Added 'EventCreationTime_t' column storing original 'CreationTime' value. Batch size changed to 500 elements from 100.
  2022-11-14    G.Berdzik   - Change to Version 4. Added support for 'schemas.json', cloud type (designed by S.Zamorano)
  2023-02-08    G.Berdzik	- File name change, change in filtering based on 'schemas.json'
  2023-03-10    G.Berdzik	- Added support for 'OutputLogs' setting in 'laconfig.json'.
  2023-09-21    G.Berdzik	- Fixes related to timeout connection on the first execution
#>

#------------------------------------------------------------------------------  
#  
#   
# This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.  
# THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
# INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  
# We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code 
# form of the Sample Code, provided that You agree: (i) to not use Our name, logo, or trademarks to market Your software product in 
# which the Sample Code is embedded; (ii) to include a valid copyright notice on Your software product in which the Sample Code is 
# embedded; and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits, 
# including attorneys fees, that arise or result from the use or distribution of the Sample Code.
#  
#------------------------------------------------------------------------------ 


#
# UseTimeParameters - if given, the provided start/end times are used instead of calculated times from timestamp file
# pFilenameCode - Code that will used in filename instead of date
#

[CmdletBinding(DefaultParameterSetName = "None")]
param(
    [Parameter(ParameterSetName="CustomParams")] 
    [Parameter(ParameterSetName="CustomParams1")] 
        [switch]$UseCustomParameters,

    [Parameter(ParameterSetName="CustomParams", Mandatory=$true)] 
        [datetime]$pStartTime,

    [Parameter(ParameterSetName="CustomParams", Mandatory=$true)] 
        [datetime]$pEndTime,

    [Parameter(ParameterSetName="CustomParams")] 
    [Parameter(ParameterSetName="CustomParams1", Mandatory=$true)] 
        [string]$pFilenameCode,

    [Parameter()] 
        [switch]$ExportToFileOnly,

    [Parameter()] 
        [switch]$ExportWithFile

#    [Parameter()]
#        [string]$OutputPath = "C:\APILogs\"
)

DynamicParam 
{
    # create dynamic parameters based on 'schema.json' entries set to 'True'
    $filePath = "$PSScriptRoot\schemas.json"
    if (Test-Path -Path $filePath)
    {
        $RunTimeDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParamAttribute = New-Object System.Management.Automation.ParameterAttribute
        $AttributeCollection.Add($ParamAttribute)

        $json = Get-Content -Raw -Path $filePath
        [PSCustomObject]$schemas = ConvertFrom-Json -InputObject $json
        foreach ($item in $schemas.psobject.Properties)
        {
            if ($schemas."$($item.Name)" -eq "True")
            {
                $ParameterName = "Filter" + $item.Name.Replace('.', '')
                $RunTimeParam = New-Object System.Management.Automation.RuntimeDefinedParameter($ParameterName, [string], $AttributeCollection)
                $RunTimeDictionary.Add($ParameterName, $RunTimeParam)
            }
        }
        return $RunTimeDictionary
    }
}


end 
{
#region Functions

# Create Function to Check content availability in all content types (inlcuding all pages) 
# and store results in $Subscription variable, also build the URI list in the correct format
function buildLog($BaseURI, $Subscription, $tenantGUID, $OfficeToken)
{
    try {
        #
        # if using custom value for start/end 
        #
        if ($UseCustomParameters)
        {
            $strt = $pStartTime.ToString("yyyy-MM-ddTHH:mm:ss")
            $end  = $pEndTime.ToString("yyyy-MM-ddTHH:mm:ss")
        }
        else
        {
            $strt = $startTime
            $end  = $endTime
        }

        Write-Verbose " Start = $strt"
        Write-Verbose " End   = $end"

        $URIstring = "$BaseURI/content?contentType=$Subscription&startTime=$strt&endTime=$end&PublisherIdentifier=$TenantGUID"
        Write-Host " "
        Write-Verbose " URI    : $uristring"

        $Log = Invoke-WebRequest -Method GET -Headers $OfficeToken `
               -Uri "$BaseURI/content?contentType=$Subscription&startTime=$strt&endTime=$end&PublisherIdentifier=$TenantGUID" `
               -UseBasicParsing -ErrorAction Stop
        
    } 
    catch {
        write-host -ForegroundColor Red "Invoke-WebRequest command has failed"
        Write-host $error[0]
        return
    }

	$TotalContentPages = @()
    #Try to find if there is a NextPage in the returned URI
    if ($Log.Headers.NextPageUri) 
    {
        $NextContentPage = $true
        $NextContentPageURI = $Log.Headers.NextPageUri
		if ($NextContentPageURI -is [array])
		{
			$NextContentPageURI = $Log.Headers.NextPageUri[0]
		}
		$oldURI = $NextContentPageURI

        Write-Verbose " NextPage is present: $NextContentPageURI"

        while ($NextContentPage -ne $false)
        {
			Write-Verbose "Retrieving page nr $($TotalContentPages.Count + 1)"
            $ThisContentPage = Invoke-WebRequest -Headers $OfficeToken -Uri $NextContentPageURI -UseBasicParsing
            $TotalContentPages += $ThisContentPage

            if ($ThisContentPage.Headers.NextPageUri)
            {
                $NextContentPage = $true    
            }
            else
            {
                $NextContentPage = $false
            }
            $NextContentPageURI = $ThisContentPage.Headers.NextPageUri
			if ($NextContentPageURI -is [array])
			{
				$NextContentPageURI = $Log.Headers.NextPageUri[0]
			}
			if ($oldURI -eq $NextContentPageURI)
			{
				$NextContentPage = $false
			}
			$oldURI = $NextContentPageURI
        }
    } 
    $TotalContentPages += $Log

    Write-Host -ForegroundColor Green "OK"
    Write-Host "***"
    return $TotalContentPages
}

#Generate the correct URI format and export  logs
function FetchData($TotalContentPages, $Officetoken, $Subscription)
{
    # Changed from "-gt 2" to "-gt 0"
    if ($TotalContentPages.content.length -gt 0)
    {
        $uris = @()
        $pages = $TotalContentPages.content.split(",")
        
        foreach($page in $pages)
        {
            if ($page -match "contenturi") {
                $uri = $page.split(":")[2] -replace """"
                $uri = "https:$uri"
                $uris += $uri
            }
        }

        $Logdata = @()
        $filterName = "Filter" + $Subscription.Replace('.', '')
        foreach($uri in $uris)
        {

            Write-Verbose " uri:$uri"

            try {

                # check for token expiration
                if ($tokenExpiresOn.AddMinutes(5) -lt (Get-Date))
                {
                    Write-Host "Refreshing access token..."
                    GetAuthToken
                }

                $result = Invoke-RestMethod -Uri $uri -Headers $Officetoken -Method Get
                if ($script:PSBoundParameters.ContainsKey($filterName))
                {
                    Write-Verbose "Applying filter '$($script:PSBoundParameters[$filterName])' on $($filterName)."
                    if ($schemas.$filterName -eq "NotContains")
                    {
                        $Logdata += $result | Where-Object {$_.Operation -notmatch $($script:PSBoundParameters[$filterName])}
                    }
                    else 
                    {
                        $Logdata += $result | Where-Object {$_.Operation -match $($script:PSBoundParameters[$filterName])}
                    }
                }
                else 
                {
                    $Logdata += $result
                }
            } 
            catch {
                write-host -ForegroundColor Red "ERROR"
                Write-host $error[0]
                return
            }      
        }
        $Logdata 
        write-host -ForegroundColor Green "OK"
    } 
    else {
        Write-Host -ForegroundColor Yellow "Nothing to output"
    }
}

# Verify output directory exists
function CheckOutputDirectory($OutputPath)
{
    # path should not be on root drive
    if ($OutputPath.EndsWith(":\"))
    {
        Write-Host -ForegroundColor Red "Path should not be on root drive. Exiting."
        exit(1)
    }

    # verify folder exists, if not try to create it
    if (!(Test-Path($OutputPath)))
    {
        Write-Host -ForegroundColor Yellow ">> Warning: '$OutputPath' does not exist. Creating one now..."
        Write-host -ForegroundColor Gray "Creating '$OutputPath': " -NoNewline
        try
        {
            New-Item -ItemType "directory" -Path $OutputPath -ErrorAction Stop | Out-Null
            Write-Host -ForegroundColor Green "Path '$OutputPath' has been created successfully"
        } catch {
            write-host -ForegroundColor Red "FAILED to create '$OutputPath'"
            Write-Host -ForegroundColor Red ">> ERROR: The directory '$OutputPath' could not be created."
            Write-Host -ForegroundColor Red $error[0]
        }
    }
    else{
        Write-Host -ForegroundColor Green "Path '$OutputPath' already exists"
    }
}

#Check folder path and construct file names
function GetFileName($Date, $Subscription, $OutputPath)
{
    if ($UseCustomParameters)
    {
        Write-Verbose " using custom parameter for filename"
        $JSONfilename = ($Subscription + "_" + $pFilenameCode + ".json")
    }
    else {
        Write-Verbose " using default for filename"
        $JSONfilename = ($Subscription + "_" + $Date + ".json")
       
    }

    Write-Verbose " filename: $jsonfilename"
    return $OutputPath + $JSONfilename
}

# get access token
function GetAuthToken
{
    $body = @{grant_type="client_credentials";resource=$APIResource;client_id=$AppClientID;client_secret=$ClientSecretValue}
    Write-Host -ForegroundColor Blue -BackgroundColor white "Obtaining authentication token..." -NoNewline
    try{
        $oauth = Invoke-RestMethod -Method Post -Uri "$loginURL/$TenantDomain/oauth2/token?api-version=1.0" -Body $body -ErrorAction Stop
        $script:tokenExpiresOn = ([DateTime]('1970,1,1')).AddSeconds($oauth.expires_on).ToLocalTime()
        $script:OfficeToken = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}
        Write-Host -ForegroundColor Green "Authentication token obtained"
    } catch {
        write-host -ForegroundColor Red "FAILED"
        write-host -ForegroundColor Red "Invoke-RestMethod failed."
        Write-host -ForegroundColor Red $error[0]
        exit
    }
}


function Export-Logs
{
    Write-Verbose " enter export-logs" 

    # Access token Request and Retrieval 
    GetAuthToken
    
    #create new Subscription (if needed)

    Write-Host -ForegroundColor Blue -BackgroundColor white "Creating Subscriptions...."

    foreach($Subscription in $Subscriptions){
        Write-Host -ForegroundColor Cyan "$Subscription : " -NoNewline
        try { 
            $response = Invoke-WebRequest -Method Post -Headers $OfficeToken `
                                          -Uri "$BaseURI/start?contentType=$Subscription" `
                                          -UseBasicParsing -ErrorAction Stop
        } catch {
            if(($error[0] | ConvertFrom-Json).error.message -like "The subscription is already enabled*"){
                Write-host -ForegroundColor Yellow "Subscription already Exists"
            } 
            else {
                write-host -ForegroundColor Red "Failed to create a subscription for $Subscription"
                Write-host -Foregroundcolor Red $error[0]
            }
        }
    }

    #Check subscription status
    $CheckSubTemp = Invoke-WebRequest -Headers $OfficeToken -Uri "$BaseURI/list" -UseBasicParsing
    Write-Host -ForegroundColor Blue -BackgroundColor white "Subscription Content Status"
    $CheckSub = $CheckSubTemp.Content | convertfrom-json
    $CheckSub | ForEach-Object {write-host $_.contenttype "--->" -nonewline; write-host $_.status -ForegroundColor Green}

    #Collecting and Exporting Log data
    Write-Host -ForegroundColor Blue -BackgroundColor white "Checking output folder path"
    
    Write-Verbose " calculated filename: $JSONfileName"

    Write-Host -ForegroundColor Blue -BackgroundColor white "Collecting and Exporting Log data"
    foreach($Subscription in $Subscriptions)
    {    
        Write-Host -ForegroundColor Cyan "`n-> Collecting log data from '" -NoNewline
        Write-Host -ForegroundColor White -BackgroundColor DarkGray $Subscription -NoNewline
        Write-Host -ForegroundColor Cyan "': " -NoNewline

        # check for token expiration
        if ($tokenExpiresOn.AddMinutes(5) -lt (Get-Date))
        {
            Write-Host "Refreshing access token..."
            GetAuthToken
        }

        $logs = buildLog $BaseURI $Subscription $TenantGUID $OfficeToken

        $JSONfileName = getFileName $Date $Subscription $outputPath
    
        $output = FetchData $logs $OfficeToken $Subscription
        if ($ExportToFileOnly)
        {
            $output | ConvertTo-Json -Depth 100 | Set-Content -Encoding UTF8 $JSONfilename
            Write-host -ForegroundColor Cyan "---> Exporting log data to '" -NoNewline
            Write-Host -ForegroundColor White -BackgroundColor DarkGray $JSONfilename -NoNewline
            Write-Host -ForegroundColor Cyan "': " -NoNewline
    
        }
        elseif ($ExportWithFile)
        {
            $output | ConvertTo-Json -Depth 100 | Set-Content -Encoding UTF8 $JSONfilename
            Write-host -ForegroundColor Cyan "---> Exporting log data to '" -NoNewline
            Write-Host -ForegroundColor White -BackgroundColor DarkGray $JSONfilename -NoNewline
            Write-Host -ForegroundColor Cyan "': " -NoNewline
            Publish-LogAnalytics $output $Subscription
        }
        else 
        {
            Publish-LogAnalytics $output $Subscription
        }
    }
}

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

    $body = [System.Text.Encoding]::UTF8.GetBytes($json)
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
        "time-generated-field" = "CreationTime"
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

function Publish-LogAnalytics
{
    param (
        $objFromJson,
        $Subscription
    )

    Write-Host "Starting export to LA..."
    $list = New-Object System.Collections.ArrayList
    $LogName = $Subscription.Replace(".", "")

    $count = 0
    $elements = 0
    foreach ($item in $objFromJson)
    {
        $elements++
        $count++
        $item | Add-Member -MemberType NoteProperty -Name "EventCreationTime" -Value ($item.CreationTime)
        [void]$list.Add($item)
        if ($elements -ge $BatchSize)
        {
            $elements = 0
            $eventJSON = $list | ConvertTo-Json -Depth 100
            $result = PostLogAnalyticsData -customerId $CustomerID -sharedKey $SharedKey -json $eventJSON -logType $LogName 
            if ([int]$result.StatusCode -ne 200)
            {
                $count -= $BatchSize 
                Write-Host "Error exporting to the Log Analytics. Exception: $($result.Exception)" -ForegroundColor Red
                $errorFile = $OutputPath + "Error_" + $Subscription + "_" + $Date + ".json"
                $eventJSON | Set-Content -Encoding utf8 -Path $errorFile
                Write-Host "Failed records were saved to the $errorFile file. Please investigate them and import with ExportAIPData2LA script."
            }
            $list.Clear()
            $list.TrimToSize()            
        }
    }
    if ($list.Count -gt 0)
    {
        $eventJSON = $list | ConvertTo-Json -Depth 100
        $result = PostLogAnalyticsData -customerId $CustomerID -sharedKey $SharedKey -json $eventJSON -logType $LogName 
        if ([int]$result.StatusCode -ne 200)
        {
            $count -= $elements 
            Write-Host "Error exporting to the Log Analytics. Exception: $($result.Exception)" -ForegroundColor Red
            $errorFile = $OutputPath + "Error_" + $Subscription + "_" + $Date + ".json"
            $eventJSON | Set-Content -Encoding utf8 -Path $errorFile
            Write-Host "Failed records were saved to the $errorFile file. Please investigate them and import with ExportAIPData2LA script."
        }
    }
    Write-Host "$count elements exported for $Subscription."
}

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

#endregion

# ******************************************************************
#Requires -Version 7.0

#region Main code

#API Endpoint URLs ---> Don't Update anything here
$CLOUDVERSIONS = @{
    Commercial = "https://manage.office.com"
    GCC = "https://manage-gcc.office.com"
    GCCH = "https://manage.office365.us"
    DOD = "https://manage.protection.apps.mil"
}

# Script variables 01  --> Update everything in this section:
$BatchSize = 500
$CONFIGFILE = "$PSScriptRoot\laconfig.json"   
$SCHEMASFILE = "$PSScriptRoot\schemas.json"   

if ($pFilenameCode -ne "" -and -not $ExportToFileOnly)
{
    Write-Warning "Custom output file name is supported only with 'ExportToFileOnly' parameter."
    exit(1)
}

# Read config file
if (-not (Test-Path -Path $CONFIGFILE))
{
    Write-Error "Missing config file."
    exit(1)
}
$json = Get-Content -Raw -Path $CONFIGFILE
[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
$EncryptedKeys = $config.EncryptedKeys
$AppClientID = $config.AppClientID
$ClientSecretValue = $config.ClientSecretValue
$TenantGUID = $config.TenantGUID
$TenantDomain = $config.TenantDomain
$CustomerID = $config.LA_CustomerID
$SharedKey = $config.LA_SharedKey
$Cloud = $config.Cloud 
if ($EncryptedKeys -eq "True")
{
    $SharedKey = DecryptSharedKey $SharedKey
    $ClientSecretValue = DecryptSharedKey $ClientSecretValue
}
$APIResource = $CLOUDVERSIONS.Commercial
if ($Cloud -ne $null)
{
    $APIResource = $CLOUDVERSIONS["$Cloud"]
    Write-Host "Connecting to $Cloud cloud."
}

$OutputPath = $config.OutPutLogs
if ($OutputPath -eq "")
{
    $OutputPath = "C:\APILogs\"
    Write-Host "'OutputLogs' has no value. Default value was assigned: $OutputPath." -ForegroundColor Yellow
}
if (-not $OutputPath.EndsWith("\"))
{
    $OutputPath += "\"
}
CheckOutputDirectory $OutputPath


# Read schemas file
$Subscriptions = @('Audit.AzureActiveDirectory','Audit.Exchange','Audit.SharePoint','Audit.General','DLP.All')
if (-not (Test-Path -Path $SCHEMASFILE))
{
    Write-Host "Schemas file is missing. Default list of subscriptions will be used."
}
else 
{
    $Subscriptions = @()
    $json = Get-Content -Raw -Path $SCHEMASFILE
    [PSCustomObject]$schemas = ConvertFrom-Json -InputObject $json
    foreach ($item in $schemas.psobject.Properties)
    {
        if ($schemas."$($item.Name)" -eq "True")
        {
            $Subscriptions += $item.Name
        }
    }
    Write-Host "Subscriptions list: $Subscriptions"    
}

# Script variables 02  ---> Don't Update anything here:
$loginURL = "https://login.microsoftonline.com/"
$BaseURI = "$APIResource/api/v1.0/$TenantGUID/activity/feed/subscriptions"

$Date = (Get-date).AddDays(-1)
$Date = $Date.ToString('MM-dd-yyyy_hh-mm-ss')

#region Timestamp/1
$timestampFile = $OutputPath + "timestamp.json"
# read startTime from the file
if (-not (Test-Path -Path $timestampFile))
{
    # if file not present create new value
    $startTime = (Get-Date).AddHours(-23).ToString("yyyy-MM-ddTHH:mm:ss")
}
else 
{
    $json = Get-Content -Raw -Path $timestampFile
    [PSCustomObject]$timestamp = ConvertFrom-Json -InputObject $json
    $startTime = $timestamp.startTime.ToString("yyyy-MM-ddTHH:mm:ss")   
    # check if startTime greater than 7 days (7 days is max value)
    if ((New-TimeSpan -Start $startTime -End ([datetime]::Now)).TotalDays -gt 7)
    {
        $startTime = (Get-Date).AddDays(-7).AddMinutes(30).ToString("yyyy-MM-ddTHH:mm:ss")
        Write-Host "StartTime is older than 7 days. Setting to the correct value: $startTime" -ForegroundColor Yellow
        Write-Host "Records with CreationTime older than two days will be ingested with current time for the TimeGenerated column!" -ForegroundColor Red
    }
}
$endTime = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
# check if difference between start and end times bigger than 24 hours 
if ((New-TimeSpan -Start $startTime -End $endTime).TotalHours -gt 24)
{
    $endTime = ([datetime]$startTime).AddHours(23).ToString("yyyy-MM-ddTHH:mm:ss")
    Write-Host "Timeframe based on StartTime is bigger than 24 hours. Setting to the correct value: $startTime" -ForegroundColor Yellow
    if ((New-TimeSpan -Start $startTime -End ([datetime]::Now)).TotalDays -gt 2)
    {
        Write-Host "Records with CreationTime older than two days will be ingested with current time for the TimeGenerated column!" -ForegroundColor Red
    }
}
$timestamp = @{"startTime" = $endTime}
ConvertTo-Json -InputObject $timestamp | Out-File -FilePath $timestampFile -Force
#endregion

Export-Logs

#endregion
}
