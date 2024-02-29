<#PSScriptInfo

.VERSION 2.0.5

.GUID 883af802-165c-4704-b4c1-352686c02f01

.AUTHOR 
https://www.linkedin.com/in/grzegorzberdzik/; Grzegorz Berdzik
https://www.linkedin.com/in/profesorkaz/; Sebastian Zamorano

.COMPANYNAME 
Microsoft Purview Advanced Rich Reports

.TAGS 
#Microsoft365 #M365 #MPARR #MicrosoftPurview #PowerBI #LogsAnalytics #Sentinel #Reporting #Dashboards #InformationProtection #MIP #Labels #DLP
#Webinar #PowerBI #DataAnalisys #Data #DataInsights #API #Office365ManagementAPI #YouTube #DataExfiltration

.PROJECTURI 
https://aka.ms/MPARR-GitHub; https://aka.ms/MPARR-LinkedIn; https://aka.ms/MPARR-YouTube 

.RELEASENOTES
The MIT License (MIT)
Copyright (c) 2015 Microsoft Corporation
Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

#>

<# 

.DESCRIPTION 
 Exports CSV file to Log Analytics. Files of size bigger than 100MB will generate high memory and CPU usage. 

#>

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
Version     : 2.0.5
Description : The script exports a CSV file as table on Logs Analytics

.NOTES

12-10-2022		S. Zamorano		- Added laconfig.json file for configuration and decryption function
12-02-2024		S. Zamorano		- Version released with EvenHub export 
01-03-2024		S. Zamorano		- Public release
#>

using module "ConfigFiles\MPARRUtils.psm1"
param (
	[Parameter()] 
        [switch]$ExportToJSONFileOnly,
	[Parameter()] 
        [switch]$ExportToEventHub,
    [Parameter(Mandatory=$true)]
        [string] $FileName,
    [Parameter(Mandatory=$true)]
        $TableName,
    $TimeGeneratedColumnName
)

function CheckPowerShellVersion
{
    # Check PowerShell version
    Write-Host "`nChecking PowerShell version... " -NoNewline
    if ($Host.Version.Major -gt 5)
    {
        Write-Host "`t`t`t`tPassed!" -ForegroundColor Green
    }
    else
    {
        Write-Host "Failed" -ForegroundColor Red
        Write-Host "`tCurrent version is $($Host.Version). PowerShell version 7 or newer is required."
        exit(1)
    }
}

function ValidateConfigurationFile
{
	#Validate laconfig.json that manage the configuration for connections
	$MPARRConfiguration = "$PSScriptRoot\ConfigFiles\laconfig.json"
	
	if (-not (Test-Path -Path $MPARRConfiguration))
	{
		Write-Host "`n##########################################################################################" -ForeGroundColor Yellow
		Write-Host "`nThe laconfig.json file is missing. Check if you are using the right path or execute MPARR_Setup.ps1 first."
		Write-Host "`nThe laconfig.json is required to continue, if you want to export the data without having MPARR installed, please execute:" -NoNewLine
		Write-Host ".\MPARR-PurviewSensitivityLabels.ps1 -ExportToFileOnly -ManualConnection" -ForeGroundColor Green
		Write-Host "`n##########################################################################################" -ForeGroundColor Yellow
		Write-Host "`n"
		if($ExportToFileOnly)
		{
			if($ManualConnection)
			{
				return
			}else
			{
				Write-Host "`n##########################################################################################" -ForeGroundColor Yellow
				Write-Host "`nThe laconfig.json is required to continue, if you want to export the data without having MPARR installed, please execute:" -NoNewLine
				Write-Host ".\MPARR-PurviewSensitivityLabels.ps1 -ExportToFileOnly -ManualConnection" -ForeGroundColor Green
				Write-Host "`n##########################################################################################" -ForeGroundColor Yellow
				Write-Host "`n"
				Write-Host "`n"
				exit
			}
		}else
		{
			exit
		}
	}else
	{
		#If the file is present we check if something is not correctly populated
		$CONFIGFILE = "$PSScriptRoot\ConfigFiles\laconfig.json"
		$json = Get-Content -Raw -Path $CONFIGFILE
		[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
		
		$EncryptedKeys = $config.EncryptedKeys
		$AppClientID = $config.AppClientID
		$WLA_CustomerID = $config.LA_CustomerID
		$WLA_SharedKey = $config.LA_SharedKey
		$CertificateThumb = $config.CertificateThumb
		$OnmicrosoftTenant = $config.OnmicrosoftURL
		
		if($AppClientID -eq "") { Write-Host "Application Id is missing! Update the laconfig.json file and run again" -ForeGroundColor Red; exit }
		if($WLA_CustomerID -eq "")  { Write-Host "Logs Analytics workspace ID is missing! Update the laconfig.json file and run again" -ForeGroundColor Red; exit }
		if($WLA_SharedKey -eq "")  { Write-Host "Logs Analytics workspace key is missing! Update the laconfig.json file and run again" -ForeGroundColor Red; exit }
		if($CertificateThumb -eq "")  { Write-Host "Certificate thumbprint is missing! Update the laconfig.json file and run again" -ForeGroundColor Red; exit }
		if($OnmicrosoftTenant -eq "")  { Write-Host "Onmicrosoft domain is missing! Update the laconfig.json file and run again" -ForeGroundColor Red; exit }
		
		Write-Host "Configuration file validation..." -NoNewLine
		Write-Host "`t`t`tPassed!" -ForeGroundColor Green
		Start-Sleep -s 1
	}
}

function CheckPrerequisites
{
    CheckPowerShellVersion
	ValidateConfigurationFile
}

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

function EventHubConnection
{
	$CONFIGFILE = "$PSScriptRoot\ConfigFiles\laconfig.json"
	$json = Get-Content -Raw -Path $CONFIGFILE
	[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
	
	$EncryptedKeys = $config.EncryptedKeys
	$AppClientID = $config.AppClientID
	$ClientSecretValue = $config.ClientSecretValue
	$TenantGUID = $config.TenantGUID
	$EventHubNamespace = $config.EventHubNamespace
	$EventHub = $config.EventHub
	
	if ($EncryptedKeys -eq "True")
	{
		$ClientSecretValue = DecryptSharedKey $ClientSecretValue
	}
	$script:EventHubInstance = [MPARREventHub]::new($TenantGUID, $EventHubNamespace, $EventHub, $AppClientID, $ClientSecretValue)
	Write-Host "EventHub connection...`t" -NoNewLine
	Write-Host "`t`t`t`tPassed!" -ForeGroundColor Green
}

function CheckExportOption
{
	$CONFIGFILE = "$PSScriptRoot\ConfigFiles\laconfig.json"
	$json = Get-Content -Raw -Path $CONFIGFILE
	[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
	$ExportOptionEventHub = $config.ExportToEventHub
	
	if($ExportToEventHub)
	{
		$ExportOptionEventHub = "True"
	}
	
	return $ExportOptionEventHub
}

function Build-Signature ($customerId, $sharedKey, $date, $contentLength, $method, $contentType, $resource) 
{
    # ---------------------------------------------------------------   
    #    Name           : Build-Signature
    #    Value          : Creates the authorization signature used in the REST API call to Log Analytics
    # ---------------------------------------------------------------

	#Original function to Logs Analytics
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

function Post-LogAnalyticsData($body, $LogAnalyticsTableName) 
{
    # ---------------------------------------------------------------   
    #    Name           : Post-LogAnalyticsData
    #    Value          : Writes the data to Log Analytics using a REST API
    #    Input          : 1) PSObject with the data
    #                     2) Table name in Log Analytics
    #    Return         : None
    # ---------------------------------------------------------------
    
	#Read configuration file
	$CONFIGFILE = "$PSScriptRoot\ConfigFiles\laconfig.json"
	$json = Get-Content -Raw -Path $CONFIGFILE
	[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
	
	$EncryptedKeys = $config.EncryptedKeys
	$WLA_CustomerID = $config.LA_CustomerID
	$WLA_SharedKey = $config.LA_SharedKey
	if ($EncryptedKeys -eq "True")
	{
		$WLA_SharedKey = DecryptSharedKey $WLA_SharedKey
	}

	# Your Log Analytics workspace ID
	$LogAnalyticsWorkspaceId = $WLA_CustomerID

	# Use either the primary or the secondary Connected Sources client authentication key   
	$LogAnalyticsPrimaryKey = $WLA_SharedKey 
	
    #Step 0: sanity checks
    if($body -isnot [array]) {return}
    if($body.Count -eq 0) {return}

    #Step 1: convert the PSObject to JSON
    $bodyJson = $body | ConvertTo-Json

    #Step 2: get the UTF8 bytestream for the JSON
    $bodyJsonUTF8 = ([System.Text.Encoding]::UTF8.GetBytes($bodyJson))

    #Step 3: build the signature        
    $method = "POST"
    $contentType = "application/json"
    $resource = "/api/logs"
    $rfc1123date = [DateTime]::UtcNow.ToString("r")
    $contentLength = $bodyJsonUTF8.Length    
    $signature = Build-Signature -customerId $LogAnalyticsWorkspaceId -sharedKey $LogAnalyticsPrimaryKey -date $rfc1123date -contentLength $contentLength -method $method -contentType $contentType -resource $resource
    
    #Step 4: create the header
    $headers = @{
        "Authorization" = $signature;
        "Log-Type" = $LogAnalyticsTableName;
        "x-ms-date" = $rfc1123date;
    };

    #Step 5: REST API call
    $uri = 'https://' + $LogAnalyticsWorkspaceId + ".ods.opinsights.azure.com" + $resource + "?api-version=2016-04-01"
    $response = Invoke-WebRequest -Uri $uri -Method Post -Headers $headers -ContentType $contentType -Body $bodyJsonUTF8 -UseBasicParsing

    if ($Response.StatusCode -eq 200) {   
        $rows = $body.Count
        Write-Information -MessageData "$rows rows written to Log Analytics workspace $uri" -InformationAction Continue
    }

}

function ExportCSVtoLA
{
	$OptionEventHub = CheckExportOption
	
	$date = Get-Date -Format "yyyyMMdd"
	$ErrorFile = "MPARR - MS Products - Error - "+$date+".json"
	
	if (Test-Path $FileName)
	{
		Write-Host "Importing CSV file..."
		$data = Import-Csv -Path $FileName
		$log_analytics_array = @()            
		foreach($i in $data) {
			$log_analytics_array += $i
		}
		$json = $log_analytics_array | ConvertTo-Json

		if($ExportToJSONFileOnly)
		{
			$ExportJSONFile = "MPARR - MS Products - "+$date+".json"
			$pathJSON = $PSScriptRoot+"\ExportedData\"+$ExportJSONFile
			$json | Set-Content -Path $pathJSON
			Write-Host "`nExport file was named as :" -NoNewLine
			Write-Host $ExportJSONFile -ForeGroundColor Green 
		}
		if($OptionEventHub -eq "True")
		{
			EventHubConnection
			$EventHubInstance.PublishToEventHub($log_analytics_array, $ErrorFile)
		}else
		{
			Write-host "Table:" $TableName
			Write-Host "Data :"$log_analytics_array.count
			Post-LogAnalyticsData -LogAnalyticsTableName $TableName -body $log_analytics_array
		}
	}else 
	{
		Write-Host "File $FileName not found. Exiting."
		return
	}
	Write-Host "`nExport finished."
}
CheckPrerequisites
ExportCSVtoLA