<#PSScriptInfo

.VERSION 2.0.5

.GUID 883af802-165c-4700-b4c1-352686c02f01

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
Script to collect data using RMS API for tracking purpose

#>

<#
.NOTES
Script to collect data using RMS API
This script needs PowerShell 5 and is called from MPARR_RMSData2.ps1

HISTORY
Script      : MPARR-GetTrackingRMSData.ps1
Author      : S. Zamorano
Version     : 2.0.5
Dependencie	: Called by MPARR_RMSData2.ps1 and uses PowerShell 5
Description : The script exports RMS Logs assigned from RMS API and pushes into a customer-specified Log Analytics table. Please note if you change the name of the table - you need to update Workbook sample that displays the report , appropriately. Do ensure the older table is deleted before creating the new table - it will create duplicates and Log analytics workspace doesn't support upserts or updates.

.NOTES 
	12-02-2024	S. Zamorano		- Version released
	01-03-2024	S. Zamorano		- Public release
#>

param(
	[Parameter(Mandatory=$true)] 
		[string]$Connection,
	[Parameter()] 
		[string]$RMSPath,
	[Parameter()] 
		[array]$ContentIds
)

function CheckCertificateInstalled($thumbprint)
{
	$var = "False"
	$certificates = @(Get-ChildItem Cert:\CurrentUser\My | Where-Object {$_.EnhancedKeyUsageList -like "*Client Authentication*"}| Select-Object Thumbprint) 
	
	foreach($certificate in $certificates)
	{
		if($thumbprint -in $certificate.Thumbprint)
		{
			$var = "True"
		}
	 }
	 if($var -eq "True")
	 {
		Write-Host "Certificate validation..." -NoNewLine
		Write-Host "`t`t`t`tPassed!" -ForegroundColor Green
		return $var
	 }else
	 {
		Write-Host "`nCertificate installed on this machine is missing!!!" -ForeGroundColor Yellow
		Write-Host "To execute this script unattended a certificate needs to be installed, the same used under Microsoft Entra App"
		Start-Sleep -s 1
		return $var
	 }
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

function connect2service
{	
	<#
	.NOTES
	If you cannot add the "Compliance Administrator" role to the Microsoft Entra App, for security reasons, you can execute with "Compliance Administrator" role 
	this script using .\MPARR-RMSData2.ps1 -ManualConnection
	#>
	if($Connection -eq "Manual")
	{
		Write-Host "`nAuthentication is required, please check your browser" -ForegroundColor Green
		Connect-AIPService
	}else
	{
		$CONFIGFILE = "$PSScriptRoot\..\ConfigFiles\laconfig.json"
		$json = Get-Content -Raw -Path $CONFIGFILE
		[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
		
		$EncryptedKeys = $config.EncryptedKeys
		$AppClientID = $config.AppClientID
		$ClientSecretValue = $config.ClientSecretValue
		$CertificateThumb = $config.CertificateThumb
		$TenantGUID = $config.TenantGUID
		$EventHubNamespace = $config.EventHubNamespace
		$EventHub = $config.EventHub
		
		if ($EncryptedKeys -eq "True")
		{
			$CertificateThumb = DecryptSharedKey $CertificateThumb
			$ClientSecretValue = DecryptSharedKey $ClientSecretValue
		}
		$status = CheckCertificateInstalled -thumbprint $CertificateThumb
		
		if($status -eq "True")
		{
			Write-Host "Certificate:"$CertificateThumb
			Write-Host "AppID:"$AppClientID
			Write-Host "TenantID:"$TenantGUID
			Connect-AIPService -CertificateThumbPrint $CertificateThumb -ApplicationId $AppClientID -TenantId $TenantGUID -ServicePrincipal
		}else
		{
			Write-Host "`nThe Certificate set in laconfig.json don't match with the certificates installed on this machine, you can try to execute using manual connection, to do that execute: "
			Write-Host ".\MPARR_RMSData2.ps1 -ManualConnection" -ForeGroundColor Green
			exit
		}
	}
}

function RMSTracking
{
	$datefile = Get-Date -Format "yyyy-MM-dd"
	$ResultNumber = 0
	$result = @()
	Write-Host "Getting RMS Tracking Logs"
	$TrackEmptyResults = "MPARR - Tracking Empty Results -"+$datefile+".csv"
	$TrackingEmpty = @()
	$ExportTracking = "$RMSPath\TrackingLogs\"
	$pathSummary = $ExportTracking+$TrackEmptyResults

	foreach ($i in $ContentIds)
    {
        try
        {
			$result = Get-AipServiceTrackingLog -ContentId $i
			
			if ($result -eq $null)
			{
				$TrackingEmpty += $i
			}
			$ExportJSONFile = "MPARR - RMS Tracking ID -"+$i+" - "+$datefile+".json"
			$result = $result | ConvertTo-Json -Depth 3
			$pathJSON = $ExportTracking+$ExportJSONFile
			$result | Set-Content -Path $pathJSON
			
        }
        catch
        {
            if ($_.Exception.Message.Contains("Connect-AipService")) 
            {
                connect2service
                $result = Get-AipServiceTrackingLog -ContentId $item
				if ($result -eq $null)
				{
					$TrackingEmpty += $i
				}
				$ExportJSONFile = "MPARR - RMS Tracking ID -"+$i+" - "+$datefile+".json"
				$result = $result | ConvertTo-Json -Depth 3
				$pathJSON = $ExportTracking+$ExportJSONFile
				$result | Set-Content -Path $pathJSON
            }
        }

    }

	$TrackingEmpty | Set-Content -Path $pathSummary
}

function MainRMSSupportingScript
{
	Write-Host "Script executed"
	
	connect2service
	RMSTracking
	
}

MainRMSSupportingScript