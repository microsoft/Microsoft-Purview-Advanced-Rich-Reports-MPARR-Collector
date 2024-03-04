<#PSScriptInfo

.VERSION 2.0.5

.GUID 883af802-165c-4708-b4c1-352686c02f01

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
This script permit to collect Microsoft Entra Users attributes existing on the tenant using Microsoft Graph API

#>

<#
HISTORY
	Script      : MPARR-MicrosoftEntraUsers.ps1
	Author      : S. Zamorano
	Version     : 2.0.5
	Description : The script exports Microsoft Entra users from Microsoft Graph and pushes into a customer-specified Log Analytics table. Please note if you change the name of the table - you need to update Workbook sample that displays the report , appropriately. Do ensure the older table is deleted before creating the new table - it will create duplicates and Log analytics workspace doesn't support upserts or updates.
	2022-10-12		S. Zamorano		- Added laconfig.json file for configuration and decryption Function
	2022-10-18		G. Berdzik		- Fix licensing information
	2023-01-03		S. Zamorano		- Added Change to use beta API capabilities, added Id for users
	2023-03-31      G. Berdzik      - Support for large tenants
	2023-03-31		S. Zamorano		- Visual improvement for progress
	2023-10-02		S. Zamorano		- Fix Progress bar
	2023-10-24		S. Zamorano		- Added Microsoft Entra filter option
	2023-11-07		S. Zamorano		- Added attribute to skip decision and use as a task
	2023-11-27		S. Zamorano		- Added Microsoft_EntraIDUsers.json file to select the Attributes required per user, added function PowerShell version check

	07-02-2024		S. Zamorano		- First release
	07-02-2024		S. Zamorano		- Added EventHub connector
	12-02-2024		S. Zamorano		- New version released
	01-03-2024		S. Zamorano		- Public release
#>

using module "ConfigFiles\MPARRUtils.psm1"
param (
    # Log Analytics table where the data is written to. Log Analytics will add an _CL to this name.
    [string]$TableName = "MicrosoftEntraUsers",
	[Parameter()] 
        [switch]$ExportToCSVFileOnly,
	[Parameter()] 
        [switch]$ExportToJSONFileOnly,
	[Parameter()] 
        [switch]$ManualConnection,
	[Parameter()] 
        [switch]$CreateTask,
	[Parameter()] 
        [switch]$ExportToEventHub

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

function CheckCertificateInstalled($thumbprint)
{
	$var = "False"
	$certificates = @(Get-ChildItem Cert:\CurrentUser\My | Where-Object {$_.EnhancedKeyUsageList -like "*Client Authentication*"}| Select-Object Thumbprint) 
	#$thumbprint -in $certificates
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

function ValidateConfigurationFile
{
	#Validate laconfig.json that manage the configuration for connections
	$MPARRConfiguration = "$PSScriptRoot\ConfigFiles\laconfig.json"
	
	if (-not (Test-Path -Path $MPARRConfiguration))
	{
		Write-Host "`n##########################################################################################" -ForeGroundColor Yellow
		Write-Host "`nThe laconfig.json file is missing. Check if you are using the right path or execute MPARR_Setup.ps1 first."
		Write-Host "`nThe laconfig.json is required to continue, if you want to export the data without having MPARR installed, please execute:" -NoNewLine
		Write-Host ".\MPARR-MicrosoftEntraUsers.ps1 -ExportToCSVFileOnly -ManualConnection" -ForeGroundColor Green
		Write-Host "`n##########################################################################################" -ForeGroundColor Yellow
		Write-Host "`n"
		if($ExportToCSVFileOnly -Or $ExportToJSONFileOnly)
		{
			if($ManualConnection)
			{
				return
			}else
			{
				Write-Host "`n##########################################################################################" -ForeGroundColor Yellow
				Write-Host "`nThe laconfig.json is required to continue, if you want to export the data without having MPARR installed, please execute:" -NoNewLine
				Write-Host ".\MPARR-MicrosoftEntraUsers.ps1 -ExportToCSVFileOnly -ManualConnection" -ForeGroundColor Green
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
}

function ReadNumber([int]$max, [string]$msg, [ref]$option)
{
    $selection = 0
    do 
    {
        $resp = Read-Host $msg
        try {
            $selection = [int]$resp
            if (($selection -gt $max) -or ($selection -lt 1))
            {
                $selection = 0
                throw 
            }            
        }
        catch {
            Write-Host "Please enter number between 1 and $max" -ForegroundColor DarkYellow 
            $selection = 0
        }

    } until ($selection -ne 0)
    $option.Value = $selection
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

function connect2service
{		
	ValidateConfigurationFile
	<#
	.NOTES
	If you cannot add the "Compliance Administrator" role to the Microsoft Entra App, for security reasons, you can execute with "Compliance Administrator" role 
	this script using .\MPARR-MicrosoftEntraUsers.ps1 -ManualConnection
	#>
	if($ManualConnection)
	{
		Write-Host "`nAuthentication is required, please check your browser" -ForegroundColor Green
		Write-Host "Please note that manual connection might not work because some additional permissions may be required." -ForegroundColor DarkYellow
		Connect-MgGraph
	}else
	{
		$CONFIGFILE = "$PSScriptRoot\ConfigFiles\laconfig.json"
		$json = Get-Content -Raw -Path $CONFIGFILE
		[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
		
		$EncryptedKeys = $config.EncryptedKeys
		$AppClientID = $config.AppClientID
		$CertificateThumb = $config.CertificateThumb
		$TenantGUID = $config.TenantGUID
		if ($EncryptedKeys -eq "True")
		{
			$CertificateThumb = DecryptSharedKey $CertificateThumb
		}
		$status = CheckCertificateInstalled -thumbprint $CertificateThumb
		
		if($status -eq "True")
		{
			Connect-MgGraph -CertificateThumbPrint $CertificateThumb -AppID $AppClientID -TenantId $TenantGUID -NoWelcome
		}else
		{
			Write-Host "`nThe Certificate set in laconfig.json don't match with the certificates installed on this machine, you can try to execute using manual connection, to do that extecute: "
			Write-Host ".\MPARR-MicrosoftEntraUsers.ps1 -ManualConnection" -ForeGroundColor Green
			exit
		}
		
	}
}

function CreateScheduledTaskFolder
{
	param([string]$taskFolder)
	
	#Main interface to select folder
	Write-Host "`n`n----------------------------------------------------------------------------------------" -ForegroundColor Yellow
	Write-Host "`n Please be aware that this list of Task Scheduler folder don't show empty folders." -ForegroundColor Red
	Write-Host "`n----------------------------------------------------------------------------------------" -ForegroundColor Yellow
	
	# Generate a unique list of parent folders under task scheduler
	$TSFolder = Get-ScheduledTask
	$uniqueTaskFolder = $TSFolder.TaskPath | Select-Object -Unique
	$tempFolder = $uniqueTaskFolder -replace '^\\(\w+)\\.*?.*','$1'
	$listTaskFolders = $tempFolder | Select-Object -Unique
	foreach ($folder in $listTaskFolders){$SchedulerTaskFolders += @([pscustomobject]@{Name=$folder})}
	
	Write-Host "`nGetting Folders..." -ForegroundColor Green
    $i = 1
    $SchedulerTaskFolders = @($SchedulerTaskFolders | ForEach-Object {$_ | Add-Member -Name "No" -MemberType NoteProperty -Value ($i++) -PassThru})
    
	#List all existing folders under Task Scheduler
    $SchedulerTaskFolders | Select-Object No, Name | Out-Host
	
	# Default folder for MPARR tasks
    $MPARRTSFolder = "MPARR2"
	$taskFolder = "\"+$MPARRTSFolder+"\"
	$choices  = '&Proceed', '&Change', '&Existing'
	Write-Host "Please consider if you want to use the default location you need select Existing and the option 1." -ForegroundColor Yellow
    $decision = $Host.UI.PromptForChoice("", "Default task Scheduler Folder is '$MPARRTSFolder'. Do you want to Proceed, Change the name or use Existing one?", $choices, 0)
    if ($decision -eq 1)
    {
        $ok = $false
        do 
        {
            $newName = Read-Host "Please enter the new name for the Task Scheduler folder"
        }
        until ($newName -ne "")
        $taskFolder = "\"+$newName+"\"
		Write-Host "The name selected for the folder under Task Scheduler is $newName." -ForegroundColor Green
		return $taskFolder
    }if ($decision -eq 0)
	{
		Write-Host "Using the default folder $MPARRTSFolder." -ForegroundColor Green
		return $taskFolder
	}else
	{
		$selection = 0
		ReadNumber -max ($i -1) -msg "Enter number corresponding to the current folder in the Task Scheduler" -option ([ref]$selection) 
		$value = $selection - 1
		$MPARRTSFolder = $SchedulerTaskFolders[$value].Name
		$taskFolder = "\"+$SchedulerTaskFolders[$value].Name+"\"
		Write-Host "Folder selected for this task $MPARRTSFolder " -ForegroundColor Green
		return $taskFolder
	}
	
}

function CreateMPARREntraUsersTask
{
	# MPARR-MicrosoftEntraUsers script
    $taskName = "MPARR-MicrosotEntraUsers"
	
	# Call function to set a folder for the task on Task Scheduler
	$taskFolder = CreateScheduledTaskFolder
	
	# Task execution
    $validDays = 30
    $choices  = '&Yes', '&No'
    $decision = $Host.UI.PromptForChoice("", "The task on task scheduler will be set for $validDays days, do you want to change?", $choices, 1)
    if ($decision -eq 0)
    {
        ReadNumber -max 31 -msg "Enter number of days (Between 1 to 31). Remember check the retention period in your workspace in Logs Analtytics." -option ([ref]$validDays)
    }

    # calculate date
    $dt = Get-Date 
    $reminder = $dt.Day % $validDays
    $dt = $dt.AddDays(-$reminder)
    $startTime = [datetime]::new($dt.Year, $dt.Month, $dt.Day, $dt.Hour, $dt.Minute, 0)

    #create task
    $trigger = New-ScheduledTaskTrigger -Once -At $startTime -RepetitionInterval (New-TimeSpan -Days $validDays)
    $action = New-ScheduledTaskAction -Execute "`"$PSHOME\pwsh.exe`"" -Argument ".\MPARR-MicrosoftEntraUsers.ps1" -WorkingDirectory $PSScriptRoot
    $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopOnIdleEnd -AllowStartIfOnBatteries `
         -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours 1)

    if (Get-ScheduledTask -TaskName $taskName -TaskPath $taskFolder -ErrorAction SilentlyContinue) 
    {
        Write-Host "`nScheduled task named '$taskName' already exists.`n" -ForegroundColor Yellow
    }
    else 
    {
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings `
        -RunLevel Highest -TaskPath $taskFolder -ErrorAction Stop | Out-Null
        Write-Host "`nScheduled task named '$taskName' was created.`nFor security reasons you have to specify run as account manually.`n" -ForegroundColor Yellow
    }
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
    $bodyJson = $body | ConvertTo-Json -Depth 4

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
        Write-Information -MessageData "   $rows rows written to Log Analytics workspace $uri" -InformationAction Continue
    }

}

function ProgressBar($TotalRows)
{
	$ProgressValue = 1
	If ($TotalRows -le 100) {
		$ProgressValue = 4
	}
	If (($TotalRows -gt 100) -AND ($TotalRows -lt 1000)){
		$ProgressValue = 20
	}
	If ($TotalRows -ge 1000) {
		$ProgressValue = $TotalRows/100
	}
}

function InitializeLAConfigFile
{
	# read config file
    $configfile = "$PSScriptRoot\ConfigFiles\laconfig.json" 
	
	if (-not (Test-Path -Path $configfile))
    {
		$config = [ordered]@{
		EncryptedKeys =  "False"
		AppClientID = ""
		ClientSecretValue = ""
		TenantGUID = ""
		TenantDomain = ""
		LA_CustomerID =  ""
		LA_SharedKey =  ""
		CertificateThumb = ""
		OnmicrosoftURL = ""
		RMSLogs = "c:\APILogs\RMSLogs\"
		OutPutLogs = "c:\APILogs\"
		Cloud = "Commercial"
		MicrosoftEntraConfig = "Not Set"
		ExportToEventHub = "False"
		EventHubNamespace = ""
		EventHub = ""
		}
		return $config
    }else
	{
		$json = Get-Content -Raw -Path $configfile
		[PSCustomObject]$configfile = ConvertFrom-Json -InputObject $json
	
		$config = [ordered]@{
		EncryptedKeys = "$($configfile.EncryptedKeys)"
		AppClientID = "$($configfile.AppClientID)"
		ClientSecretValue = "$($configfile.ClientSecretValue)"
		TenantGUID = "$($configfile.TenantGUID)"
		TenantDomain = "$($configfile.TenantDomain)"
		LA_CustomerID = "$($configfile.LA_CustomerID)"
		LA_SharedKey = "$($configfile.LA_SharedKey)"
		CertificateThumb = "$($configfile.CertificateThumb)"
		OnmicrosoftURL = "$($configfile.OnmicrosoftURL)"
		RMSLogs = "$($configfile.RMSLogs)"
		OutPutLogs = "$($configfile.OutPutLogs)"
		Cloud = "$($configfile.Cloud)"
		MicrosoftEntraConfig = "$($configfile.MicrosoftEntraConfig)"
		ExportToEventHub = "$($configfile.ExportToEventHub)"
		EventHubNamespace = "$($configfile.EventHubNamespace)"
		EventHub = "$($configfile.EventHub)"
		}
		return $config
	}
}

function WriteToJsonFile
{
	if (Test-Path "$PSScriptRoot\ConfigFiles\laconfig.json")
    {
        $date = Get-Date -Format "yyyyMMddHHmmss"
        Move-Item "$PSScriptRoot\ConfigFiles\laconfig.json" "$PSScriptRoot\BackupScripts\laconfig_$date.json"
        Write-Host "`nThe old config file moved to 'laconfig_$date.json'"
    }
    $config | ConvertTo-Json | Out-File "$PSScriptRoot\ConfigFiles\laconfig.json"
    Write-Host "Setup completed. New config file was created." -ForegroundColor Yellow
}

function MicrosoftEntraIDAttributes
{
	$EntraAttributes = "$PSScriptRoot\ConfigFiles\MPARR_EntraIDUsers.json"
	
	$DefaultAttributes = "userPrincipalName,displayName,signInActivity,assignedLicenses,assignedPlans,city,createdDateTime,department,jobTitle,mail,officeLocation,userType"
	
	if (-not (Test-Path -Path $EntraAttributes))
	{
		Write-Host "MPARR_EntraIDUsers.json file is missing. Default list of subscriptions will be used."
		return $DefaultAttributes
	}
	else 
	{
		$DefaultAttributes = $Null
		$json = Get-Content -Raw -Path $EntraAttributes
		[PSCustomObject]$attributes = ConvertFrom-Json -InputObject $json
		foreach ($attribute in $attributes.psobject.Properties)
		{
			if ($attributes."$($Attribute.Name)" -eq "True")
			{
				$DefaultAttributes += $Attribute.Name+","
			}
		}
		#Write-Host "Subscriptions list: $DefaultAttributes"   
		return $DefaultAttributes
	}
	
}

function SelectImportFilter
{
	
	$CONFIGFILE = "$PSScriptRoot\ConfigFiles\laconfig.json"
	$json = Get-Content -Raw -Path $CONFIGFILE
	[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
	$MicrosoftEntraConfig = $config.MicrosoftEntraConfig
	
	if($MicrosoftEntraConfig -eq $Null)
	{
		$config = InitializeLAConfigFile
		$config.MicrosoftEntraConfig = "Not Set"
		WriteToJsonFile
	}
	Start-Sleep -s 1
	
	$CONFIGFILE = "$PSScriptRoot\ConfigFiles\laconfig.json"
	$json = Get-Content -Raw -Path $CONFIGFILE
	[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
	$MicrosoftEntraConfig = $config.MicrosoftEntraConfig
	
	$DefaultAttributes = MicrosoftEntraIDAttributes
	
	if($MicrosoftEntraConfig -eq "Not Set")
	{
		#This Function is used to select the kind of filter for the users from MIcrosoft Entra ID
		Write-Host "`n##########################################################################################" -ForegroundColor Blue
		Write-Host "`nBy default this script import the data only from licensed users and as a members of Tenant, any other kind of users like as guest or unlicensed are not imported." -ForegroundColor Yellow
		Write-Host "If you want to collect all the information for all users in your tenant including guest and unlicensed users select Change." -ForegroundColor Yellow
		$choices  = '&Proceed', '&Change'
		Write-Host "If you are ok with this you can select Proceed, if you want to download all users including guest and unlicensed users please select Change." -ForegroundColor Yellow
		$decision = $Host.UI.PromptForChoice("", "Default filter only members with licenses assigned. Do you want to Proceed or Change?", $choices, 0)
		if ($decision -eq 1)
		{
			Write-Host "Importing all your users..." -ForegroundColor Green
			Write-Host "Fetching data from Microsoft Entra ID..."
			$body = @{
			select=$DefaultAttributes
			count="true"
			}
			$config.MicrosoftEntraConfig = "Total"
			WriteToJsonFile
			return $body
		}elseif ($decision -eq 0)
		{
			Write-Host "Using the default filter to import your users" -ForegroundColor Green
			Write-Host "Fetching data from Microsoft Entra ID..."
			$body = @{
			select=$DefaultAttributes
			filter="assignedLicenses/count ne 0 and userType eq 'Member'"
			count="true"
			}
			$config.MicrosoftEntraConfig = "Filtered"
			WriteToJsonFile
			return $body
		}
	}else
	{
		if ($MicrosoftEntraConfig -eq "Total")
		{
			Write-Host "Importing all your users, including not licensed and guests..." -ForegroundColor Green
			Write-Host "Fetching data from Microsoft Entra ID..."
			$body = @{
			select=$DefaultAttributes
			count="true"
			}
			return $body
		}elseif ($MicrosoftEntraConfig -eq "Filtered")
		{
			Write-Host "Using the default filter to import your users" -ForegroundColor Green
			Write-Host "Fetching data from Microsoft Entra ID..."
			$body = @{
			select=$DefaultAttributes
			filter="assignedLicenses/count ne 0 and userType eq 'Member'"
			count="true"
			}
			return $body
		}
	}
}

Function Export-MicrosoftData
{
    # ---------------------------------------------------------------   
    #    Name           : Export-MicrosoftData
    #    Desc           : Extracts data from Get-MgUsers into Log analytics workspace tables for reporting purposes
    #    Return         : None
    # ---------------------------------------------------------------
    
	$OptionEventHub = CheckExportOption
	if($OptionEventHub -eq "True")
	{
		EventHubConnection
	}
	connect2service

	$body = SelectImportFilter
	$DefaultAttributes = MicrosoftEntraIDAttributes
	$DefaultAttributes = $DefaultAttributes -replace “.$”
	$AttributesArray = $DefaultAttributes.Split(",")
	
	Write-Host "`n################################################################################" -ForegroundColor DarkGreen
	Write-Host "`nThe current Microsoft Entra ID attributes was selected:" -ForegroundColor Green
	foreach($attribute in $AttributesArray)
	{
		Write-Host "`t* '$($attribute)'" -ForegroundColor Green
	}
	Write-Host "`n################################################################################" -ForegroundColor DarkGreen
    
    $headers = @{
        ConsistencyLevel="eventual"
    }

    $usersAL = New-Object System.Collections.ArrayList     
    $bufferSize = 10MB
    $size = 0        
    $page = 1
	$stop = $false
	$Progress = 0
	$perc = 0

    $response = Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/v1.0/users" -Body $body -Headers $headers
	$TotalRows = $response["@odata.count"]
    Write-Host "Total number of records found: $($response["@odata.count"])." 
	ProgressBar
	
	if($ExportToCSVFileOnly -Or $ExportToJSONFileOnly)
	{
		$ExportPath = $PSScriptRoot+"\ExportedData"
		if(-Not (Test-Path $ExportPath ))
		{
			Write-Host "Export data directory is missing, creating a new folder called ExportedData"
			New-Item -ItemType Directory -Force -Path "$PSScriptRoot\ExportedData" | Out-Null
		}
		$date = Get-Date -Format "yyyyMMdd"

		if($ExportToCSVFileOnly)
		{
			$ExportCSVFile = "MPARR - Microsoft Entra Users - "+$date+".csv"
			$pathCSV = $PSScriptRoot+"\ExportedData\"+$ExportCSVFile
			do
		{
			$Progress += $ProgressValue
			$perc = $Progress/$TotalRows
			Write-Progress -Activity "Data received. Processing page no. [$page]" -PercentComplete $perc
			$page++

			foreach($user in $response.value) 
			{
				$newAttributeItem = @{}
				foreach ($attribute in $AttributesArray)
					{
						$newAttributeItem[$attribute] = $user.$attribute
						#Write-Host "'$($user.$attribute)'"
					}
				$newItem = [PSCustomObject]$newAttributeItem

				[void]$usersAL.Add($newitem)
				$size += [System.Text.Encoding]::UTF8.GetByteCount(($newitem | ConvertTo-Json -Depth 100))
				if ($size -gt $bufferSize)
				{
					$log_analytics_array = $usersAL.ToArray()
					$log_analytics_array | Export-CSV -Path $pathCSV -NTI -Append
					$log_analytics_array = $null
					$usersAL.Clear()
					$size = 0
				}
			}
			if ($response["@odata.nextLink"] -ne $null)
			{
				$response = Invoke-MgGraphRequest -Method Get -Uri $response["@odata.nextLink"] #-Body $body -Headers $headers
			}
			else 
			{
				$stop = $true
				Write-Host "   Work completed!!! $TotalRows elements imported to Logs Analytics" -ForegroundColor Green
			} 

		} while (-not $stop)

		# Push remaining data to Log Analytics
		if ($usersAL.Count -gt 0)
		{
			$log_analytics_array = $usersAL.ToArray()
			$log_analytics_array | Export-CSV -Path $pathCSV -NTI -Append			
		}
			Write-Host "`nExport file was named as :" -NoNewLine
			Write-Host $ExportCSVFile -ForeGroundColor Green 
		}
		if($ExportToJSONFileOnly)
		{
			$ExportJSONFile = "MPARR - Microsoft Entra Users - "+$date+".json"
			$pathJSON = $PSScriptRoot+"\ExportedData\"+$ExportJSONFile
			do
			{
				$Progress += $ProgressValue
				$perc = $Progress/$TotalRows
				Write-Progress -Activity "Data received. Processing page no. [$page]" -PercentComplete $perc
				$page++

				foreach($user in $response.value) 
				{
					$newAttributeItem = @{}
					foreach ($attribute in $AttributesArray)
						{
							$newAttributeItem[$attribute] = $user.$attribute
							#Write-Host "'$($user.$attribute)'"
						}
					$newItem = [PSCustomObject]$newAttributeItem

					[void]$usersAL.Add($newitem)
					$size += [System.Text.Encoding]::UTF8.GetByteCount(($newitem | ConvertTo-Json -Depth 100))
					if ($size -gt $bufferSize)
					{
						$log_analytics_array = $usersAL.ToArray()
						$json = $log_analytics_array | ConvertTo-Json -Depth 4
						$json | Add-Content -Path $pathJSON 
						$log_analytics_array = $null
						$usersAL.Clear()
						$size = 0
					}
				}
				if ($response["@odata.nextLink"] -ne $null)
				{
					$response = Invoke-MgGraphRequest -Method Get -Uri $response["@odata.nextLink"] #-Body $body -Headers $headers
				}
				else 
				{
					$stop = $true
					Write-Host "   Work completed!!! $TotalRows elements imported to Logs Analytics" -ForegroundColor Green
				} 

			} while (-not $stop)

			# Push remaining data to Log Analytics
			if ($usersAL.Count -gt 0)
			{
				$log_analytics_array = $usersAL.ToArray()
				$json = $log_analytics_array | ConvertTo-Json -Depth 4
				$json | Add-Content -Path $pathJSON		
			}
			
			Write-Host "`nExport file was named as :" -NoNewLine
			Write-Host $ExportJSONFile -ForeGroundColor Green 
		}

		Write-Host "`nFile was copied at :" -NoNewLine
		Write-Host $PSScriptRoot"\ExportedData" -ForeGroundColor Green 
		Write-Host "`n"
	}elseif($OptionEventHub -eq "True")
	{
		# Else format for Log Analytics
		do
		{
			$Progress += $ProgressValue
			$perc = $Progress/$TotalRows
			Write-Progress -Activity "Data received. Processing page no. [$page]" -PercentComplete $perc
			$page++

			foreach($user in $response.value) 
			{
				$newAttributeItem = @{}
				foreach ($attribute in $AttributesArray)
					{
						$newAttributeItem[$attribute] = $user.$attribute
						#Write-Host "'$($user.$attribute)'"
					}
				$newItem = [PSCustomObject]$newAttributeItem

				[void]$usersAL.Add($newitem)
				$size += [System.Text.Encoding]::UTF8.GetByteCount(($newitem | ConvertTo-Json -Depth 100))
				if ($size -gt $bufferSize)
				{
					$log_analytics_array = $usersAL.ToArray()
					$EventHubInstance.PublishToEventHub($log_analytics_array, $ErrorFile)
					$log_analytics_array = $null
					$usersAL.Clear()
					$size = 0
				}
			}
			if ($response["@odata.nextLink"] -ne $null)
			{
				$response = Invoke-MgGraphRequest -Method Get -Uri $response["@odata.nextLink"] #-Body $body -Headers $headers
			}
			else 
			{
				$stop = $true
				Write-Host "   Work completed!!! $TotalRows elements imported to Logs Analytics" -ForegroundColor Green
			} 

		} while (-not $stop)

		# Push remaining data to Log Analytics
		if ($usersAL.Count -gt 0)
		{
			$log_analytics_array = $usersAL.ToArray()
			$EventHubInstance.PublishToEventHub($log_analytics_array, $ErrorFile)			
		}
		
	}else
	{
		# Else format for Log Analytics
		do
		{
			$Progress += $ProgressValue
			$perc = $Progress/$TotalRows
			Write-Progress -Activity "Data received. Processing page no. [$page]" -PercentComplete $perc
			$page++

			foreach($user in $response.value) 
			{
				$newAttributeItem = @{}
				foreach ($attribute in $AttributesArray)
					{
						$newAttributeItem[$attribute] = $user.$attribute
						#Write-Host "'$($user.$attribute)'"
					}
				$newItem = [PSCustomObject]$newAttributeItem

				[void]$usersAL.Add($newitem)
				$size += [System.Text.Encoding]::UTF8.GetByteCount(($newitem | ConvertTo-Json -Depth 100))
				if ($size -gt $bufferSize)
				{
					$log_analytics_array = $usersAL.ToArray()
					Post-LogAnalyticsData -LogAnalyticsTableName $TableName -body $log_analytics_array
					$log_analytics_array = $null
					$usersAL.Clear()
					$size = 0
				}
			}
			if ($response["@odata.nextLink"] -ne $null)
			{
				$response = Invoke-MgGraphRequest -Method Get -Uri $response["@odata.nextLink"] #-Body $body -Headers $headers
			}
			else 
			{
				$stop = $true
				Write-Host "   Work completed!!! $TotalRows elements imported to Logs Analytics" -ForegroundColor Green
			} 

		} while (-not $stop)

		# Push remaining data to Log Analytics
		if ($usersAL.Count -gt 0)
		{
			$log_analytics_array = $usersAL.ToArray()
			Post-LogAnalyticsData -LogAnalyticsTableName $TableName -body $log_analytics_array			
		}
	}
}
     
#Run the script.
CheckPrerequisites
if($CreateTask)
{
	CreateMPARREntraUsersTask
	exit
}
Export-MicrosoftData
