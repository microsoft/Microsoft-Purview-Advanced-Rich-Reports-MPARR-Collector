<#PSScriptInfo

.VERSION 2.1.1

.GUID 883af802-165c-4703-b4c1-352686c02f01

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
The script exports Content Explorer from Export-ContentExplorerData and pushes into a customer-specified Log Analytics table. 

#>

<#
HISTORY
Script      : MPARR-ContentExplorerData.ps1
Author      : Sebastian Zamorano
Co-Author   : 
Version     : 2.1.1
Date		: 09-04-2024
Description : The script exports Content Explorer from Export-ContentExplorerData and pushes into a customer-specified Log Analytics table. 
			Please note if you change the name of the table - you need to update Workbook sample that displays the report , appropriately. 
			Do ensure the older table is deleted before creating the new table - it will create duplicates and Log analytics workspace doesn't support upserts or updates.
			
.NOTES 
	26-12-2023	S. Zamorano		- MPARR-ContentExplorerData-BasicReturn.ps1 used as base
	26-12-2023	S. Zamorano		- Added Tablename, Export 2 file only, export to Logs analytics, configuration files.
	29-12-2023	S. Zamorano		- First Release
	02-01-2024  S. Zamorano		- Columns added to the results, TagType and TagName for Logs Analytics, to improve the reports on Power BI
	03-01-2024	S. Zamorano		- Organize the Json files in alphabetical order, my thanks to G. Berdzik
	04-01-2024	S. Zamorano		- Some additional information added to logs for errors and summary, added logs to export to Logs Analtytics
	05-01-2024	S. Zamorano		- Improve how to manage page size, and how the data is exported to CSV or Logs Analytics
	05-02-2024	S. Zamorano		- Buffer added to send to Logs Analytics and Trap cmdlet to fix some breaks
	01-03-2024	S. Zamorano		- Public release
	08-04-2024	S. Zamorano		- Change on the logic used to get the data from Content Explorer Data and reduce errors from that PowerShell Module. Added a list of users and sites
	09-04-2024	S. Zamorano		- Fixes for MassExport and simple export. Added Event Hub option
#>

using module "ConfigFiles\MPARRUtils.psm1"
[CmdletBinding(DefaultParameterSetName = "None")]
param(
	[string]$TableName = "ContentExplorer",
	#Export-ContentExplorerData cmdlet requires a PageSize that can be between 1 to 5000, by default is set to 100, you can change the number below or use the parameter -ChangePageSize to modify during the execution
	[int]$InitialPageSize = 200,
	[Parameter()] 
        [switch]$SimpleExportToFile,
    [Parameter()] 
        [switch]$ChangePageSize,
	[Parameter()] 
        [switch]$MassExportToCsv,
	[Parameter()] 
        [switch]$MassExportToJson,
	[Parameter()] 
        [switch]$CreateConfigFiles,
	[Parameter()] 
        [switch]$ManualConnection,
	[Parameter()] 
        [switch]$CheckDependencies,
	[Parameter()] 
        [switch]$ExportToEventHub,
	[Parameter()] 
        [switch]$CreateTask
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

function CheckIfElevated
{
    $IsElevated = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (!$IsElevated)
    {
        Write-Host "`nPlease start PowerShell as Administrator.`n" -ForegroundColor Yellow
        exit(1)
    }
}

function CheckRequiredModules 
{
    # Check PowerShell modules
    Write-Host "Checking PowerShell modules..."
    $requiredModules = @(
		@{Name="Microsoft.Graph.Sites"; MinVersion="0.0"},
		@{Name="Microsoft.Graph.Reports"; MinVersion="0.0"},
        @{Name="ExchangeOnlineManagement"; MinVersion="0.0"}
        )

    $modulesToInstall = @()
    foreach ($module in $requiredModules)
    {
        Write-Host "`t$($module.Name) - " -NoNewline
        $installedVersions = Get-Module -ListAvailable $module.Name
        if ($installedVersions)
        {
            if ($installedVersions[0].Version -lt [version]$module.MinVersion)
            {
                Write-Host "`t`t`tNew version required" -ForegroundColor Red
                $modulesToInstall += $module.Name
            }
            else 
            {
                Write-Host "`t`t`tInstalled" -ForegroundColor Green
            }
        }
        else
        {
            Write-Host "`t`t`tNot installed" -ForegroundColor Red
            $modulesToInstall += $module.Name
        }
    }

    if ($modulesToInstall.Count -gt 0)
    {
        CheckIfElevated
		$choices  = '&Yes', '&No'

        $decision = $Host.UI.PromptForChoice("", "Misisng required modules. Proceed with installation?", $choices, 0)
        if ($decision -eq 0) 
        {
            Write-Host "Installing modules..."
            foreach ($module in $modulesToInstall)
            {
                Write-Host "`t$module"
				Install-Module $module -ErrorAction Stop
                
            }
            Write-Host "`nModules installed. Please start the script again."
            exit(0)
        } 
        else 
        {
            Write-Host "`nExiting setup. Please install required modules and re-run the setup."
            exit(1)
        }
    }
}

function ValidateConfigurationFile
{
	#Validate laconfig.json that manage the configuration for connections
	$MPARRConfiguration = "$PSScriptRoot\ConfigFiles\laconfig.json"
	
	if (-not (Test-Path -Path $MPARRConfiguration))
	{
		Write-Host "`nThe laconfig.json file is missing. Check if you are using the right path or execute MPARR_Setup.ps1 first."
		exit
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
	
	#To export data from Trainable Classifiers this is mandatory
	$TCSelected = "$PSScriptRoot\ConfigFiles\MPARR-TrainableClassifiersList.json"
	if (-not (Test-Path -Path $TCSelected))
	{
		Write-Host "MPARR-TrainableClassifiersList.json file is missing, you will not be available to download data related to Trainable Classifiers" -ForeGroundColor DarkYellow
		Write-Host "File need to be located at "$PSScriptRoot"\ConfigFiles"
		Write-Host "You can find the file in our GitHub repo at https://aka.ms/MPARR-GitHub"
		exit
	}else
	{
		Write-Host "Checking MPARR-TrainableClassifiersList.json file..." -NoNewLine
		Write-Host "`tPassed!" -ForeGroundColor Green
		$jsonTC = Get-Content -Raw -Path $TCSelected
		[PSCustomObject]$tcs = ConvertFrom-Json -InputObject $jsonTC
		$CountTCs = 0
		$CountTCselected = 0
		
		foreach ($tcd in $tcs.psobject.Properties)
		{
			if ($tcs."$($tcd.Name)" -eq "True")
			{
				$CountTCselected++
			}
			$CountTCs++
		}
		
		Write-Host "`t`tTrainable Classifiers selected : `t$CountTCselected of $CountTCs"
		Start-Sleep -s 1
		Start-Sleep -s 1
	}
	
	#Check configuration file for tags
	$TagsSelected = "$PSScriptRoot\ConfigFiles\MPARR-CETagtype.json"
	if (-not (Test-Path -Path $TagsSelected))
	{
		Write-Host "`nMPARR-CETagtype.json file is missing, you will not be available to download data related to Trainable Classifiers" -ForeGroundColor DarkYellow
		Write-Host "File need to be located at "$PSScriptRoot"\ConfigFiles"
		Write-Host "If the file is not present, default values will be used."
		Write-Host "You can find the file in our GitHub repo at https://aka.ms/MPARR-GitHub"
		Start-Sleep -s 2
	}else
	{
		Write-Host "Checking MPARR-CETagtype.json file..." -NoNewLine
		Write-Host "`t`t`tPassed!" -ForeGroundColor Green
		Start-Sleep -s 1
	}
	
	#Check configuration file for Workloads
	$WorkloadsSelected = "$PSScriptRoot\ConfigFiles\MPARR-CEWorkload.json"
	if (-not (Test-Path -Path $WorkloadsSelected))
	{
		Write-Host "`nMPARR-CEWorkload.json file is missing, you will not be available to download data related to Trainable Classifiers" -ForeGroundColor DarkYellow
		Write-Host "File need to be located at "$PSScriptRoot"\ConfigFiles"
		Write-Host "If the file is not present, default values will be used."
		Write-Host "You can find the file in our GitHub repo at https://aka.ms/MPARR-GitHub"
		Start-Sleep -s 2
	}else
	{
		Write-Host "Checking MPARR-CEWorkload.json file..." -NoNewLine
		Write-Host "`t`t`tPassed!" -ForeGroundColor Green
		Start-Sleep -s 1
	}
}

function ValidateAdditionalConfigurationFiles
{
	#The next files can be created through execute .\MPARR-ContentExplorerData2.ps1 -CreateConfigFiles
	
	#To export data from Sensitive Information Types
	$SITsSelected = "$PSScriptRoot\ConfigFiles\MPARR-SensitiveInfoTypesList.json"
	if (-not (Test-Path -Path $SITsSelected))
	{
		Write-Host "MPARR-SensitiveInfoTypesList.json file is not set, all Sensitive Information Types will be used" -ForeGroundColor DarkYellow
		Write-Host "If you want to use this configuration file, can  be created executing .\MPARR-ContentExplorerData2.ps1 -CreateConfigFiles"
		Write-Host "File will be created at "$PSScriptRoot"\ConfigFiles"
	}else
	{
		Write-Host "Checking MPARR-SensitiveInfoTypesList.json file..." -NoNewLine
		Write-Host "`tAvailable!" -ForeGroundColor Green
		$jsonSIT = Get-Content -Raw -Path $SITsSelected
		[PSCustomObject]$sitss = ConvertFrom-Json -InputObject $jsonSIT
		$CountSITs = 0
		$CountSITselected = 0
		
		foreach ($sitd in $sitss.psobject.Properties)
		{
			if ($sitss."$($sitd.Name)" -eq "True")
			{
				$CountSITselected++
			}
			$CountSITs++
		}
		
		Write-Host "`t`tSensitive Information Types selected : `t$CountSITselected of $CountSITs"
		Start-Sleep -s 2
	}
	
	#To export data from Sensitivity Labels
	$SLSelected = "$PSScriptRoot\ConfigFiles\MPARR-SensitivityLabelsList.json"
	if (-not (Test-Path -Path $SLSelected))
	{
		Write-Host "MPARR-SensitivityLabelsList.json file is not set, all Sensitivity Labels will be used" -ForeGroundColor DarkYellow
		Write-Host "If you want to use this configuration file, can  be created executing .\MPARR-ContentExplorerData2.ps1 -CreateConfigFiles"
		Write-Host "File will be created at "$PSScriptRoot"\ConfigFiles"
	}else
	{
		Write-Host "Checking MPARR-SensitivityLabelsList.json file..." -NoNewLine
		Write-Host "`tAvailable!" -ForeGroundColor Green
		$jsonSL = Get-Content -Raw -Path $SLSelected
		[PSCustomObject]$sls = ConvertFrom-Json -InputObject $jsonSL
		$CountSLs = 0
		$CountSLselected = 0
		
		foreach ($sld in $sls.psobject.Properties)
		{
			if ($sls."$($sld.Name)" -eq "True")
			{
				$CountSLselected++
			}
			$CountSLs++
		}
		
		Write-Host "`t`tSensitivity Labels selected : `t`t$CountSLselected of $CountSLs"
		Start-Sleep -s 1
	}
	
	#To export data from Retention Labels
	$RLSelected = "$PSScriptRoot\ConfigFiles\MPARR-RetentionLabelsList.json"
	if (-not (Test-Path -Path $RLSelected))
	{
		Write-Host "MPARR-RetentionLabelsList.json file is not set, all Retention Labels will be used" -ForeGroundColor DarkYellow
		Write-Host "If you want to use this configuration file, can  be created executing .\MPARR-ContentExplorerData2.ps1 -CreateConfigFiles"
		Write-Host "File will be created at "$PSScriptRoot"\ConfigFiles"
	}else
	{
		Write-Host "Checking MPARR-RetentionLabelsList.json file..." -NoNewLine
		Write-Host "`t`tAvailable!" -ForeGroundColor Green
		$jsonRL = Get-Content -Raw -Path $RLSelected
		[PSCustomObject]$rls = ConvertFrom-Json -InputObject $jsonRL
		$CountRLs = 0
		$CountRLselected = 0
		
		foreach ($rld in $rls.psobject.Properties)
		{
			if ($rls."$($rld.Name)" -eq "True")
			{
				$CountRLselected++
			}
			$CountRLs++
		}
		
		Write-Host "`t`tRetention Labels selected : `t`t$CountRLselected of $CountRLs"
		Start-Sleep -s 1
	}
}

function CheckPrerequisites
{
    CheckPowerShellVersion
}

function CheckContentExplorerPermissions
{
	 if (-not (Get-Command -Name Export-ContentExplorerData -ErrorAction SilentlyContinue)) 
	 {
		Write-Host "You don´t have the permissions required to execute the cmdlet Export-ContentExplorerData"
		Write-Host "Please sign-in again with an account with these permissions assigned :"
		Write-Host "`t* Content Explorer Content Viewer"
		Write-Host "`t* Content Explorer List Viewer"
		Write-Host "`nYou can connect manually running " -NoNewline
		Write-Host ".\MPARR-ContentExplorerData2.ps1 -ManualConnection"
		exit
	 }
}

function UpdateMPARREntraApp
{
	Connect-MgGraph -Scopes "Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All", "Directory.ReadWrite.All", "User.ReadWrite.All" -NoWelcome
	Clear-Host
	
	Write-Host "`n`n----------------------------------------------------------------------------------------"
	Write-Host "`nMPARR Microsoft Entra App update!" -ForegroundColor DarkGreen
	Write-Host "This menu helps to validate that the Microsoft Entra App previously created have all the API permissions required." -ForegroundColor DarkGreen
	Write-Host "You will need to consent permissions Under Microsoft Entra portal to the app and the new permissions." -ForegroundColor DarkGreen
	Write-Host "`n----------------------------------------------------------------------------------------"
	
	$CONFIGFILE = "$PSScriptRoot\ConfigFiles\laconfig.json"
	$json = Get-Content -Raw -Path $CONFIGFILE
	[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
	$AppID = $config.AppClientID
	
    $filter = "AppId eq '$AppId'"
    $servicePrincipal = Get-MgServicePrincipal -All -Filter $filter
    $roles = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId ($servicePrincipal.Id)
    if ($roles.AppRoleId -notcontains "230c1aed-a721-4c5d-9cb4-a90514e508ef")
    {
        Write-Host "Microsoft Graph API permission 'Reports.Read.All'" -NoNewLine
        Write-Host "`tNot Found!" -ForegroundColor Red
		Write-Host "App ID used:" $AppId
        Write-Host "Press any key to continue..."
        $key = ([System.Console]::ReadKey($true))
        Write-Host "Adding permission"
        # app parameters and API permissions definition
        $params = @{
            AppId = $AppID
            RequiredResourceAccess = @(
                @{
                    ResourceAppId = "00000003-0000-0000-c000-000000000000"
                    ResourceAccess = @(
                        @{
                            Id = "230c1aed-a721-4c5d-9cb4-a90514e508ef"
                            Type = "Role"
                        }
                    )
                }
        
            )
        }
        Update-MgApplicationByAppId @params
        Write-Host "Permission added." -ForegroundColor Green
        Write-Host "`nPlease go to the Azure portal to manually grant admin consent:"
        Write-Host "https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$($AppId)`n" -ForegroundColor Cyan    
    }
    else 
    {
        Write-Host "Microsoft Graph API permission..." -NoNewLine
		Write-Host "`t'Reports.Read.All'" -NoNewLine -ForegroundColor Green
        Write-Host "`tpermission already in place." 
		Start-Sleep -s 3
    }
	if ($roles.AppRoleId -notcontains "332a536c-c7ef-4017-ab91-336970924f0d")
    {
        Write-Host "Microsoft Graph API permission 'Sites.Read.All'" -NoNewLine
        Write-Host "`tNot Found!" -ForegroundColor Red
		Write-Host "App ID used:" $AppId
        Write-Host "Press any key to continue..."
        $key = ([System.Console]::ReadKey($true))
        Write-Host "Adding permission"
        # app parameters and API permissions definition
        $params = @{
            AppId = $AppID
            RequiredResourceAccess = @(
                @{
                    ResourceAppId = "00000003-0000-0000-c000-000000000000"
                    ResourceAccess = @(
                        @{
                            Id = "332a536c-c7ef-4017-ab91-336970924f0d"
                            Type = "Role"
                        }
                    )
                }
        
            )
        }
        Update-MgApplicationByAppId @params
        Write-Host "Permission added." -ForegroundColor Green
        Write-Host "`nPlease go to the Azure portal to manually grant admin consent:"
        Write-Host "https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$($AppId)`n" -ForegroundColor Cyan    
    }
    else 
    {
        Write-Host "Microsoft Graph API permission..." -NoNewLine
		Write-Host "`t'Sites.Read.All'" -NoNewLine -ForegroundColor Green
        Write-Host "`tpermission already in place." 
		Start-Sleep -s 3
    }
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

function GetAuthToken
{
    $CONFIGFILE = "$PSScriptRoot\ConfigFiles\laconfig.json" 
	$json = Get-Content -Raw -Path $CONFIGFILE
	[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
	$EncryptedKeys = $config.EncryptedKeys
	$AppClientID = $config.AppClientID
	$ClientSecretValue = $config.ClientSecretValue
	$TenantDomain = $config.TenantDomain
	$loginURL = "https://login.microsoftonline.com"

	if ($EncryptedKeys -eq "True")
	{
		$ClientSecretValue = DecryptSharedKey $ClientSecretValue
	}
	
	$body = @{grant_type="client_credentials";scope="https://graph.microsoft.com/.default";client_id=$AppClientID;client_secret=$ClientSecretValue}
    Write-Host -ForegroundColor Blue -BackgroundColor white "Obtaining authentication token..." -NoNewline
    try{
        $oauth = Invoke-RestMethod -Method Post -Uri "$loginURL/$TenantDomain/oauth2/v2.0/token" -Body $body -ErrorAction Stop
        $script:tokenExpiresOn = (Get-Date).AddSeconds($oauth.expires_in).ToLocalTime()
        $script:GraphToken = "$($oauth.token_type) $($oauth.access_token)"
        Write-Host -ForegroundColor Green "Authentication token obtained"
    } catch {
        write-host -ForegroundColor Red "FAILED"
        write-host -ForegroundColor Red "Invoke-RestMethod failed."
        Write-host -ForegroundColor Red $error[0]
        exit
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
	Write-Host "Passed!" -ForeGroundColor Green
}

function connect2service($ReadExport)
{
	$ExportTo = $ReadExport
	
	if($ExportTo -eq 'File')
	{
		Write-Host "`nAuthentication is required, please check your browser" -ForegroundColor Green
		Connect-IPPSSession -UseRPSSession:$false
	}else
	{
		ValidateConfigurationFile
		
		$CONFIGFILE = "$PSScriptRoot\ConfigFiles\laconfig.json"
		$json = Get-Content -Raw -Path $CONFIGFILE
		[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
		
		$EncryptedKeys = $config.EncryptedKeys
		$AppClientID = $config.AppClientID
		$CertificateThumb = $config.CertificateThumb
		$OnmicrosoftTenant = $config.OnmicrosoftURL
		if ($EncryptedKeys -eq "True")
		{
			$CertificateThumb = DecryptSharedKey $CertificateThumb
		}
		
		<#
		.NOTES
		If you cannot add the "Compliance Administrator" role to the Microsoft Entra App, for security reasons, you can execute with "Compliance Administrator" role 
		this script using .\MPARR-ContentExplorer.ps1 -ManualConnection
		#>
		if($ManualConnection)
		{
			Write-Host "`nAuthentication is required, please check your browser" -ForegroundColor Green
			Connect-IPPSSession -UseRPSSession:$false
		}else
		{
			Connect-IPPSSession -CertificateThumbPrint $CertificateThumb -AppID $AppClientID -Organization $OnmicrosoftTenant
		}
	}
}

function connect2MicrosoftGraph
{		
	<#
	.NOTES
	If you cannot add the "Compliance Administrator" role to the Microsoft Entra App, for security reasons, you can execute with "Compliance Administrator" role 
	this script using .\MPARR-PurviewSensitivityLabels.ps1 -ManualConnection
	#>
	if($ManualConnection)
	{
		Write-Host "`nAuthentication is required, please check your browser" -ForegroundColor Green
		Import-Module Microsoft.Graph.Sites
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
			Write-Host ".\MPARR-SPOSites.ps1 -ManualConnection" -ForeGroundColor Green
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

function CreateMPARRContentExplorerTask
{
	# MPARR-ContentExplorerData script
    $taskName = "MPARR-ContentExplorerData"
	
	# Call function to set a folder for the task on Task Scheduler
	$taskFolder = CreateScheduledTaskFolder
	
	# Task execution
    $validDays = 30
    $choices  = '&Yes', '&No'
    $decision = $Host.UI.PromptForChoice("", "The task on task scheduler will be set for 30 days, do you want to change?", $choices, 1)
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
    $action = New-ScheduledTaskAction -Execute "`"$PSHOME\pwsh.exe`"" -Argument ".\MPARR-ContentExplorerData.ps1" -WorkingDirectory $PSScriptRoot
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
        Write-Information -MessageData "   $rows rows written to Log Analytics workspace $uri" -InformationAction Continue
    }

}

function ExportToJsonFiles
{
	<#
		.NOTES
		Trainable classifiers currently works with a Json file that is released with this script, 
		unfortunately doesn't exist yet a PowerShell cmdlet to obtain that data, 
		and the json for Trainable Classifiers is created manually
	#>
	
	cls
	Write-Host "`nJson files will be created to filter the data from Sensitivity Labels, Retention Labels and Sensitive Information Types." -ForeGroundColor DarkYellow
	Write-Host "Json files will be stored at $PSScriptRoot" -ForeGroundColor DarkYellow
	Write-Host "`nTo filter any of this kind of classifiers you need to change the value 'True' for 'False'" -ForeGroundColor DarkYellow
	New-Item -ItemType Directory -Force -Path "$PSScriptRoot\ConfigFiles" | Out-Null
	
	#Create Json for Sensitivity Labels
	$SensitivityLabels = Get-Label | select DisplayName,ParentLabelDisplayName
	$ListSensitivityLabels = @()
	
	foreach($label in $SensitivityLabels)
	{
		if($label.ParentLabelDisplayName -ne $Null)
		{
			$ListSensitivityLabels += $label.ParentLabelDisplayName+"/"+$label.DisplayName		
		}else
		{
			$ListSensitivityLabels += $label.DisplayName
		}
	}
	
	$tempFolder = $ListSensitivityLabels
	$results = @()
	$SortedResults = @()
	
	foreach ($label in $tempFolder){$results += @([pscustomobject]@{Name=$label})}
	Write-Host "`nTotal Sensitivity Labels found it :" -NoNewLine
	Write-Host "`t" $results.count -ForeGroundColor Green
	$SortedResults = $results | Sort-Object -Property Name -Unique
	
	$ArraySL = [ordered]@{}
	foreach($result in $results)
	{
		$ArraySL[$result.Name] = "True"
	}
	$ExportSL = "MPARR-SensitivityLabelsList.json"
	$pathSL = $PSScriptRoot+"\ConfigFiles\"+$ExportSL
	$jsonSL = $ArraySL | ConvertTo-Json
	$jsonSL | Set-Content -Path $pathSL
	Write-Host "`nA new configuration file was created at $pathSL"
	
	#Create Json for Sensitive Information Types
	$results = @()
	$SortedResults = @()
	$results = Get-DlpSensitiveInformationType | select Name
	$SortedResults = $results | Sort-Object -Property Name -Unique
	Start-Sleep -s 1
	Write-Host "`nTotal Sensitive Information Types found it :" -NoNewLine
	Write-Host "`t" $SortedResults.count -ForeGroundColor Green
	$ArraySIT = [ordered]@{}
	foreach($result in $SortedResults)
	{
		$ArraySIT[$result.Name] = "True"
	}
	$ExportSIT = "MPARR-SensitiveInfoTypesList.json"
	$pathSIT = $PSScriptRoot+"\ConfigFiles\"+$ExportSIT
	$jsonSIT = $ArraySIT | ConvertTo-Json
	$jsonSIT | Set-Content -Path $pathSIT
	Write-Host "`nA new configuration file was created at $pathSIT"
	
	#Create Json for Retention Labels
	$results = @()
	$SortedResults = @()
	$results = Get-ComplianceTag | select Name
	$SortedResults = $results | Sort-Object -Property Name -Unique
	Write-Host "`nTotal Retention Labels found it :" -NoNewLine
	Write-Host "`t" $SortedResults.count -ForeGroundColor Green
	$ArrayRL = [ordered]@{}
	foreach($result in $SortedResults)
	{
		$ArrayRL[$result.Name] = "True"
	}
	$ExportRL = "MPARR-RetentionLabelsList.json"
	$pathRL = $PSScriptRoot+"\ConfigFiles\"+$ExportRL
	$jsonRL = $ArrayRL | ConvertTo-Json
	$jsonRL | Set-Content -Path $pathRL
	Write-Host "`nA new configuration file was created at $pathRL"
	
	Start-Sleep -s 2
}

function GetM365AllSites($service)
{
	connect2MicrosoftGraph
	GetAuthToken
	
	if ($tokenExpiresOn.AddMinutes(5) -lt (Get-Date))
	{
		Write-Host "Refreshing access token..."
		GetAuthToken
	}

	$headers = @{
		'Content-Type' = 'application/json'
		Accept = 'application/json'
		Authorization = $GraphToken
	}
	
	$results = @()
	$tempArray = @()
	$GetSPOSitesResults = @()
	$GetODBSitesResults = @()
	
	$BaseURI = "https://graph.microsoft.com/v1.0/sites"
	$URI = "$BaseURI/getAllSites"

	# Run the cmdlet to get Sites
	$results = Invoke-RestMethod -Method Get -Uri $URI -Headers $headers -ErrorAction Stop
	$tempArray += $results.value
	
	foreach($item in $tempArray)
	{
		if($item.isPersonalSite -eq "true")
		{
			$GetODBSitesResults += $item.webUrl
		}else
		{
			$GetSPOSitesResults += $item.webUrl
		}
	}

	# Status update
	$recordsSPOCount = $GetSPOSitesResults.Count
	$recordsODBCount = $GetODBSitesResults.Count

	# If there is no data, skip
	if ($GetSPOSitesResults.Count -eq 0)
	{
		Write-Host "`nNo SharePoint Online Sites was found on your Tenant" -ForeGroundColor Yellow
		exit 
	}elseif ($GetODBSitesResults.Count -eq 0)
	{
		Write-Host "`nNo OneDrive for Business accounts was found on your Tenant" -ForeGroundColor Yellow
		exit 
	}
	
	if($service -eq "SharePoint")
	{
		Write-Information -MessageData "$recordsSPOCount records returned from SharePoint Online Sites" -InformationAction Continue
		return $GetSPOSitesResults
	}
	if($service -eq "OneDrive")
	{
		Write-Information -MessageData "$recordsODBCount records returned from OneDrive for Business" -InformationAction Continue
		return $GetODBSitesResults
	}

}

function GetM365Accounts($service)
{
	connect2MicrosoftGraph
	GetAuthToken
	
	if ($tokenExpiresOn.AddMinutes(5) -lt (Get-Date))
	{
		Write-Host "Refreshing access token..."
		GetAuthToken
	}

	$headers = @{
		'Content-Type' = 'application/json'
		Accept = 'application/json'
		Authorization = $GraphToken
	}
	
	$results = @()
	$tempArray = @()
	$GetExoM365Accounts = @()
	$GetTeamsM365Accounts = @()
	$ExoM365AccountsNotLicensed = @()
	$TeamsM365AccountsNotLicensed = @()
	
	$ReportName = "Office365ActiveUserDetail"
	$BaseURI = "https://graph.microsoft.com/v1.0/reports"
	$URI = "$BaseURI/get$($ReportName)"
	$URI += "(period='D7')"

	# Run the cmdlet to get Sites
	$results = Invoke-RestMethod -Method Get -Uri $URI -Headers $headers -ErrorAction Stop
	$tempArray += $results | ConvertFrom-Csv
	$EXOLicense = "Has Exchange License"
	$TeamsLicense = "Has Teams License"
	$M365UPN = "User Principal Name"
	
	
	foreach($item in $tempArray)
	{
		if($item.$EXOLicense -eq "true")
		{
			$GetExoM365Accounts += $item."User Principal Name"
		}else
		{
			$ExoM365AccountsNotLicensed += $item."User Principal Name"
		}
		if($item.$TeamsLicense -eq "true")
		{
			$GetTeamsM365Accounts += $item."User Principal Name"
		}else
		{
			$TeamsM365AccountsNotLicensed += $item."User Principal Name"
		}
	}

	# Status update
	$recordsEXOCount = $GetExoM365Accounts.Count
	$recordsTeamsCount = $GetTeamsM365Accounts.Count

	# If there is no data, skip
	if ($GetExoM365Accounts.Count -eq 0)
	{
		Write-Host "`nNo Exchange Online licensed accounts was found on your Tenant" -ForeGroundColor Yellow
		exit 
	}elseif ($GetTeamsM365Accounts.Count -eq 0)
	{
		Write-Host "`nNo Microsoft Teams licensed accounts was found on your Tenant" -ForeGroundColor Yellow
		exit 
	}
	
	if($service -eq "Exchange")
	{
		Write-Host "$recordsEXOCount licensed accounts returned from Exchange Online"
		return $GetExoM365Accounts		
	}elseif($service -eq "Teams")
	{
		Write-Host "$recordsTeamsCount licensed accounts returned from Microsoft Teams"
		return $GetTeamsM365Accounts
	}

}

function ReadWorkload($ReadExport)
{
	$ExportTo = $ReadExport
	
	if($ExportTo -eq 'File')
	{
		$choices  = '&Exchange','&SharePoint','&OneDrive','&Teams'
		$decision = $Host.UI.PromptForChoice("", "`nPlease select the service that you want to use in your query", $choices, 1)
		if ($decision -eq 0)
		{
			$workload = 'Exchange'
			return $workload
		}
		if ($decision -eq 1)
		{
			$workload = 'SharePoint'
			return $workload
		}
		if ($decision -eq 2)
		{
			$workload = 'OneDrive'
			return $workload
		}
		if ($decision -eq 3)
		{
			$workload = 'Teams'
			return $workload
		}
	}else
	{
		$ContentExplorerWorkload = "$PSScriptRoot\ConfigFiles\MPARR-CEWorkload.json"
		$workload = @("Exchange","OneDrive","SharePoint","Teams")
		
		if (-not (Test-Path -Path $ContentExplorerWorkload))
		{
			Write-Host "MPARR-CEWorkload file is missing. All workloads will be used."
			return $workload
		}
		else
		{
			$workload = @()
			$json = Get-Content -Raw -Path $ContentExplorerWorkload
			[PSCustomObject]$workloads = ConvertFrom-Json -InputObject $json
			foreach ($service in $workloads.psobject.Properties)
			{
				if ($workloads."$($service.Name)" -eq "True")
				{
					$workload += $service.Name
				}
			}
			return $workload
		}

	}
}

function ReadTagType($ReadExport)
{
	$ExportTo = $ReadExport
	
	if($ExportTo -eq 'File')
	{
		cls
		$choices  = '&Retention Labels','Sensitive &Information Type','&Sensitivity Labels','&Trainable Classifiers'
		$decision = $Host.UI.PromptForChoice("", "`nPlease select the classifier that you want to use in your query :", $choices, 2)
		if ($decision -eq 0)
		{
			$TagType = 'Retention'
			return $TagType
		}
		if ($decision -eq 1)
		{
			$TagType = 'SensitiveInformationType'
			return $TagType
		}
		if ($decision -eq 2)
		{
			$TagType = 'Sensitivity'
			return $TagType
		}
		if ($decision -eq 3)
		{
			$TagType = 'TrainableClassifier'
			return $TagType
		}
	}else
	{
		$ContentExplorerTagType = "$PSScriptRoot\ConfigFiles\MPARR-CETagtype.json"
		$TagType = @("Retention","Sensitivity","SensitiveInformationType","TrainableClassifier")
		
		if (-not (Test-Path -Path $ContentExplorerTagType))
		{
			Write-Host "MPARR-CEWorkload file is missing. All workloads will be used."
			return $TagType
		}
		else
		{
			$TagType = @()
			$json = Get-Content -Raw -Path $ContentExplorerTagType
			[PSCustomObject]$tagtypes = ConvertFrom-Json -InputObject $json
			foreach ($tag in $tagtypes.psobject.Properties)
			{
				if ($tagtypes."$($tag.Name)" -eq "True")
				{
					$TagType += $tag.Name
				}
			}
			return $TagType
		}
	}
}

function GetSensitivityLabelList
{
	Write-Host "`nGetting Sensitivity Labels..." -ForegroundColor Green
	Write-Host "`nThe list can be long, check your PowerShell buffer and set at least on 500." -ForeGroundColor DarkYellow
	$SensitivityLabels = Get-Label | select DisplayName,ParentLabelDisplayName
	$ListSensitivityLabels = @()
	
	foreach($label in $SensitivityLabels)
	{
		if($label.ParentLabelDisplayName -ne $Null)
		{
			$ListSensitivityLabels += $label.ParentLabelDisplayName+"/"+$label.DisplayName		
		}else
		{
			$ListSensitivityLabels += $label.DisplayName
		}
	}
	
	$tempFolder = $ListSensitivityLabels
	$SensitivityLabelsSelection = @()
	
	foreach ($label in $tempFolder){$SensitivityLabelsSelection += @([pscustomobject]@{Name=$label})}
	
	$i = 1
    $SensitivityLabelsSelection = @($SensitivityLabelsSelection| ForEach-Object {$_ | Add-Member -Name "No" -MemberType NoteProperty -Value ($i++) -PassThru})
	
	#List all existing folders under Task Scheduler
    $SensitivityLabelsSelection | Select-Object No, Name | Out-Host
	
	# Select label
    $selection = 0
    ReadNumber -max ($i -1) -msg "Enter number corresponding to the Sensitivity Label name" -option ([ref]$selection)
    $LabelSelected = $SensitivityLabelsSelection[$selection - 1].Name
	
	return $LabelSelected
}

function GetRetentionLabelList
{
	Write-Host "`nGetting Retention Labels..." -ForegroundColor Green
	Write-Host "`nThe list can be long, check your PowerShell buffer and set at least on 500." -ForeGroundColor DarkYellow
	$RetentionLabels = Get-ComplianceTag | select Name
	$ListRetentionLabels = @()
	
	foreach($label in $RetentionLabels)
	{
		$ListRetentionLabels += $label.Name
	}
	
	$tempFolder = $ListRetentionLabels
	$RetentionLabelsSelection = @()
	
	foreach ($label in $tempFolder){$RetentionLabelsSelection += @([pscustomobject]@{Name=$label})}
	
	$i = 1
    $RetentionLabelsSelection = @($RetentionLabelsSelection| ForEach-Object {$_ | Add-Member -Name "No" -MemberType NoteProperty -Value ($i++) -PassThru})
	
	#List all existing folders under Task Scheduler
    $RetentionLabelsSelection | Select-Object No, Name | Out-Host
	
	# Select label
    $selection = 0
    ReadNumber -max ($i -1) -msg "Enter number corresponding to the Retention Label name" -option ([ref]$selection)
    $LabelSelected = $RetentionLabelsSelection[$selection - 1].Name
	
	return $LabelSelected
}

function GetSensitiveInformationType
{
	$choices  = '&Enter Name','&Select from a list'
	$decision = $Host.UI.PromptForChoice("", "`nPlease, select how to you want identify the Sensitive Information Type to be used in your query", $choices, 0)
	if ($decision -eq 0)
    {
		$SIT = Read-Host "`nPlease enter the Sensitive Information Type name that will be used in your query "
		return $SIT
	}
	if ($decision -eq 1)
    {
	
		Write-Host "`nGetting Sensitive Information Types..." -ForegroundColor Green
		Write-Host "`nThe list can be long, check your PowerShell buffer and set at least on 500." -ForeGroundColor DarkYellow
		$SITs = Get-DlpSensitiveInformationType | select Name
		$SITlist = @()
		
		foreach($SITd in $SITs)
		{
			$SITlist += $SITd.Name
		}
		
		$tempFolder = $SITlist
		$SITSelection = @()
		
		foreach ($SITd in $tempFolder){$SITSelection += @([pscustomobject]@{Name=$SITd})}
		
		$i = 1
		$SITSelection = @($SITSelection| ForEach-Object {$_ | Add-Member -Name "No" -MemberType NoteProperty -Value ($i++) -PassThru})
		
		#List all SITs
		$SITSelection | Select-Object No, Name | Out-Host
		
		# Select Sensitive Information Type
		$selection = 0
		ReadNumber -max ($i -1) -msg "Enter number corresponding to the Sensitive Information Type name" -option ([ref]$selection)
		$SIT = $SITSelection[$selection - 1].Name
		
		return $SIT
	}
}

function GetTrainableClassifiers
{
	$choices  = '&Enter Name','&Select from a list'
	$decision = $Host.UI.PromptForChoice("", "`nPlease, select how to you want identify the Trainable Classifier to be used in your query", $choices, 0)
	if ($decision -eq 0)
    {
		$TC = Read-Host "`nPlease enter the Trainable Classifier name that will be used in your query "
		return $TC
	}
	if ($decision -eq 1)
    {
		$TCSelected = "$PSScriptRoot\ConfigFiles\MPARR-TrainableClassifiersList.json"
		
		if (-not (Test-Path -Path $TCSelected))
		{
			Write-Host "`nThe file MPARR_TrainableClassifiersList.json is missing at $PSScriptRoot." -ForeGroundColor DarkYellow
			Write-Host "You can found it in the GitHub site at https://aka.ms/MPARR-GitHub"
			GetTrainableClassifiers
		}else
		{
			Write-Host "`nGetting Trainable Classifiers..." -ForegroundColor Green
			Write-Host "`nThe list can be long, check your PowerShell buffer and set at least on 500." -ForeGroundColor DarkYellow
			
			$json = Get-Content -Raw -Path $TCSelected
			[PSCustomObject]$tcs = ConvertFrom-Json -InputObject $json
			$TClist = @()
			
			foreach ($tcd in $tcs.psobject.Properties)
			{
				if ($tcs."$($tcd.Name)" -eq "True")
				{
					$TClist += $tcd.Name
				}
			}
			
			$tempFolder = $TClist
			$TCSelection = @()
			
			foreach ($tcd in $tempFolder){$TCSelection += @([pscustomobject]@{Name=$tcd})}
			
			$i = 1
			$TCSelection = @($TCSelection| ForEach-Object {$_ | Add-Member -Name "No" -MemberType NoteProperty -Value ($i++) -PassThru})
			
			#List all Trainable classifiers
			$TCSelection | Select-Object No, Name | Out-Host
			
			# Select Trainable Classifier
			$selection = 0
			ReadNumber -max ($i -1) -msg "Enter number corresponding to the Trainable Classifier name" -option ([ref]$selection)
			$TC = $TCSelection[$selection - 1].Name
			
			return $TC
		}
	}
}

function ExportPageSize($PageSize)
{
	$Size = $PageSize

	$choices  = '&Yes', '&No'
    $decision = $Host.UI.PromptForChoice("", "`nThe default Page Size used in your query is: '$($Size)', do you want to change?", $choices, 0)
    if ($decision -eq 0)
    {
        ReadNumber -max 5000 -msg "Enter a page size number (Between 1 to 5000)." -option ([ref]$Size)
		return $Size
    }
	
	return $Size
}

function ExecuteExportCmdlet($TagType, $Workload, $Tag, $PageSize)
{
	if($Workload -eq "SharePoint")
	{
		$DetailedData = GetM365AllSites -service $Workload
		$DetailedDataCount = $DetailedData.count
		Write-Host "Records found for $Workload :" -NoNewline
		Write-Host "`t$DetailedDataCount" -ForeGroundColor Green
	}
	if($Workload -eq "OneDrive")
	{
		$DetailedData = GetM365AllSites -service $Workload
		$DetailedDataCount = $DetailedData.count
		Write-Host "Records found for $Workload :" -NoNewline
		Write-Host "`t$DetailedDataCount" -ForeGroundColor Green
	}
	if($Workload -eq "Exchange")
	{
		$DetailedData = GetM365Accounts -service $Workload
		$DetailedDataCount = $DetailedData.count
		Write-Host "Records found for $Workload :" -NoNewline
		Write-Host "`t$DetailedDataCount" -ForeGroundColor Green
	}
	if($Workload -eq "Teams")
	{
		$DetailedData = GetM365Accounts -service $Workload
		$DetailedDataCount = $DetailedData.count
		Write-Host "Records found for $Workload :" -NoNewline
		Write-Host "`t$DetailedDataCount" -ForeGroundColor Green
	}
	
	if($TagType -eq 'Sensitivity')
	{
		$tagname = $tag.replace('/','-')
	}else
	{
		$tagname = $tag
	}
	
	$date = Get-Date -Format "yyyyMMddHHmm"
	$ExportFile = "ContentExplorerExport - "+$TagType+" - "+$tagname+" - "+$Workload+" - "+$date+".csv"
	$date2 = Get-Date -Format "yyyyMMdd"
	$ExportError = "ContentExplorerExport-Error"+$date2+".csv"
	$ExportSummary = "ContentExplorerExport-Summary"+$date2+".csv"
	$path = $PSScriptRoot+"\ContentExplorerExport\"+$ExportFile
	
	foreach($GranularValue in $DetailedData)
	{
		$CEResults = @()
		if($Workload -eq "SharePoint")
		{
			$query = Export-ContentExplorerData -TagType $TagType -TagName $tag -PageSize $PageSize -Workload $Workload -SiteUrl $GranularValue
			$CmdletUsed = "Export-ContentExplorerData -TagType $TagType -TagName '$($tag)' -PageSize $PageSize -Workload $Workload -SiteUrl $GranularValue"
		}
		if($Workload -eq "OneDrive")
		{
			$query = Export-ContentExplorerData -TagType $TagType -TagName $tag -PageSize $PageSize -Workload $Workload -SiteUrl $GranularValue
			$CmdletUsed = "Export-ContentExplorerData -TagType $TagType -TagName '$($tag)' -PageSize $PageSize -Workload $Workload -SiteUrl $GranularValue"
		}
		if($Workload -eq "Exchange")
		{
			$query = Export-ContentExplorerData -TagType $TagType -TagName $tag -PageSize $PageSize -Workload $Workload -UserPrincipalName $GranularValue
			$CmdletUsed = "Export-ContentExplorerData -TagType $TagType -TagName '$($tag)' -PageSize $PageSize -Workload $Workload -UserPrincipalName $GranularValue"
		}
		if($Workload -eq "Teams")
		{
			$query = Export-ContentExplorerData -TagType $TagType -TagName $tag -PageSize $PageSize -Workload $Workload -UserPrincipalName $GranularValue
			$CmdletUsed = "Export-ContentExplorerData -TagType $TagType -TagName '$($tag)' -PageSize $PageSize -Workload $Workload -UserPrincipalName $GranularValue"
		}
		Write-Host "Cmdlet used : " -NoNewLine
		Write-Host "$CmdletUsed" -ForeGroundColor Green
		$TotalResults = ($query.count) - 1
		Write-Host "Total returned : $TotalResults"
		
		$var = $query.count
		$Total = $query[0].TotalCount
		$TotalExported = 0
		$remaining = $Total
		$ErrorExportArray = @()
		$SummaryExportArray = @()
		
		if($Total -eq 0)
		{
			Write-Host "`n### Your query don't returned records. ###" -ForeGroundColor Blue
			Write-Host "Query tested with:"
			Write-Host "Service `t: "$Workload
			Write-Host "Classifier type : "$TagType
			Write-Host "Classifier name : "$tag
			Write-Host "Value used`t:"$GranularValue
			Write-Host "### File was not created." -ForeGroundColor Blue
			$path2 = $PSScriptRoot+"\ContentExplorerExport\"+$ExportError
			$ErrorExportArray = @(
				[pscustomobject]@{TagType=$TagType;TagName=$tag;Workload=$Workload;ExportedFiles=$Total;TotalMatches=$Total;CmdletUsed=$CmdletUsed}
			)
			$ErrorExportArray | Export-Csv -Path $path2 -Force -Append | Out-Null
		}else
		{
			Write-Host "Total matches returned :" -NoNewLine
			Write-Host $remaining -ForeGroundColor Green	
		}

		While ($query[0].MorePagesAvailable -eq 'True') {
			$CEResults += $query[1..$var]
			if($Workload -eq "SharePoint")
			{
				$query = Export-ContentExplorerData -TagType $TagType -TagName $tag -PageSize $PageSize -Workload $Workload -SiteUrl $GranularValue -PageCookie $query[0].PageCookie
				$CmdletUsed = "Export-ContentExplorerData -TagType $TagType -TagName '$($tag)' -PageSize $PageSize -Workload $Workload -SiteUrl $GranularValue"
			}
			if($Workload -eq "OneDrive")
			{
				$query = Export-ContentExplorerData -TagType $TagType -TagName $tag -PageSize $PageSize -Workload $Workload -SiteUrl $GranularValue -PageCookie $query[0].PageCookie
				$CmdletUsed = "Export-ContentExplorerData -TagType $TagType -TagName '$($tag)' -PageSize $PageSize -Workload $Workload -SiteUrl $GranularValue"
			}
			if($Workload -eq "Exchange")
			{
				$query = Export-ContentExplorerData -TagType $TagType -TagName $tag -PageSize $PageSize -Workload $Workload -UserPrincipalName $GranularValue -PageCookie $query[0].PageCookie
				$CmdletUsed = "Export-ContentExplorerData -TagType $TagType -TagName '$($tag)' -PageSize $PageSize -Workload $Workload -UserPrincipalName $GranularValue"
			}
			if($Workload -eq "Teams")
			{
				$query = Export-ContentExplorerData -TagType $TagType -TagName $tag -PageSize $PageSize -Workload $Workload -UserPrincipalName $GranularValue -PageCookie $query[0].PageCookie
				$CmdletUsed = "Export-ContentExplorerData -TagType $TagType -TagName '$($tag)' -PageSize $PageSize -Workload $Workload -UserPrincipalName $GranularValue"
			}
			
			$remaining -= ($var - 1)
			Write-Host "Total matches remaining to process :" -NoNewLine
			trap { 'Error processing json, it is continue processing...'; continue }
			Write-Host $remaining -ForeGroundColor Green
			$TotalExported += ($query.count - 1)
			$CEResults | Export-Csv -Path $path -NTI -Force -Append | Out-Null
			$CEResults = @()
		}

		if ($query.count -gt 0)
		{
			$CEResults += $query[1..$remaining]
			$TotalExported += ($query.count - 1)
			$CEResults | Export-Csv -Path $path -NTI -Force -Append | Out-Null
		}
	}
	
	#Generate a summary with the total results
	$pathsummary = $PSScriptRoot+"\ContentExplorerExport\"+$ExportSummary
	$SummaryExportArray = @(
		[pscustomobject]@{TagType=$TagType;TagName=$tag;Workload=$Workload;MatchedFiles=$Total;ExportedFiles=$TotalExported;FileName=$ExportFile;CmdletUsed=$CmdletUsed}
	)
	$SummaryExportArray | Export-Csv -Path $pathsummary -Force -Append
}

function ExportContentExplorerDetailsTo($TagType, $Workload, $Tag, $PageSize, $GranularValue, $DataConnector)
{
	#Generate the query to collect the data
	Write-Host "`nData Connector set to: $DataConnector `n" -ForeGroundColor DarkBlue
	$date2 = Get-Date -Format "yyyyMMdd"
	$ExportError = "ContentExplorerExport-ErrorLogsAnalytics-"+$date2+".csv"
	$ExportSummary = "ContentExplorerExport-SummaryLogsAnalytics"+$date2+".csv"
	$ExportFile = "ContentExplorerExport - "+$TagType+" - "+$Tag+" - "+$Workload+" - "+$date2+".csv"
	$ExportJSONFile = "ContentExplorerExport - "+$TagType+" - "+$Tag+" - "+$Workload+" - "+$date2+".json"
	$path = $PSScriptRoot+"\ContentExplorerExport\"+$ExportFile
	$pathJSON = $PSScriptRoot+"\ContentExplorerExport\"+$ExportJSONFile
	$CEResults = @()
	Write-Host "Getting data from... "$GranularValue -ForeGroundColor Magenta
	if($Workload -eq "SharePoint")
	{
		$query = Export-ContentExplorerData -TagType $TagType -TagName $tag -PageSize $PageSize -Workload $Workload -SiteUrl $GranularValue
		$CmdletUsed = "Export-ContentExplorerData -TagType $TagType -TagName '$($tag)' -PageSize $PageSize -Workload $Workload -SiteUrl $GranularValue"
	}
	if($Workload -eq "OneDrive")
	{
		$query = Export-ContentExplorerData -TagType $TagType -TagName $tag -PageSize $PageSize -Workload $Workload -SiteUrl $GranularValue
		$CmdletUsed = "Export-ContentExplorerData -TagType $TagType -TagName '$($tag)' -PageSize $PageSize -Workload $Workload -SiteUrl $GranularValue"
	}
	if($Workload -eq "Exchange")
	{
		$query = Export-ContentExplorerData -TagType $TagType -TagName $tag -PageSize $PageSize -Workload $Workload -UserPrincipalName $GranularValue
		$CmdletUsed = "Export-ContentExplorerData -TagType $TagType -TagName '$($tag)' -PageSize $PageSize -Workload $Workload -UserPrincipalName $GranularValue"
	}
	if($Workload -eq "Teams")
	{
		$query = Export-ContentExplorerData -TagType $TagType -TagName $tag -PageSize $PageSize -Workload $Workload -UserPrincipalName $GranularValue
		$CmdletUsed = "Export-ContentExplorerData -TagType $TagType -TagName '$($tag)' -PageSize $PageSize -Workload $Workload -UserPrincipalName $GranularValue"
	}
	
	Write-Host "Cmdlet used : "$CmdletUsed -ForeGroundColor Green
	
	$var = $query.count
	$Total = $query[0].TotalCount
	$TotalExported = 0
	$remaining = $Total
	$ErrorExportArray = @()
	$SummaryExportArray = @()
	$pathsummary = $PSScriptRoot+"\ContentExplorerExport\"+$ExportSummary
	
	#Add additional columns to simplify reports
	$i = 1
	While($i -lt $var)
	{
		$query[$i] | Add-Member -MemberType NoteProperty -Name 'TagType' -Value $TagType
		$query[$i] | Add-Member -MemberType NoteProperty -Name 'TagName' -Value $tag
		$i++
	}
	
	if($Total -eq 0)
	{
		$path2 = $PSScriptRoot+"\ContentExplorerExport\"+$ExportError
		$ErrorExportArray = @(
			[pscustomobject]@{TagType=$TagType;TagName=$tag;Workload=$Workload;ExportedFiles=$Total;TotalMatches=$Total;CmdletUsed=$CmdletUsed}
		)
		$ErrorExportArray | Export-Csv -Path $path2 -Force -Append | Out-Null
		return
	}else
	{
		Write-Host "Total matches returned :" -NoNewLine
		Write-Host $remaining -ForeGroundColor Green	
	}

	While ($query[0].MorePagesAvailable -eq 'True') {
		$CEResults += $query[1..$var]
		if($Workload -eq "SharePoint")
		{
			$query = Export-ContentExplorerData -TagType $TagType -TagName $tag -PageSize $PageSize -Workload $Workload -SiteUrl $GranularValue -PageCookie $query[0].PageCookie
			$CmdletUsed = "Export-ContentExplorerData -TagType $TagType -TagName '$($tag)' -PageSize $PageSize -Workload $Workload -SiteUrl $GranularValue"
		}
		if($Workload -eq "OneDrive")
		{
			$query = Export-ContentExplorerData -TagType $TagType -TagName $tag -PageSize $PageSize -Workload $Workload -SiteUrl $GranularValue -PageCookie $query[0].PageCookie
			$CmdletUsed = "Export-ContentExplorerData -TagType $TagType -TagName '$($tag)' -PageSize $PageSize -Workload $Workload -SiteUrl $GranularValue"
		}
		if($Workload -eq "Exchange")
		{
			$query = Export-ContentExplorerData -TagType $TagType -TagName $tag -PageSize $PageSize -Workload $Workload -UserPrincipalName $GranularValue -PageCookie $query[0].PageCookie
			$CmdletUsed = "Export-ContentExplorerData -TagType $TagType -TagName '$($tag)' -PageSize $PageSize -Workload $Workload -UserPrincipalName $GranularValue"
		}
		if($Workload -eq "Teams")
		{
			$query = Export-ContentExplorerData -TagType $TagType -TagName $tag -PageSize $PageSize -Workload $Workload -UserPrincipalName $GranularValue -PageCookie $query[0].PageCookie
			$CmdletUsed = "Export-ContentExplorerData -TagType $TagType -TagName '$($tag)' -PageSize $PageSize -Workload $Workload -UserPrincipalName $GranularValue"
		}

		$i = 1
		While($i -lt $query.count)
		{
			$query[$i] | Add-Member -MemberType NoteProperty -Name 'TagType' -Value $TagType
			$query[$i] | Add-Member -MemberType NoteProperty -Name 'TagName' -Value $tag
			$i++
		}
		$remaining -= ($var - 1)
		trap { 'Error processing json, it is continue processing...'; continue }
		Write-Host "Total matches remaining to process :" -NoNewLine
		Write-Host $remaining -ForeGroundColor Green
		$TotalExported += ($query.count - 1)
	}

	if ($query.count -gt 0)
	{
		$CEResults += $query[1..$remaining]
	}

	# Push data to Log Analytics
	if($Workload -eq 'Exchange')
	{
		$TableLA = $TableName+"_EXO"
		# Else format for Log Analytics
        $log_analytics_array = @()            
        foreach($i in $CEResults) 
		{
			$log_analytics_array += $i
        }    

        if($DataConnector -eq "Event Hub")
		{
			$EventHubInstance.PublishToEventHub($log_analytics_array, $ErrorFile)
		}elseif($DataConnector -eq "MassCSVExport")
		{
			$CEResults | Export-Csv -Path $path -NTI -Force -Append | Out-Null
			Write-Host "Data exported to..." -NoNewline
			Write-Host "`n$path" -ForeGroundColor Cyan
		}elseif($DataConnector -eq "MassJsonExport")
		{
			$json = $log_analytics_array | ConvertTo-Json -Depth 3
			$json | Add-Content -Path $pathJSON
			Write-Host "`nData exported to... :" -NoNewLine
			Write-Host $pathJSON -ForeGroundColor Cyan
		}else
		{
			Post-LogAnalyticsData -LogAnalyticsTableName $TableLA -body $log_analytics_array
		}
		
		#Generate a summary with the total results
		$SummaryExportArray = @(
			[pscustomobject]@{TagType=$TagType;TagName=$tag;Workload=$Workload;MatchedFiles=$Total;ExportedFiles=$log_analytics_array.count;TableName=$TableLA;CmdletUsed=$CmdletUsed}
		)
		$SummaryExportArray | Export-Csv -Path $pathsummary -Force -Append
	}
	if($Workload -eq 'SharePoint')
	{
		$TableLA = $TableName+"_SPO"
		# Else format for Log Analytics
        $log_analytics_array = @()            
        foreach($i in $CEResults) 
		{
			$log_analytics_array += $i
        }    

        if($DataConnector -eq "Event Hub")
		{
			$EventHubInstance.PublishToEventHub($log_analytics_array, $ErrorFile)
		}elseif($DataConnector -eq "MassCSVExport")
		{
			$CEResults | Export-Csv -Path $path -NTI -Force -Append | Out-Null
		}elseif($DataConnector -eq "MassJsonExport")
		{
			$json = $log_analytics_array | ConvertTo-Json -Depth 3
			$json | Add-Content -Path $pathJSON
			Write-Host "`nData exported to... :" -NoNewLine
			Write-Host $pathJSON -ForeGroundColor Cyan
		}else
		{
			Post-LogAnalyticsData -LogAnalyticsTableName $TableLA -body $log_analytics_array
		}
		
		#Generate a summary with the total results
		$SummaryExportArray = @(
			[pscustomobject]@{TagType=$TagType;TagName=$tag;Workload=$Workload;MatchedFiles=$Total;ExportedFiles=$log_analytics_array.count;TableName=$TableLA;CmdletUsed=$CmdletUsed}
		)
		$SummaryExportArray | Export-Csv -Path $pathsummary -Force -Append
	}
	if($Workload -eq 'OneDrive')
	{
		$TableLA = $TableName+"_ODB"
		# Else format for Log Analytics
        $log_analytics_array = @()            
        foreach($i in $CEResults) 
		{
			$log_analytics_array += $i
        }    

        if($DataConnector -eq "Event Hub")
		{
			$EventHubInstance.PublishToEventHub($log_analytics_array, $ErrorFile)
		}elseif($DataConnector -eq "MassCSVExport")
		{
			$CEResults | Export-Csv -Path $path -NTI -Force -Append | Out-Null
		}elseif($DataConnector -eq "MassJsonExport")
		{
			$json = $log_analytics_array | ConvertTo-Json -Depth 3
			$json | Add-Content -Path $pathJSON
			Write-Host "`nData exported to... :" -NoNewLine
			Write-Host $pathJSON -ForeGroundColor Cyan
		}else
		{
			Post-LogAnalyticsData -LogAnalyticsTableName $TableLA -body $log_analytics_array
		}
		
		#Generate a summary with the total results
		$SummaryExportArray = @(
			[pscustomobject]@{TagType=$TagType;TagName=$tag;Workload=$Workload;MatchedFiles=$Total;ExportedFiles=$log_analytics_array.count;TableName=$TableLA;CmdletUsed=$CmdletUsed}
		)
		$SummaryExportArray | Export-Csv -Path $pathsummary -Force -Append
	}
	if($Workload -eq 'Teams')
	{
		$TableLA = $TableName+"_Teams"
		# Else format for Log Analytics
        $log_analytics_array = @()            
        foreach($i in $CEResults) 
		{
			$log_analytics_array += $i
        }    

        if($DataConnector -eq "Event Hub")
		{
			$EventHubInstance.PublishToEventHub($log_analytics_array, $ErrorFile)
		}elseif($DataConnector -eq "MassCSVExport")
		{
			$CEResults | Export-Csv -Path $path -NTI -Force -Append | Out-Null
		}elseif($DataConnector -eq "MassJsonExport")
		{
			$json = $log_analytics_array | ConvertTo-Json -Depth 3
			$json | Add-Content -Path $pathJSON
			Write-Host "`nData exported to... :" -NoNewLine
			Write-Host $pathJSON -ForeGroundColor Cyan
		}else
		{
			Post-LogAnalyticsData -LogAnalyticsTableName $TableLA -body $log_analytics_array
		}
		#Generate a summary with the total results
		$SummaryExportArray = @(
			[pscustomobject]@{TagType=$TagType;TagName=$tag;Workload=$Workload;MatchedFiles=$Total;ExportedFiles=$log_analytics_array.count;TableName=$TableLA;CmdletUsed=$CmdletUsed}
		)
		$SummaryExportArray | Export-Csv -Path $pathsummary -Force -Append
	}
}

function CollectData($TagType, $Workload, $PageSize, $ReadExport, $ExportDataTo)
{
	$ExportTo = $ReadExport
	$Connector = $ExportDataTo
	
	if($ExportTo -eq 'File')
	{
		#Step 1: Collect all the variables	
		if($TagType -contains 'Retention')
		{
			$RetentionLabels = GetRetentionLabelList
			$textvalue = $RetentionLabels
			$tagname = $RetentionLabels
		}
		if($TagType -contains 'SensitiveInformationType')
		{
			$SITs = GetSensitiveInformationType
			$textvalue = $SITs
			$tagname = $SITs
		}
		if($TagType -contains 'Sensitivity')
		{
			$SensitivityLabels = GetSensitivityLabelList
			$textvalue = $SensitivityLabels.replace('/','-')
			$tagname = $SensitivityLabels
		}
		if($TagType -contains 'TrainableClassifier')
		{
			$TrainableClassifiers = GetTrainableClassifiers
			$textvalue = $TrainableClassifiers
			$tagname = $TrainableClassifiers
		}
		
		# Set the default configuration for Export-ContentExplorer
		$PageSize = $PageSize
		
		#Step 2: Show the configuration set
		cls
		Write-Host "`n#################################################################################"
		Write-Host "`t`t`tConfiguration Set:"
		Write-Host "`nTag Types selected:" -NoNewLine
			Write-Host "`t`t`t"$TagType -ForegroundColor Green
		Write-Host "Workloads selected:" -NoNewLine
			Write-Host "`t`t`t"$Workload -ForegroundColor Green
		if($TagType -contains 'SensitiveInformationType')
		{
			Write-Host "Sensitive Information Type selected:" -NoNewLine
			Write-Host "`t'$($SITs)' " -ForegroundColor Green
		}
		if($TagType -contains 'Sensitivity')
		{
			Write-Host "Sensitivity Labels selected:" -NoNewLine
			Write-Host "`t`t'$($SensitivityLabels)' " -ForegroundColor Green
		}
		if($TagType -contains 'Retention')
		{
			Write-Host "Retention Labels selected:" -NoNewLine
			Write-Host "`t`t'$($RetentionLabels)' " -ForegroundColor Green
		}
		if($TagType -contains 'TrainableClassifier')
		{
			Write-Host "Trainable Classifier selected:" -NoNewLine
			Write-Host "`t`t'$($TrainableClassifiers)' " -ForegroundColor Green
		}
		Write-Host "Page size set:" -NoNewLine
			Write-Host "`t`t`t`t"$PageSize -ForegroundColor Green
		Write-Host "`n#################################################################################"
		
		$ExportPath = $PSScriptRoot+"\ContentExplorerExport"
		if(-Not (Test-Path $ExportPath ))
		{
			Write-Host "Export data directory is missing, creating a new folder called ContentExplorerExport"
			New-Item -ItemType Directory -Force -Path "$PSScriptRoot\ContentExplorerExport" | Out-Null
		}
		
		ExecuteExportCmdlet -TagType $TagType -Workload $Workload -Tag $tagname -PageSize $PageSize
	}else
	{
		# Set the default configuration for Export-ContentExplorer
		$PageSize = $PageSize
		$ExportPath = $PSScriptRoot+"\ContentExplorerExport"
		if(-Not (Test-Path $ExportPath ))
		{
			Write-Host "Export data directory is missing, creating a new folder called ContentExplorerExport"
			New-Item -ItemType Directory -Force -Path "$PSScriptRoot\ContentExplorerExport" | Out-Null
		}
		
		#Step 2: Show the configuration set
		cls
		Write-Host "`n#################################################################################"
		Write-Host "`t`t`tConfiguration Set:"
		Write-Host "`nTag Types selected:"
		foreach($tag in $TagType)
		{
			Write-Host "`t"$tag -ForeGroundColor Green
		}
		Write-Host "Workloads selected:"
		foreach($service in $Workload)
		{
			Write-Host "`t"$service -ForeGroundColor Green
		}
		Write-Host "Page size set:"
		Write-Host "`t"$PageSize -ForegroundColor Green
		Write-Host "`n#################################################################################"
		Write-Host "`n"
		
		#Initiate arrays
		
		foreach($service in $Workload)
		{
			$DetailedData = @()
			$DetailedDataCount = 0
			$Counter = 1
			if($service -eq "SharePoint")
			{
				$DetailedData = GetM365AllSites -service $service
				$DetailedDataCount = $DetailedData.count
			}
			if($service -eq "OneDrive")
			{
				$DetailedData = GetM365AllSites -service $service
				$DetailedDataCount = $DetailedData.count
			}
			if($service -eq "Exchange")
			{
				$DetailedData = GetM365Accounts -service $service
				$DetailedDataCount = $DetailedData.count 
			}
			if($service -eq "Teams")
			{
				$DetailedData = GetM365Accounts -service $service
				$DetailedDataCount = $DetailedData.count
			}
			
			foreach($value in $DetailedData)
			{
				Write-Host "Progress..." -NoNewLine
				Write-Host "$Counter of $DetailedDataCount" -ForeGroundColor Blue
				foreach($tag in $TagType)
				{
					if($tag -eq 'Retention')
					{
						$tag = $tag
						
						if (Test-Path "$PSScriptRoot\ConfigFiles\MPARR-RetentionLabelsList.json")
						{
							$RetentionSelected = "$PSScriptRoot\ConfigFiles\MPARR-RetentionLabelsList.json"
						
							$jsonRL = Get-Content -Raw -Path $RetentionSelected
							[PSCustomObject]$rls = ConvertFrom-Json -InputObject $jsonRL
							$RetentionLabels = @()
							
							foreach ($rld in $rls.psobject.Properties)
							{
								if ($rls."$($rld.Name)" -eq "True")
								{
									$RetentionLabels += $rld.Name
									#Write-Host $rld.Name
								}
							}
							$RetentionLabels = @($RetentionLabels | ForEach-Object {[PSCustomObject]@{'Name' = $_}})
						}else
						{
							$RetentionLabels = @()
							$RetentionLabels = Get-ComplianceTag | select Name
						}
						$TotalRT = $RetentionLabels.count
						$ProgressRT = 1
						
						Write-Host "`nTotal Retention Labels found:" -NoNewLine
						Write-Host "`t"$TotalRT -ForeGroundColor Green
						
						foreach($rl in $RetentionLabels)
						{
							ExportContentExplorerDetailsTo -TagType $tag -Workload $service -Tag $rl.name -GranularValue $value -PageSize $PageSize -DataConnector $Connector
						}
					}
					if($tag -eq 'Sensitivity')
					{
						if (Test-Path "$PSScriptRoot\ConfigFiles\MPARR-SensitivityLabelsList.json")
						{
							$SensitivitySelected = "$PSScriptRoot\ConfigFiles\MPARR-SensitivityLabelsList.json"
						
							$jsonSL = Get-Content -Raw -Path $SensitivitySelected
							[PSCustomObject]$sls = ConvertFrom-Json -InputObject $jsonSL
							$ListSensitivityLabels = @()
							
							foreach ($sld in $sls.psobject.Properties)
							{
								if ($sls."$($sld.Name)" -eq "True")
								{
									$ListSensitivityLabels += $sld.Name
								}
							}
						}else
						{
							$SensitivityLabels = Get-Label | select DisplayName,ParentLabelDisplayName
							$ListSensitivityLabels = @()
							
							foreach($label in $SensitivityLabels)
							{
								if($label.ParentLabelDisplayName -ne $Null)
								{
									$ListSensitivityLabels += $label.ParentLabelDisplayName+"/"+$label.DisplayName		
								}else
								{
									$ListSensitivityLabels += $label.DisplayName
								}
							}
						}
						$tempFolder = $ListSensitivityLabels
						$SensitivityLabelsSelection = @()
						
						foreach ($label in $tempFolder){$SensitivityLabelsSelection += @([pscustomobject]@{Name=$label})}
						
						$TotalSL = $SensitivityLabelsSelection.count
						$ProgressSL = 1
						
						Write-Host "`nTotal Sensitivity Labels found:" -NoNewLine
						Write-Host "`t"$TotalSL -ForeGroundColor Green
						
						foreach($sl in $SensitivityLabelsSelection)
						{
							ExportContentExplorerDetailsTo -TagType $tag -Workload $service -Tag $sl.name -GranularValue $value -PageSize $PageSize -DataConnector $Connector
						}
						
					}
					if($tag -eq 'SensitiveInformationType')
					{
						if (Test-Path "$PSScriptRoot\ConfigFiles\MPARR-SensitiveInfoTypesList.json")
						{
							$SITsSelected = "$PSScriptRoot\ConfigFiles\MPARR-SensitiveInfoTypesList.json"
						
							$jsonSIT = Get-Content -Raw -Path $SITsSelected
							[PSCustomObject]$sitss = ConvertFrom-Json -InputObject $jsonSIT
							$SITs = @()
							
							foreach ($sitd in $sitss.psobject.Properties)
							{
								if ($sitss."$($sitd.Name)" -eq "True")
								{
									$SITs += $sitd.Name
								}
							}
							$SITs = @($SITs | ForEach-Object {[PSCustomObject]@{'Name' = $_}})
						}else
						{
							$SITs = Get-DlpSensitiveInformationType | select Name
						}
						$TotalSIT = $SITs.count
						$ProgressSIT = 1
						
						Write-Host "`nTotal Sensitive Information Types found:" -NoNewLine
						Write-Host "`t"$TotalSIT -ForeGroundColor Green
						
						foreach($sit in $SITs)
						{
							ExportContentExplorerDetailsTo -TagType $tag -Workload $service -Tag $sit.name -GranularValue $value -PageSize $PageSize -DataConnector $Connector
						}
					}
					if($tag -eq 'TrainableClassifier')
					{
						$TCSelected = "$PSScriptRoot\ConfigFiles\MPARR-TrainableClassifiersList.json"
						
						$json = Get-Content -Raw -Path $TCSelected
						[PSCustomObject]$tcs = ConvertFrom-Json -InputObject $json
						$TClist = @()
						
						foreach ($tcd in $tcs.psobject.Properties)
						{
							if ($tcs."$($tcd.Name)" -eq "True")
							{
								$TClist += $tcd.Name
							}
						}
						
						$tempFolder = $TClist
						$TCSelection = @()
						
						foreach ($tcd in $tempFolder){$TCSelection += @([pscustomobject]@{Name=$tcd})}
						
						$TotalTC = $TCSelection.count
						$ProgressTC = 1
						
						Write-Host "`nTotal Trainable Classifiers found:" -NoNewLine
						Write-Host "`t"$TotalTC -ForeGroundColor Green
						
						foreach($tc in $TCSelection)
						{
							ExportContentExplorerDetailsTo -TagType $tag -Workload $service -Tag $tc.name -GranularValue $value -PageSize $PageSize -DataConnector $Connector
						}
					}
				}
				$Counter++
			}
		}
		
	}
}

function SelectContinuity
{
	$choices  = '&Yes','&No'
	$decision = $Host.UI.PromptForChoice("", "`nDo you want to export more data? ", $choices, 1)
	
	if ($decision -eq 0)
    {
		MainFunction
	}
	if ($decision -eq 1)
	{
		exit
	}
	
}

function MainFunction() 
{
    # ---------------------------------------------------------------   
    #    Name           : Export-ContentExplorerData
    #    Desc           : Extracts data from Content ExplorerData into Logs Analytics
    #    Return         : None
    # ---------------------------------------------------------------
		<#
		.NOTES
		If you cannot add the "Compliance Administrator" role to the Microsoft Entra App, for security reasons, you can comment the line 167 and uncomment the line 166, in that case
		Someone with "Compliance Administrator" role needs to execute this script, this script is executed on-demand to refresh the SITs names
		#>
		
		UpdateMPARREntraApp
		CheckRequiredModules
		ValidateAdditionalConfigurationFiles
		
		#Connectio to Service
		if($SimpleExportToFile)
		{
			$DefaultExport = "File"
			connect2service -ReadExport $DefaultExport
			CheckContentExplorerPermissions
			$TagType = ReadTagType -ReadExport $DefaultExport
			$Workload = ReadWorkload -ReadExport $DefaultExport
		}else
		{
			$DefaultExport = "Logs Analytics"
			connect2service -ReadExport $DefaultExport
			CheckContentExplorerPermissions
			$TagType = ReadTagType -ReadExport $DefaultExport
			$Workload = ReadWorkload -ReadExport $DefaultExport
		}
		
		Start-Sleep -s 5
		#Clean screen after connection
		cls	
		
		#Welcome screen
		Write-Host "`n#################################################################################" -ForeGroundColor Green
		Write-Host "`n"
		Write-Host "This script is thought to export Content Explorer data to MPARR."
		Write-Host "Remember check that you have the right permissions."
		Write-Host "`n#################################################################################" -ForeGroundColor Green
		Write-Host "`n"
		
		#PageSize to be used
		if($ChangePageSize)
		{
			$Size = ExportPageSize -PageSize $InitialPageSize
		}else
		{
			$Size = $InitialPageSize
		}
		
		#Execute the query
		$ExportOptionTo = "Logs Analytics"
		$OptionEventHub = CheckExportOption
		if($OptionEventHub -eq "True")
		{
			EventHubConnection
			$ExportOptionTo = "Event Hub"
		}
		if($MassExportToCsv)
		{
			$ExportOptionTo = "MassCSVExport"
		}
		if($MassExportToJson)
		{
			$ExportOptionTo = "MassJsonExport"
		}
		CollectData -TagType $TagType -Workload $Workload -PageSize $Size -ReadExport $DefaultExport -ExportDataTo $ExportOptionTo
		
		#Check if you want to finish or request a new export
		if($SimpleExportToFile)
		{
			SelectContinuity
		}else
		{
			exit
		}
}  
 
#Main Code - Run as required. Do ensure older table is deleted before creating the new table - as it will create duplicates.
CheckPrerequisites

if($CreateConfigFiles)
{
	ExportToJsonFiles
	exit
}
if($CreateTask)
{
	CreateMPARRContentExplorerTask
	exit
}
if($CheckDependencies)
{
	CheckIfElevated
	CheckRequiredModules
	connect2service
	ValidateAdditionalConfigurationFiles
	CheckContentExplorerPermissions
	UpdateMPARREntraApp
	exit
}

MainFunction
