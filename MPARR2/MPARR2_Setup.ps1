<#PSScriptInfo

.VERSION 2.0.6

.GUID 883af802-165c-4702-b4c1-352686c02f01

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
MPARR installer. 

#>

<#
.SYNOPSIS
    Script to setup MPARR data collector.

.DESCRIPTION
    Script is designed to simplify MPARR setup.
    
    To automate setup, simply run the script and choose one of the following options:

    [1] - Full setup (select Subscription, Log Analytics workspace, create Azure app registration, specify required parameters)
    [2] - Encrypt secrets
	[3] - Create scheduled task for Core Scripts (MPARR Collector and RMS)
	[4] - Create scheduled task for users information
    [5] - Create scheduled task for domains information
    [6] - Create scheduled task for administrator roles information
    [7] - Create scheduled task for Purview Sensitivity Labels and SITs information
	[8] - Sign MPARR scripts
    [0] - Exit 
    
.NOTES
    Version 2.0.6
    Current version - 14.03.2024
#> 

<#
HISTORY
  2023-09-06    G.Berdzik 	- Initial version (partial functionality implemented)
  2023-09-12	G.Berdzik 	- Fixes
  2023-09-14    G.Berdzik 	- All planned functionalities implemented
  2023-09-19    G.Berdzik 	- Fixes
  2023-09-21    G.Berdzik 	- Fixes
  2023-09-22    G.Berdzik 	- Fixes
  2023-09-26    G.Berdzik 	- Fixes
  2023-09-27	S.Zamorano	- QA and some comments
  2023-09-28    G.Berdzik   - Fixes
  2023-10-02	S.Zamorano	- Fix some descriptions
  2023-10-03	S.Zamorano	- Added new tasks on task scheduler creation for supporting scripts (Users, Domains, Roles, Labels, SITs)
  2023-10-03	S.Zamorano	- Added digital signature for MPARR scripts
  2023-10-05	S.Zamorano	- Added comment in the configuration menu
  2023-10-20	S.Zamorano	- Folder selection added for Task Scheduler, permit to create or use existing.
  
  2024-03-01	S.Zamorano	- Public release supporting all the new scripts for MPARR 2 
  2024-03-14	S.Zamorano	- Minor fixes related to sign scripts with extension psd1 and psm1 located on the ConfigFiles folder
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
        @{Name="AIPService"; MinVersion="0.0"},
        @{Name="Az.Accounts"; MinVersion="2.9.0"}, 
        @{Name="Az.OperationalInsights"; MinVersion="0.0"},
        @{Name="Az.Resources"; MinVersion="0.0"},
        @{Name="Microsoft.Graph.Applications"; MinVersion="0.0"},
        @{Name="Microsoft.Graph.Users"; MinVersion="0.0"}, 
        @{Name="Microsoft.Graph.Identity.DirectoryManagement"; MinVersion="0.0"}, 
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
                Write-Host "New version required" -ForegroundColor Red
                $modulesToInstall += $module.Name
            }
            else 
            {
                Write-Host "Installed" -ForegroundColor Green
            }
        }
        else
        {
            Write-Host "Not installed" -ForegroundColor Red
            $modulesToInstall += $module.Name
        }
    }

    if ($modulesToInstall.Count -gt 0)
    {
        $choices  = '&Yes', '&No'

        $decision = $Host.UI.PromptForChoice("", "Misisng required modules. Proceed with installation?", $choices, 0)
        if ($decision -eq 0) 
        {
            Write-Host "Installing modules..."
            foreach ($module in $modulesToInstall)
            {
                Write-Host "`t$module"
                if ($module -ne "AIPService")
                {
                    Install-Module $module -ErrorAction Stop
                }
                else
                {
                    Start-Process "C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe" -Wait -UseNewEnvironment `
                    -ArgumentList '-Command "&{Write-Host "Installing module AIPService..."; [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Import-Module PowerShellGet; Install-Module AIPService -RequiredVersion 2.0.0.3 -Force; Write-Host "Exiting Windows PowerShell session..."; Start-Sleep -Seconds 2}"'

                }
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

function CheckPowerShellVersion
{
    # Check PowerShell version
    Write-Host "`nChecking PowerShell version... " -NoNewline
    if ($Host.Version.Major -gt 5)
    {
        Write-Host "Passed" -ForegroundColor Green
        Write-Host "`tCurrent version is $($Host.Version). Please note that MPARR-RMSData2.ps1 script is executed on PowerShell 7"
		Write-Host "but run services in the background in PowerShell 5.1."
    }
    else
    {
        Write-Host "Failed" -ForegroundColor Red
        Write-Host "`tCurrent version is $($Host.Version). PowerShell version 6 or newer is required."
        exit(1)
    }
}

function CheckPrerequisites
{
    CheckIfElevated
    CheckPowerShellVersion
    CheckRequiredModules
}

# function to get option number
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

# Connect to Log Analtytics
function SConnectToLA 
{
    $CONFIGFILE = $PSScriptRoot+"\ConfigFiles\laconfig.json"
	$config = InitializeLAConfigFile -DirRoot $CONFIGFILE

	#Write-Host "`n*** Executing 'Connect to Log Analytics'.`n"

    Write-Host "`nGetting subscriptions..."
    $i = 1
    $subscriptions = @(Get-AzSubscription -TenantId (Get-AzContext).Tenant -ErrorAction Stop | Select-Object Name, Id | 
        ForEach-Object {$_ | Add-Member -Name "No" -MemberType NoteProperty -Value ($i++) -PassThru})
    
    if ($subscriptions.Count -eq 0)
    {
        Write-Host "`nNo subscriptions found. Reconnect with Connect-AzAccount cmdlet." -ForegroundColor Red
        Write-Host "Exiting..."
        exit(2)
    }
    elseif ($subscriptions.Count -gt 1)
    {
        $subscriptions | Select-Object No, Name, Id | Out-Host

        $selection = 0
        ReadNumber -max ($i -1) -msg "Enter number corresponding to the subscription" -option ([ref]$selection)
        Set-AzContext -SubscriptionId $subscriptions[$selection - 1].Id -ErrorAction Stop 
    }
    else 
    {
        Write-Host "`nOnly one subscription available. '$($subscriptions[0].Name)' selected.`n"    
    }

    $i = 1
    try 
    {
        Write-Host "`nGetting workspaces..."
        $workspaces = @(Get-AzOperationalInsightsWorkspace -ErrorAction Stop |
            ForEach-Object {
                $_ | Add-Member -Name "No" -MemberType NoteProperty -Value ($i++) -PassThru
            }
        )    
    }
    catch 
    {
        Write-Host "$($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Exiting..."
        exit(2)
    }
    $workspaces | Format-Table No, Name, ResourceGroupName, Location, Sku, Tags | Out-Host

    Write-Host "In case workspace recently created is not listed, please stop the script with Ctrl+C and run it again.`n"
    $selection = 0
    ReadNumber -max ($i -1) -msg "Enter number corresponding to the Log Analytics workspace" -option ([ref]$selection)
    $primaryKey = (Get-AzOperationalInsightsWorkspaceSharedKey -ResourceGroupName $workspaces[$selection - 1].ResourceGroupName `
        -Name $workspaces[$selection -1].Name -ErrorAction Stop).PrimarySharedKey
    
    $config.LA_CustomerID = ($workspaces[$selection - 1].CustomerId).ToString()
    $config.LA_SharedKey = ($primaryKey).ToString()

	WriteToConfigFile -DirRoot $CONFIGFILE
    Write-Host "`n"
}

# function to create Azure App
function NewApp
{
    Connect-MgGraph -Scopes "Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All", "Directory.ReadWrite.All", "User.ReadWrite.All" -NoWelcome

	$CONFIGFILE = $PSScriptRoot+"\ConfigFiles\laconfig.json"
	$config = InitializeLAConfigFile -DirRoot $CONFIGFILE
	
    $appName = "MPARR2-DataCollector"
    Get-MgApplication -ConsistencyLevel eventual -Count appCount -Filter "startsWith(DisplayName, 'MPARR2-DataCollector')" | Out-Null
    if ($appCount -gt 0)
    {   
        $sufix = ((New-Guid) -split "-")[0]
        $appName = "MPARR2-DataCollector-$sufix"
        Write-Host "'MPARR2-DataCollector' app already exists. New name was generated: '$appName'`n"
    }

    # ask for the app name
    $choices  = '&Proceed', '&Change'

    $decision = $Host.UI.PromptForChoice("", "'$appName' application will be registered. Do you want to proceed or change the name?", $choices, 0)
    if ($decision -eq 1)
    {
        $ok = $false
        do 
        {
            $newName = Read-Host "Please enter the new name"
            if ($newName -ne "")
            {
                Get-MgApplication -ConsistencyLevel eventual -Count appCount -Filter "DisplayName eq '$newName'" | Out-Null
                if ($appCount -eq 0)
                {
                    $appName = $newName
                    $ok = $true
                }
                else 
                {
                    Write-Host "Selected name already exists."
                }
            }
        }
        until ($ok)
    }

    # app parameters and API permissions definition
    $params = @{
        DisplayName = $appName
        SignInAudience = "AzureADMyOrg"
        RequiredResourceAccess = @(
            @{
            ResourceAppId = "00000003-0000-0000-c000-000000000000"
            ResourceAccess = @(
                @{
                    Id = "e1fe6dd8-ba31-4d61-89e7-88639da4683d"
                    Type = "Scope"
                },
                @{
                    Id = "b0afded3-3588-46d8-8b3d-9842eff778da"
                    Type = "Role"
                },
                @{
                    Id = "7ab1d382-f21e-4acd-a863-ba3e13f7da61"
                    Type = "Role"
                },
                @{
                    Id = "5b567255-7703-4780-807c-7be8301ae99b"
                    Type = "Role"
                },
                @{
                    Id = "498476ce-e0fe-48b0-b801-37ba7e2685c6"
                    Type = "Role"
                },
                @{
                    Id = "df021288-bdef-4463-88db-98f22de89214"
                    Type = "Role"
                },
                @{
                    Id = "230c1aed-a721-4c5d-9cb4-a90514e508ef"
                    Type = "Role"
                }
            )
        },
        @{
            ResourceAppId = "00000012-0000-0000-c000-000000000000"
            ResourceAccess = @(
                @{
                    Id = "e23bd57d-bfd5-4906-867f-89fb5ed8cd43"
                    Type = "Role"
                }
            )
        },
        @{
            ResourceAppId = "00000002-0000-0ff1-ce00-000000000000"
            ResourceAccess = @(
                @{
                    Id = "dc50a0fb-09a3-484d-be87-e023b12c6440"
                    Type = "Role"
                }
            )
        },
        @{
            ResourceAppId = "c5393580-f805-4401-95e8-94b7a6ef2fc2"
            ResourceAccess = @(
                @{
                    Id = "4807a72c-ad38-4250-94c9-4eabfe26cd55"
                    Type = "Role"
                },
                @{
                    Id = "594c1fb6-4f81-4475-ae41-0c394909246c"
                    Type = "Role"
                },
                @{
                    Id = "e2cea78f-e743-4d8f-a16a-75b629a038ae"
                    Type = "Role"
                }
            )
        }
        )
    }
    # create application
    $app = New-MgApplication @params
    $appId = $app.Id

    # assign owner
    $userId = (Get-MgUser -UserId (Get-MgContext).Account).Id
    $params = @{
        "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$userId"
    }
    New-MgApplicationOwnerByRef -ApplicationId $appId -BodyParameter $params

    # ask for certificate name
    $certName = "MPARR-DataCollector"
    $choices  = '&Proceed', '&Change'
    $decision = $Host.UI.PromptForChoice("", "Default certificate name is '$certName'. Do you want to proceed or change the name?", $choices, 0)
    if ($decision -eq 1)
    {
        #$ok = $false
        do 
        {
            $newName = Read-Host "Please enter the new name"
        }
        until ($newName -ne "")
        $certName = $newName
    }

    # certificate life
    $validMonths = 12
    $choices  = '&Yes', '&No'
    $decision = $Host.UI.PromptForChoice("", "Certificate is valid for 12 months. Do you want to change this value?", $choices, 1)
    if ($decision -eq 0)
    {
        ReadNumber -max 36 -msg "Enter number of months (max. 36)" -option ([ref]$validMonths)
    }

    # create key
    $cert = New-SelfSignedCertificate -DnsName $certName -CertStoreLocation "cert:\CurrentUser\My" -NotAfter (Get-Date).AddMonths($validMonths)
    $certBase64 = [System.Convert]::ToBase64String($cert.RawData)
    $keyCredential = @{
        type = "AsymmetricX509Cert"
        usage = "Verify"
        key = [System.Text.Encoding]::ASCII.GetBytes($certBase64)
    }
    while (-not (Get-MgApplication -ApplicationId $appId -ErrorAction SilentlyContinue)) 
    {
        Write-Host "Waiting while app is being created..."
        Start-Sleep -Seconds 5
    }
    Update-MgApplication -ApplicationId $appId -KeyCredentials $keyCredential -ErrorAction Stop
    $choices  = '&Yes', '&No'
    $decision = $Host.UI.PromptForChoice("", "Do you want to backup certificate to file?", $choices, 1)
    if ($decision -eq 0)
    {
        if ((Get-Module -Name PKI).Version.Build -eq -1)
        {
            Write-Host "`nThis system uses old version of PKI module that is not able to proceed with certificate export." -ForegroundColor Yellow
            Write-Host "Please export certificate manually using 'certmgr.msc' console.`n" -ForegroundColor Yellow
        }
        else
        {
            $pass = Read-Host -Prompt "Please enter password to secure certificate" -AsSecureString
            Export-PfxCertificate -Cert $cert -FilePath ".\Certs\$certname.pfx" -Password $pass | Out-Null
            Remove-Variable pass
        }
    }
    
    # ask for client secret name
    $keyName = "MPARR Collector App Secret key"
    $choices  = '&Proceed', '&Change'
    $decision = $Host.UI.PromptForChoice("", "Default client description for secret key is '$keyName'. Do you want to proceed or change the name?", $choices, 0)
    if ($decision -eq 1)
    {
        #$ok = $false
        do 
        {
            $newName = Read-Host "Please enter the new name"
        }
        until ($newName -ne "")
        $keyName = $newName
    }

    # create client secret
    $passwordCred = @{
        displayName = $keyName
        endDateTime = (Get-Date).AddMonths(12)
     }
     
    $secret = Add-MgApplicationPassword -applicationId $appId -PasswordCredential $passwordCred

    Write-Host "`nAzure application was created."
    Write-Host "App Name: $appName"
    Write-Host "App ID: $($app.AppId)"
    Write-Host "Secret password: $($secret.SecretText)"
    Write-Host "Certificate thumbprint: $($cert.Thumbprint)"

    Write-Host "`nPlease go to the Azure portal to manually grant admin consent:"
    Write-Host "https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$($app.AppId)`n" -ForegroundColor Cyan

    $config.AppClientID = $app.AppId
    $config.CertificateThumb = $cert.Thumbprint
    $config.ClientSecretValue = $secret.SecretText
	
	WriteToConfigFile -DirRoot $CONFIGFILE

    Remove-Variable cert
    Remove-Variable certBase64
    Remove-Variable secret
}

function GetTenantInfo
{
    $CONFIGFILE = $PSScriptRoot+"\ConfigFiles\laconfig.json"
	$config = InitializeLAConfigFile -DirRoot $CONFIGFILE
	$tenant = Get-MgDomain
    $config.TenantGUID = (Get-MgContext).TenantId
    $config.TenantDomain = ($tenant | Where-Object IsDefault).Id
    $config.OnmicrosoftURL = ($tenant | Where-Object IsInitial).Id
	WriteToConfigFile -DirRoot $CONFIGFILE
}

function SelectCloud
{
    $CONFIGFILE = $PSScriptRoot+"\ConfigFiles\laconfig.json"
	$config = InitializeLAConfigFile -DirRoot $CONFIGFILE
	$choices = '&Commercial', '&GCC', 'GCC&H', '&DOD'
    $decision = $Host.UI.PromptForChoice("", "`nPlease select cloud version:", $choices, 0)
    switch ($decision) {
        0 {$config.Cloud = "Commercial"; break}
        1 {$config.Cloud = "GCC"; break}
        2 {$config.Cloud = "GCCH"; break}
        3 {$config.Cloud = "DOD"; break}
    }
	WriteToConfigFile -DirRoot $CONFIGFILE
}

# function to choose destination directory for logs
function SelectLogPath
{
    $CONFIGFILE = $PSScriptRoot+"\ConfigFiles\laconfig.json"
	$LogsDirectory = $PSScriptRoot+"\Logs\"
	$RMSLogsDirectory = $PSScriptRoot+"\RMSLogs\"
	$config = InitializeLAConfigFile -DirRoot $CONFIGFILE
	$choices  = '&Yes', '&No'
    $decision = $Host.UI.PromptForChoice("", "Default locations for logs are '$($RMSLogsDirectory)' and '$($LogsDirectory)'. Do you want change the location?", $choices, 1)
    if ($decision -eq 0)
    {
        [System.Reflection.Assembly]::Load("System.Windows.Forms") | Out-Null
        $folder = New-Object System.Windows.Forms.FolderBrowserDialog
        $folder.Description = "Select folder to store logs"
        $folder.rootFolder = 'MyComputer'
        $folder.UseDescriptionForTitle = $true
        # main log directory
        if ($folder.ShowDialog() -eq "OK")
        {
            $config.OutPutLogs = $folder.SelectedPath + "\"
            Write-Host "`nOutput logs set to '$($config.OutPutLogs)'."
        }

        # RMS logs dir
        $folder.Description = "Select folder to store RMS logs"
        $folder.rootFolder = 'MyComputer'
        $folder.InitialDirectory = $config.OutPutLogs
        if ($folder.ShowDialog() -eq "OK")
        {
            $config.RMSLogs = $folder.SelectedPath + "\"
            Write-Host "`nRMS logs set to '$($config.RMSLogs)'."
        }
    }elseif ($decision -eq 1)
	{
		Write-Host "Same default folders was selected."
		if(-Not (Test-Path $LogsDirectory ))
		{
			Write-Host "Export data directory is missing, creating a new folder called Logs"
			New-Item -ItemType Directory -Force -Path "$PSScriptRoot\Logs" | Out-Null
		}
		if(-Not (Test-Path $RMSLogsDirectory ))
		{
			Write-Host "Export data directory is missing, creating a new folder called RMSLogs"
			New-Item -ItemType Directory -Force -Path "$PSScriptRoot\RMSLogs" | Out-Null
		}
		$config.OutPutLogs = $LogsDirectory
		$config.RMSLogs = $RMSLogsDirectory
	}
	
	WriteToConfigFile -DirRoot $CONFIGFILE
}

function InitializeLAConfigFile($DirRoot)
{
	# read config file
    $configfile = "$DirRoot"
	
	if(-Not (Test-Path $configfile ))
	{
		Write-Host "Export data directory is missing, creating a new folder called ConfigFiles"
		New-Item -ItemType Directory -Force -Path "$PSScriptRoot\ConfigFiles" | Out-Null
	}
	
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

# write configuration data to json file
function WriteToJsonFile
{
    $BackupPath = $PSScriptRoot+"\BackupScripts"
	if(-Not (Test-Path $BackupPath ))
	{
		Write-Host "Export data directory is missing, creating a new folder called BackupScripts"
		New-Item -ItemType Directory -Force -Path "$PSScriptRoot\BackupScripts" | Out-Null
	}
	
	$MPARRConfigFolder = "$PSScriptRoot\ConfigFiles\"
	$MPARRConfigFile = "$MPARRConfigFolder"+"laconfig.json"
	if (Test-Path -Path $MPARRConfigFile)
    {
        $date = Get-Date -Format "yyyyMMddHHmmss"
		$BackupFile = $PSScriptRoot+"\BackupScripts\"+"laconfig_"+$date+".backup.json"
        Move-Item $MPARRConfigFile $BackupFile
        Write-Host "`nThe old config file moved to 'laconfig_$date.backup.json'"
    }
    $config | ConvertTo-Json | Out-File $MPARRConfigFile
    Write-Host "Setup completed. New config file was created." -ForegroundColor Green
}

function UpdateMPARREventHub
{
	Clear-Host
	cls
	
	$CONFIGFILE = "$PSScriptRoot\ConfigFiles\laconfig.json"
	$json = Get-Content -Raw -Path $CONFIGFILE
	[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
	
	Write-Host "`n`n----------------------------------------------------------------------------------------"
	Write-Host "`nMPARR Event Hub configuration menu!" -ForegroundColor DarkGreen
	Write-Host "If you enable Event Hub connector Logs Analytics connector will not be work by default." -ForegroundColor DarkYellow
	Write-Host "The current menu permit configure the connector to be used manually," -ForegroundColor DarkGreen
	Write-Host "to do that you can execute all the scripts using the attribute -ExportToEventHub" -ForegroundColor DarkGreen
	Write-Host "The current configuration is:"
	Write-Host "`tExport to Event Hub enabled `t:`t"$config.ExportToEventHub
	Write-Host "`tEvent Hub Namespace is set to `t:`t"$config.EventHubNamespace
	Write-Host "`tEvent Hub Instance is set to `t:`t"$config.EventHub
	Write-Host "`n----------------------------------------------------------------------------------------`n`n"
	
	$choices  = '&Yes', '&No'
	$decision = $Host.UI.PromptForChoice("", "Do you want to change the current configuration?", $choices, 1)
	
	if ($decision -eq 0)
	{
		Write-Host "`nYou decide to change the current configuration."
		Write-Host "You can change the configuration at any time using this option."
		Write-Host "Remember that setting Event Hub to run automatically disable the option to send the data to Logs Analytics."
		Write-Host "All the scripts can be executed manually to send the data to Event Hub using the attribute -ExportToEventHub"
		Write-Host "As a sample: .\MPARR_Collector2.ps1 -ExportToEventHub"
		$ExecutionChoices  = '&Automatically', '&Manual'
		$ExecutionDecision = $Host.UI.PromptForChoice("", "`nDo you want to change the current configuration?", $ExecutionChoices, 1)
		if ($ExecutionDecision -eq 0)
		{
			$config.ExportToEventHub = "True"
		}elseif ($ExecutionDecision -eq 1)
		{
			$config.ExportToEventHub = "False"
		}
		
		Write-Host "`n"
		do 
        {
            $newEventHubName = Read-Host "Please enter the Event Hub Namespace"
        }
        until ($newEventHubName -ne "")
        $config.EventHubNamespace = $newEventHubName 
		Write-Host "The Event Hub Namespace established is :" -NoNewLine
		Write-Host "`t$newEventHubName." -ForegroundColor Green
		
		Write-Host "`n"
		do 
        {
            $newEventHubInstance = Read-Host "Please enter the Event Hub Instance"
        }
        until ($newEventHubInstance -ne "")
        $config.EventHub = $newEventHubInstance 
		Write-Host "The Event Hub Instance established is  :" -NoNewLine
		Write-Host "`t$newEventHubInstance." -ForegroundColor Green
		
		$config | ConvertTo-Json | Out-File $CONFIGFILE
		Start-Sleep -s 1
		
	}else
	{
		return
	}
	
	Write-Host "New configuration for Event Hub added to laconfig.json file"
	Write-Host "Press any key to continue..."
	$key = ([System.Console]::ReadKey($true)) | Out-Null
}

function WriteToConfigFile($DirRoot)
{
    $config | ConvertTo-Json | Out-File "$DirRoot"
    Write-Host "Setup completed. New config file was created." -ForegroundColor Yellow
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
        #$ok = $false
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

function CreateMPARRCollectorTask
{
    # main data collector script
    $taskName = "MPARR-DataCollector"
	
	# Call function to set a folder for the task on Task Scheduler
	$taskFolder = CreateScheduledTaskFolder
	
	<#
	.NOTES
	This function create both task,MPARR_Collector and MPARR-RMSData, to run every 30 minutes, that time can be changed on the same task scheduler, is not recommended less time.
	MPARR_Collector use PowerShell 7 
	#>
	Write-Host "`n`n----------------------------------------------------------------------------------------" -ForegroundColor Yellow
	Write-Host "`n Please be aware that the scripts MPARR_Collector is set to execute every 30 minutes" -ForegroundColor DarkYellow
	Write-Host "` You can change directly on task scheduler and change the execution period" -ForegroundColor DarkYellow
	Write-Host "` Depend on your logs volume cannot be recommend use less time," -ForegroundColor DarkYellow
	Write-Host "` to give time to the scripts to be execute correctly." -ForegroundColor DarkYellow
	Write-Host "`n----------------------------------------------------------------------------------------" -ForegroundColor Yellow
	Write-Host "`n`n"

    # calculate date
    $dt = Get-Date
    $nearestMinutes = 30 
    $reminder = $dt.Minute % $nearestMinutes
    $dt = $dt.AddMinutes(-$reminder)
    $startTime = [datetime]::new($dt.Year, $dt.Month, $dt.Day, $dt.Hour, $dt.Minute, 0)

    #create task
    $trigger = New-ScheduledTaskTrigger -Once -At $startTime -RepetitionInterval (New-TimeSpan -Minutes $nearestMinutes)
    $action = New-ScheduledTaskAction -Execute "`"$PSHOME\pwsh.exe`"" -Argument ".\MPARR_Collector2.ps1" -WorkingDirectory $PSScriptRoot
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

function CreateMPARRRMSDataTask
{
	# MPARR-ContentExplorerData script
    $taskName = "MPARR-RMSData"
	
	# Call function to set a folder for the task on Task Scheduler
	$taskFolder = CreateScheduledTaskFolder
	
	# Task execution
    $validDays = 1
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
    $StartTime = [datetime]::new($dt.Year, $dt.Month, $dt.Day, $dt.Hour, $dt.Minute, 0)

    #create task
    $trigger = New-ScheduledTaskTrigger -Once -At $StartTime -RepetitionInterval (New-TimeSpan -Days $validDays)
    $action = New-ScheduledTaskAction -Execute "`"$PSHOME\pwsh.exe`"" -Argument ".\MPARR_RMSData2.ps1" -WorkingDirectory $PSScriptRoot
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

function CreateMPARRUsersTask
{
	# MPARR-MicrosoftEntraUsers script
    $taskName = "MPARR-MicrosoftEntraUsers"
	
	# Call function to set a folder for the task on Task Scheduler
	$taskFolder = CreateScheduledTaskFolder
	
	# Task execution
    $validDays = 15
    $choices  = '&Yes', '&No'
    $decision = $Host.UI.PromptForChoice("", "The task on task scheduler will be set for 15 days, do you want to change?", $choices, 1)
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
	Write-Host "`nYou need to execute this script at least once manually." -ForegroundColor DarkRed
	Write-Host "`nPress any key to continue..."
	$key = ([System.Console]::ReadKey($true)) | Out-Null
}

function CreateMPARRDomainsTask
{
	# MPARR-MicrosoftEntraDomains script
    $taskName = "MPARR-MicrosoftEntraDomains"
	
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
    $action = New-ScheduledTaskAction -Execute "`"$PSHOME\pwsh.exe`"" -Argument ".\MPARR-MicrosoftEntraDomains.ps1" -WorkingDirectory $PSScriptRoot
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

function CreateMPARRRolesTask
{
	# MPARR-MicrosoftEntraRoles script
    $taskName = "MPARR-MicrosoftEntraRoles"
	
	# Call function to set a folder for the task on Task Scheduler
	$taskFolder = CreateScheduledTaskFolder
	
	# Task execution
    $validDays = 7
    $choices  = '&Yes', '&No'
    $decision = $Host.UI.PromptForChoice("", "The task on task scheduler will be set for 7 days, do you want to change?", $choices, 1)
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
    $action = New-ScheduledTaskAction -Execute "`"$PSHOME\pwsh.exe`"" -Argument ".\MPARR-MicrosoftEntraRoles.ps1" -WorkingDirectory $PSScriptRoot
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

function CreateMPARRM365StatisticsTask
{
	# MPARR-ContentExplorerData script
    $taskName = "MPARR-Microsoft365CloudStatistics"
	
	# Call function to set a folder for the task on Task Scheduler
	$taskFolder = CreateScheduledTaskFolder
	
	# Task execution
    $validDays = 7
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
    $action = New-ScheduledTaskAction -Execute "`"$PSHOME\pwsh.exe`"" -Argument ".\MPARR-M365CloudStatistics.ps1" -WorkingDirectory $PSScriptRoot
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

function CreateMPARRPurviewLabelsTask
{
	# MPARR-PurviewSensitivityLabels script
    $taskName = "MPARR-MicrosoftPurviewSensitivityLabel"
	
	# Call function to set a folder for the task on Task Scheduler
	$taskFolder = CreateScheduledTaskFolder
	
	# Task execution
    $validDays = 7
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
    $action = New-ScheduledTaskAction -Execute "`"$PSHOME\pwsh.exe`"" -Argument ".\MPARR-PurviewSensitivityLabels.ps1" -WorkingDirectory $PSScriptRoot
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

function CreateMPARRPurviewSITsTask
{
	# MPARR-PurviewSITs script
    $taskName = "MPARR-MicrosoftPurviewSITs"
	
	# Call function to set a folder for the task on Task Scheduler
	$taskFolder = CreateScheduledTaskFolder
	
	# Task execution
    $validDays = 7
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
    $action = New-ScheduledTaskAction -Execute "`"$PSHOME\pwsh.exe`"" -Argument ".\MPARR-PurviewSITs.ps1" -WorkingDirectory $PSScriptRoot
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

function CreateMPARRPurviewRolesTask
{
	# MPARR-PurviewRoles script
    $taskName = "MPARR-PurviewRoles"
	
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
    $action = New-ScheduledTaskAction -Execute "`"$PSHOME\pwsh.exe`"" -Argument ".\MPARR-PurviewRoles.ps1" -WorkingDirectory $PSScriptRoot
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

function CreateMPARRMicrosoftLicensesTask
{
	# MPARR-PurviewRoles script
    $taskName = "MPARR-MSLicenses"
	
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
    $action = New-ScheduledTaskAction -Execute "`"$PSHOME\pwsh.exe`"" -Argument ".\MPARR-MicrosoftLicenses.ps1" -WorkingDirectory $PSScriptRoot
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

function SelfSignScripts
{
	#Menu for self signed or use an own certificate
	<#
	.NOTES
	MPARR scripts can request change your Execution Policy to bypass to be executed, using PS:\> Set-ExecutionPolicy -ExecutionPolicy bypass.
	In some organizations for security concerns this cannot be set, and the script need to be digital signed.
	This function permit to use a self-signed certificate or use an external one. 
	BE AWARE : The external certificate needs to be for a CODE SIGNING is not a coomon SSL certificate.
	#>
	
	Write-Host "`n`n----------------------------------------------------------------------------------------" -ForegroundColor Yellow
	Write-Host "`nThis option will be digital sign all MPARR scripts." -ForegroundColor DarkYellow
	Write-Host "The certificate used is the kind of CodeSigning not a SSL certificate" -ForegroundColor DarkYellow
	Write-Host "If you choose to select your own certificate be aware of this." -ForegroundColor DarkYellow
	Write-Host "`n----------------------------------------------------------------------------------------" -ForegroundColor Yellow
	Write-Host "`n`n" 
	
	# Decide if you want to progress or not
	$choices  = '&Yes', '&No', '&Install new certificate'
    $decision = $Host.UI.PromptForChoice("", "Do you want to proceed with the digital signature for all the scripts?", $choices, 1)
	if ($decision -eq 1)
	{
		Write-Host "`nYou decide don't proceed with the digital signature." -ForegroundColor DarkYellow
		Write-Host "Remember to use MPARR scripts set permissions with Administrator rigths on Powershel using:." -ForegroundColor DarkYellow
		Write-Host "`nSet-ExecutionPolicy -ExecutionPolicy bypass." -ForegroundColor Green
	}elseif($decision -eq 2)
	{
		Write-Host "`nNew certificate will be created and installed." -ForegroundColor Blue
		Write-Host "Proceeding to create one..."
		Write-Host "This can take a minute and a pop-up will appear, please accept to install the certificate."
		Write-Host "After finish you'll be forwarded to the initial Certificate menu."
		CreateCodeSigningCertificate
		SelfSignScripts
	}else
	{
		
		#Review if some certificate was installed previously
		Write-Host "`nGetting Code Signing certificates..." -ForegroundColor Green
		$i = 1
		$certificates = @(Get-ChildItem Cert:\CurrentUser\My | Where-Object {$_.EnhancedKeyUsageList -like "*Code Signing*"}| Select-Object Subject, Thumbprint, NotBefore, NotAfter | ForEach-Object {$_ | Add-Member -Name "No" -MemberType NoteProperty -Value ($i++) -PassThru})
		$certificates | Format-Table No, Subject, Thumbprint, NotBefore, NotAfter | Out-Host
		
		if ($certificates.Count -eq 0)
		{
			Write-Host "`nNo certificates for Code Signing was found." -ForegroundColor Red
			Write-Host "Proceeding to create one..."
			Write-Host "This can take a minute and a pop-up will appear, please accept to install the certificate."
			Write-Host "After finish you'll be forwarded to the initial Certificate menu."
			CreateCodeSigningCertificate
			SelfSignScripts
		} else{
			$selection = 0
			ReadNumber -max ($i -1) -msg "Enter number corresponding to the certificate to use" -option ([ref]$selection)
			#Obtain certificate from local store
			$cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert | Where-Object {$_.Thumbprint -eq $certificates[$selection - 1].Thumbprint}
			
			#Sign MPARR Scripts
			$files = Get-ChildItem -Path .\MPARR*.ps1
			$SupportFiles = Get-ChildItem -Path .\ConfigFiles\MPARR*.ps*
			
			foreach($file in $files)
			{
				Write-Host "`Signing..."
				Write-Host "$($file.Name)" -ForegroundColor Green
				Set-AuthenticodeSignature -FilePath ".\$($file.Name)" -Certificate $cert
			}
			
			foreach($SupportFile in $SupportFiles)
			{
				Write-Host "`Signing..."
				Write-Host "$($SupportFile.Name)" -ForegroundColor Green
				Set-AuthenticodeSignature -FilePath ".\ConfigFiles\$($SupportFile.Name)" -Certificate $cert
			}
			
			Write-Host "`nPress any key to continue..."
			$key = ([System.Console]::ReadKey($true))
		}
	}
}

function CreateCodeSigningCertificate
{
	#CMDLET to create certificate
	$MPARRcert = New-SelfSignedCertificate -Subject "CN=MPARR PowerShell Code Signing Cert" -Type "CodeSigning" -CertStoreLocation "Cert:\CurrentUser\My" -HashAlgorithm "sha256"
		
	### Add Self Signed certificate as a trusted publisher (details here https://adamtheautomator.com/how-to-sign-powershell-script/)
		
		# Add the self-signed Authenticode certificate to the computer's root certificate store.
		## Create an object to represent the CurrentUser\Root certificate store.
		$rootStore = [System.Security.Cryptography.X509Certificates.X509Store]::new("Root","CurrentUser")
		## Open the root certificate store for reading and writing.
		$rootStore.Open("ReadWrite")
		## Add the certificate stored in the $authenticode variable.
		$rootStore.Add($MPARRcert)
		## Close the root certificate store.
		$rootStore.Close()
			 
		# Add the self-signed Authenticode certificate to the computer's trusted publishers certificate store.
		## Create an object to represent the CurrentUser\TrustedPublisher certificate store.
		$publisherStore = [System.Security.Cryptography.X509Certificates.X509Store]::new("TrustedPublisher","CurrentUser")
		## Open the TrustedPublisher certificate store for reading and writing.
		$publisherStore.Open("ReadWrite")
		## Add the certificate stored in the $authenticode variable.
		$publisherStore.Add($MPARRcert)
		## Close the TrustedPublisher certificate store.
		$publisherStore.Close()	
}

function EncryptSecrets
{
    # read config file
    $CONFIGFILE = "$PSScriptRoot\ConfigFiles\laconfig.json"  
    if (-not (Test-Path -Path $CONFIGFILE))
    {
        Write-Host "`nMissing config file '$CONFIGFILE'." -ForegroundColor Yellow
        return
    }
    $json = Get-Content -Raw -Path $CONFIGFILE
    [PSCustomObject]$config = ConvertFrom-Json -InputObject $json
    $EncryptedKeys = $config.EncryptedKeys

    # check if already encrypted
    if ($EncryptedKeys -eq "True")
    {
        Write-Host "`nAccording to the configuration settings (EncryptedKeys: True), secrets are already encrypted." -ForegroundColor Yellow
        Write-Host "No actions taken."
        return
    }

    # encrypt secrets
    $ClientSecretValue = $config.ClientSecretValue
    $SharedKey = $config.LA_SharedKey
    $CertificateThumb = $config.CertificateThumb

    $ClientSecretValue = $ClientSecretValue | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString
    $SharedKey = $SharedKey | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString
    $CertificateThumb = $CertificateThumb | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString

    # write results to the file
    $config.EncryptedKeys = "True"
    $config.ClientSecretValue = $ClientSecretValue
    $config.LA_SharedKey = $SharedKey
    $config.CertificateThumb = $CertificateThumb

    $date = Get-Date -Format "yyyyMMddHHmmss"
    Move-Item "laconfig.json" "$PSScriptRoot\ConfigFiles\laconfig_$date.json"
    Write-Host "`nSecrets encrypted."
    Write-Host "The old config file moved to 'laconfig_$date.json'" -ForegroundColor Green
    $config | ConvertTo-Json | Out-File $CONFIGFILE

    Write-Host "Warning!" -ForegroundColor Yellow
    Write-Host "Please note that encrypted keys can be decrypted only on this machine, using the same account." -ForegroundColor Yellow
	Write-Host "`nPress any key to continue..."
	$key = ([System.Console]::ReadKey($true)) | Out-Null
}

function MicrosoftLicensing($LicenseOption)
{
	$MSProductsTableName = "MSProducts"
	$CommandPath = $PSScriptRoot
	$CommandSupportInfoFolder = $CommandPath+"\Support\"
	if(-Not (Test-Path $CommandSupportInfoFolder ))
	{
		Write-Host "Export data directory is missing, creating a new folder called Support"
		New-Item -ItemType Directory -Force -Path "$PSScriptRoot\Support" | Out-Null
		$M365LicenseURI = "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0"
		$M365FileName = "Product names and service plan identifiers for licensing.csv"
		$SupportFolder = "$PSScriptRoot\Support"
		$result = Invoke-WebRequest -Uri "$M365LicenseURI/$M365FileName" -OutFile $SupportFolder
	}
	$CommandScript = "MPARR-ExportCSV2LA.ps1"
	$CommandData = gci $CommandSupportInfoFolder -Filter *.csv | select -last 1
	if($LicenseOption -eq 1)
	{
		&"$CommandPath\$CommandScript" -Filename $CommandData -TableName $MSProductsTableName
		Start-Sleep -s 3
	}
	if($LicenseOption -eq 2)
	{
		$CommandName = "MPARR-MicrosoftLicenses.ps1"
		$scriptFile ="$PSScriptRoot\$CommandName" 
		if (-not (Test-Path -Path $scriptFile))
		{
			$Command0 = '$MSProductsTableName = "MSProducts"'
			$Command1 = '$CommandPath = $PSScriptRoot'
			$Command2 = '$CommandSupportInfoFolder = $CommandPath+"\Support\"'
			$Command3 = '$CommandScript = "MPARR-ExportCSV2LA.ps1"'
			$Command4 = '$CommandData = gci $CommandSupportInfoFolder -Filter *.csv | select -last 1'
			$Command5 = '&"$CommandPath\$CommandScript" -Filename $CommandData -TableName $MSProductsTableName'
			
			$Command0 | Out-File -FilePath ".\$CommandName" -Append
			$Command1 | Out-File -FilePath ".\$CommandName" -Append
			$Command2 | Out-File -FilePath ".\$CommandName" -Append
			$Command3 | Out-File -FilePath ".\$CommandName" -Append
			$Command4 | Out-File -FilePath ".\$CommandName" -Append
			$Command5 | Out-File -FilePath ".\$CommandName" -Append
		}
		
		Write-Host "A new script was created at $PSScriptRoot called : " -NoNewline
		Write-Host $CommandName -ForegroundColor Green
		Start-Sleep -s 3
		CreateMPARRMicrosoftLicensesTask	
	}
	if($LicenseOption -eq 3)
	{
		$M365LicenseURI = "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0"
		$M365FileName = "Product names and service plan identifiers for licensing.csv"
		$SupportFolder = "$PSScriptRoot\Support"
		$result = Invoke-WebRequest -Uri "$M365LicenseURI/$M365FileName" -OutFile $SupportFolder
		Write-Host "File updated at $SupportFolder"
		Start-Sleep -s 3
	}
	Write-Host "Press any key to continue..."
	$key = ([System.Console]::ReadKey($true))
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
	
	Write-Host "`nYou will prompted to select the right path where the laconfig.json file is located."
	Write-Host "Press any key to continue..."
	$key = ([System.Console]::ReadKey($true))
	
                
	#Here you start selecting each folder
	[System.Reflection.Assembly]::Load("System.Windows.Forms") | Out-Null
	$file = New-Object System.Windows.Forms.OpenFileDialog
	# Start selecting laconfig.json location  
	$file.Title = "Select folder where laconfig.json is located"
	$file.InitialDirectory = 'ProgramFiles'
	$file.Filter = 'MPARR Config file|laconfig.json'
	# main log directory
	if ($file.ShowDialog() -eq "OK")
	{
		$CONFIGFILE = $file.FileName
		$json = Get-Content -Raw -Path $CONFIGFILE
		[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
		$AppID = $config.AppClientID
	}
	
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
        Write-Host "`tPermission already in place" -ForegroundColor Green
    }

	Write-Host "Press any key to continue..." 
    $key = ([System.Console]::ReadKey($true))
}

function MPARRFolderStructure
{
	Clear-Host
	cls
	Write-Host "`n`n----------------------------------------------------------------------------------------"
	Write-Host "`nMPARR configuration to set folder structure to prepare migration!" -ForegroundColor DarkGreen
	Write-Host "This menu helps to validate the folder structure required for migration." -ForegroundColor DarkGreen
	Write-Host "`n----------------------------------------------------------------------------------------"
	
	Write-Host "`nYou will prompted to select the right path where MPARR will be allocated."
	Write-Host "Press any key to continue..." -NoNewLine
	$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
	Write-Host "`n"
	
	[System.Reflection.Assembly]::Load("System.Windows.Forms") | Out-Null
	$folder = New-Object System.Windows.Forms.FolderBrowserDialog
	$folder.UseDescriptionForTitle = $true
	
	# Select MPARR data folder
	$folder.Description = "Select folder where MPARR solution data will be located"
	$folder.rootFolder = 'Recent'
	if ($folder.ShowDialog() -eq "OK")
	{
		$MPARRRootFolder = $folder.SelectedPath 
	}
	
	$BackupFolder = $MPARRRootFolder+"\BackupScripts"
	$CertificateFolder = $MPARRRootFolder+"\Certs"
	$ConfigurationFolder = $MPARRRootFolder+"\ConfigFiles"
	$MPARRLogs = $MPARRRootFolder+"\Logs"
	$MPARRRMSLogs = $MPARRRootFolder+"\RMSLogs"
	$SupportFolder = $MPARRRootFolder+"\Support"
	
	if(-not(Test-Path -Path $BackupFolder))
	{
		Write-Host "Backup data directory is missing, creating a new folder called BackupScripts" -ForegroundColor Blue
		New-Item -ItemType Directory -Force -Path "$MPARRRootFolder\BackupScripts" | Out-Null
	}else
	{
		Write-Host "Folder BackupScripts is already available!" -ForegroundColor Green
	}
	
	if(-not(Test-Path -Path $CertificateFolder))
	{
		Write-Host "Certificate data directory is missing, creating a new folder called Certs" -ForegroundColor Blue
		New-Item -ItemType Directory -Force -Path "$MPARRRootFolder\Certs" | Out-Null
	}else
	{
		Write-Host "Folder Certs is already available!" -ForegroundColor Green
	}
	
	if(-not(Test-Path -Path $ConfigurationFolder))
	{
		Write-Host "Configuration Files directory is missing, creating a new folder called ConfigFiles" -ForegroundColor Blue
		New-Item -ItemType Directory -Force -Path "$MPARRRootFolder\ConfigFiles" | Out-Null
	}else
	{
		Write-Host "Folder ConfigFiles is already available!" -ForegroundColor Green
	}

	if(-not(Test-Path -Path $MPARRLogs))
	{
		Write-Host "MPARR Logs directory is missing, creating a new folder called Logs" -ForegroundColor Blue
		New-Item -ItemType Directory -Force -Path "$MPARRRootFolder\Logs" | Out-Null
	}else
	{
		Write-Host "Folder Logs is already available!" -ForegroundColor Green
	}
	
	if(-not(Test-Path -Path $MPARRRMSLogs))
	{
		Write-Host "MPARR Logs directory is missing, creating a new folder called RMSLogs" -ForegroundColor Blue
		New-Item -ItemType Directory -Force -Path "$MPARRRootFolder\RMSLogs" | Out-Null
	}else
	{
		Write-Host "Folder RMSLogs is already available!" -ForegroundColor Green
	}
	
	if(-not(Test-Path -Path $SupportFolder))
	{
		Write-Host "Support directory is missing, creating a new folder called Support" -ForegroundColor Blue
		New-Item -ItemType Directory -Force -Path "$MPARRRootFolder\Support" | Out-Null
	}else
	{
		Write-Host "Folder Support is already available!" -ForegroundColor Green
	}
	Write-Host "`n"
}

function MPARRCopyFiles
{
	Clear-Host
	cls
	Write-Host "`n`n----------------------------------------------------------------------------------------"
	Write-Host "`nMPARR configuration to copy files from previous configuration!" -ForegroundColor DarkGreen
	Write-Host "This menu helps to copy all the files to the new folder structure required for migration." -ForegroundColor DarkGreen
	Write-Host "`n----------------------------------------------------------------------------------------"
	
	Write-Host "`nYou will prompted to select the right path where MPARR is located."
	Write-Host "Press any key to continue..." -NoNewLine
	$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
	Write-Host "`n"
	
	# Select MPARR source data folder
	[System.Reflection.Assembly]::Load("System.Windows.Forms") | Out-Null
	$folder = New-Object System.Windows.Forms.FolderBrowserDialog
	$folder.UseDescriptionForTitle = $true
	
	Add-Type -AssemblyName PresentationCore,PresentationFramework
	$msgBody = "Please, take care to select the right folder where MPARR is installed - SOURCE -"
	[System.Windows.MessageBox]::Show($msgBody) |Out-Null
	
	$folder.Description = "Select folder where MPARR solution is currently installed"
	$folder.rootFolder = 'Recent'
	if ($folder.ShowDialog() -eq "OK")
	{
		$MPARRSourceFolder = $folder.SelectedPath 
	}
	
	# Select MPARR destination data folder
	$folderDestination = New-Object System.Windows.Forms.FolderBrowserDialog
	$folderDestination.UseDescriptionForTitle = $true
	
	$msgBody2 = "Please, take care to select the right folder where MPARR will be installed - DESTINATION -"
	[System.Windows.MessageBox]::Show($msgBody2) |Out-Null
	
	$folderDestination.Description = "Select folder where MPARR solution will be installed"
	$folderDestination.rootFolder = 'Recent'
	if ($folderDestination.ShowDialog() -eq "OK")
	{
		$MPARRDestinationFolder = $folderDestination.SelectedPath 
	}
	
	$MPARRConfigFile = $MPARRSourceFolder+"\laconfig.json"
	$MPARRSchemasFile = $MPARRSourceFolder+"\schemas.json"
	$CertificateFolder = $MPARRSourceFolder+"\Certs"
	$ConfigurationFolder = $MPARRSourceFolder+"\ConfigFiles"
	$SupportFolder = $MPARRSourceFolder+"\Support"
	
	$ConfigFilesFolder = $MPARRDestinationFolder+"\ConfigFiles"
	$CertsFolder = $MPARRDestinationFolder+"\Certs"
	$LogsFolder = $MPARRDestinationFolder+"\Logs"
	$RMSLogsFolder = $MPARRDestinationFolder+"\RMSLogs"
	$SupportDestFolder = $MPARRDestinationFolder+"\Support"
	
	if(-not(Test-Path -Path $MPARRConfigFile))
	{
		Write-Host "laconfig.json file is not located on the root folder, please check that you selected the right folder." -ForegroundColor DarkYellow
		Write-Host "Please check that the laconfig.json file is located at the root folder."
		Write-Host "Press any key to continue..." -NoNewLine
		$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
		MPARRCopyFiles
	}else
	{
		Copy-Item $MPARRConfigFile -Destination $ConfigFilesFolder
		Write-Host "laconfig.json file copied to : "$ConfigFilesFolder -ForegroundColor Green
	}
	
	if(-not(Test-Path -Path $CertificateFolder))
	{
		Write-Host "Certificate folder was not found, nothing will be copied from this path." -ForegroundColor DarkYellow
	}else
	{
		$CertSource = $CertificateFolder+"\*"
		Copy-Item -Path $CertSource -Destination $CertsFolder -Recurse
		Write-Host "Content from Cert folder copied to : "$CertsFolder -ForegroundColor Green
	}
	
	if(-not(Test-Path -Path $MPARRSchemasFile))
	{
		Write-Host "schemas.json file was not found, nothing will be copied from this path." -ForegroundColor DarkYellow
	}else
	{
		Copy-Item $MPARRSchemasFile -Destination $ConfigFilesFolder
		Write-Host "schemas.json file copied to : "$ConfigFilesFolder -ForegroundColor Green
	}

	if(-not(Test-Path -Path $ConfigurationFolder))
	{
		Write-Host "ConfigFiles directory is missing, nothing will be copied from this path." -ForegroundColor DarkYellow
	}else
	{
		$ConfigFilesSource = $ConfigurationFolder+"\*"
		Copy-Item -Path $ConfigFilesSource -Destination $ConfigFilesFolder -Recurse
		Write-Host "Content from ConfigFiles folder copied to : "$ConfigFilesFolder -ForegroundColor Green
	}
	
	$CONFIGFILE = $MPARRConfigFile 
	$json = Get-Content -Raw -Path $CONFIGFILE
	[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
	$LogsRootFolder = $config.OutPutLogs
	$RMSLogsRootFolder = $config.RMSLogs
	
	if(-not(Test-Path -Path $LogsRootFolder))
	{
		Write-Host "MPARR Logs directory is missing, nothing will be copied from this path." -ForegroundColor DarkYellow
		Write-Host "Please check laconfig.json and the path set for OutPutLogs" -ForegroundColor Red
	}else
	{
		$LogsSource = $LogsRootFolder+"\*"
		Copy-Item -Path $LogsSource -Destination $LogsFolder -Recurse
		Write-Host "Content from Logs folder copied to : "$LogsFolder -ForegroundColor Green
	}
	
	if(-not(Test-Path -Path $RMSLogsRootFolder))
	{
		Write-Host "MPARR RMS Logs directory is missing, nothing will be copied from this path." -ForegroundColor DarkYellow
		Write-Host "Please check laconfig.json and the path set for RMSLogs" -ForegroundColor Red
	}else
	{
		$RMSLogsSource = $RMSLogsRootFolder+"\*"
		Copy-Item -Path $RMSLogsSource -Destination $RMSLogsFolder -Recurse
		Write-Host "Content from RMSLogs folder copied to : "$RMSLogsFolder -ForegroundColor Green
	}
	
	if(-not(Test-Path -Path $SupportFolder))
	{
		Write-Host "Support directory is missing, nothing will be copied from this path." -ForegroundColor DarkYellow
	}else
	{
		$SupportFolderSource = $SupportFolder+"\*"
		Copy-Item -Path $SupportFolderSource -Destination $SupportDestFolder -Recurse
		Write-Host "Content from Support folder copied to : "$SupportDestFolder -ForegroundColor Green
	}
	
	$MPARRFiles = $MPARRSourceFolder+"\MPARR*"
	Copy-Item -Path $MPARRFiles -Destination $MPARRDestinationFolder -Recurse -Force
	Write-Host "All MPARR scripts copied to : "$SupportDestFolder -ForegroundColor Green
	
	###Update laconfig.json file copied to the new MPARR installation
	$ConfigDestinationFile = "$MPARRDestinationFolder\ConfigFiles\laconfig.json"  
    $MPARRjson = Get-Content -Raw -Path $ConfigDestinationFile
    [PSCustomObject]$MPARRconfig = ConvertFrom-Json -InputObject $MPARRjson
	$MPARRconfig.RMSLogs = $RMSLogsFolder+"\"
	$MPARRconfig.OutPutLogs = $LogsFolder+"\"
	$MPARRconfig | ConvertTo-Json | Out-File $ConfigDestinationFile
	Write-Host "laconfig.json updated at the destination folder" -ForegroundColor Green
	Start-Sleep -s 3
	
	Write-Host "`n"
}

function MPARRCopyConfigFilesOnly
{
	Clear-Host
	cls
	
	Write-Host "`n`n----------------------------------------------------------------------------------------"
	Write-Host "`nMPARR config files migration!" -ForegroundColor DarkGreen
	Write-Host "This menu helps to migrate laconfig and schemas from a previous MPARR installation and to apply any new change required." -ForegroundColor DarkGreen
	Write-Host "`n----------------------------------------------------------------------------------------"
	
	Write-Host "`nYou will prompted to select the right path where MPARR is located."
	Write-Host "Press any key to continue..." -NoNewLine
	$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
	Write-Host "`n"
	
	# Select MPARR source data folder
	[System.Reflection.Assembly]::Load("System.Windows.Forms") | Out-Null
	$folder = New-Object System.Windows.Forms.FolderBrowserDialog
	$folder.UseDescriptionForTitle = $true
	
	Add-Type -AssemblyName PresentationCore,PresentationFramework
	$msgBody = "Please, take care to select the right folder where MPARR is installed - SOURCE -"
	[System.Windows.MessageBox]::Show($msgBody) |Out-Null
	
	$folder.Description = "Select folder where MPARR solution is currently installed"
	$folder.rootFolder = 'Recent'
	if ($folder.ShowDialog() -eq "OK")
	{
		$MPARRSourceFolder = $folder.SelectedPath  
	}
	
	# Select MPARR destination data folder
	$folderDestination = New-Object System.Windows.Forms.FolderBrowserDialog
	$folderDestination.UseDescriptionForTitle = $true
	
	$msgBody2 = "Please, take care to select the right folder where MPARR will be installed - DESTINATION -"
	[System.Windows.MessageBox]::Show($msgBody2) |Out-Null
	
	$folderDestination.Description = "Select folder where MPARR solution will be installed"
	$folderDestination.rootFolder = 'Recent'
	if ($folderDestination.ShowDialog() -eq "OK")
	{
		$MPARRDestinationFolder = $folderDestination.SelectedPath 
	}
	
	$MPARRConfigFile = $MPARRSourceFolder+"\laconfig.json"
	$MPARRSchemasFile = $MPARRSourceFolder+"\schemas.json" 
	
	$ConfigFilesFolder = $MPARRDestinationFolder+"\ConfigFiles"
	$LogsFolder = $MPARRDestinationFolder+"\Logs"

	
	if(-not(Test-Path -Path $MPARRConfigFile))
	{
		Write-Host "laconfig.json file is not located on the root folder, please check that you selected the right folder." -ForegroundColor DarkYellow
		Write-Host "Please check that the laconfig.json file is located at the root folder."
		Write-Host "$MPARRConfigFile"
		Write-Host "Press any key to continue..." -NoNewLine
		$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
		MPARRCopyFiles
	}else
	{
		Copy-Item $MPARRConfigFile -Destination $ConfigFilesFolder
		Write-Host "laconfig.json file copied to : "$ConfigFilesFolder -ForegroundColor Green
	}
	
	if(-not(Test-Path -Path $MPARRSchemasFile))
	{
		Write-Host "schemas.json file was not found, nothing will be copied from this path." -ForegroundColor DarkYellow
	}else
	{
		Copy-Item $MPARRSchemasFile -Destination $ConfigFilesFolder
		Write-Host "schemas.json file copied to : "$ConfigFilesFolder -ForegroundColor Green
	}
	
	$CONFIGFILE = $MPARRConfigFile 
	$json = Get-Content -Raw -Path $CONFIGFILE
	[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
	$LogsRootFolder = $config.OutPutLogs
	$MPARRTimeStampFile = $LogsRootFolder+"timestamp.json"
	
	if(-not(Test-Path -Path $LogsRootFolder))
	{
		Write-Host "MPARR Logs directory is missing, nothing will be copied from this path." -ForegroundColor DarkYellow
		Write-Host "Please check laconfig.json and the path set for OutPutLogs" -ForegroundColor Red
	}else
	{
		Copy-Item $MPARRTimeStampFile -Destination $LogsFolder
		Write-Host "timestamp.son file copied to : "$LogsFolder -ForegroundColor Green
	}

	Write-Host "`n"
}

function UpdateMPARRlaconfigFile
{
	Clear-Host
	cls
	
	Write-Host "`n`n----------------------------------------------------------------------------------------"
	Write-Host "`nMPARR laconfigjson file check and update!" -ForegroundColor DarkGreen
	Write-Host "This menu helps to check laconfig.json previously created and to apply any new change required." -ForegroundColor DarkGreen
	Write-Host "`n----------------------------------------------------------------------------------------"
	
	Write-Host "`nYou will prompted to select the right path where the laconfig.json file is located."
	Write-Host "Press any key to continue..."
	$key = ([System.Console]::ReadKey($true)) | Out-Null
	
                
	#Here you start selecting each folder
	[System.Reflection.Assembly]::Load("System.Windows.Forms") | Out-Null
	$file = New-Object System.Windows.Forms.OpenFileDialog
	# Start selecting laconfig.json location  
	$file.Title = "Select folder where laconfig.json is located"
	$file.InitialDirectory = 'ProgramFiles'
	$file.Filter = 'MPARR Config file|laconfig.json'
	# main log directory
	if ($file.ShowDialog() -eq "OK")
	{
		$CONFIGFILE = $file.FileName
		$json = Get-Content -Raw -Path $CONFIGFILE
		[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
	}
	
	$Changes = "False"
	$MicrosoftEntraConfig = $config.MicrosoftEntraConfig
	$ExportToEventHub = $config.ExportToEventHub
	
	if($MicrosoftEntraConfig -eq $Null)
	{
		Write-Host "`n`nMPARR Entra Users is not set, this configuration is required to collect your user data."
		Write-Host "`n`nPlease execute .\MPARR-MicrosoftEntraUsers.ps1 once to set this configuration."
		$config = InitializeLAConfigFile -DirRoot $CONFIGFILE
		$config.MicrosoftEntraConfig = "Not Set"
		WriteToConfigFile -DirRoot $CONFIGFILE
		$Changes = "True"
		Start-Sleep -s 1
	}
	
	
	if($ExportToEventHub -eq $Null)
	{
		Write-Host "MPARR Event Hub connector is not set, this configuration is not mandatory, but is required if you want to use."
		$config = InitializeLAConfigFile -DirRoot $CONFIGFILE
		$config.ExportToEventHub = "False"
		$config.EventHubNamespace = ""
		$config.EventHub = ""
		WriteToConfigFile -DirRoot $CONFIGFILE
		$Changes = "True"
		Start-Sleep -s 1
	}

	if($Changes -eq "True")
	{
		Write-Host "`nlaconfig.json file was" -NoNewLine
		Write-Host "`t`tupdated!`n" -ForegroundColor Green
	}else
	{
		Write-Host "`nlaconfig.json file is up to date, no changes applied.`n" -ForegroundColor Green
	}
	
	$config | ConvertTo-Json | Out-File $CONFIGFILE
	
	Write-Host "Press any key to continue..."
	$key = ([System.Console]::ReadKey($true)) | Out-Null
}

function UpdateMPARRScripts
{
	$MPARRUri = "https://raw.githubusercontent.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/main"

    $result = Invoke-WebRequest -Uri "$MPARRUri/UpdateInfo/update.json"
    $update = $result.Content | ConvertFrom-Json
	$BackupPath = $PSScriptRoot+"\BackupScripts"
	
	if(-Not (Test-Path $BackupPath ))
	{
		Write-Host "Export data directory is missing, creating a new folder called BackupScripts"
		New-Item -ItemType Directory -Force -Path "$PSScriptRoot\BackupScripts" | Out-Null
	}

    foreach ($item in $update.files)
    {
        if($item.format -eq "ps1")
		{
			$destDir = "."
			if ($item.directory -ne "ROOT")
			{
				$destDir = ".\$($item.directory)"
				if (-not (Test-Path $item.directory))
				{
					Write-Host "Creating '$($item.directory)' directory." -ForegroundColor Cyan
					New-Item -Name ($item.directory) -ItemType Directory | Out-Null
				}
			}
			
			Write-Host "`nThe file $($item.file) located at GitHub repo is set to version $($item.Version)"
			$ScriptName = $item.file
			if ($item.directory -ne "ROOT")
			{
				$SupportFolder = $item.directory
				$MPARRFile = $PSScriptRoot+"\"+$SupportFolder+"\"+$ScriptName
			}else
			{
				$MPARRFile = "$PSScriptRoot\$ScriptName"
			}

			if (-not (Test-Path -Path $MPARRFile))
			{
				Write-Host "`nFile $ScriptName was not found" -ForegroundColor Blue
				Write-Host "Downloading $($item.file)..."
				Invoke-WebRequest -Uri "$MPARRUri/$($item.URI)" -OutFile "$destDir\$($item.file)"
			}else
			{
				$validatefile = Test-PSScriptFileInfo -Path $MPARRFile
				$date = Get-Date -Format "yyyyMMdd"
				$BackupFile = $PSScriptRoot+"\BackupScripts\"+$ScriptName+"_"+$date+".backup"
				$SupportFolder = $item.directory
				$SupportFile = "$SupportFolder\$ScriptName"
				if($validatefile -eq "True")
				{
					$var = Test-ScriptFileInfo -Path $MPARRFile | select Version
					if($var.Version -eq $item.Version)
					{
						$VersionValue = $var.Version
						Write-Host "You already have the latest version, version $VersionValue!" -ForegroundColor Green
					}else
					{
						$CloudVersionValue = $item.Version
						$VersionValue = $var.Version
						Write-Host "You have an old version, version $VersionValue, updating to $CloudVersionValue..." -ForegroundColor DarkYellow 
						if ($item.directory -ne "ROOT")
						{
							Move-Item "$SupportFile" "$BackupFile" -Force
						}else
						{
							Move-Item "$ScriptName" "$BackupFile" -Force
						}
						Write-Host "Downloading $($item.file)..."
						Invoke-WebRequest -Uri "$MPARRUri/$($item.URI)" -OutFile "$destDir\$($item.file)"
					}
				}else
				{
					if ($item.directory -ne "ROOT")
					{
						Move-Item "$SupportFile" "$BackupFile" -Force
					}else
					{
						Move-Item "$ScriptName" "$BackupFile" -Force
					}
					Write-Host "`nThe old script file was moved to '$BackupFile'"
				}
			}
		}else
		{
			$ScriptName = $item.file
			$destDir = "."
			if ($item.directory -ne "ROOT")
			{
				$destDir = ".\$($item.directory)"
				if (-not (Test-Path $item.directory))
				{
					Write-Host "Creating '$($item.directory)' directory." -ForegroundColor Cyan
					New-Item -Name ($item.directory) -ItemType Directory | Out-Null
				}
			}
			$SupportFolder = $item.directory
			$MPARRFile = $PSScriptRoot+"\"+$SupportFolder+"\"+$ScriptName
			if (-not (Test-Path -Path $MPARRFile))
			{
				Write-Host "`nSupporting file $ScriptName was not found" -ForegroundColor Blue
				Write-Host "Downloading $($item.file)..."
				Invoke-WebRequest -Uri "$MPARRUri/$($item.URI)" -OutFile "$destDir\$($item.file)"
			}else
			{
				Write-Host "Supporting file $ScriptName is already available!!" -ForegroundColor Cyan
			}
			
		}

    }
}

function CheckMPARROnTheWeb
{
	$MPARRUri = "https://raw.githubusercontent.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/main"

    $result = Invoke-WebRequest -Uri "$MPARRUri/UpdateInfo/update.json"
    $update = $result.Content | ConvertFrom-Json
	$ItemNumber = 1

    foreach ($item in $update.files)
    {
        if($item.format -eq "ps1")
		{
			$destDir = "."
			if ($item.directory -ne "ROOT")
			{
				$destDir = ".\$($item.directory)"
				if (-not (Test-Path $item.directory))
				{
					Write-Host "`n'$($item.directory)' directory was not found." -ForegroundColor DarkYellow
					Write-Host "Please validate the MPARR folder structure at Menu 6 and then Menu 2"
				}
			}
			
			Write-Host "`n$ItemNumber.- The file $($item.file) located at GitHub repo is set to version $($item.Version)" -ForegroundColor DarkBlue
			$ScriptName = $item.file
			if ($item.directory -ne "ROOT")
			{
				$SupportFolder = $item.directory
				$MPARRFile = $PSScriptRoot+"\"+$SupportFolder+"\"+$ScriptName
			}else
			{
				$MPARRFile = "$PSScriptRoot\$ScriptName"
			}

			if (-not (Test-Path -Path $MPARRFile))
			{
				Write-Host "File $ScriptName was not found" -ForegroundColor DarkYellow
				Write-Host "File description in the version '$($item.Version)' is :"$item.changes
				Write-Host "Remember that you can update all the files using this same setup script(Menu 7 and then menu 2)."

			}else
			{
				$validatefile = Test-PSScriptFileInfo -Path $MPARRFile
				$SupportFolder = $item.directory
				$SupportFile = "$SupportFolder\$ScriptName"
				if($validatefile -eq "True")
				{
					$var = Test-ScriptFileInfo -Path $MPARRFile | select Version
					if($var.Version -eq $item.Version)
					{
						$VersionValue = $var.Version
						Write-Host "You already have the latest version, version $VersionValue!" -ForegroundColor Green
					}else
					{
						$CloudVersionValue = $item.Version
						$VersionValue = $var.Version
						Write-Host "You have an old version, version $VersionValue." -ForegroundColor DarkYellow 
						Write-Host "File description in the new version '$($item.Version)' is :"$item.changes
						Write-Host "Remember that you can update all the files using this same setup script(Menu 7 and then menu 2)."
					}
				}else
				{
					Write-Host "You have an old version, without versioning." -ForegroundColor DarkYellow 
					Write-Host "File description in the new version '$($item.Version)' is :"$item.changes
					Write-Host "Remember that you can update all the files using this same setup script(Menu 7 and then menu 2)."
				}
			}
		}else
		{
			Write-Host "`n$ItemNumber.- The file $($item.file) located at GitHub repo is set to version $($item.Version)" -ForegroundColor DarkBlue
			$ScriptName = $item.file
			$destDir = "."
			if ($item.directory -ne "ROOT")
			{
				$destDir = ".\$($item.directory)"
				if (-not (Test-Path $item.directory))
				{
					Write-Host "`n'$($item.directory)' directory was not found." -ForegroundColor DarkYellow
					Write-Host "Please validate the MPARR folder structure at Menu 6 and then Menu 2"
				}
			}
			$SupportFolder = $item.directory
			$MPARRFile = $PSScriptRoot+"\"+$SupportFolder+"\"+$ScriptName
			if (-not (Test-Path -Path $MPARRFile))
			{
				Write-Host "Supporting file $ScriptName was not found" -ForegroundColor DarkYellow
				Write-Host "File description in the version '$($item.Version)' is :"$item.changes
				Write-Host "Remember that you can update all the files using this same setup script(Menu 7 and then menu 2)."
			}else
			{
				Write-Host "Supporting file $ScriptName is already available!!" -ForegroundColor Cyan
			}
			
		}
		$ItemNumber++
    }
	Write-Host "`nPress any key to continue..." -ForegroundColor DarkYellow
	$key = ([System.Console]::ReadKey($true)) | Out-Null
}

function SubMenuMPARRCoreScripts
{
	$choice = 1
	while ($choice -ne "0")
	{
		Clear-Host
		cls
		Write-Host "`n`n----------------------------------------------------------------------------------------"
		Write-Host "`nMPARR Core scripts schedule tasks!" -ForegroundColor DarkBlue
		Write-Host "Here you can set the tasks for Task Scheduler from MPARR Core Scripts." -ForegroundColor Blue
		Write-Host "`n----------------------------------------------------------------------------------------"
		Write-Host "`n### MPARR Core scripts ###" -ForegroundColor Blue
		Write-Host "`nWhat do you want to do?"
		Write-Host "`t[1] - Create MPARR Collector task"
		Write-Host "`t[2] - Create MPARR RMS task"	
		Write-Host "`t[0] - Back to main menu"
		Write-Host "`n"
		Write-Host "`nPlease choose option:"
		
		$choice = ([System.Console]::ReadKey($true)).KeyChar
		switch ($choice) {
        "1" {CreateMPARRCollectorTask; break}
		"2" {CreateMPARRRMSDataTask; break}
		"0" {cls;return}
		}
	
	}
}

function SubMenuMicrosoftGraphAPIScripts
{
	$choice = 1
	while ($choice -ne "0")
	{
		Clear-Host
		cls
		Write-Host "`n`n----------------------------------------------------------------------------------------"
		Write-Host "`nMicrosoft Graph API scripts schedule tasks!" -ForegroundColor DarkBlue
		Write-Host "Here you can set the tasks for Task Scheduler from Microsoft Graph API Scripts." -ForegroundColor Blue
		Write-Host "`n----------------------------------------------------------------------------------------"
		Write-Host "`n### Microsoft Graph API scripts ###" -ForegroundColor Blue
		Write-Host "`nWhat do you want to do?"
		Write-Host "`t[1] - Create Microsoft Entra Users task"
		Write-Host "`t[2] - Create Microsoft Entra Domains task"	
		Write-Host "`t[3] - Create Microsoft Entra Roles task"
		Write-Host "`t[4] - Create Microsoft 365 Cloud Statistics task"
		Write-Host "`t[0] - Back to main menu"
		Write-Host "`n"
		Write-Host "`nPlease choose option:"
		
		$choice = ([System.Console]::ReadKey($true)).KeyChar
		switch ($choice) {
        "1" {CreateMPARRUsersTask; break}
		"2" {CreateMPARRDomainsTask; break}
		"3" {CreateMPARRRolesTask; break}
		"4" {CreateMPARRM365StatisticsTask;break}
		"0" {cls;return}
		}
	
	}
}

function SubMenuMicrosoftPurviewAPIScripts
{
	$choice = 1
	while ($choice -ne "0")
	{
		Clear-Host
		cls
		Write-Host "`n`n----------------------------------------------------------------------------------------"
		Write-Host "`nMicrosoft Purview API scripts schedule tasks!" -ForegroundColor DarkBlue
		Write-Host "Here you can set the tasks for Task Scheduler from Microsoft Graph API Scripts." -ForegroundColor Blue
		Write-Host "Remember that all these scripts require elevate privileges." -ForegroundColor Blue
		Write-Host "If you want to avoid give elevate privileges the scripts can be execute manually." -ForegroundColor Blue
		Write-Host "You can execute manually using .\MPARR-PurviewXXXXXXX -ManualConnection ." -ForegroundColor Blue
		Write-Host "`n----------------------------------------------------------------------------------------"
		Write-Host "`n### Microsoft Graph API scripts ###" -ForegroundColor Blue
		Write-Host "`nWhat do you want to do?"
		Write-Host "`t[1] - Create Microsoft Purview Sensitivity Labels task"
		Write-Host "`t[2] - Create Microsoft Purview Sensitive Information Types task"	
		Write-Host "`t[3] - Create Microsoft Purview Roles task"
		Write-Host "`t[4] - Create Microsoft Purview Content Explorer task"
		Write-Host "`t[0] - Back to main menu"
		Write-Host "`n"
		Write-Host "`nPlease choose option:"
		
		$choice = ([System.Console]::ReadKey($true)).KeyChar
		switch ($choice) {
        "1" {CreateMPARRPurviewLabelsTask; break}
		"2" {CreateMPARRPurviewSITsTask; break}
		"3" {CreateMPARRPurviewRolesTask; break}
		"0" {cls;return}
		}
	
	}
}

function SubMenuTaskSchedulerScripts
{
	$choice = 1
	while ($choice -ne "0")
	{
		Clear-Host
		cls
		Write-Host "`n`n----------------------------------------------------------------------------------------"
		Write-Host "`nMPARR Task scheduler menu!" -ForegroundColor DarkBlue
		Write-Host "Here you can set the tasks under Task Scheduler to run automatically MPARR." -ForegroundColor Blue
		Write-Host "`n----------------------------------------------------------------------------------------"
		Write-Host "`n### Microsoft Graph API scripts ###" -ForegroundColor Blue
		Write-Host "`nWhat do you want to do?"
		Write-Host "`t[1] - Create scheduled task for Core Scripts (MPARR Collector and MPARR RMS)"
		Write-Host "`t[2] - Microsoft Graph API Scripts(Users, Domains, Roles, Cloud Statistics)"
		Write-Host "`t[3] - Microsoft Purview API(Sensitivity Labels, Sensitive Info Types, Purview Roles, Content Explorer )"
		Write-Host "`t[0] - Back to main menu"
		Write-Host "`n"
		Write-Host "`nPlease choose option:"
		
		$choice = ([System.Console]::ReadKey($true)).KeyChar
		switch ($choice) {
			"1" {SubMenuMPARRCoreScripts; break}
			"2" {SubMenuMicrosoftGraphAPIScripts; break}
			"3" {SubMenuMicrosoftPurviewAPIScripts; break}
			"0" {cls;return}
		}
	
	}
}

function SubMenuMicrosoftLicensingScripts
{
	$choice = 1
	while ($choice -ne "0")
	{
		Clear-Host
		cls
		Write-Host "`n`n----------------------------------------------------------------------------------------"
		Write-Host "`nMicrosoft Licensing scripts!" -ForegroundColor DarkBlue
		Write-Host "Here you can execute the script related to Microsoft Licensing friendly names." -ForegroundColor Blue
		Write-Host "Or set a task for that activity." -ForegroundColor Blue
		Write-Host "`n----------------------------------------------------------------------------------------"
		Write-Host "`n### Microsoft Graph API scripts ###" -ForegroundColor Blue
		Write-Host "`nWhat do you want to do?"
		Write-Host "`t[1] - Execute export of Microsoft licensing friendly names"
		Write-Host "`t[2] - Create Microsoft licensing friendly name task"
		Write-Host "`t[3] - Update Microsoft licensing friendly name file"
		Write-Host "`t[0] - Back to main menu"
		Write-Host "`n"
		Write-Host "`nPlease choose option:"
		
		$choice = ([System.Console]::ReadKey($true)).KeyChar
		switch ($choice) {
        "1" {MicrosoftLicensing -LicenseOption 1; break}
		"2" {MicrosoftLicensing -LicenseOption 2; break}
		"3" {MicrosoftLicensing -LicenseOption 3; break}
		"0" {cls;return}
		}
	
	}
}

function SubMenuMigrateMPARR
{
	$choice = 1
	while ($choice -ne "0")
	{
		Clear-Host
		cls
		Write-Host "`n`n----------------------------------------------------------------------------------------"
		Write-Host "`nMPARR migration menu!" -ForegroundColor DarkBlue
		Write-Host "Here you can execute the script related to Migrate MPARR from a previous installation." -ForegroundColor Blue
		Write-Host "`n----------------------------------------------------------------------------------------"
		Write-Host "`n### Microsoft Graph API scripts ###" -ForegroundColor Blue
		Write-Host "`nWhat do you want to do?"
		Write-Host "`t[1] - Check Microsoft Entra App permissions"
		Write-Host "`t[2] - Set MPARR Path and Folder Structure"
		Write-Host "`t[3] - Migrate from a previous MPARR installation"
		Write-Host "`t[4] - Migrate only configuration files"
		Write-Host "`t[5] - Check laconfig.json consistency and fix"
		Write-Host "`t[6] - Configure Event Hub connector"
		Write-Host "`t[0] - Back to main menu"
		Write-Host "`n"
		Write-Host "`nPlease choose option:"
		
		$choice = ([System.Console]::ReadKey($true)).KeyChar
		switch ($choice) {
        "1" {UpdateMPARREntraApp; break}
		"2" {MPARRFolderStructure; break}
		"3" {MPARRCopyFiles; break}
		"4" {MPARRCopyConfigFilesOnly; break}
		"5" {UpdateMPARRlaconfigFile; break}
		"6" {UpdateMPARREventHub; break}
		"0" {cls; return}
		}
	
	}
}

function SubMenuMPARROnTheWeb
{
	$choice = 1
	while ($choice -ne "0")
	{
		Clear-Host
		cls
		Write-Host "`n`n----------------------------------------------------------------------------------------"
		Write-Host "`nMPARR migration menu!" -ForegroundColor DarkBlue
		Write-Host "Here you can execute the script related to Microsoft Licensing friendly names." -ForegroundColor Blue
		Write-Host "Or set a task for that activity." -ForegroundColor Blue
		Write-Host "`n----------------------------------------------------------------------------------------"
		Write-Host "`n### Microsoft Graph API scripts ###" -ForegroundColor Blue
		Write-Host "`nWhat do you want to do?"
		Write-Host "`t[1] - Check about MPARR new scripts"
		Write-Host "`t[2] - Update MPARR from Web"
		Write-Host "`t[0] - Back to main menu"
		Write-Host "`n"
		Write-Host "`nPlease choose option:"
		
		$choice = ([System.Console]::ReadKey($true)).KeyChar
		switch ($choice) {
        "1" {CheckMPARROnTheWeb; break}
		"2" {UpdateMPARRScripts; break}
		"0" {cls;return}
		}
	
	}
}

############
# Main code
############
function MainMenu
{

	Write-Host "`nRunning prerequisites check..."
	CheckPrerequisites

	# disable breaking changes message
	Update-AzConfig -DisplayBreakingChangeWarning $false -Scope Process | Out-Null

	Start-Sleep -s 2

	$choice = 1
	while ($choice -ne "0")
	{
		cls
		Write-Host "`n`n----------------------------------------------------------------------------------------"
		Write-Host "`nWelcome to the MPARR setup script!" -ForegroundColor Green
		Write-Host "Script allows to automatically execute setup steps."
		Write-Host "`n----------------------------------------------------------------------------------------"
		Write-Host "`nWhat do you want to do?"
		Write-Host "`t[1] - Setup MPARR 2(New installation)"
		Write-Host "`t[2] - Encrypt secrets"
		Write-Host "`t[3] - Set MPARR scripts on Task Scheduler"
		Write-Host "`t[4] - Microsoft Licensing"
		Write-Host "`t[5] - Migrate from previous MPARR setup to MPARR 2"
		Write-Host "`t[6] - MPARR on the Web(check new versions, update MPARR)"
		Write-Host "`t[7] - Sign MPARR scripts"
		Write-Host "`t[0] - Exit"
		Write-Host "`n"
		Write-Host "`nPlease choose option:"

		$choice = ([System.Console]::ReadKey($true)).KeyChar
		switch ($choice) {
			"1" {
					Connect-AzAccount -WarningAction SilentlyContinue | Out-Null
					SConnectToLA
					NewApp
					GetTenantInfo
					SelectCloud
					SelectLogPath
					UpdateMPARREventHub
					break
				}
			"2" {EncryptSecrets; break}
			"3" {SubMenuTaskSchedulerScripts; break}
			"4" {SubMenuMicrosoftLicensingScripts; break}
			"5" {SubMenuMigrateMPARR; break}
			"6" {SubMenuMPARROnTheWeb; break}
			"7" {SelfSignScripts; break}
			"0" {cls; exit}
			
		}
	}
}
MainMenu
