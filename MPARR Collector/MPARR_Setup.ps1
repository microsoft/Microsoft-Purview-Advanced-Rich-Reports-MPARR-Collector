<#
.SYNOPSIS
    Script to setup MPARR data collector.

.DESCRIPTION
    Script is designed to simplify MPARR setup.
    
    To automate setup, simply run the script and choose one of the following options:

    [1] - Full setup (select Subscription, Log Analytics workspace, create Azure app registration, specify required parameters)
    [2] - Encrypt secrets
	[3] - Create scheduled task
    [0] - Exit 
    
.NOTES
    Version 0.36
    Current version - 28.09.2023
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
                    -ArgumentList '-Command "&{Write-Host "Installing module AIPService..."; [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Import-Module PowerShellGet; Install-Module AIPService -Force; Write-Host "Exiting Windows PowerShell session..."; Start-Sleep -Seconds 2}"'

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
        Write-Host "`tCurrent version is $($Host.Version). Please note that MPARR-RMSData.ps1 script must be executed under PowerShell 5.1."
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

    Write-Host "`n"
}

# Function to create Azure App
function NewApp
{
    Connect-MgGraph -Scopes "Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All", "Directory.ReadWrite.All", "User.ReadWrite.All"

    $appName = "MPARR-DataCollector"
    Get-MgApplication -ConsistencyLevel eventual -Count appCount -Filter "startsWith(DisplayName, 'MPARR-DataCollector')" | Out-Null
    if ($appCount -gt 0)
    {   
        $sufix = ((New-Guid) -split "-")[0]
        $appName = "MPARR-DataCollector-$sufix"
        Write-Host "'MPARR-DataCollector' app already exists. New name was generated: '$appName'`n"
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
        $ok = $false
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
        $ok = $false
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

    Remove-Variable cert
    Remove-Variable certBase64
    Remove-Variable secret
}

function GetTenantInfo
{
    $tenant = Get-MgDomain
    $config.TenantGUID = (Get-MgContext).TenantId
    $config.TenantDomain = ($tenant | Where-Object IsDefault).Id
    $config.OnmicrosoftURL = ($tenant | Where-Object IsInitial).Id
}

function SelectCloud
{
    $choices = '&Commercial', '&GCC', 'GCC&H', '&DOD'
    $decision = $Host.UI.PromptForChoice("", "`nPlease select cloud version:", $choices, 0)
    switch ($decision) {
        0 {$config.Cloud = "Commercial"; break}
        1 {$config.Cloud = "GCC"; break}
        2 {$config.Cloud = "GCCH"; break}
        3 {$config.Cloud = "DOD"; break}
    }
}

# function to choose destination directory for logs
function SelectLogPath
{
    $choices  = '&Yes', '&No'
    $decision = $Host.UI.PromptForChoice("", "Default locations for logs are '$($config.RMSLogs)' and '$($config.OutPutLogs)'. Do you want change the location?", $choices, 1)
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
    }
}

# write configuration data to json file
function WriteToJsonFile
{
    if (Test-Path "laconfig.json")
    {
        $date = Get-Date -Format "yyyyMMddHHmmss"
        Move-Item "laconfig.json" "laconfig_$date.json"
        Write-Host "`nThe old config file moved to 'laconfig_$date.json'"
    }
    $config | ConvertTo-Json | Out-File "laconfig.json"
    Write-Host "Setup completed. New config file was created." -ForegroundColor Green
}


function CreateScheduledTask
{
    # main data collector script
    $taskName = "MPARR-DataCollector"

    # calculate date
    $dt = Get-Date
    $nearestMinutes = 15 
    $reminder = $dt.Minute % $nearestMinutes
    $dt = $dt.AddMinutes(-$reminder)
    $startTime = [datetime]::new($dt.Year, $dt.Month, $dt.Day, $dt.Hour, $dt.Minute, 0)

    #create task
    $trigger = New-ScheduledTaskTrigger -Once -At $startTime -RepetitionInterval (New-TimeSpan -Minutes $nearestMinutes)
    #$filePath = Join-Path $PSScriptRoot "MPARR_Collector.ps1"
    #$workingDir = "{0}" -f ("$PSScriptRoot", "`"$PSScriptRoot`"")[$PSScriptRoot.Contains(" ")]
    $action = New-ScheduledTaskAction -Execute "`"$PSHOME\pwsh.exe`"" -Argument ".\MPARR_Collector.ps1" -WorkingDirectory $PSScriptRoot
    $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopOnIdleEnd -AllowStartIfOnBatteries `
         -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours 1)

    if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) 
    {
        Write-Host "`nScheduled task named '$taskName' already exists.`n" -ForegroundColor Yellow
    }
    else 
    {
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings `
        -RunLevel Highest -ErrorAction Stop | Out-Null
        Write-Host "`nScheduled task named '$taskName' was created.`nFor security reasons you have to specify run as account manually.`n" -ForegroundColor Yellow
    }

    # RMS data script
    $taskName = "MPARR-DataCollector-RMS"
    if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) 
    {
        Write-Host "`nScheduled task named '$taskName' already exists.`n" -ForegroundColor Yellow
        return
    }
    $dt = $dt.AddMinutes(5)
    $startTime = [datetime]::new($dt.Year, $dt.Month, $dt.Day, $dt.Hour, $dt.Minute, 0)
    $trigger = New-ScheduledTaskTrigger -Once -At $startTime -RepetitionInterval (New-TimeSpan -Minutes $nearestMinutes)
    #$filePath = Join-Path $PSScriptRoot "MPARR-RMSData.ps1"
    $action = New-ScheduledTaskAction -Execute '"C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe"' -Argument ".\MPARR-RMSData.ps1" -WorkingDirectory $PSScriptRoot
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings `
        -RunLevel Highest -ErrorAction Stop | Out-Null
    Write-Host "`nScheduled task named '$taskName' was created.`nFor security reasons you have to specify run as account manually.`n" -ForegroundColor Yellow
}

function EncryptSecrets
{
    # read config file
    $CONFIGFILE = "$PSScriptRoot\laconfig.json"  
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
    Move-Item "laconfig.json" "laconfig_$date.json"
    Write-Host "`nSecrets encrypted."
    Write-Host "The old config file moved to 'laconfig_$date.json'" -ForegroundColor Green
    $config | ConvertTo-Json | Out-File $CONFIGFILE

    Write-Host "Warning!" -ForegroundColor Yellow
    Write-Host "Please note that encrypted keys can be decrypted only on this machine, using the same account." -ForegroundColor Yellow
}

############
# Main code
############
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
}

Write-Host "`nRunning prerequisites check..."
CheckPrerequisites

# disable breaking changes message
Update-AzConfig -DisplayBreakingChangeWarning $false -Scope Process | Out-Null

Write-Host "`n`n----------------------------------------------------------------------------------------"
Write-Host "`nWelcome to the MPARR setup script!"
Write-Host "Script allows to automatically execute setup steps.`n"


$choice = 1
while ($choice -ne "0")
{
    Write-Host "`n`n----------------------------------------------------------------------------------------"
    Write-Host "`nWhat do you want to do?"
    Write-Host "`t[1] - Setup MPARR (select LA, register Azure app...)"
    Write-Host "`t[2] - Encrypt secrets"
    Write-Host "`t[3] - Create scheduled task"
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
                WriteToJsonFile
                break
            }
        "3" {CreateScheduledTask; break}
        "2" {EncryptSecrets; break}
    }
}

