<#PSScriptInfo

.VERSION 2.0.5

.GUID 883af802-165c-4705-b4c1-352686c02f01

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
This script permit to collect some Cloud statistics, total active files on SPO/OD and total number of sites, using Microsoft Graph API

#>

<#
HISTORY
Script      : MPARR-M365CloudStatistics.ps1
Author      : S. Zamorano
Version     : 2.0.5
Description : This script permit to collect some Cloud statistics, total active files on SPO/OD and total number of sites, using Microsoft Graph API

.NOTES 
	04-01-2024	S. Zamorano		- baseline first release
	07-02-2024	S. Zamorano		- Added EventHub connector
	07-02-2024  G. Berdzik  	- Set Graph connection functions
	07-02-2024	S. Zamorano		- Set integration
	12-02-2024	S. Zamorano		- Version released
	01-03-2024	S. Zamorano		- Public release
#>

using module "ConfigFiles\MPARRUtils.psm1"
param (
    # Log Analytics table where the data is written to. Log Analytics will add an _CL to this name.
    [string]$TableName = "M365Statistics",
	[Parameter()] 
        [switch]$ExportToCSVFileOnly,
	[Parameter()] 
        [switch]$ExportToJSONFileOnly,
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

function ValidateConfigurationFile
{
	#Validate laconfig.json that manage the configuration for connections
	$MPARRConfiguration = "$PSScriptRoot\ConfigFiles\laconfig.json"
	
	if (-not (Test-Path -Path $MPARRConfiguration))
	{
		Write-Host "`n##########################################################################################" -ForeGroundColor Yellow
		Write-Host "`nThe laconfig.json file is missing. Check if you are using the right path or execute MPARR_Setup.ps1 first."
		Write-Host "`nThe laconfig.json is required to continue."
		Write-Host "`n##########################################################################################" -ForeGroundColor Yellow
		Write-Host "`n"
		if($ExportToCSVFileOnly -Or $ExportToJSONFileOnly)
		{
			
			Write-Host "`n##########################################################################################" -ForeGroundColor Yellow
			Write-Host "`n"
			Write-Host "`nThe laconfig.json is required to continue." -ForeGroundColor DarkYellow
			Write-Host "`n"
			Write-Host "`n##########################################################################################" -ForeGroundColor Yellow
			Write-Host "`n"
			Write-Host "`n"
			exit
			
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

function GetM365CloudStatistics($OptionEventHub)
{
	$BaseURI = "https://graph.microsoft.com/v1.0/reports"
	$MGREPORTSFILE = "$PSScriptRoot\ConfigFiles\mgreports.json"
	$json = Get-Content -Raw -Path $MGREPORTSFILE
    [PSCustomObject]$reportsJSON = ConvertFrom-Json -InputObject $json
    $ReportNames = $reportsJSON.reports
	
	Write-Host "Option :"$OptionEventHub -ForeGroundColor Blue
	
	Write-Host -ForegroundColor Blue -BackgroundColor white "Collecting and Exporting Log data"
	foreach($ReportName in $ReportNames)
    {    
        Write-Host -ForegroundColor Cyan "`n-> Collecting log data from '" -NoNewline
        Write-Host -ForegroundColor White -BackgroundColor DarkGray $($ReportName.report) -NoNewline
        Write-Host -ForegroundColor Cyan "': " -NoNewline

        # check for token expiration
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

        $URI = "$BaseURI/get$($ReportName.report)"
        if ($ReportName.parameters -eq "period")
        {
            $URI += "(period='D7')"
        }

        Write-Host $URI -ForegroundColor Yellow
        <#try {#>
            $results = @()
			$tempArray = @()
			$date = Get-Date -Format "yyyyMMdd"
			$results = Invoke-RestMethod -Method Get -Uri $URI -Headers $headers -ErrorAction Stop
			$tempArray += $results | ConvertFrom-Csv

			$i = 0
			While($i -lt $tempArray.count)
			{
				$tempArray[$i] | Add-Member -MemberType NoteProperty -Name 'CloudStatistics' -Value $ReportName.report
				$i++
			}
			
			if($ExportToCSVFileOnly -Or $ExportToJSONFileOnly)
			{
				$ExportPath = $PSScriptRoot+"\ExportedData"
				if(-Not (Test-Path $ExportPath ))
				{
					Write-Host "Export data directory is missing, creating a new folder called ExportedData"
					New-Item -ItemType Directory -Force -Path "$PSScriptRoot\ExportedData" | Out-Null
				}
				$date = Get-Date -Format "yyyyMMdd"
				$ErrorFile = "MPARR - M365 Cloud Statistics - Error - "+$date+".json"
				
				if($ExportToCSVFileOnly)
				{
					$ExportCSVFile = "MPARR - M365 Cloud Statistics - "+$date+".csv"
					$pathCSV = $PSScriptRoot+"\ExportedData\"+$ExportCSVFile
					$tempArray | Export-CSV $pathCSV -Append
					Write-Host "`nExport file was named as :" -NoNewLine
					Write-Host $ExportCSVFile -ForeGroundColor Green 
				}
				if($ExportToJSONFileOnly)
				{
					$ExportJSONFile = "MPARR - M365 Cloud Statistics - "+$date+".json"
					$pathJSON = $PSScriptRoot+"\ExportedData\"+$ExportJSONFile
					$log_analytics_array = @()            
					foreach($i in $tempArray) {
						$log_analytics_array += $i
					}
					$json = $log_analytics_array | ConvertTo-Json
					$json | Set-Content -Path $pathJSON
					Write-Host "`nExport file was named as :" -NoNewLine
					Write-Host $ExportJSONFile -ForeGroundColor Green 
				}
				
				Write-Host "`nFile was copied at :" -NoNewLine
				Write-Host $PSScriptRoot"\ExportedData" -ForeGroundColor Green 
				Write-Host "`n"
			}elseif($OptionEventHub -eq "True")
			{
				# Else format for Log Analytics
				EventHubConnection
				$ErrorFile = "MPARR - M365 Cloud Statistics - Error - "+$date+".json"
				$log_analytics_array = @()            
				foreach($i in $tempArray) {
					$log_analytics_array += $i
				}
				$EventHubInstance.PublishToEventHub($log_analytics_array, $ErrorFile)
			}else
			{
				# Else format for Log Analytics
				$log_analytics_array = @()            
				foreach($i in $tempArray) {
					$log_analytics_array += $i
				}

				# Push data to Log Analytics
				Post-LogAnalyticsData -LogAnalyticsTableName $TableName -body $log_analytics_array
			}
        <#}
        catch {
            Write-Host "$(($_ | ConvertFrom-Json).error.message)" -ForegroundColor Red
        }#>
	}
}

function Export-M365CloudStatistics 
{
    # ---------------------------------------------------------------   
    #    Name           : Export-M365CloudStatistics
    #    Desc           : Extracts data from Microsoft Graph related a total files on SPO and ODB, plus total sites on SPO into Log analytics workspace tables for reporting purposes
    #    Return         : None
    # ---------------------------------------------------------------
    
    GetAuthToken
	
	$CONFIGFILE = "$PSScriptRoot\ConfigFiles\laconfig.json"   
	$MGREPORTSFILE = "$PSScriptRoot\ConfigFiles\mgreports.json"

	$json = Get-Content -Raw -Path $CONFIGFILE
	[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
	$OutputPath = $config.OutPutLogs
	
	if ($OutputPath -eq "")
	{
		New-Item -ItemType Directory -Force -Path "$PSScriptRoot\Logs" | Out-Null
		$OutputPath = "$PSScriptRoot\Logs\"
		Write-Host "'OutputLogs' has no value. Default value was assigned: $OutputPath." -ForegroundColor Yellow
	}
	if (-not $OutputPath.EndsWith("\"))
	{
		$OutputPath += "\"
	}
	CheckOutputDirectory $OutputPath
	
	if (-not (Test-Path -Path $MGREPORTSFILE))
	{
		Write-Host "mgreports.json file is missing or is not located on the righ path. Exiting..." -ForegroundColor Yellow
		Write-Host "Check that the file is located in the ConfigFiles folder"
		exit(1)
	}
	else 
	{
		$json = Get-Content -Raw -Path $MGREPORTSFILE
		[PSCustomObject]$reportsJSON = ConvertFrom-Json -InputObject $json
		$ReportNames = $reportsJSON.reports
		Write-Host "ReportNames list: $($ReportNames.report)"    
	}
	
	$OptionEventHub = CheckExportOption
	
	GetM365CloudStatistics -OptionEventHub $OptionEventHub
	
}
    
#Run the script.
CheckPrerequisites
if($CreateTask)
{
	CreateMPARRM365StatisticsTask
	exit
}
Export-M365CloudStatistics 
