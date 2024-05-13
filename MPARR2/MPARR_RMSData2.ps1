<#PSScriptInfo

.VERSION 2.0.7

.GUID 883af802-165c-4701-b4c1-352686c02f01

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
The script exports Aipservice Log Data from Microsoft AADRM API and pushes into a customer-specified Log Analytics table. Please note if you change the name of the table - you need to update Workbook sample that displays the report , appropriately. Do ensure the older table is deleted before creating the new table - it will create duplicates and Log analytics workspace doesn't support upserts or updates.
 
#>

<#
.NOTES 

2022-10-19		S. Zamorano		- Added laconfig.json file for configuration and decryption function
2022-11-18      G.Berdzik       - Fixed issue with data parsing
2022-12-21      G.Berdzik       - Changed logic to avoid data duplicates
2022-12-28      S. Zamorano     
2023-01-02      G.Berdzik       - Minor change (check for output directory)
2023-01-25      G.Berdzik       - Added code for Get-AipServiceTrackingLog data
2023-01-26      G.Berdzik       - Added support for multithreading
2023-11-06      G.Berdzik       - Fixes
2023-12-06      G.Berdzik       
2023-12-12      G.Berdzik       - Optional export to file for RMS Details
2023-12-13      G.Berdzik

HISTORY
Script      : MPARR-RMSData2.ps1
Author      : S. Zamorano
Version     : 2.0.7

.NOTES (Version 2)
	02-02-2024	S. Zamorano		- Script was re written and EventHub connector added
	02-02-2024	S. Zamorano		- First release
	14-02-2024	Berdzik\Zamorano- Added function to call support scripts using PowerShell 5
	01-03-2024	S. Zamorano		- Public release
	18-03-2024	S. Zamorano		- Logic to process the array associated to ContenID field is changed to process by day.
	10-05-2024 	G.Berdzik		- Fix applied to array management for Content ID
#> 

using module "ConfigFiles\MPARRUtils.psm1"
param (
    # Log Analytics table where the data is written to. Log Analytics will add an _CL to this name.
    [string]$TableName = "RMSData",
	[int]$NumberOfThreads = 4,
	[string]$batchSize = 8MB,
	[Parameter()] 
        [switch]$ExportToJSONFileOnly,
	[Parameter()] 
        [switch]$ExportToCSVFileOnly,
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

function CheckOutputDirectory
{
	# path should not be on root drive
    if ($RMSLogs.EndsWith(":\"))
    {
        Write-Host -ForegroundColor Red "Path should not be on root drive. Exiting."
        exit(1)
    }

    # verify folder exists, if not try to create it
    if (!(Test-Path($RMSLogs)))
    {
        Write-Host -ForegroundColor Yellow ">> Warning: '$RMSLogs' does not exist. Creating one now..."
        Write-host -ForegroundColor Gray "Creating '$RMSLogs': " -NoNewline
        try
        {
            New-Item -ItemType "directory" -Path $RMSLogs -ErrorAction Stop | Out-Null
            Write-Host -ForegroundColor Green "Path '$RMSLogs' has been created successfully"
        } catch {
            write-host -ForegroundColor Red "FAILED to create '$RMSLogs'"
            Write-Host -ForegroundColor Red ">> ERROR: The directory '$RMSLogs' could not be created."
            Write-Host -ForegroundColor Red $error[0]
        }
    }
    else{
        Write-Host -ForegroundColor Green "Path '$RMSLogs' already exists"
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
		Write-Host ".\MPARR_RMSData2.ps1 -ExportToCSVFileOnly -ManualConnection" -ForeGroundColor Green
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
				Write-Host ".\MPARR_RMSData2.ps1 -ExportToCSVFileOnly -ManualConnection" -ForeGroundColor Green
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
	CheckOutputDirectory
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
	$EventHubNamespace = $config.EventHubNamespace
	$EventHub = $config.EventHub

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
    $MPARRTSFolder = "MPARR"
	$taskFolder = "\"+$MPARRTSFolder+"\"
	$choices  = '&Proceed', '&Change', '&Existing'
	Write-Host "Please consider if you want to use the default location you need select Existing and the option 1." -ForegroundColor Yellow
    $decision = $Host.UI.PromptForChoice("", "Default task Scheduler Folder is '$MPARRTSFolder'. Do you want to Proceed, Change the name or use Existing one?", $choices, 0)
    if ($decision -eq 1)
    {
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

function CreateMPARRRMSDataTask
{
	# MPARR-ContentExplorerData script
    $taskName = "MPARR-RMSDatav2"
	
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

function CheckExportOption
{
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

function ExecuteRemoteRMSScript
{
	$ContentIds = $ContentIds
	
	# get the newest log file and set StartTime
	$processedFiles = Get-ChildItem "$RMSLogs\*.zip" | Sort-Object -Property Name -Descending

	if ($processedFiles.Count -gt 0)
	{
		$lastFile = $processedFiles[0].FullName
		$lastFile -match ".*(?<date>\d{4}-\d{2}-\d{2}).*" | Out-Null
		$fileDate = $Matches.date
		$StartTime = ([datetime]$fileDate).AddDays(1)
		if ($StartTime -eq (Get-Date).Date)
		{
			Write-Host "Logs are up to date - nothing to download. Exiting."
			Disconnect-AipService
			exit(0)
		}
	}else 
	{
		$StartTime = (Get-Date).AddDays(-1)   
	}

	$timeOffset = [System.TimezoneInfo]::Local.BaseUtcOffset.TotalMinutes
	$StartTime = $StartTime.AddMinutes($timeOffset)
	$EndTime = [datetime]::Today.AddMinutes($timeOffset -1)

	$ea = $ErrorActionPreference
	$ErrorActionPreference = "SilentlyContinue"
	
	###Start call script using PowerShell 5
	Write-Host "Calling MPARR-GetRMSData.ps1 script to get Users logs..."

	$StartTime = $StartTime.ToString("MM/dd/yyyy")
	$EndTime = $EndTime.ToString("MM/dd/yyyy")
	$RMSPath = $RMSLogs
	$ScriptPath = $PSScriptRoot
	$Connection = "Auto"
	if($ManualConnection)
	{
		$Connection = "Manual"
	}
		
	$sb = 
	{
		param(
			$NumberOfThreads,
			$Connection,
			$StartTime,
			$EndTime,
			$RMSPath,
			$ScriptPath
			)
		Set-Location $ScriptPath
		&"$ScriptPath\ConfigFiles\MPARR-GetRMSData.ps1" -NumberOfThreads $NumberOfThreads -Connection $Connection -StartTime $StartTime -EndTime $EndTime -RMSPath $RMSPath #-ContentIds $ContentIds
	}

	Write-Host "Starting script" -ForegroundColor Green
	$job = Start-Job -ScriptBlock $sb -ArgumentList $NumberOfThreads,$Connection,$StartTime,$EndTime,$RMSPath,$PSScriptRoot -PSVersion 5.1 -Verbose
	
	$var = $job | Select-Object -Property State
	Write-Host $var.State -NoNewline

	while ($var.State -ne "Completed")
	{
		Write-Host "." -NoNewline
		Start-Sleep -s 3
		$var = $job | Select-Object -Property State
	}

	Write-Host "`nCollector script execution "
	Write-Host $var.State"!!!" -ForeGroundColor Green
	### End call script using PowerShell 5
	
	$ErrorActionPreference = $ea
	if ($MyError.Count -gt 0)
	{
		Write-Host $MyError[0].ErrorRecord -ForegroundColor Red
		Write-Host "Exiting..."
		exit(2)
	}
}

function ExecuteRemoteTrackingScript($ContentIds)
{
	$ContentIds = $ContentIds
	
	$ea = $ErrorActionPreference
	$ErrorActionPreference = "SilentlyContinue"
	
	###Start call script using PowerShell 5
	Write-Host "`n`nCalling MPARR-GetTrackingRMSData.ps1 script to get Tracking logs..."
	Write-Host "This process can take time depend on the logged volume of access protected data." -ForegroundColor DarkYellow

	$RMSPath = $RMSLogs
	$ScriptPath = $PSScriptRoot
	$Connection = "Auto"
	if($ManualConnection)
	{
		$Connection = "Manual"
	}
		
	$sb = 
	{
		param(
			$Connection,
			$RMSPath,
			$ContentIds,
			$ScriptPath
			)
		Set-Location $ScriptPath
		&"$ScriptPath\ConfigFiles\MPARR-GetTrackingRMSData.ps1" -Connection $Connection -RMSPath $RMSPath -ContentIds $ContentIds
	}

	Write-Host "Starting script" -ForegroundColor Green
	$job = Start-Job -ScriptBlock $sb -ArgumentList $Connection,$RMSPath,$ContentIds,$PSScriptRoot -PSVersion 5.1 -Verbose
	
	$var = $job | Select-Object -Property State
	Write-Host $var.State -NoNewline

	while ($var.State -ne "Completed")
	{
		Write-Host "." -NoNewline
		Start-Sleep -s 3
		$var = $job | Select-Object -Property State
	}

	Write-Host "`nTracking Collector script execution "
	Write-Host $var.State"!!!" -ForeGroundColor Green
	### End call script using PowerShell 5
	
	$ErrorActionPreference = $ea
	if ($MyError.Count -gt 0)
	{
		Write-Host $MyError[0].ErrorRecord -ForegroundColor Red
		Write-Host "Exiting..."
		exit(2)
	}
}

function TrackingRMSDetails($ContentIds)
{
	$ExportTracking = "$RMSLogs\TrackingLogs"
	$ExportPath = $PSScriptRoot+"\ExportedData"
	$datefile = Get-Date 
	$datefile = $datefile.AddDays(-1).ToString("yyyy-MM-dd")
	
	if(-Not (Test-Path $ExportTracking ))
	{
		Write-Host "Export Tracking Log Data directory is missing, creating a new folder called TrackingLogs"
		New-Item -ItemType Directory -Force -Path $ExportTracking | Out-Null
	}
	if(-Not (Test-Path $ExportPath ))
	{
		Write-Host "Export Tracking Log Data directory is missing, creating a new folder called TrackingLogs"
		New-Item -ItemType Directory -Force -Path $ExportPath | Out-Null
	}
	
	ExecuteRemoteTrackingScript -ContentIds $ContentIds
	
	$TrackingFiles = Get-ChildItem $ExportTracking -Filter *.json

	# If there is no data, skip
	if ($TrackingFiles.Count -eq 0) { continue; }

	# Else format for Log Analytics
	$ExportTrackingPath = $ExportPath+"\Tracking" 

	foreach ($TrackingFile in $TrackingFiles)
	{
		if($ExportToCSVFileOnly -Or $ExportToJSONFileOnly)
		{
			if(-Not (Test-Path $ExportTrackingPath ))
			{
				Write-Host "Export data directory is missing, creating a new folder called ExportedData"
				New-Item -ItemType Directory -Force -Path $ExportTrackingPath | Out-Null
			}
			
			if($ExportToCSVFileOnly)
			{
				$TrackingName = $TrackingFile.BaseName+".csv"
				$TrackingDestination = $ExportTrackingPath+"\"+$TrackingName
				$CSVContent = Get-Content -Path $TrackingFile | ConvertFrom-Json
				$CSVContent | Export-Csv $TrackingDestination -Append
				
			}
			if($ExportToJSONFileOnly)
			{
				Copy-Item $TrackingFile -Destination $ExportTrackingPath -Force
			}
			Write-Host "File was copied at :" -NoNewLine
			Write-Host $ExportTrackingPath -ForeGroundColor Green 
		} elseif($OptionEventHub -eq "True")
		{
			$data = Get-Content -Raw -Path $TrackingFile | ConvertFrom-Json
			ExportDataAsCurrentOption -sourcedata $data -ExportOption 3
		} else
		{
			# Else format for Log Analytics
			# Push data to Log Analytics
			$data = Get-Content -Raw -Path $TrackingFile | ConvertFrom-Json

			Post-LogAnalyticsData -LogAnalyticsTableName ($TableName + "Details") -body $data
		}

		Write-Host "File '$TrackingFile' was processed."
		
		Move-Item $TrackingFile "$($TrackingFile).processed" -Force
	}
	$compress = @{
		Path = $ExportTracking+"\*.processed"
		CompressionLevel = "Fastest"
		DestinationPath = $ExportTracking+"\MPARR - RMS Tracking Logs - "+$datefile+".zip"
	}
	Compress-Archive @compress -Force
	$RemoveFiles = $ExportTracking+"\*.processed"
	Remove-Item -Path $RemoveFiles -Force
	
}

function ExportDataAsCurrentOption([array]$sourcedata, $ExportOption, $exportname)
{
	$dateError = Get-Date -Format "yyyy-MM-dd"
	if($ExportOption -eq 2)
	{
		$ExportPath = $PSScriptRoot+"\ExportedData"
		$FinalDestination = $ExportPath+"\"+$exportname
		$json = $sourcedata | ConvertTo-Json
		$json | Add-Content -Path $FinalDestination
	}
	if($ExportOption -eq 3)
	{
		Write-Host "`nExporting to Event Hub..." -Foregroundcolor Green
		EventHubConnection
		$ErrorFile = "MPARR - RMS Logs - Error - "+$dateError+".json"

		$EventHubInstance.PublishToEventHub($sourcedata, $ErrorFile)
	}
	if($ExportOption -eq 4)
	{
		Write-Host "`nExporting to Logs Analytics..." -Foregroundcolor Green
		# Push data to Log Analytics
		Post-LogAnalyticsData -LogAnalyticsTableName $TableName -body $sourcedata
	}
	if($ExportOption -eq 5)
	{
		Write-Host "`nExporting to Logs Analytics..." -Foregroundcolor Green
		# Push data to Log Analytics
		Post-LogAnalyticsData -LogAnalyticsTableName ($TableName+"Details") -body $sourcedata
	}
}

function FillContentIdsTable($source)
{
    $inputRows = $source | Where-Object {$_."content-id" -ne "-"} | Select-Object "content-id"
    if ($inputRows.Count -gt 0)
    {
        $inputRows = @($inputRows."content-id" -replace "[{}]", "")
        [void]$RMSContentID.AddRange($inputRows)
    }
}

function Export-RMSDatav2 
{
    # ---------------------------------------------------------------   
    #    Name           : Export-RMSDatav2
    #    Desc           : Extracts data from Get-AIPServiceLogs and Get-AIPTrackingLogs into Log analytics workspace tables for reporting purposes
    #    Return         : None
    # ---------------------------------------------------------------
	
	#$RMSContentID = @()
	$RMSContentID = New-Object System.Collections.ArrayList
	$ExportPath = $PSScriptRoot+"\ExportedData"
	$datefile = Get-Date 
	$datefile = $datefile.AddDays(-1).ToString("yyyy-MM-dd")
	if(-Not (Test-Path $ExportPath ))
	{
		Write-Host "Export data directory is missing, creating a new folder called ExportedData"
		New-Item -ItemType Directory -Force -Path "$PSScriptRoot\ExportedData" | Out-Null
	}
	
	ExecuteRemoteRMSScript
	
	$files = Get-ChildItem $RMSLogs -Filter *.log

	# If there is no data, skip
	if ($files.Count -eq 0) { continue; }
	
	#Advice related to the execution time
	Write-Host "`nPlease be aware that this process can take several minutes or inclusive hours depends on the volume data, the tokens are refreshed to maintain the connections." -ForeGroundColor DarkYellow
	Start-Sleep -s 2

	# Else format for Log Analytics
	$loopNumber = 0
	#$ContentIDData = New-Object PSObject
	foreach ($file in $files)
	{
		$csv = New-Object System.Collections.ArrayList
		$GetLogs = Get-Content -Path $file -TotalCount 4
		$csvHeader = ($GetLogs | Select-String "^#Fields:").ToString().Replace("#Fields: ", "")
        [void]$csv.Add($csvHeader)
		$srFile = [System.IO.StreamReader]::new($file)
		# skip first 4 lines (header)
		for ($i=0; $i -lt 4; $i++)
		{
			$srFile.ReadLine() | Out-Null
		}
		
		$ContentSize = 0
		while ($line = $srFile.ReadLine())
		{
			[void]$csv.Add($line)
			$ContentSize += $line.Length 
			$TotalContentSize += $line.Length 
			if($ContentSize -gt $batchSize)
			{
				Start-Sleep -s 1
				$ContentSize = 0
				$data = $csv | ConvertFrom-Csv -Delimiter "`t"
				FillContentIdsTable $data

				if($ExportToCSVFileOnly)
				{
					continue
				}elseif($ExportToJSONFileOnly)
				{
					$ExportedName = "MPARR - "+$file.BaseName+".json"
					ExportDataAsCurrentOption -sourcedata $data -ExportOption 2 -exportname $ExportedName
				}elseif($OptionEventHub -eq "True")
				{
					ExportDataAsCurrentOption -sourcedata $data -ExportOption 3
				}else
				{
					ExportDataAsCurrentOption -sourcedata $data -ExportOption 4
				}
				$loopNumber++
				$csv.Clear()
				$csv = New-Object System.Collections.ArrayList
				[void]$csv.Add($csvHeader)

			}
		}
		$srFile.Close()
		$data = $csv | ConvertFrom-Csv -Delimiter "`t"
		FillContentIdsTable $data
		
		if($ExportToCSVFileOnly)
		{
			$DestinationName = "MPARR - "+$file.BaseName+".csv"
			$FinalDestination = $ExportPath+"\"+$DestinationName
			Copy-Item $file -Destination $FinalDestination -Force
		}elseif($ExportToJSONFileOnly)
		{
			$ExportedName = "MPARR - "+$file.BaseName+".json"
			ExportDataAsCurrentOption -sourcedata $data -ExportOption 2 -exportname $ExportedName
		}elseif($OptionEventHub -eq "True")
		{
			ExportDataAsCurrentOption -sourcedata $data -ExportOption 3
		}else
		{
			ExportDataAsCurrentOption -sourcedata $data -ExportOption 4
		}	
		
		$loopNumber++
		Write-Host "File '$file' was processed with $($csv.count - 1) records."
		Move-Item $file "$($file).processed" -Force
	}
	$compress = @{
		Path = $RMSLogs+"*.processed"
		CompressionLevel = "Fastest"
		DestinationPath = $RMSLogs+"\MPARR - RMS Logs - "+$datefile+".zip"
	}
	Compress-Archive @compress -Force
	$RemoveFiles = $RMSLogs+"*.processed"
	Remove-Item -Path $RemoveFiles -Force

	$RMSContentID = $RMSContentID | Select-Object -Unique
	TrackingRMSDetails -ContentIds $RMSContentID
}

#Run the script.

# Script settings
$CONFIGFILE = "$PSScriptRoot\ConfigFiles\laconfig.json"
$json = Get-Content -Raw -Path $CONFIGFILE
[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
$RMSLogs = $config.RMSLogs
$EncryptedKeys = $config.EncryptedKeys
$AppClientID = $config.AppClientID
$WLA_CustomerID = $config.LA_CustomerID
$WLA_SharedKey = $config.LA_SharedKey
$CertificateThumb = $config.CertificateThumb
$OnmicrosoftTenant = $config.OnmicrosoftURL
$ExportOptionEventHub = $config.ExportToEventHub
$ClientSecretValue = $config.ClientSecretValue
$TenantGUID = $config.TenantGUID

if ($EncryptedKeys -eq "True")
{
	$WLA_SharedKey = DecryptSharedKey $WLA_SharedKey
	$ClientSecretValue = DecryptSharedKey $ClientSecretValue
}

# Your Log Analytics workspace ID
$LogAnalyticsWorkspaceId = $WLA_CustomerID

# Use either the primary or the secondary Connected Sources client authentication key   
$LogAnalyticsPrimaryKey = $WLA_SharedKey 

$OptionEventHub = CheckExportOption

ValidateConfigurationFile

CheckPrerequisites
if($CreateTask)
{
	CreateMPARRRMSDataTask
	exit
}
Export-RMSDatav2