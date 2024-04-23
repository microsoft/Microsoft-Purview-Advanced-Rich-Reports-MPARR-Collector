<#PSScriptInfo

.VERSION 2.0.7

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
Exports Microsoft 365 Audit logs data to Log Analytics or Event Hub. Optionaly data from files can be created. 

#>

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
			"MicrosoftEntraConfig": "Set on 1st configuration",
		    "ExportToEventHub": "False",
		    "EventHubNamespace": "EventHubNamespace",
		    "EventHub": "EventHub"
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

.PARAMETER ExportToCSVFileOnly
    Switch to disable export to Log Analytics.
	
.PARAMETER ExportToJSONFileOnly
    Switch to disable export to Log Analytics.

.PARAMETER ExportWithFile
    Switch to export to Log Analytics creating output files at the same time.

.EXAMPLE
    mparr_collector2.ps1 -FilterAuditSharepoint "Accessed"

    Exports compliance data to LA with filtering enabled for Sharepoint data. Please note that list of the filters depends on the 'schema.json' content.

.EXAMPLE
    "your_secret" | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString

    Encrypts secret string. Resulting string should be pasted to the "laconfig.json" file.
    When enabling secret encryption, both secrets are required be encrypted - "ClientSecretValue" and "LA_SharedKey" from the "laconfig.json" file 
    (replace "your_secret" with these values and put results into the corresponding fields of the config file).
    Value of "EncryptedKeys" must be set to "True".

#> 

<#
.NOTES
HISTORY
Script      : MPARR-Collector2.ps1
Authors     : G.Berdzik / S. Zamorano
Version     : 2.0.7
Purpose		: Collects Logs from Office 365 Management API and send to Logs Analytcs, Event Hub or File

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
  
.NOTES 
	07-02-2024	S. Zamorano		- First release
	07-02-2024	S. Zamorano		- Added EventHub connector
	12-02-2024	S. Zamorano		- New version released
	01-03-2024	S. Zamorano		- Public release
	03-04-2024	Marco van Doorn	- Fix related to an odd behavior with UTC time
	19-04-2024	G.Berdzik		- Get Authentication token fix
#>

#
# UseTimeParameters - if given, the provided start/end times are used instead of calculated times from timestamp file
# pFilenameCode - Code that will used in filename instead of date
#

using module "ConfigFiles\MPARRUtils.psm1"

[CmdletBinding(DefaultParameterSetName = "None")]
param(
    [int]$MPARRBatchSize = 500,
	[Parameter(ParameterSetName="CustomParams")] 
    [Parameter(ParameterSetName="CustomParams1")] 
        [switch]$UseCustomParameters,
    [Parameter(ParameterSetName="CustomParams", Mandatory=$true)] 
        [datetime]$pStartTime,
    [Parameter(ParameterSetName="CustomParams", Mandatory=$true)] 
        [datetime]$pEndTime,
    [Parameter()] 
        [switch]$ExportToCSVFileOnly,
	[Parameter()] 
        [switch]$ExportToJSONFileOnly,
    [Parameter()] 
        [switch]$ExportWithFile,
	[Parameter()] 
        [switch]$CreateTask,
	[Parameter()] 
        [switch]$ExportToEventHub
)

DynamicParam 
{
    # create dynamic parameters based on 'schemas.json' entries set to 'True'
    $filePath = "$PSScriptRoot\ConfigFiles\schemas.json"
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
			Write-Host ".\MPARR_Collector2.ps1 -ExportToCSVFileOnly -ManualConnection" -ForeGroundColor Green
			Write-Host "`n##########################################################################################" -ForeGroundColor Yellow
			Write-Host "`n"
			if($ExportToCSVFileOnly -Or $ExportToJSONFileOnly)
			{
				Write-Host "`n##########################################################################################" -ForeGroundColor Yellow
				Write-Host "`nThe laconfig.json is required to continue" -ForeGroundColor DarkYellow
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
			$ClientSecretValue = $config.ClientSecretValue
			$WLA_CustomerID = $config.LA_CustomerID
			$WLA_SharedKey = $config.LA_SharedKey
			$TenantGUID = $config.TenantGUID
			$OnmicrosoftTenant = $config.OnmicrosoftURL
			$TenantDomain = $config.TenantDomain
			$Cloud = $config.Cloud
			
			if($AppClientID -eq "") { Write-Host "Application Id is missing! Update the laconfig.json file and run again" -ForeGroundColor Red; exit }
			if($WLA_CustomerID -eq "")  { Write-Host "Logs Analytics workspace ID is missing! Update the laconfig.json file and run again" -ForeGroundColor Red; exit }
			if($WLA_SharedKey -eq "")  { Write-Host "Logs Analytics workspace key is missing! Update the laconfig.json file and run again" -ForeGroundColor Red; exit }
			if($ClientSecretValue -eq "")  { Write-Host "Microsoft Entra App Secret is missing! Update the laconfig.json file and run again" -ForeGroundColor Red; exit }
			if($TenantGUID -eq "")  { Write-Host "Tenant ID is missing! Update the laconfig.json file and run again" -ForeGroundColor Red; exit }
			if($TenantDomain -eq "")  { Write-Host "Main Tenant domain is missing! Update the laconfig.json file and run again" -ForeGroundColor Red; exit }
			if($Cloud -eq "")  { Write-Host "Tenant cloud is missing! Update the laconfig.json file and run again" -ForeGroundColor Red; exit }
			
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
		# function to get option number
		$selection = 0
		do 
		{
			$resp = Read-Host $msg
			try {
				$selection = [int]$resp
				if (($selection -gt $max) -or ($selection -lt 15))
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
		$loginURL = "https://login.microsoftonline.com/"

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

	function CreateMPARRCollectorTask
	{
		# MPARR-ContentExplorerData script
		$taskName = "MPARR-DataCollector2"
		
		# Call function to set a folder for the task on Task Scheduler
		$taskFolder = CreateScheduledTaskFolder
		
		<#
		.NOTES
		This function create both task,MPARR_Collector, to run every 30 minutes, that time can be changed on the same task scheduler, is not recommended less time.
		MPARR_Collector use PowerShell 7
		#>
		Write-Host "`n`n----------------------------------------------------------------------------------------" -ForegroundColor Yellow
		Write-Host "`nPlease be aware that the scripts MPARR_Collector is set to execute every 30 minutes" -ForegroundColor DarkYellow
		Write-Host "You can change directly on task scheduler and change the execution period" -ForegroundColor DarkYellow
		Write-Host "Depend on your logs volume cannot be recommend use less time," -ForegroundColor DarkYellow
		Write-Host "to give time to the scripts to be execute correctly." -ForegroundColor DarkYellow
		Write-Host "`n----------------------------------------------------------------------------------------" -ForegroundColor Yellow
		Write-Host "`n`n"
		
		# Task execution
		$validMinutes = 30
		$choices  = '&Yes', '&No'
		$decision = $Host.UI.PromptForChoice("", "The task on task scheduler will be set for $validMinutes minutes, do you want to change?", $choices, 1)
		if ($decision -eq 0)
		{
			ReadNumber -max 120 -msg "Enter number of days (Between 15 to 120). Remember check the retention period in your workspace in Logs Analtytics." -option ([ref]$validDays)
		}

		# calculate date
		# calculate date
		$dt = Get-Date
		$nearestMinutes = $validMinutes
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

	function CheckOutputDirectory($OutputPath)
	{
		### Verify output directory exists
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
				New-Item -ItemType "directory" -Path $OutputPath -Force -ErrorAction Stop | Out-Null
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

	function buildLog($BaseURI, $Subscription, $tenantGUID, $OfficeToken)
	{
		# Create Function to Check content availability in all content types (inlcuding all pages) 
		# and store results in $Subscription variable, also build the URI list in the correct format
		
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
				$end = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ss") 
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

	function FetchData($TotalContentPages, $Officetoken, $Subscription)
	{
		##Generate the correct URI format and export  logs
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
		if($body -isnot [array]) {Write-host "1A"; return}
		if($body.Count -eq 0) {return}

		#Step 1: convert the PSObject to JSON
		$bodyJson = $body | ConvertTo-Json -Depth 100

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
		$response = Invoke-WebRequest -Uri $uri -Method $method -Headers $headers -ContentType $contentType -Body $bodyJsonUTF8 -UseBasicParsing

		if ($Response.StatusCode -eq 200) {   
			$rows = $body.Count
			Write-Information -MessageData "$rows rows written to Log Analytics workspace $uri" -InformationAction Continue
		}

	}
	
	function Publish-LogAnalytics
	{
		param (
			$objFromJson,
			$Subscription
		)
		
		#Read configuration file
		$CONFIGFILE = "$PSScriptRoot\ConfigFiles\laconfig.json"
		$json = Get-Content -Raw -Path $CONFIGFILE
		[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
		$OutputPath = $config.OutPutLogs
		
		$EncryptedKeys = $config.EncryptedKeys
		$CustomerID = $config.LA_CustomerID
		$SharedKey = $config.LA_SharedKey
		if ($EncryptedKeys -eq "True")
		{
			$SharedKey = DecryptSharedKey $SharedKey
		}

		Write-Host "Starting export to LA..." -NoNewLine
		$list = New-Object System.Collections.ArrayList
		$LogName = $Subscription.Replace(".", "")
		$BatchSize = $MPARRBatchSize

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
				$log_analytics_array = @()            
				foreach($i in $list) {
					$log_analytics_array += $i
				}
				Post-LogAnalyticsData -body $log_analytics_array -LogAnalyticsTableName $LogName 
				$list.Clear()
				$list.TrimToSize()            
			}
		}
		if ($list.Count -gt 0)
		{
			$log_analytics_array = @()            
				foreach($i in $list) {
					$log_analytics_array += $i
				}
			Post-LogAnalyticsData -body $log_analytics_array -LogAnalyticsTableName $LogName
		}
		Write-Host "$count elements exported for $Subscription."
	}

	function Export-Logs($Subscriptions)
	{
		Write-Verbose " enter export-logs" 
		# Script variables  ---> Don't Update anything here:
		#$loginURL = "https://login.microsoftonline.com/"
		$BaseURI = "$APIResource/api/v1.0/$TenantGUID/activity/feed/subscriptions"
		
		#Folders and files needed
		$Date = (Get-date).AddDays(-1)
		$Date = $Date.ToString('MM-dd-yyyy_hh-mm-ss')
		$ExportPath = $PSScriptRoot+"\ExportedData"
		if(-Not (Test-Path $ExportPath ))
		{
			Write-Host "Export data directory is missing, creating a new folder called ExportedData"
			New-Item -ItemType Directory -Force -Path "$PSScriptRoot\ExportedData" | Out-Null
		}

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
			Write-Host "`n`n"
			Write-Host "#######################################################################"
			Write-Host -ForegroundColor Cyan "-> Collecting log data from '" -NoNewline
			Write-Host -ForegroundColor White -BackgroundColor DarkGray $Subscription -NoNewline
			Write-Host -ForegroundColor Cyan "': " -NoNewline

			# check for token expiration
			if ($tokenExpiresOn.AddMinutes(5) -lt (Get-Date))
			{
				Write-Host "Refreshing access token..."
				GetAuthToken
			}

			$logs = buildLog $BaseURI $Subscription $TenantGUID $OfficeToken

			$JSONfilename = ("MPARR-"+$Subscription+"-"+$Date+".json")
			$ExportJson = "$ExportPath\$JSONfilename"
			$CSVfilename = ("MPARR-"+$Subscription+"-"+$Date+".csv")
			$ExportCsv = "$ExportPath\$CSVfilename"
		
			$output = FetchData $logs $OfficeToken $Subscription
			Write-Host "Total amount of records returned : "$output.count
			if ($ExportToJSONFileOnly)
			{
				if($output.count -eq 0)
				{
					Write-Host "No data returned from : "$Subscription -ForegroundColor DarkYellow
					Write-Host "No export file was created." -ForegroundColor DarkYellow
				}else
				{
					$output | ConvertTo-Json -Depth 100 | Set-Content -Encoding UTF8 $ExportJson
					Write-host -ForegroundColor Cyan "---> Exporting log data to '" -NoNewline
					Write-Host -ForegroundColor White -BackgroundColor DarkGray $JSONfilename
				}
			}elseif ($ExportToCSVFileOnly)
			{
				if($output.count -eq 0)
				{
					Write-Host "No data returned from : "$Subscription -ForegroundColor DarkYellow
					Write-Host "No export file was created." -ForegroundColor DarkYellow
				}else
				{
					$output | Export-CSV $ExportCsv -Append
					Write-host -ForegroundColor Cyan "---> Exporting log data to '" -NoNewline
					Write-Host -ForegroundColor White -BackgroundColor DarkGray $CSVfilename
				}
			}elseif ($ExportWithFile)
			{
				if($output.count -eq 0)
				{
					Write-Host "No data returned from : "$Subscription -ForegroundColor DarkYellow
					Write-Host "No data was exported to Logs Analytics and no export file was created." -ForegroundColor DarkYellow
				}else
				{
					$output | ConvertTo-Json -Depth 100 | Set-Content -Encoding UTF8 $ExportJson
					Write-host -ForegroundColor Cyan "---> Exporting log data to '" -NoNewline
					Write-Host -ForegroundColor White -BackgroundColor DarkGray $JSONfilename
					Write-Host -ForegroundColor Cyan "': " -NoNewline
					Publish-LogAnalytics $output $Subscription
				}
			}elseif ($ExportToEventHub)
			{
				if($output.count -eq 0)
				{
					Write-Host "No data returned from : "$Subscription -ForegroundColor DarkYellow
					Write-Host "No data was exported to Event Hub." -ForegroundColor DarkYellow
				}else
				{
					EventHubConnection
					$ErrorFile = "MPARRCollector-"+$Subscription+"-Error-"+$Date+".json"
					$EventHubInstance.PublishToEventHub($output, $ErrorFile)
				}
			}else 
			{
				if($output.count -eq 0)
				{
					Write-Host "No data returned from : "$Subscription -ForegroundColor DarkYellow
					Write-Host "No data was exported to Logs Analytics." -ForegroundColor DarkYellow
				}else
				{
					Publish-LogAnalytics $output $Subscription
				}
			}
		}
	}

	function MainCollector
	{
		#region Main code
		# Script variables 01  --> Update everything in this section:
		$CONFIGFILE = "$PSScriptRoot\ConfigFiles\laconfig.json"   
		$SCHEMASFILE = "$PSScriptRoot\ConfigFiles\schemas.json" 
		
		# Read config file
		if (-not (Test-Path -Path $CONFIGFILE))
		{
			Write-Error "Missing config file."
			exit(1)
		}
		
		# Load laconfig.json into variables
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
		$OutputPath = $config.OutPutLogs
		
		#API Endpoint URLs ---> Don't Update anything here
		$CLOUDVERSIONS = @{
			Commercial = "https://manage.office.com"
			GCC = "https://manage-gcc.office.com"
			GCCH = "https://manage.office365.us"
			DOD = "https://manage.protection.apps.mil"
		}
		
		if ($EncryptedKeys -eq "True")
		{
			$ClientSecretValue = DecryptSharedKey $ClientSecretValue
			$SharedKey = DecryptSharedKey $SharedKey
		}
		
		$APIResource = $CLOUDVERSIONS.Commercial
		if ($Cloud -ne $null)
		{
			$APIResource = $CLOUDVERSIONS["$Cloud"]
			Write-Host "Connecting to $Cloud cloud."
		}

		if ($OutputPath -eq "")
		{
			$OutputPath = "$PSScriptRoot\Logs\"
			Write-Host "'OutputLogs' has no value. Default value was assigned: $OutputPath." -ForegroundColor Yellow
			Write-Host "Logs directory is missing, creating a new folder called Logs"
			New-Item -ItemType Directory -Force -Path "$PSScriptRoot\ExportedData" | Out-Null
			
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

		#region Timestamp/1
		$timestampFile = "$OutputPath"+"timestamp.json"
		
		# read startTime from the file
		if (-not (Test-Path -Path $timestampFile))
		{
			# if file not present create new value
			Write-Host "Time stamp not present, a new one will be created..." -ForeGroundColor DarkYellow
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
		#$endTime = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
		$endTime = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ss") 
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

		Export-Logs -Subscriptions $Subscriptions
	}

	CheckPrerequisites
	if($CreateTask)
	{
		CreateMPARRCollectorTask
		exit
	}
	MainCollector
#endregion
}
