<#
.SYNOPSIS
    Script to setup EDM.

.DESCRIPTION
    Script is designed to simplify EDM configuration as a task.
    
.NOTES
    Version 0.9
    Current version - 27.10.2023
#> 

<#
HISTORY
  2023-09-06    G.Berdzik 	- Initial version (used MPARR_Setup script as a base)

  2023-10-27	S.Zamorano	- New version using the original script as a base for EDM
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

function CheckPowerShellVersion
{
    # Check PowerShell version
    Write-Host "`nChecking PowerShell version... " -NoNewline
    if ($Host.Version.Major -gt 5)
    {
        Write-Host "Passed" -ForegroundColor Green
    }
    else
    {
        Write-Host "Failed" -ForegroundColor Red
        Write-Host "`tCurrent version is $($Host.Version). PowerShell version 7 or newer is required."
        exit(1)
    }
}

function CheckPrerequisites
{
    CheckIfElevated
    CheckPowerShellVersion
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

function InitializeHostName
{
	$config = "$PSScriptRoot\EDMConfig.json"
	
	if (-not (Test-Path -Path $config))
    {
		Write-Host "Working on remote host." -ForegroundColor Red
		return
	}elseif ($EDMHostName -eq "Localhost")
	{
		$json = Get-Content -Raw -Path $config
		[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
		$EDMHostName = $config.EDMHostName
		$EDMHostName = hostname
		$config.EDMHostName = $EDMHostName
		WriteToJsonFile
	}
}

function TakeAPause
{
	$choices  = '&Continue'
	$decision = $Host.UI.PromptForChoice("", "`nDo you want to Continue? If you see an error above, validate your credentials.", $choices, 0)
	if ($decision -eq 0)
    {
		return
	}
}

# function to decrypt shared key
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

function Connect2EDM
{
	$CONFIGFILE = "$PSScriptRoot\EDMConfig.json"
	if (-not (Test-Path -Path $CONFIGFILE))
	{
		$CONFIGFILE = "$PSScriptRoot\EDM_RemoteConfig.json"
	}
	
	$json = Get-Content -Raw -Path $CONFIGFILE
	[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
	$EncryptedKeys = $config.EncryptedKeys
	$EDMFolder = $config.EDMAppFolder
	$user = $config.User
	$SharedKey = $config.Password
	
	if ($EncryptedKeys -eq "True")
	{
		$SharedKey = DecryptSharedKey $SharedKey
		cd $EDMFolder | cmd
		cls
		Write-Host "Validating connection to EDM..." -ForegroundColor Green
		.\EdmUploadAgent.exe /Authorize /Username $user /Password $SharedKey 
	}else{
		cd $EDMFolder | cmd
		cls
		Write-Host "Validating connection to EDM..." -ForegroundColor Green
		.\EdmUploadAgent.exe /Authorize /Username $user /Password $SharedKey
	}
}

# function to identify the paths used by EDM
function SelectEDMPaths
{
    $choices  = '&Yes', '&No'
	Write-Host "`n`n##########################################"
	Write-Host "`nThe current configuration for EDM is:"
	Write-Host "* EDM appplication location '$($config.EDMAppFolder)'."
	Write-Host "* EDM root folder '$($config.EDMrootFolder)'."
	Write-Host "* Hash folder location '$($config.HashFolder)'."
	Write-Host "* Schema data folder '$($config.SchemaFolder)'."
    $decision = $Host.UI.PromptForChoice("", "`nDo you want change the locations?", $choices, 1)
    if ($decision -eq 0)
    {
        [System.Reflection.Assembly]::Load("System.Windows.Forms") | Out-Null
        $folder = New-Object System.Windows.Forms.FolderBrowserDialog
		$folder.UseDescriptionForTitle = $true
        
		#Here you start selecting each folder
		
		# Start selecting first EDM App location
		$folder.Description = "Select folder where EdmUploadAgent.exe is located"
        $folder.rootFolder = 'ProgramFiles'
        # main log directory
        if ($folder.ShowDialog() -eq "OK")
        {
            $config.EDMAppFolder = $folder.SelectedPath + "\"
            Write-Host "`nEDM App folder set to '$($config.EDMAppFolder)'."
        }

        # EDM data root folder
        $folder.Description = "Select where the data for EDM will be located"
        $folder.rootFolder = 'ProgramFiles'
        if ($folder.ShowDialog() -eq "OK")
        {
            $config.EDMrootFolder = $folder.SelectedPath + "\"
            Write-Host "`nData root folder set to '$($config.EDMrootFolder)'."
        }
		
		# Hash data folder
        $folder.Description = "Select folder where Hash data will be located"
        $folder.rootFolder = 'Recent'
        $folder.InitialDirectory = $config.EDMrootFolder
        if ($folder.ShowDialog() -eq "OK")
        {
            $config.HashFolder = $folder.SelectedPath + "\"
            Write-Host "`nHash folder set to '$($config.HashFolder)'."
        }
		
		# Schema data folder
        $folder.Description = "Select folder where Schema data will be located"
        $folder.rootFolder = 'Recent'
        $folder.InitialDirectory = $config.EDMrootFolder
        if ($folder.ShowDialog() -eq "OK")
        {
            $config.SchemaFolder = $folder.SelectedPath + "\"
            Write-Host "`nData folder set to '$($config.SchemaFolder)'."
        }
    }
}

function SelectEDMRemotePaths
{
	$choices  = '&Yes', '&No'
	Write-Host "`n`n##########################################"
	Write-Host "`nThe current configuration for EDM Remote activities is:"
	Write-Host "* EDM appplication location '$($RemoteConfig.EDMAppFolder)'."
	Write-Host "* EDM root folder '$($RemoteConfig.EDMrootFolder)'."
	Write-Host "* Hash folder location '$($RemoteConfig.HashFolder)'."
    $decision = $Host.UI.PromptForChoice("", "`nDo you want change the locations?", $choices, 0)
    if ($decision -eq 0)
    {
        [System.Reflection.Assembly]::Load("System.Windows.Forms") | Out-Null
        $folder = New-Object System.Windows.Forms.FolderBrowserDialog
		$folder.UseDescriptionForTitle = $true
        
		#Here you start selecting each folder
		
		# Start selecting first EDM App location
		$folder.Description = "Select folder where EdmUploadAgent.exe is located"
        $folder.rootFolder = 'ProgramFiles'
        # main log directory
        if ($folder.ShowDialog() -eq "OK")
        {
            $RemoteConfig.EDMAppFolder = $folder.SelectedPath + "\"
            Write-Host "`nEDM App folder set to '$($RemoteConfig.EDMAppFolder)'."
        }

        # EDM data root folder
        $folder.Description = "Select the root folder used by EDM scripts"
        $folder.rootFolder = 'ProgramFiles'
        if ($folder.ShowDialog() -eq "OK")
        {
            $RemoteConfig.EDMrootFolder = $folder.SelectedPath + "\"
            Write-Host "`nData root folder set to '$($RemoteConfig.EDMrootFolder)'."
        }
		
		# Hash data folder
        $folder.Description = "Select folder where Hash data will be located"
        $folder.rootFolder = 'Recent'
        $folder.InitialDirectory = $RemoteConfig.EDMrootFolder
        if ($folder.ShowDialog() -eq "OK")
        {
            $RemoteConfig.HashFolder = $folder.SelectedPath + "\"
            Write-Host "`nHash folder set to '$($RemoteConfig.HashFolder)'."
        }
    }
}

function GetEDMUserCredentials
{
	$Credential = $host.ui.PromptForCredential("Your credentials are needed", "Please validate that your user is part of EDM_DataUploaders group", "", "")
	$Ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToCoTaskMemUnicode($Credential.Password)
	$config.Password = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($Ptr)
	$config.User = $Credential.Username
	$config.EncryptedKeys = "False"
}

function GetEDMRemoteUserCredentials
{
	$Credential = $host.ui.PromptForCredential("Your credentials are needed", "Please validate that your user is part of EDM_DataUploaders group", "", "")
	$Ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToCoTaskMemUnicode($Credential.Password)
	$RemoteConfig.Password = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($Ptr)
	$RemoteConfig.User = $Credential.Username
	$RemoteConfig.EncryptedKeys = "False"
}

function GetDataStores
{
	Connect2EDM | Out-Null
	
	$config = "$PSScriptRoot\EDMConfig.json"
	$json = Get-Content -Raw -Path $config
	[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
	
	$EDMFolder = $config.EDMAppFolder
	cd $EDMFolder | cmd
	
	$DataStores = .\EdmUploadAgent.exe /GetDataStore
	$DS = $DataStores | Where-Object { $_ –ne $DataStores[0] }
	$DS = $DS | Where-Object { $_ –ne $DS[0] }
	$DS = $DS | Where-Object { $_ –ne $DS[-1] }
	$tempFolder = $DS -replace '(\, ).*','$1'
	$tempFolder = $tempFolder -replace ', ',''
	
	foreach ($DStore in $tempFolder){$DataStoresEDM += @([pscustomobject]@{Name=$DStore})}
	
	Write-Host "`nGetting Data Stores..." -ForegroundColor Green
    $i = 1
    $DataStoresEDM = @($DataStoresEDM | ForEach-Object {$_ | Add-Member -Name "No" -MemberType NoteProperty -Value ($i++) -PassThru})
	
	#List all existing folders under Task Scheduler
    $DataStoresEDM | Select-Object No, Name | Out-Host
	
	# Select EDM datastore tasks
	Write-Host "In case the EDM Schema was recently created and is not listed, please stop the script with Ctrl+C and run it again.`n"
    $selection = 0
    ReadNumber -max ($i -1) -msg "Enter number corresponding to the DataStore name" -option ([ref]$selection)
    $config.DataStoreName = $DataStoresEDM[$selection - 1].Name
	
	Write-Host "`nData Store selected '$($config.DataStoreName)'" -ForegroundColor Green

	WriteToJsonFile
}

function GetSchemaFile
{
	Connect2EDM | Out-Null
	
	$configfile = "$PSScriptRoot\EDMConfig.json"
	$json = Get-Content -Raw -Path $configfile
	[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
	$EDMDSName = $config.DataStoreName
	$SchemaFolder = $config.SchemaFolder
	
	$EDMFolder = $config.EDMAppFolder
	cd $EDMFolder | cmd	
	
	if ($EDMDSName -eq "Not set")
	{
		Write-Host "Missing DataStore name." -ForegroundColor Red
		$choices  = '&Yes', '&No'
		$decision = $Host.UI.PromptForChoice("", "`nDo you want to select the data store?", $choices, 0)
		if ($decision -eq 0)
		{
			GetDataStores
			WriteToJsonFile
		}return
	}else{
		.\EdmUploadAgent.exe /SaveSchema /DataStoreName $config.DataStoreName /OutputDir $SchemaFolder
		$XMLfile = gci $SchemaFolder | select -last 1
		$config.SchemaFile = $XMLfile.Name
		WriteToJsonFile
	}
}

function ValidateEDMData
{
	Connect2EDM | Out-Null
	$configfile = "$PSScriptRoot\EDMConfig.json"
	$json = Get-Content -Raw -Path $configfile
	[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
	
	$EDMFolder = $config.EDMAppFolder
	cd $EDMFolder | cmd
	
	$choices  = '&Yes', '&No'
	Write-Host "`n`n##########################################"
	Write-Host "`nThe current data files for EDM is:"
	Write-Host "* '$($config.DataFile)'" -ForegroundColor Green
	
	$decision = $Host.UI.PromptForChoice("", "`nDo you want to change the file set?", $choices, 0)
	if ($decision -eq 0)
    {
        [System.Reflection.Assembly]::Load("System.Windows.Forms") | Out-Null
        $file = New-Object System.Windows.Forms.OpenFileDialog
		# Start selecting Data file location used for EDM 
		$file.Title = "Select folder where Data file is located"
        $file.InitialDirectory = 'MyComputer'
		$file.Filter = 'CSV format |*.CSV|TSV format|*.TSV'
        # main log directory
        if ($file.ShowDialog() -eq "OK")
        {
            $config.DataFile = $file.FileName
			$EDMDataPath = $file.FileName
			$config.EDMDataFolder = (Get-Item $EDMDataPath).DirectoryName+"\"
			WriteToJsonFile
            Write-Host "`nOutput logs set to '$($config.DataFile)'."
        }
	}
	
	$SchemaLocation = $config.SchemaFolder+$config.SchemaFile
	Write-Host "`nSchema location is '$($SchemaLocation)'." -ForegroundColor Green
	.\EdmUploadAgent.exe /ValidateData /DataFile $config.DataFile /Schema $SchemaLocation
}

function EDMHashCreation
{
	Connect2EDM | Out-Null
	$configfile = "$PSScriptRoot\EDMConfig.json"
	$json = Get-Content -Raw -Path $configfile
	[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
	$HashFolder = $config.HashFolder
	
	$EDMFolder = $config.EDMAppFolder
	cd $EDMFolder | cmd
	
	$EDMData = $config.DataFile
	$EDMHash = $config.HashFolder
	$EDMSchema = $config.SchemaFolder+$config.SchemaFile
	
	
	.\EdmUploadAgent.exe /CreateHash /DataFile $EDMData /HashLocation $EDMHash /Schema $EDMSchema  /AllowedBadLinesPercentage 5
	$Hashfile = gci $HashFolder -Filter *.edmhash | select -last 1
	$config.HashFile = $Hashfile.Name
	WriteToJsonFile
	
	Write-Host "`nHash and Salt files created at:" -ForegroundColor Green
	Write-Host "* Schema file '$($config.HashFolder)'." -ForegroundColor Green
}

function EDMHashUpload
{
	Connect2EDM | Out-Null
	
	$configfile = "$PSScriptRoot\EDMConfig.json"
	If (-not (Test-Path -Path $configfile))
	{
		$configfile = "$PSScriptRoot\EDM_RemoteConfig.json"
	}
	$json = Get-Content -Raw -Path $configfile
	[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
	
	$EDMDSName = $config.DataStoreName
	$HashName = $config.HashFolder+$config.HashFile
	
	$EDMFolder = $config.EDMAppFolder
	cd $EDMFolder | cmd	
	
	.\EdmUploadAgent.exe /UploadHash /DataStoreName $EDMDSName /HashFile $HashName
	Write-Host "`nHash is uploading, you can validate the state in the -EDM Hash Upload Status- menu" -ForegroundColor Green
	Write-Host "`nREMEMBER: You can update your EDM data only 5 times per day." -ForegroundColor RED
}

function EDMUploadStatus
{
	Connect2EDM | Out-Null
	
	$configfile = "$PSScriptRoot\EDMConfig.json"
	If (-not (Test-Path -Path $configfile))
	{
		$configfile = "$PSScriptRoot\EDM_RemoteConfig.json"
	}
	$json = Get-Content -Raw -Path $configfile
	[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
	$DSName = $config.DataStoreName
	
	$EDMFolder = $config.EDMAppFolder
	cd $EDMFolder | cmd
	
	cls
	Write-Host "`nChecking the Hash upload status" -ForegroundColor Green
	.\EdmUploadAgent.exe /GetSession /DataStoreName employeesdataschema
}

function EDMCopyDataNeeded
{
	cls
	CreateRemoteConfigFile
	$configfile = "$PSScriptRoot\EDMConfig.json"
	$json = Get-Content -Raw -Path $configfile
	[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
	$HashData = $config.HashFolder
	$HashData = $HashData.Substring(0,$HashData.Length-1)
	$HashData = $HashData+"*"
	$SupportScripts = $config.EDMSupportFolder
	$SupportScripts = $SupportScripts+"EDM_*"
	$Destination = $config.EDMremoteFolder

	$EDMScripts = "$PSScriptRoot\EDM_*"
	
	#Here we ned to select the destination folder
	$choices  = '&Yes', '&No'
	Write-Host "`n`n##########################################"
	Write-Host "`nThe current configuration for remote folder for hash EDM is:"
	Write-Host "EDM remote path '$($config.EDMremoteFolder)'."
	Write-Host "REMEMBER: This first copy needs to be done in an empty folder." -ForegroundColor Red
	Write-Host "`n##########################################"

    $decision = $Host.UI.PromptForChoice("", "`nDo you want change the locations?", $choices, 1)
    if ($decision -eq 0)
    {
        [System.Reflection.Assembly]::Load("System.Windows.Forms") | Out-Null
        $folder = New-Object System.Windows.Forms.FolderBrowserDialog
		$folder.UseDescriptionForTitle = $true
        
		#Here you start selecting each folder
		# Start selecting first EDM remote location
		$folder.Description = "Select folder where Hash data will be copied"
        $folder.rootFolder = 'ProgramFiles'
        # main log directory
        if ($folder.ShowDialog() -eq "OK")
        {
            $config.EDMremoteFolder = $folder.SelectedPath + "\"
			$Destination = $config.EDMremoteFolder
            Write-Host "`nEDM App folder set to '$($config.EDMremoteFolder)'."
			WriteToJsonFile
        }
	}
	
	Write-Host "`n###################################################" -ForegroundColor Red
	Write-Host "These files will be copy to '$($Destination)'." -ForegroundColor Green
	Write-Host "`n`tHash y Salt files located at '$($config.HashFolder)' " -ForegroundColor Green
	Write-Host "`tEDM_RemoteConfig.json file (Password was decrypted)  " -ForegroundColor Green
	Write-Host "`tThis EDM_Setup file " -ForegroundColor Green
	Write-Host "`tSupport script for upload task " -ForegroundColor Green
	Write-Host "`n###################################################" -ForegroundColor Red
	
	Copy-Item $HashData $Destination -recurse -force
	Copy-Item $EDMScripts $Destination -recurse -force
	Copy-Item $SupportScripts $Destination -recurse -force	
}

function InitializeEDMConfigFile
{
	# read config file
    $configfile = "$PSScriptRoot\EDMConfig.json" 
	
	if (-not (Test-Path -Path $configfile))
    {
		$config = [ordered]@{
		EncryptedKeys =  "False"
		SchemaFile = "Not set"
		Password = ""
		User = ""
		HashFile = "Not set"
		DataFile = "Not set"
		BadLinesPercentage = "5"
		DataStoreName = "Not set"
		EDMAppFolder = "c:\Program Files\Microsoft\EdmUploadAgent\"
		EDMrootFolder = "C:\EDM data\"
		HashFolder = "C:\EDM data\Hash\"
		SchemaFolder = "C:\EDM data\Schemas\"
		EDMremoteFolder = "\\localhost\c$\"
		EDMSupportFolder = "C:\EDM data\Support\"
		EDMDataFolder = "C:\EDM data\Data\"
		EDMHostName = "Localhost"
		}
		return $config
    }else
	{
		$json = Get-Content -Raw -Path $configfile
		[PSCustomObject]$configfile = ConvertFrom-Json -InputObject $json
	
		$config = [ordered]@{
		EncryptedKeys = "$($configfile.EncryptedKeys)"
		SchemaFile = "$($configfile.SchemaFile)"
		Password = "$($configfile.Password)"
		User = "$($configfile.User)"
		HashFile = "$($configfile.HashFile)"
		DataFile = "$($configfile.DataFile)"
		BadLinesPercentage = "$($configfile.BadLinesPercentage)"
		DataStoreName = "$($configfile.DataStoreName)"
		EDMAppFolder = "$($configfile.EDMAppFolder)"
		EDMrootFolder = "$($configfile.EDMrootFolder)"
		HashFolder = "$($configfile.HashFolder)"
		SchemaFolder = "$($configfile.SchemaFolder)"
		EDMremoteFolder = "$($configfile.EDMremoteFolder)"
		EDMSupportFolder = "$($configfile.EDMSupportFolder)"
		EDMDataFolder = "$($configfile.EDMDataFolder)"
		EDMHostName = "$($configfile.EDMHostName)"
		}
		return $config
	}
}

function InitializeEDMRemoteConfigFile
{
	# read config file
    $configfile = "$PSScriptRoot\EDMConfig.json" 
	
	if (-not (Test-Path -Path $configfile))
    {
		$config = [ordered]@{
		EncryptedKeys =  "False"
		Password = ""
		User = ""
		HashFile = "Not set"
		DataStoreName = "Not set"
		EDMAppFolder = "c:\Program Files\Microsoft\EdmUploadAgent\"
		EDMrootFolder = "C:\EDM data\"
		HashFolder = "C:\EDM data\Hash\"
		EDMHostName = "Localhost"
		}
		return $config
    }else
	{
		$json = Get-Content -Raw -Path $configfile
		[PSCustomObject]$configfile = ConvertFrom-Json -InputObject $json
		$EncryptedKeys = $configfile.EncryptedKeys
		$SharedKey = $configfile.Password
	
		if ($EncryptedKeys -eq "True")
		{
			$SharedKey = DecryptSharedKey $SharedKey 
		}
	
		$config = [ordered]@{
		EncryptedKeys = "$($configfile.EncryptedKeys)"
		Password = "$($SharedKey)"
		User = "$($configfile.User)"
		HashFile = "$($configfile.HashFile)"
		DataStoreName = "$($configfile.DataStoreName)"
		EDMAppFolder = "$($configfile.EDMAppFolder)"
		EDMrootFolder = "$($configfile.EDMrootFolder)"
		HashFolder = "$($configfile.HashFolder)"
		EDMHostName = "$($configfile.EDMHostName)"
		}
		return $config
	}
}

function WriteToRemoteJsonFile
{
	if (Test-Path "$PSScriptRoot\EDM_RemoteConfig.json")
    {
        $date = Get-Date -Format "yyyyMMddHHmmss"
        Move-Item "$PSScriptRoot\EDM_RemoteConfig.json" "$PSScriptRoot\bck_EDM_RemoteConfig_$date.json"
        Write-Host "`nThe old config file moved to 'bck_EDM_RemoteConfig_$date.json'"
    }
	$RemoteConfig | ConvertTo-Json | Out-File "$PSScriptRoot\EDM_RemoteConfig.json"
    Write-Host "Setup completed. New config file was created." -ForegroundColor Green
}

function CreateRemoteConfigFile
{
	$RemoteConfig = InitializeEDMRemoteConfigFile
	WriteToRemoteJsonFile
}

# write configuration data to json file
function WriteToJsonFile
{
	if (Test-Path "$PSScriptRoot\EDMConfig.json")
    {
        $date = Get-Date -Format "yyyyMMddHHmmss"
        Move-Item "$PSScriptRoot\EDMConfig.json" "$PSScriptRoot\bck_EDMConfig_$date.json"
        Write-Host "`nThe old config file moved to 'bck_EDMConfig_$date.json'"
    }
    $config | ConvertTo-Json | Out-File "$PSScriptRoot\EDMConfig.json"
    Write-Host "Setup completed. New config file was created." -ForegroundColor Green
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
	
	# Default folder for EDM tasks
    $EDMTSFolder = "EDM"
	$taskFolder = "\"+$EDMTSFolder+"\"
	$choices  = '&Proceed', '&Change', '&Existing'
	Write-Host "Please consider if you want to use the default location you need select Existing and the option 1." -ForegroundColor Yellow
    $decision = $Host.UI.PromptForChoice("", "Default task Scheduler Folder is '$EDMTSFolder'. Do you want to Proceed, Change the name or use Existing one?", $choices, 0)
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
		Write-Host "Using the default folder $EDMTSFolder." -ForegroundColor Green
		return $taskFolder
	}else
	{
		$selection = 0
		ReadNumber -max ($i -1) -msg "Enter number corresponding to the current folder in the Task Scheduler" -option ([ref]$selection) 
		$value = $selection - 1
		$EDMTSFolder = $SchedulerTaskFolders[$value].Name
		$taskFolder = "\"+$SchedulerTaskFolders[$value].Name+"\"
		Write-Host "Folder selected for this task $EDMTSFolder " -ForegroundColor Green
		return $taskFolder
	}
	
}

function CreateEDMHashUploadScheduledTask
{
	# EDM task script
    $taskName = "EDM-HashUpload"
	
	# Call function to set a folder for the task on Task Scheduler
	$taskFolder = CreateScheduledTaskFolder
	
	$config = "$PSScriptRoot\EDMConfig.json"
	$json = Get-Content -Raw -Path $config
	[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
	$EDMSupportFolder = $config.EDMSupportFolder
	
	# Task execution
    $validDays = 1
    $choices  = '&Yes', '&No'
    $decision = $Host.UI.PromptForChoice("", "The task on task scheduler will be set for '$($validDays)' day(s), do you want to change?", $choices, 1)
	Write-Host "`nYou can change later in the task '$($taskName)' under Task Scheduler`n" -ForegroundColor Yellow
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
    $action = New-ScheduledTaskAction -Execute "`"$PSHOME\pwsh.exe`"" -Argument ".\EDMHashUpload.ps1" -WorkingDirectory $EDMSupportFolder
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

function CreateEDMRemoteHashUploadScheduledTask
{
	# EDM remote task script
    $taskName = "EDM-RemoteHashUpload"
	
	# Call function to set a folder for the task on Task Scheduler
	$taskFolder = CreateScheduledTaskFolder
	
	$config = "$PSScriptRoot\EDM_RemoteConfig.json"
	$json = Get-Content -Raw -Path $config
	[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
	
	# Task execution
    $validDays = 1
    $choices  = '&Yes', '&No'
    $decision = $Host.UI.PromptForChoice("", "The task on task scheduler will be set for '$($validDays)' day(s), do you want to change?", $choices, 1)
	Write-Host "`nYou can change later in the task '$($taskName)' under Task Scheduler`n" -ForegroundColor Yellow
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
    $action = New-ScheduledTaskAction -Execute "`"$PSHOME\pwsh.exe`"" -Argument ".\EDM_RemoteHashUpload.ps1" -WorkingDirectory $EDMSupportFolder
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

function CreateEDMHashCreateScheduledTask
{
	# MPARR-AzureADUsers script
    $taskName = "EDM-CreateHash"
	
	# Call function to set a folder for the task on Task Scheduler
	$taskFolder = CreateScheduledTaskFolder
	
	$config = "$PSScriptRoot\EDMConfig.json"
	$json = Get-Content -Raw -Path $config
	[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
	$EDMSupportFolder = $config.EDMSupportFolder
	
	# Task execution
    $validDays = 1
    $choices  = '&Yes', '&No'
    $decision = $Host.UI.PromptForChoice("", "The task on task scheduler will be set for '$($validDays)' day(s), do you want to change?", $choices, 1)
	Write-Host "`nYou can change later in the task '$($taskName)' under Task Scheduler`n" -ForegroundColor Yellow
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
    $action = New-ScheduledTaskAction -Execute "`"$PSHOME\pwsh.exe`"" -Argument ".\EDMCreateHash.ps1" -WorkingDirectory $EDMSupportFolder
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

function CreateEDMHashCopyScheduledTask
{
	# MPARR-AzureADUsers script
    $taskName = "EDM-Hash"
	
	# Call function to set a folder for the task on Task Scheduler
	$taskFolder = CreateScheduledTaskFolder
	
	$config = "$PSScriptRoot\EDMConfig.json"
	$json = Get-Content -Raw -Path $config
	[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
	$EDMSupportFolder = $config.EDMSupportFolder
	
	# Task execution
    $validDays = 1
    $choices  = '&Yes', '&No'
    $decision = $Host.UI.PromptForChoice("", "The task on task scheduler will be set for '$($validDays)' day(s), do you want to change?", $choices, 1)
	Write-Host "`nYou can change later in the task '$($taskName)' under Task Scheduler`n" -ForegroundColor Yellow
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
    $action = New-ScheduledTaskAction -Execute "`"$PSHOME\pwsh.exe`"" -Argument ".\EDMCopyHash.ps1" -WorkingDirectory $EDMSupportFolder
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
	EDM scripts can request change your Execution Policy to bypass to be executed, using PS:\> Set-ExecutionPolicy -ExecutionPolicy bypass.
	In some organizations for security concerns this cannot be set, and the script need to be digital signed.
	This function permit to use a self-signed certificate or use an external one. 
	BE AWARE : The external certificate needs to be for a CODE SIGNING is not a coomon SSL certificate.
	#>
	
	Write-Host "`n`n----------------------------------------------------------------------------------------" -ForegroundColor Yellow
	Write-Host "`nThis option will be digital sign all EDM scripts." -ForegroundColor DarkYellow
	Write-Host "The certificate used is the kind of CodeSigning not a SSL certificate" -ForegroundColor DarkYellow
	Write-Host "If you choose to select your own certificate be aware of this." -ForegroundColor DarkYellow
	Write-Host "`n----------------------------------------------------------------------------------------" -ForegroundColor Yellow
	Write-Host "`n`n" 
	
	# Decide if you want to progress or not
	$choices  = '&Yes', '&No'
    $decision = $Host.UI.PromptForChoice("", "Do you want to proceed with the digital signature for all the scripts?", $choices, 1)
	if ($decision -eq 1)
	{
		Write-Host "`nYou decide don't proceed with the digital signature." -ForegroundColor DarkYellow
		Write-Host "Remember to use EDM scripts set permissions with Administrator rigths on Powershel using:." -ForegroundColor DarkYellow
		Write-Host "`nSet-ExecutionPolicy -ExecutionPolicy bypass." -ForegroundColor Green
	} else
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
			
			#Sign EDM Scripts
			$config = "$PSScriptRoot\EDMConfig.json"
			if(-not (Test-Path -Path $config))
			{
				$config = "$PSScriptRoot\EDM_RemoteConfig.json"
				$json = Get-Content -Raw -Path $config
				[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
				$EDMrootFolder = $config.EDMrootFolder+"EDM*.ps1"
				
				$files = Get-ChildItem -Path $EDMrootFolder
				foreach($file in $files)
				{
					Write-Host "`Signing..."
					Write-Host "$($file.Name)" -ForegroundColor Green
					Set-AuthenticodeSignature -FilePath ".\$($file.Name)" -Certificate $cert
				}
				
			}else
			{
			$json = Get-Content -Raw -Path $config
			[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
			$EDMrootFolder = $config.EDMrootFolder+"EDM*.ps1"
			$EDMrootFolder2Sign = $config.EDMSupportFolder
			$EDMSupportFolder = $config.EDMSupportFolder+"EDM*.ps1"
			
			$files = Get-ChildItem -Path $EDMrootFolder
			$SupportFiles = Get-ChildItem -Path $EDMSupportFolder
			
			foreach($file in $files)
			{
				Write-Host "`Signing..."
				Write-Host "$($file.Name)" -ForegroundColor Green
				Set-AuthenticodeSignature -FilePath ".\$($file.Name)" -Certificate $cert
			}
			foreach($SupportFile in $SupportFiles)
			{
				Write-Host "`Signing..."
				$FileName = $SupportFile.Name
				$File2Sign = $EDMrootFolder2Sign+$FileName
				Write-Host "$($File2Sign)" -ForegroundColor Green
				Set-AuthenticodeSignature -FilePath $File2Sign -Certificate $cert
			}
			}
		}
	}
}

function CreateCodeSigningCertificate
{
	#CMDLET to create certificate
	$EDMcert = New-SelfSignedCertificate -Subject "CN=EDM PowerShell Code Signing Cert" -Type "CodeSigning" -CertStoreLocation "Cert:\CurrentUser\My" -HashAlgorithm "sha256"
		
	### Add Self Signed certificate as a trusted publisher (details here https://adamtheautomator.com/how-to-sign-powershell-script/)
		
		# Add the self-signed Authenticode certificate to the computer's root certificate store.
		## Create an object to represent the CurrentUser\Root certificate store.
		$rootStore = [System.Security.Cryptography.X509Certificates.X509Store]::new("Root","CurrentUser")
		## Open the root certificate store for reading and writing.
		$rootStore.Open("ReadWrite")
		## Add the certificate stored in the $authenticode variable.
		$rootStore.Add($EDMcert)
		## Close the root certificate store.
		$rootStore.Close()
			 
		# Add the self-signed Authenticode certificate to the computer's trusted publishers certificate store.
		## Create an object to represent the CurrentUser\TrustedPublisher certificate store.
		$publisherStore = [System.Security.Cryptography.X509Certificates.X509Store]::new("TrustedPublisher","CurrentUser")
		## Open the TrustedPublisher certificate store for reading and writing.
		$publisherStore.Open("ReadWrite")
		## Add the certificate stored in the $authenticode variable.
		$publisherStore.Add($EDMcert)
		## Close the TrustedPublisher certificate store.
		$publisherStore.Close()	
}

function EncryptPasswords
{
    # read config file
    $CONFIGFILE = "$PSScriptRoot\EDMConfig.json"  
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
        Write-Host "`nAccording to the configuration settings (EncryptedKeys: True), passwords are already encrypted." -ForegroundColor Yellow
        Write-Host "No actions taken."
        return
    }

    # encrypt secrets
    $ClientSecretValue = $config.Password

    $ClientSecretValue = $ClientSecretValue | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString

    # write results to the file
    $config.EncryptedKeys = "True"
    $config.Password = $ClientSecretValue

    $date = Get-Date -Format "yyyyMMddHHmmss"
    Move-Item "EDMConfig.json" "EDMConfig_$date.json"
    Write-Host "`nPasswords encrypted."
    Write-Host "The old config file moved to 'EDMConfig_$date.json'" -ForegroundColor Green
    $config | ConvertTo-Json | Out-File $CONFIGFILE

    Write-Host "Warning!" -ForegroundColor Yellow
    Write-Host "Please note that encrypted passwords can be decrypted only on this machine, using the same account." -ForegroundColor Yellow
}

function EncryptRemotePasswords
{
    # read config file
    $CONFIGFILE = "$PSScriptRoot\EDM_RemoteConfig.json"  
    if (-not (Test-Path -Path $CONFIGFILE))
    {
        Write-Host "`nMissing config file '$CONFIGFILE'." -ForegroundColor Yellow
        return
    }
    $json = Get-Content -Raw -Path $CONFIGFILE
    [PSCustomObject]$RemoteConfig = ConvertFrom-Json -InputObject $json
    $EncryptedKeys = $RemoteConfig.EncryptedKeys

    # check if already encrypted
    if ($EncryptedKeys -eq "True")
    {
        Write-Host "`nAccording to the configuration settings (EncryptedKeys: True), passwords are already encrypted." -ForegroundColor Yellow
        Write-Host "No actions taken."
        return
    }

    # encrypt secrets
    $ClientSecretValue = $RemoteConfig.Password

    $ClientSecretValue = $ClientSecretValue | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString

    # write results to the file
    $RemoteConfig.EncryptedKeys = "True"
    $RemoteConfig.Password = $ClientSecretValue

    $date = Get-Date -Format "yyyyMMddHHmmss"
    Move-Item "EDM_RemoteConfig.json" "bck_EDM_RemoteConfig_$date.json"
    Write-Host "`nPasswords encrypted."
    Write-Host "The old config file moved to 'EDMConfig_$date.json'" -ForegroundColor Green
    $RemoteConfig | ConvertTo-Json | Out-File $CONFIGFILE

    Write-Host "Warning!" -ForegroundColor Yellow
    Write-Host "Please note that encrypted passwords can be decrypted only on this machine, using the same account." -ForegroundColor Yellow
}

function SubMenuInitialization
{
	cls
	Write-Host "`n`n----------------------------------------------------------------------------------------"
	Write-Host "`nWelcome to the Initilization Menu!" -ForegroundColor Green
	Write-Host "This is a first configuration about folders, credentials, encrypt password, sign the scrips and validate the EDM Connection" -ForegroundColor Green
	$choice = 1
	while ($choice -ne "0")
	{
		Write-Host "`n----------------------------------------------------------------------------------------"
		Write-Host "`nWhat do you want to do?" -ForegroundColor Blue
		Write-Host "`t[1] - Initial Setup for EDM (principal folders used)"
		Write-Host "`t[2] - Get credentials for connection"	
		Write-Host "`t[3] - Encrypt passwords"
		Write-Host "`t[4] - Connect to EDM service"
		Write-Host "`t[0] - Back to principal menu"
		Write-Host "`n"
		Write-Host "`nPlease choose option:"
		
		$choice = ([System.Console]::ReadKey($true)).KeyChar
		switch ($choice) {
        "1" {
                SelectEDMPaths
                WriteToJsonFile
                break
            }
		"2" {
				GetEDMUserCredentials
				WriteToJsonFile
				break
			 }
		"3" {EncryptPasswords; break}
		"4" {
				Connect2EDM
				TakeAPause
				break
			}
		"0" {return}
		}
	
	}
}

function SubMenuEDMGeneration
{
	cls
	Write-Host "`n`n----------------------------------------------------------------------------------------"
	Write-Host "`nWelcome to the EDM Generation Menu!" -ForegroundColor DarkYellow
	Write-Host "This is the 2nd step to set DataStore name, get the Schema file, select your data and hash the data." -ForegroundColor DarkYellow
	$choice = 1
	while ($choice -ne "0")
	{
		Write-Host "`n----------------------------------------------------------------------------------------"
		Write-Host "`nWhat do you want to do?" -ForegroundColor Blue
		Write-Host "`t[1] - Get EDM Datastores"
		Write-Host "`t[2] - Get Schema file"
		Write-Host "`t[3] - Validate EDM Data"
		Write-Host "`t[4] - Create Hash for your data"
		Write-Host "--- If you want to use another server to upload your data go back and select the next menu---" -ForegroundColor Blue
		Write-Host "`t[5] - Upload Hash data"
		Write-Host "`t[6] - EDM Hash Upload Status"
		Write-Host "`t[7] - Crete task to create Hash files"
		Write-Host "`t[0] - Back to principal menu"
		Write-Host "`n"
		Write-Host "`nPlease choose option:"
		
		$choice = ([System.Console]::ReadKey($true)).KeyChar
		switch ($choice) {
        "1" {GetDataStores;break}
		"2" {GetSchemaFile; break}
		"3" {ValidateEDMData; break}
		"4" {EDMHashCreation; break}
		"5" {EDMHashUpload; break}
		"6" {EDMUploadStatus; break}
		"7" {CreateEDMHashCreateScheduledTask; break}
		"8" {CreateEDMHashUploadScheduledTask; break}
		"0" {return}
		}
	
	}	
}

function SubMenuRemoteUpload
{
	cls
	Write-Host "`n`n----------------------------------------------------------------------------------------"
	Write-Host "`nEDM for remote upload menu" -ForegroundColor Magenta
	Write-Host "If you want to upload from another server this menu is for that." -ForegroundColor Magenta
	$choice = 1
	while ($choice -ne "0")
	{
		Write-Host "`n----------------------------------------------------------------------------------------"
		Write-Host "`nWhat do you want to do?" -ForegroundColor Blue
		Write-Host "`t[1] - Copy the data needed to a remote server"
		Write-Host "`t[2] - Create task to copy Hash data daily"
		Write-Host "`t[0] - Back to principal menu"
		Write-Host "`n"
		Write-Host "`nPlease choose option:"
		
		$choice = ([System.Console]::ReadKey($true)).KeyChar
		switch ($choice) {
		"1" {EDMCopyDataNeeded; break}
		"2" {CreateEDMHashCopyScheduledTask; break}
		"0" {return}
		}
	
	}
}

function SubMenuRemoteConfig
{
	cls
	Write-Host "`n`n----------------------------------------------------------------------------------------"
	Write-Host "`nWelcome to Remote Menu!"
	Write-Host "This menu is to be used only on the remote server used to upload hash data to Microsoft 365"
	Write-Host "Used only if you are using a 2nd Server to upload Hash data" -ForegroundColor Red
	$choice = 1
	while ($choice -ne "0")
	{
		Write-Host "`n----------------------------------------------------------------------------------------"
		Write-Host "`nWhat do you want to do?" -ForegroundColor Blue
		Write-Host "`t[1] - Plase validate your new folders."
		Write-Host "`t[2] - Sign the scripts again."
		Write-Host "`t[3] - Change credentials, only if you want to use another account."
		Write-Host "`t[4] - Encrypt password."
		Write-Host "`t[5] - Upload Hash to Microsoft 365."
		Write-Host "`t[6] - Create a task to upload Hash to Microsoft 365."
		Write-Host "`t[7] - Check Hash upload status."
		Write-Host "`t[0] - Back to principal menu"
		Write-Host "`n"
		Write-Host "`nPlease choose option:"
		
		$choice = ([System.Console]::ReadKey($true)).KeyChar
		switch ($choice) {
		"1" {
				SelectEDMRemotePaths
				WriteToRemoteJsonFile
				break
			}
		"2" {SelfSignScripts; break}
		"3" {
				GetEDMRemoteUserCredentials
				WriteToRemoteJsonFile
				break
			}
		"4" {EncryptRemotePasswords; break}
		"5" {EDMHashUpload; break}
		"6" {CreateEDMRemoteHashUploadScheduledTask; break}
		"7" {EDMUploadStatus; break}
		"0" {return}
		}
	
	}
}

function SubMenuSupportingElements
{
	cls
	Write-Host "`n`n----------------------------------------------------------------------------------------"
	Write-Host "`nWelcome to Support element Menu!" -ForegroundColor Magenta
	Write-Host "Here you can sign the scripts" -ForegroundColor Magenta
	$choice = 1
	while ($choice -ne "0")
	{
		Write-Host "`n----------------------------------------------------------------------------------------"
		Write-Host "`nWhat do you want to do?" -ForegroundColor Blue
		Write-Host "`t[1] - Sign EDM scripts"
		Write-Host "`t[0] - Back to principal menu"
		Write-Host "`n"
		Write-Host "`nPlease choose option:"
		
		$choice = ([System.Console]::ReadKey($true)).KeyChar
		switch ($choice) {
		"1" {SelfSignScripts; break}
		"0" {return}
		}
	
	}
}

############
# Main code
############

cls
$config = InitializeEDMConfigFile
InitializeHostName

Write-Host "`nRunning prerequisites check..."
CheckPrerequisites

Write-Host "`n`n----------------------------------------------------------------------------------------"
Write-Host "`nWelcome to the EDM setup script!" -ForegroundColor Blue
Write-Host "Script allows to automatically execute setup steps." -ForegroundColor Blue
Write-Host "`n----------------------------------------------------------------------------------------"

### Validate hostname
$config = "$PSScriptRoot\EDMConfig.json"
$config2 = "$PSScriptRoot\EDMConfig.json"
if (-not (Test-Path -Path $config))
{
	$config = "$PSScriptRoot\EDM_RemoteConfig.json"
	$json = Get-Content -Raw -Path $config
	[PSCustomObject]$RemoteConfig = ConvertFrom-Json -InputObject $json
	$EDMHostName = $RemoteConfig.EDMHostName
	$EDMHostExecuting = hostname
}else
{
	$json = Get-Content -Raw -Path $config
	[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
	$EDMHostName = $config.EDMHostName
	$EDMHostExecuting = hostname
}


If($EDMHostName -ne $EDMHostExecuting)
{
	Write-Host "`n####################################################################################" -ForegroundColor Red
	Write-Host "`nYou are executing in a remote server" -ForegroundColor DarkCyan
	Write-Host "Work with the menu 4 (Remote server activities) " -ForegroundColor DarkCyan
	Write-Host "`n####################################################################################" -ForegroundColor Red
}

$choice = 1
while ($choice -ne "0")
{
    Write-Host "`nWhat do you want to do?"
    Write-Host "`t[1] - Initial Setup for EDM"
	Write-Host "`t[2] - Generate EDM Hash & upload(optional) "
	Write-Host "`t[3] - Copy files needed and Hash to a remote server"
	if (-not (Test-Path -Path $config2))
	{
		Write-Host "`t[4] - Remote server activities" -ForegroundColor Green
	}
	Write-Host "`t[4] - Remote server activities"
	Write-Host "`t[9] - Supporting elements"
    Write-Host "`t[0] - Exit"
	Write-Host "`n"
	Write-Host "`nPlease choose option:"

    $choice = ([System.Console]::ReadKey($true)).KeyChar
    switch ($choice) {
        "1" {SubMenuInitialization;break}
		"2" {SubMenuEDMGeneration;break}
		"3" {SubMenuRemoteUpload;break}
		"4" {SubMenuRemoteConfig;break}
		"9" {SubMenuSupportingElements; break}
		"0" {
				$OriginalPath = $PSScriptRoot
				cd $OriginalPath | cmd
			exit}
    }
}


# SIG # Begin signature block
# MIIFywYJKoZIhvcNAQcCoIIFvDCCBbgCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAfHh9P7rA+2RB8
# IfdVQSZ9QgPUszyzkn62SQx84FpRe6CCAy4wggMqMIICEqADAgECAhB8ncNW0y2J
# mUf2EeuGEwoOMA0GCSqGSIb3DQEBCwUAMC0xKzApBgNVBAMMIk1QQVJSIFBvd2Vy
# U2hlbGwgQ29kZSBTaWduaW5nIENlcnQwHhcNMjMxMDA1MTA1MzEwWhcNMjQxMDA1
# MTExMzEwWjAtMSswKQYDVQQDDCJNUEFSUiBQb3dlclNoZWxsIENvZGUgU2lnbmlu
# ZyBDZXJ0MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA4Pb3hHiflVlv
# fWMNz2SqHCT/xq/wgzncd4j9MX/d5jIQ9Ln312297R/d+GVdVVBOsi1+OuDB5UWO
# XTxL+NlCeulHU4ye4JBE30y6XmDC9J2ygrUlSc2ClurNThRHNc15kd3lurR1Y6VI
# 8y0yHN7ijH/N/z9HPsyov8EdCLUmfKcc7ibKcyxCZz3Nnzd8YcEdHAwGgeGrOLen
# /ptfv4Cs1GbNA+FKzWk1g4eBLfHoMA+d2FjEJ/VHz+kzLr+oUyaR2NvkdNHNWHw5
# 8A49AdOMRTTBUFh5owqb/Eg2RmxzxAUeYT2xWsDeUF100/F2hF9ueeaQfkMJ56H8
# dXAUUt87SQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYB
# BQUHAwMwHQYDVR0OBBYEFLUt3inp0JwA7tNAP31Sx2moiu1vMA0GCSqGSIb3DQEB
# CwUAA4IBAQB+kfOCUotSSR8mJGz+WMu04UcXSulRN/YZIBMdq9cRQkirn6upBAF3
# jFQQ5DIHQATEJUrB/tnQYbLXIDbHUmzKeD0mXd37Pw8fZdpOehvqia9wg0fAPW7r
# /haIVCas4Q4qmTcDFCeU9f8Yf0E0ZDzezL5IV6m2LGDmQOP+uzbIupnkQyPPD+Y8
# HvHlW+rz85MmDUNAMJZ5Un6jIaX6vAfbRY/nOMgDLp/LN0xNX8GwBe4nkMfJyFq3
# r0p0yBWmySXDjoaIuDDFdZ0j8EZ8rLKRPJy0ALjlU+pgSjNAvJkZJwsd2Homwi/0
# bilMd4E9V/ext2xRMg4Qd/VuKpEZ3dDfMYIB8zCCAe8CAQEwQTAtMSswKQYDVQQD
# DCJNUEFSUiBQb3dlclNoZWxsIENvZGUgU2lnbmluZyBDZXJ0AhB8ncNW0y2JmUf2
# EeuGEwoOMA0GCWCGSAFlAwQCAQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKEC
# gAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwG
# CisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEII0JL4htCusoyjiep1OVklj8Vzbi
# 5KmsKUCX0+WsT3YsMA0GCSqGSIb3DQEBAQUABIIBAJLgs+NPDFWx9vE1aRfkaE3B
# Ca4MrVd4Q4GWsFSjd9J+7jn3RI7SMkJf2PWlkteggGoEaVYp7OfGxOSaOxIGoenc
# FsYshBLLDJqUocLwzpqOcbe3KTextQPHaGiGeLRkIgk+tu7qGVNAn1uw9bF4gJXq
# 3Xjutz5ZeP1pN1DIRyPoD/4ESVN1XYJUE1+iSlPSwMsHpyLLo1LBuKLkrClOltAs
# 11vvNSiJFEhwwcOS7OFhNf8qGBoioTl0fqtPyow3k2bULklcxG8si+crtw/8Ibqf
# zQzAn4R1hhNK3NVhRxGqCFDqdqTHwHXX6rA5WLUnSZRKBuVI/2cPo9MigHy7cF8=
# SIG # End signature block
