<#
.SYNOPSIS
    Script to Upload the Hash used on EDM from data file to Microsoft 365.

.DESCRIPTION
    Script is designed to simplify EDM configuration as a task.
	Takes the hash an upload to Microsoft 365
    
.NOTES
    Version 0.9
    Current version - 27.10.2023
#> 

<#
HISTORY
  2023-10-27	S.Zamorano	- Initial script to create Hash locally
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
    CheckPowerShellVersion
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

function Connect2EDM
{
	$CONFIGFILE = "$PSScriptRoot\..\EDMConfig.json"
	if (-not (Test-Path -Path $CONFIGFILE))
	{
		Write-Error "Missing config file." -ForegroundColor Red
		exit(1)
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

function HashDate
{
	$configfile = "$PSScriptRoot\..\EDMConfig.json"
	$json = Get-Content -Raw -Path $configfile
	[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
	$OutputPath = $config.EDMSupportFolder
	$HashFolder = $config.HashFolder
	
	$Hashfile = gci $HashFolder -Filter *.edmhash | select -last 1
	
	$timestampFile = $OutputPath + "CopyHash_timestamp.json"
	# read LastWriteTime from the file
	if (-not (Test-Path -Path $timestampFile))
	{
		# if file not present create new value
		$Hashtimestamp = $Hashfile.LastWriteTime.ToString("yyyy-MM-ddTHH:mm:ss")
	}else{
		$json = Get-Content -Raw -Path $timestampFile
		[PSCustomObject]$timestamp = ConvertFrom-Json -InputObject $json
		$Hashtimestamp = $timestamp.LastWriteTime.ToString("yyyy-MM-ddTHH:mm:ss")
	}
	$Hashtimestamp = @{"LastWriteTime" = $Hashtimestamp}
	ConvertTo-Json -InputObject $Hashtimestamp | Out-File -FilePath $timestampFile -Force
}

function CreateHash
{
	CheckPrerequisites
	HashDate
	$configfile = "$PSScriptRoot\..\EDMConfig.json"
	$json = Get-Content -Raw -Path $configfile
	[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
	$HashData = $config.HashFolder
	$HashData = $HashData.Substring(0,$HashData.Length-1)
	$HashData = $HashData+"*"
	$HashFolder = $config.HashFolder
	$OutputPath = $config.EDMSupportFolder
	$Destination = $config.EDMremoteFolder
	
	$timestampFile = $OutputPath + "CopyHash_timestamp.json"
	$jsonHash = Get-Content -Raw -Path $timestampFile
	[PSCustomObject]$timestamp = ConvertFrom-Json -InputObject $jsonHash
	$Hashtimestamp = $timestamp.LastWriteTime.ToString("yyyy-MM-ddTHH:mm:ss")
	#Write-Host "Hashtimestamp '$($Hashtimestamp)'." -ForegroundColor Green
	$Hashfile = gci $HashFolder -Filter *.edmhash | select -last 1
	$HashfileTime = $Hashfile.LastWriteTime.ToString("yyyy-MM-ddTHH:mm:ss")
	#Write-Host "Hashfile '$($Hashfile.LastWriteTime.ToString("yyyy-MM-ddTHH:mm:ss"))'." -ForegroundColor Green
	
	if($HashfileTime -eq $Hashtimestamp)
	{
		Write-Host "Hash file is still the same, nothing was Uploaded." -ForegroundColor DarkYellow
	}else{
		Connect2EDM | Out-Null
	
		$configfile = "$PSScriptRoot\..\EDMConfig.json"
		$json = Get-Content -Raw -Path $configfile
		[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
		
		$EDMDSName = $config.DataStoreName
		$HashName = $config.HashFolder+$config.HashFile
		$SchemaFolder = $config.SchemaFolder
		
		$EDMFolder = $config.EDMAppFolder
		cd $EDMFolder | cmd	
		
		.\EdmUploadAgent.exe /UploadHash /DataStoreName $EDMDSName /HashFile $HashName
		Write-Host "`nREMEMBER: You can update your EDM data only 5 times per day." -ForegroundColor RED
		cd $OutputPath | cmd
	}
	
}

CreateHash