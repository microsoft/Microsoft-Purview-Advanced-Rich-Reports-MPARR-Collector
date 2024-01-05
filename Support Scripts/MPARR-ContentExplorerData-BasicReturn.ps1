<#   
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
HISTORY
Script      : MPARR-ContentExplorerData-BasicReturn.ps1
Author      : Sebastian Zamorano
Co-Author   : 
Version     : 1.0.4
Date		: 22-12-2023
Description : The script exports Content Explorer from Export-ContentExplorerData and pushes into a customer-specified Log Analytics table. 
			Please note if you change the name of the table - you need to update Workbook sample that displays the report , appropriately. 
			Do ensure the older table is deleted before creating the new table - it will create duplicates and Log analytics workspace doesn't support upserts or updates.
			
.NOTES 
	22-12-2023	S. Zamorano		- First released
	26-12-2023	S. Zamorano		- Added functions to support list of SITs, list of Trainable Classifiers and capability to set Page Size
	05-01-2024	S. Zamorano		- Change page size adjustment, export to csv modified to append data to CSV on each query, avoiding to collect all in one array previous to export(memory management)
#>

[CmdletBinding(DefaultParameterSetName = "None")]
param(
    [Parameter()] 
        [switch]$ChangePageSize,
	#Export-ContentExplorerData cmdlet requires a PageSize that can be between 1 to 5000, by default is set to 100, you can change the number below or use the parameter -ChangePageSize to modify during the execution
	[int]$InitialPageSize = 100
)

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

function connect2service
{
	Write-Host "`nAuthentication is required, please check your browser" -ForegroundColor Green
	Connect-IPPSSession -UseRPSSession:$false
}

function ReadWorkload
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
}

function ReadTagType
{
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
		$TCSelected = "$PSScriptRoot\MPARR-TrainableClassifiersList.json"
		
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

function CollectData($TagType, $Workload, $PageSize)
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
	
	$date = Get-Date -Format "yyyyMMddHHmmss"
	$ExportFile = "ContentExplorerExport - "+$TagType+" - "+$textvalue+" - "+$Workload+" - "+$date+".csv"
	
	Write-Host "`nFile to be written :" -NoNewLine
	Write-Host $ExportFile -ForeGroundColor Green 
	
	Write-Host "`nFile to be copied at :" -NoNewLine
	Write-Host $PSScriptRoot -ForeGroundColor Green 

	Write-Host "`n"
	
	$CEResults = @()
	$query = Export-ContentExplorerData -TagType $TagType -TagName $tagname -PageSize $PageSize -Workload $Workload
	$var = $query.count
	$Total = $query[0].TotalCount
	$remaining = $Total
	
	if($Total -eq 0)
	{
		Write-Host "`n### Your query don't returned records. ###" -ForeGroundColor Blue
		Write-Host -NoNewLine "`nPress any key to continue..."
		$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
		MainFunction
	}else
	{
		Write-Host "Total matches returned :" -NoNewLine
		Write-Host $remaining -ForeGroundColor Green	
	}
	


	While ($query[0].MorePagesAvailable -eq 'True') {
		$CEResults += $query[1..$var]
		$query = Export-ContentExplorerData -TagType $TagType -TagName $tagname -PageSize $PageSize -Workload $Workload -PageCookie $query[0].PageCookie
		$remaining -= ($var - 1)
		Write-Host "Total matches remaining to process :" -NoNewLine
		Write-Host $remaining -ForeGroundColor Green
		$CEResults | Export-Csv -Path $ExportFile -NTI -Force -Append | Out-Null
		$CEResults = @()
	}

	if ($remaining -gt 0)
	{
		$CEResults += $query[1..$remaining]
		$CEResults | Export-Csv -Path $ExportFile -NTI -Force -Append | Out-Null
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
    #    Desc           : Extracts data from Content ExplorerData into a CSV
    #    Return         : None
    # ---------------------------------------------------------------
		<#
		.NOTES
		If you cannot add the "Compliance Administrator" role to the Microsoft Entra App, for security reasons, you can comment the line 167 and uncomment the line 166, in that case
		Someone with "Compliance Administrator" role needs to execute this script, this script is executed on-demand to refresh the SITs names
		#>
		
		#Clean screen after connection
		cls
		
		#Welcome screen
		Write-Host "`n#################################################################################" -ForeGroundColor Green
		Write-Host "`n"
		Write-Host "This script was thought to help to you to execute the cmdlet Export-ContentExplorerData and exporting the data to a CSV file."
		Write-Host "Remember check that you have the right permissions."
		Write-Host "`n#################################################################################" -ForeGroundColor Green
		Write-Host "`n"
		
		#Here need to be set the function to read the TagType configuration
		$TagType = ReadTagType
		
		#Read workloads to be used with Export-ContentExplorerData
		$Workload = ReadWorkload	

		#PageSize to be used
		if($ChangePageSize)
		{
			$Size = ExportPageSize -PageSize $InitialPageSize
		}else
		{
			$Size = $InitialPageSize
		}

		#Execute the query
		CollectData -TagType $TagType -Workload $Workload -PageSize $Size
		
		#Check if you want to finish or request a new export
		SelectContinuity
}  
 
#Main Code - Run as required. Do ensure older table is deleted before creating the new table - as it will create duplicates.
CheckPrerequisites
connect2service 
MainFunction
