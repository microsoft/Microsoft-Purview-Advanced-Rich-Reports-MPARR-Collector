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
Script      : Get-RMSData.ps1
Author      : S. Zamorano
Version     : 1.2.1
Description : The script exports Aipservice Log Data from Microsoft AADRM API and pushes into a customer-specified Log Analytics table. Please note if you change the name of the table - you need to update Workbook sample that displays the report , appropriately. Do ensure the older table is deleted before creating the new table - it will create duplicates and Log analytics workspace doesn't support upserts or updates.
2022-10-19		S. Zamorano		- Added laconfig.json file for configuration and decryption function
2022-11-18      G.Berdzik       - Fixed issue with data parsing
2022-12-21      G.Berdzik       - Changed logic to avoid data duplicates
2022-12-28      S. Zamorano     
2023-01-02      G.Berdzik       - Minor change (check for output directory)
2023-01-25      G.Berdzik       - Added code for Get-AipServiceTrackingLog data
2023-01-26      G.Berdzik       - Added support for multithreading
#>

[CmdletBinding()]
param (
    # Log Analytics table where the data is written to. Log Analytics will add an _CL to this name.
    [string]$TableName = "RMSData"
)

# Function to decrypt shared key
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


# Verify output directory exists
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


function Build-Signature ($customerId, $sharedKey, $date, $contentLength, $method, $contentType, $resource) {
    # ---------------------------------------------------------------   
    #    Name           : Build-Signature
    #    Value          : Creates the authorization signature used in the REST API call to Log Analytics
    # ---------------------------------------------------------------

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

function Post-LogAnalyticsData($body, $LogAnalyticsTableName) {
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
    $response = Invoke-WebRequest -Uri $uri -Method Post -Headers $headers -ContentType $contentType -Body $bodyJsonUTF8 -UseBasicParsing

    if ($Response.StatusCode -eq 200) {   
        $rows = $body.Count
        Write-Information -MessageData "   $rows rows written to Log Analytics workspace $uri" -InformationAction Continue
    }

}

function GetAipServiceTrackingLogData($source, $array)
{
    # ---------------------------------------------------------------   
    #    Name           : GetAipServiceTrackingLogData
    #    Desc           : Gets data from Get-AipServiceTrackingLog
    # ---------------------------------------------------------------

    Write-Host "Processing Get-AipServiceTrackingLog..." -ForegroundColor Cyan
    $inputRows = $source | Where-Object {$_."content-id" -ne "-"}
    $count = $array.Count

    #region threads
    $runspacePool= [runspacefactory]::CreateRunspacePool(1, $threads)
    $runspacePool.Open()
    $jobs= New-Object System.Collections.ArrayList
    # worker script (just executing AIP cmdlet)
    $worker = {
        param($item)

        $id = $item."content-id" -replace "[{}]", ""
        Get-AipServiceTrackingLog -ContentId $id
    }

    foreach ($item in $inputRows)
    {
        $powerShell= [powershell]::Create()
        $powerShell.RunspacePool = $runspacePool
        $powerShell.AddScript($worker).AddArgument($item) | Out-Null
        $jobObj = New-Object -TypeName PSObject -Property @{
            Runspace = $powerShell.BeginInvoke()
            PowerShell = $powerShell
        }
        [void]$jobs.Add($jobObj)
    }

    # wait for jobs to complete
    while ($jobs.Runspace.IsCompleted -contains $false)
    {
        Start-Sleep -Milliseconds 10
    }
    # receive results
    $results = $jobs | ForEach-Object {
        $_.PowerShell.EndInvoke($_.Runspace)
        $_.PowerShell.Dispose()
    }
    $jobs.Clear()
    [void]$runspacePool.Close()
    [void]$runspacePool.Dispose()
    [GC]::Collect()
    #endregion

    # copy results to the destination array
    foreach ($item in $results)
    {
        if ($item -ne $null)
        {
            foreach ($element in $item)
            {
                [void]$array.Add($element)
            }
        }
    }
    Write-Host "$($array.count - $count) elements retrieved." -ForegroundColor Cyan
}


function Export-RMSusersLogs {
    # ---------------------------------------------------------------   
    #    Name           : Export-RMSusersLogs
    #    Desc           : Extracts data from Get-MgUser into Log analytics workspace tables for reporting purposes
    #    Return         : None
    # ---------------------------------------------------------------
    
        Connect-AIPService -CertificateThumbPrint $CertificateThumb -ApplicationId $AppClientID -TenantId $TenantGUID -ServicePrincipal

        # get the newest log file and set startTime
        $processedFiles = Get-ChildItem "$RMSLogs\*.processed" | Sort-Object -Property Name -Descending
        if ($processedFiles.Count -gt 0)
        {
            $lastFile = $processedFiles[0].FullName
            $lastFile -match ".*aipservicelog-(?<date>\d{4}-\d{2}-\d{2}).*" | Out-Null
            $fileDate = $Matches.date
            $startTime = $fileDate
        }
        else 
        {
            $startTime = (Get-Date).ToString("yyyy-MM-dd") 
        }

        # Calculate number of threads
        $threads = (Get-WmiObject win32_computersystem).NumberOfLogicalProcessors * 4
        $ea = $ErrorActionPreference
        $ErrorActionPreference = "SilentlyContinue"
        Write-Host "Fetching logs..."
		$response = Get-AipServiceUserLog -FromDate $startTime -NumberOfThreads $threads -Path $RMSLogs -force -ErrorVariable MyError
        $ErrorActionPreference = $ea
        if ($MyError.Count -gt 0)
        {
            Write-Host $MyError[0].ErrorRecord -ForegroundColor Red
            Write-Host "Exiting..."
            exit(2)
        }
        $result = $response -split "`n"
        $files = New-Object System.Collections.ArrayList
        foreach ($line in $result)
        {
            if ($line -match ".*The log is available at (?<fileName>.*aipservice.*)\.")
            {
                [void]$files.Add($Matches.fileName)
            }
        }

        # If there is no data, skip
        if ($files.Count -eq 0) { continue; }

        # define batch size
        $batchSize = 10MB
        # Else format for Log Analytics
        foreach ($file in $files)
        {
            # table to store info from Get-AipServiceTrackingLog
            $rmsDetails = New-Object System.Collections.ArrayList
            
            $currentBufferSize = 0
            $csv = New-Object System.Collections.ArrayList
            # parse file header (field's names)
            $logData = Get-Content -Path $file -TotalCount 4
            $csvHeader = ($logData | Select-String "^#Fields:").ToString().Replace("#Fields: ", "")
            [void]$csv.Add($csvHeader)
        
            if (Test-Path -Path "$($file).processed")
            {
                Write-Host "Checking file $file..."
                # check if processed file is the same as the new one
                $oldFileHash = (Get-FileHash -Path "$($file).processed" -Algorithm SHA256).Hash
                $newFileHash = (Get-FileHash -Path $file -Algorithm SHA256).Hash
                if ($oldFileHash -eq $newFileHash)
                {
                    Remove-Item -Path $file -Force
                    continue
                }

                # compare files
                $newContent = [RMSDataTools.FileComparer]::Compare("$($file).processed", $file)
                foreach ($line in $newContent)
                {
                    $currentBufferSize += $line.Length
                    if ($currentBufferSize -lt $batchSize)
                    {
                        [void]$csv.Add($line)
                    }
                    else 
                    {
                        $data = $csv | ConvertFrom-Csv -Delimiter "`t"
                        GetAipServiceTrackingLogData $data $rmsDetails
                        Post-LogAnalyticsData -LogAnalyticsTableName $TableName -body $data
                        $currentBufferSize = 0
                        $csv.Clear()
                        [void]$csv.Add($csvHeader)
                    }
                }
            }
            else 
            {
                $srFile = [System.IO.StreamReader]::new($file)
                # skip first 4 lines (header)
                for ($i=0; $i -lt 4; $i++)
                {
                    $x = $srFile.ReadLine()
                }
                while ($line = $srFile.ReadLine())
                {
                    $currentBufferSize += $line.Length
                    if ($currentBufferSize -lt $batchSize)
                    {
                        [void]$csv.Add($line)
                    }
                    else 
                    {
                        $data = $csv | ConvertFrom-Csv -Delimiter "`t"
                        GetAipServiceTrackingLogData $data $rmsDetails
                        Post-LogAnalyticsData -LogAnalyticsTableName $TableName -body $data
                        $currentBufferSize = 0
                        $csv.Clear()
                        [void]$csv.Add($csvHeader)
                    }
                }    
                $srFile.Close()
            }
            $data = $csv | ConvertFrom-Csv -Delimiter "`t"

            # Push data to Log Analytics
            GetAipServiceTrackingLogData $data $rmsDetails
            Post-LogAnalyticsData -LogAnalyticsTableName $TableName -body $data
            Write-Host "File '$file' was processed with $($csv.count - 1) records."
            Move-Item $file "$($file).processed" -Force

            if ($rmsDetails.Count -gt 0)
            {
                Post-LogAnalyticsData -LogAnalyticsTableName ($TableName + "Details") -body ($rmsDetails.ToArray())
            }
            $rmsDetails.Clear()
        }
		Disconnect-AIPService
    }
    
 
#Main Code - Run as required. Do ensure older table is deleted before creating the new table - as it will create duplicates.
if ($PSVersionTable.PSVersion.Major -gt 5)
{
    Write-Host "Windows PowerShell is required, cannot run on PowerShell Core." -ForegroundColor Yellow
    exit(3)
}

#region CSharp code
$cSharp = @"
using System;
using System.Linq;
using System.IO;
using System.Collections.Generic;

namespace RMSDataTools
{
    public class FileComparer 
    {
        public static Array Compare(string firstFile, string secondFile)
        {
            var file1Lines = File.ReadLines(firstFile);
            var file2Lines = File.ReadLines(secondFile);
            IEnumerable<String> inSecondNotInFirst = file2Lines.Except(file1Lines);
            return inSecondNotInFirst.ToArray();
        }
    }
}
"@

Add-Type -Language CSharp -TypeDefinition $cSharp
#endregion

#region Config file read
$CONFIGFILE = "$PSScriptRoot\laconfig.json"
$json = Get-Content -Raw -Path $CONFIGFILE
[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
$EncryptedKeys = $config.EncryptedKeys
$AppClientID = $config.AppClientID
$TenantGUID = $config.TenantGUID
$ClientSecretValue = $config.ClientSecretValue
$WLA_CustomerID = $config.LA_CustomerID
$WLA_SharedKey = $config.LA_SharedKey
$CertificateThumb = $config.CertificateThumb
$OutputPath = $config.OutPutLogs
$RMSLogs = $config.RMSLogs

CheckOutputDirectory $RMSLogs

if ($EncryptedKeys -eq "True")
{
    $WLA_SharedKey = DecryptSharedKey $WLA_SharedKey
    $ClientSecretValue = DecryptSharedKey $ClientSecretValue
	$CertificateThumb = DecryptSharedKey $CertificateThumb
}
#endregion

# Your Log Analytics workspace ID
$LogAnalyticsWorkspaceId = $WLA_CustomerID

# Use either the primary or the secondary Connected Sources client authentication key   
$LogAnalyticsPrimaryKey = $WLA_SharedKey 

if($LogAnalyticsWorkspaceId -eq "") { throw "Log Analytics workspace Id is missing! Update the script and run again" }
if($LogAnalyticsPrimaryKey -eq "")  { throw "Log Analytics primary key is missing! Update the script and run again" }

Export-RMSusersLogs 
