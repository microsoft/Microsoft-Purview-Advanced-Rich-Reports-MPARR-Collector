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
Script      : Get-AzureADData.ps1
Author      : S. Zamorano
Version     : 1.1.0
Description : The script exports Azure AD users from Microsoft Graph and pushes into a customer-specified Log Analytics table. Please note if you change the name of the table - you need to update Workbook sample that displays the report , appropriately. Do ensure the older table is deleted before creating the new table - it will create duplicates and Log analytics workspace doesn't support upserts or updates.
2022-10-12		S. Zamorano		- Added laconfig.json file for configuration and decryption function
2022-10-18		G. Berdzik		- Fix licensing information
2023-01-03		S. Zamorano		- Added Change to use beta API capabilities, added Id for users
2023-03-31      G. Berdzik      - Support for large tenants
2023-03-31		S. Zamorano		- Visual improvement for progress
#>


param (
    # Log Analytics table where the data is written to. Log Analytics will add an _CL to this name.
    [string]$TableName = "AzureADUsers"

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

$CONFIGFILE = "$PSScriptRoot\laconfig.json"
$json = Get-Content -Raw -Path $CONFIGFILE
[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
$EncryptedKeys = $config.EncryptedKeys
$AppClientID = $config.AppClientID
$ClientSecretValue = $config.ClientSecretValue
$TenantGUID = $config.TenantGUID
$TenantDomain = $config.TenantDomain
$WLA_CustomerID = $config.LA_CustomerID
$WLA_SharedKey = $config.LA_SharedKey
$CertificateThumb = $config.CertificateThumb
$OnmicrosoftTenant = $config.OnmicrosoftURL
if ($EncryptedKeys -eq "True")
{
    $WLA_SharedKey = DecryptSharedKey $WLA_SharedKey
    $ClientSecretValue = DecryptSharedKey $ClientSecretValue
	$CertificateThumb = DecryptSharedKey $CertificateThumb
}

# Your Log Analytics workspace ID
$LogAnalyticsWorkspaceId = $WLA_CustomerID

# Use either the primary or the secondary Connected Sources client authentication key   
$LogAnalyticsPrimaryKey = $WLA_SharedKey 

if($LogAnalyticsWorkspaceId -eq "") { throw "Log Analytics workspace Id is missing! Update the script and run again" }
if($LogAnalyticsPrimaryKey -eq "")  { throw "Log Analytics primary key is missing! Update the script and run again" }

 

Function Build-Signature ($customerId, $sharedKey, $date, $contentLength, $method, $contentType, $resource) {
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

Function Post-LogAnalyticsData($body, $LogAnalyticsTableName) {
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
        #Write-Information -MessageData "   $rows rows written to Log Analytics workspace $uri" -InformationAction Continue
    }

}

Function Export-AzureADData() {
    # ---------------------------------------------------------------   
    #    Name           : Export-AzureADData
    #    Desc           : Extracts data from Get-MgUser into Log analytics workspace tables for reporting purposes
    #    Return         : None
    # ---------------------------------------------------------------
    
    Connect-MgGraph -CertificateThumbPrint $CertificateThumb -AppID $AppClientID -TenantId $TenantGUID
    #Connect-MgGraph -Scopes 'User.Read.All', 'AuditLog.Read.All'  # for testing purposes only
    Select-MgProfile -Name "beta"

    Write-Host "Fetching data from Azure AD..."
    $body = @{
        select='userPrincipalName,displayName,signInActivity,assignedLicenses,assignedPlans,city,createdDateTime,department,jobTitle,mail,officeLocation'
        filter="assignedLicenses/count ne 0 and userType eq 'Member'"
        count="true"
    }
    $headers = @{
        ConsistencyLevel="eventual"
    }


    $usersAL = New-Object System.Collections.ArrayList     
    $bufferSize = 10MB
    $size = 0        
    $page = 1
	$stop = $false
	$Progress = 0
	$perc = 0

    $response = Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/v1.0/users" -Body $body -Headers $headers
	$TotalRows = $response["@odata.count"]
    Write-Host "Total number of records found: $($response["@odata.count"])." 
    do
    {
		$Progress += 90
		$perc = ($Progress*100)/$TotalRows
		Write-Progress -Activity "Data received. Processing page no. [$page]" -PercentComplete $perc
		$page++

        foreach($user in $response.value) 
        {
            $newitem = [PSCustomObject]@{    
                UserPrincipalName		= $user.UserPrincipalName
                DisplayName             = $user.DisplayName
                City                    = $user.City
                Country                 = $user.Country
                Department              = $user.Department
                JobTitle                = $user.JobTitle
                Mail                    = $user.Mail
                OfficeLocation          = $user.OfficeLocation
                AssignedLicenses		= $user.AssignedLicenses
                AssignedPlans			= $user.AssignedPlans
                CreateDateTime			= $user.CreateDateTime
                LastAccess				= $user.SignInActivity.LastSignInDateTime
                UserID					= $user.Id
            }

            [void]$usersAL.Add($newitem)
            $size += [System.Text.Encoding]::UTF8.GetByteCount(($newitem | ConvertTo-Json -Depth 100))
            if ($size -gt $bufferSize)
            {
                $log_analytics_array = $usersAL.ToArray()
                Post-LogAnalyticsData -LogAnalyticsTableName $TableName -body $log_analytics_array
                $log_analytics_array = $null
                $usersAL.Clear()
                $size = 0
            }
        }
        if ($response["@odata.nextLink"] -ne $null)
        {
            $response = Invoke-MgGraphRequest -Method Get -Uri $response["@odata.nextLink"] #-Body $body -Headers $headers
        }
        else 
        {
            $stop = $true
			Write-Host "   Work completed!!! $TotalRows elements imported to Logs Analytics" -ForegroundColor Green
        } 

    } while (-not $stop)

    # Push remaining data to Log Analytics
    if ($usersAL.Count -gt 0)
    {
        $log_analytics_array = $usersAL.ToArray()
		Post-LogAnalyticsData -LogAnalyticsTableName $TableName -body $log_analytics_array			
    }
}
     
#Main Code - Run as required. Do ensure older table is deleted before creating the new table - as it will create duplicates.
Export-AzureADData 