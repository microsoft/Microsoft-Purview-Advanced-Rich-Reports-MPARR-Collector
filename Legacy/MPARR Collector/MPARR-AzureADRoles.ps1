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
Script      : Get-AzureADRoles.ps1
Author      : S. Zamorano
Version     : 1.0.0
Description : The script exports Azure AD Roles assigned from Microsoft Graph and pushes into a customer-specified Log Analytics table. Please note if you change the name of the table - you need to update Workbook sample that displays the report , appropriately. Do ensure the older table is deleted before creating the new table - it will create duplicates and Log analytics workspace doesn't support upserts or updates.
#>


param (
    # Log Analytics table where the data is written to. Log Analytics will add an _CL to this name.
    [string]$TableName = "AzureADRoles"

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
        Write-Information -MessageData "   $rows rows written to Log Analytics workspace $uri" -InformationAction Continue
    }

}

Function Export-AzureADRoles() {
    # ---------------------------------------------------------------   
    #    Name           : Export-AzureADRoles
    #    Desc           : Extracts data from Get-MgDirectoryRole into Log analytics workspace tables for reporting purposes
    #    Return         : None
    # ---------------------------------------------------------------
    
        Connect-MgGraph -CertificateThumbPrint $CertificateThumb -AppID $AppClientID -TenantId $TenantGUID
		#Select-MgProfile -Name "beta"
        # Run the commandlet to search through the Audit logs and get the AIP events in the specified timeframe
		$GetAzureADRole = Get-MgDirectoryRole -All -ExpandProperty Members| select DisplayName, Description, Id, Members
		$Max = $GetAzureADRole.Count

        # Status update
        $recordsCount = $GetAzureADRole.Count
        Write-Information -MessageData "   $recordsCount rows returned by Get-MgDirectoryRole" -InformationAction Continue

        # If there is no data, skip
        if ($GetAzureADRole.Count -eq 0) { continue; }

        # Else format for Log Analytics
		$P = 0
        $log_analytics_array = @()            
        foreach($i in $GetAzureADRole) {
			$UnusedAccountWarning = "OK"; $P++
			Write-Host ("Processing account {0} {1}/{2}" -f $i.DisplayName, $P, $Max)
            $newitem = [PSCustomObject]@{    
				DisplayName             = $i.DisplayName
                Description             = $i.Description
				Members                 = $i.Members
				RoleID					= $i.Id
            }
            $log_analytics_array += $newitem
        }

        # Push data to Log Analytics
        Post-LogAnalyticsData -LogAnalyticsTableName $TableName -body $log_analytics_array
    }
    
 
#Main Code - Run as required. Do ensure older table is deleted before creating the new table - as it will create duplicates.
Export-AzureADRoles 