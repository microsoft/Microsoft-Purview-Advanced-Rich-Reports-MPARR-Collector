<#
.NOTES
# MPARRUtils module
# v.1.0 2023-04-28
	01-02-2024		G.Berdzik	:	First release for MPARR Solution
#>

enum HubTier {
    Basic = 256KB
    Standard = 1MB
    Premium = 1MB
    Dedicated = 1MB
}


class MPARREventHub
{
    hidden [string] $AuthURL
    hidden [string] $AccessToken
    [string] $TenantID
    [string] $HubNamespace 
    [string] $HubQueue
    [string] $ClientID
    [string]$ClientSecret
    [HubTier] $Tier = [HubTier]::Basic


    # constructor
    MPARREventHub([string] $TenantID, [string] $HubNamespace, [string] $HubName, [string] $ClientID, [string]$ClientSecret)
    {
        $this.TenantID = $TenantID
        $this.HubNamespace = $HubNamespace
        $this.HubQueue = $HubName
        $this.ClientID = $ClientID
        $this.ClientSecret = $ClientSecret

        $this.AccessToken = $this.Authenticate()
    }

    # method for authentication
    hidden [string] Authenticate()
    {
        $this.AuthURL = "https://login.microsoftonline.com/$($this.TenantID)/oauth2/token"
        $headers = @{
            "Content-Type" = "application/x-www-form-urlencoded"
        }
        $body = @{
            grant_type = "client_credentials"
            client_id = $this.ClientID
            client_secret = $this.ClientSecret
            resource = "https://eventhubs.azure.net"
        }
        try 
        {
            $response = Invoke-RestMethod -Method Post -Uri ($this.AuthURL) -Headers $headers -Body $body
        }
        catch
        {
            $err = $_.Exception.Message
            if ($_.ErrorDetails.Message -ne $null)
            {
                $err += "`n" + ($_.ErrorDetails.Message | ConvertFrom-Json -Depth 10).error_description
            }
            Write-Host "Error authentication to '$($this.ClientID)' app." -ForegroundColor Red
            Write-Host $err -ForegroundColor Red
            Write-Host "Exiting..." -ForegroundColor Red
            exit(1)
        }
        return $response.access_token
    }

    # method to send single batch to Event Hub
    hidden [bool] SendBatchToEventHub($message) 
    {
        $body = $message
        $txt = $body.ToString()
        $size = [System.Text.Encoding]::UTF8.GetByteCount($txt) 
        if ($size -gt $this.Tier.value__)
        {
            Write-Host "Message size too big." -ForegroundColor Yellow
            return $false
        }
        $URL = "https://$($this.HubNamespace).servicebus.windows.net/$($this.HubQueue)/messages"
        $headers = @{
            Authorization = "Bearer $($this.AccessToken)"
            "Content-Type" = "application/atom+xml;type=entry;charset=utf-8"
        }

        $status = ""
        Invoke-RestMethod -Method Post -Uri $URL -Headers $headers -Body $body -StatusCodeVariable status
        if ($status -ne 201)
        {
            Write-Host "Error sending data to the Event Hub: $status."
            return $false
        }

        return $true
    }


    # method to send multiple batches to Event Hub
    [bool] PublishToEventHub($messages, $ErrorLogName)
    {
        Write-Host "Starting export to Event Hub of multiple batches..."
        $currentBatch = New-Object System.Collections.ArrayList

        $count = 0
        $currentBatchSize = 0

		foreach ($message in $messages) 
        {
            $message | Add-Member -MemberType NoteProperty -Name "EventCreationTime" -Value ($message.CreationTime)
            $count++

            $txt = ($message | ConvertTo-Json -Depth 100).ToString()
            $messageSize = [System.Text.Encoding]::UTF8.GetByteCount($txt)
			
			if (($currentBatchSize + $messageSize)*1.1 -gt $this.Tier.value__) 
            {
                $body = $currentBatch | ConvertTo-Json -Depth 100
                $result = $this.SendBatchToEventHub($body)
                if (-not $result)
                {
                    # error
                    Write-Host "Error exporting to Event Hub. Source data written to $ErrorLogName file." -ForegroundColor Red
                    $body | Out-File -FilePath $ErrorLogName -Append
                }
                $currentBatch.Clear()
                $currentBatch.TrimToSize()            
                $currentBatchSize = 0
                Write-Host "Batch written to the $($this.HubQueue)."
            }

            [void]$currentBatch.Add($message)
            $currentBatchSize += $messageSize
        }

        if ($currentBatch.Count -gt 0) {
            $body = $currentBatch | ConvertTo-Json -Depth 100
            $result = $this.SendBatchToEventHub($body)
            if (-not $result)
            {
                # error
                Write-Host "Error exporting to Event Hub. Source data written to $ErrorLogName file." -ForegroundColor Red
                $body | Out-File -FilePath $ErrorLogName -Append
        }
        }
        
        Write-Host "$count elements exported."

        return $true
    }

}
