# Input bindings are passed in via param block.
param($Timer, $myOutputBlob, $TriggerMetadata)

# Debug statement to check the output blob path
Write-Host "PowerShell Blob trigger function Processed blob Name: $($myOutputBlob)"

# Get the current universal time in the default string format.
$currentUTCtime = (Get-Date).ToUniversalTime()

# The 'IsPastDue' property is 'true' when the current function invocation is later than scheduled.
if ($Timer.IsPastDue) {
    Write-Host "PowerShell timer is running late!"
}

# Write an information log with the current time.
Write-Host "PowerShell timer trigger function ran! TIME: $currentUTCtime"

# Authenticate using the service principal
$tenantId = "2f250fde-b995-4c3b-a347-79dab0d3311b"
$clientId = "12b5f3d9-ba52-4505-be6b-307f58271e39"
$clientSecret = "<YOURSECRETHERE>"

try {
    $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" `
        -Method Post `
        -ContentType "application/x-www-form-urlencoded" `
        -Body @{
        client_id     = $clientId
        scope         = "https://analysis.windows.net/powerbi/api/.default"
        client_secret = $clientSecret
        grant_type    = "client_credentials"
    }

    $accessToken = $tokenResponse.access_token
    Write-Host "Successfully obtained access token."

    # Set the authorization header
    $headers = @{
        Authorization = "Bearer $accessToken"
    }

    # Input values before running the script:
    $NbrDaysDaysToExtract = 7    
    $ExportFileName = 'PBIActivityEvents'
    #--------------------------------------------

    # Start with yesterday for counting back to ensure full day results are obtained:
    [datetime]$DayUTC = (([datetime]::Today.ToUniversalTime()).Date).AddDays(-1)

    # Suffix for file name so we know when it was written:
    [string]$DateTimeFileWrittenUTCLabel = ([datetime]::Now.ToUniversalTime()).ToString("yyyyMMddHHmm")

    # Loop through each of the days to be extracted (<Initialize> ; <Condition> ; <Repeat>)
    For ($LoopNbr = 0 ; $LoopNbr -lt $NbrDaysDaysToExtract ; $LoopNbr++) {
        [datetime]$DateToExtractUTC = $DayUTC.AddDays(-$LoopNbr).ToString("yyyy-MM-dd")
        [string]$DateToExtractLabel = $DateToExtractUTC.ToString("yyyy-MM-dd")
        
        # Create full file name:
        [string]$FullExportFileName = $ExportFileName + '-' + ($DateToExtractLabel -replace '-', '') + '-' + $DateTimeFileWrittenUTCLabel + '.json'

        $continuationUri = $null
        $allEvents = @()

        do {
            # Obtain activity events and store intermediary results:
            $uri = "https://api.powerbi.com/v1.0/myorg/admin/activityevents?startDateTime='$($DateToExtractLabel+'T00:00:00.000Z')'&endDateTime='$($DateToExtractLabel+'T23:59:59.999Z')'" 
            if ($continuationUri) {
                $uri = $continuationUri               
                # $uri += "&continuationToken=$continuationToken"
            }

            Write-Host $uri

            $response = Invoke-RestMethod -Uri $uri `
                -Headers $headers `
                -Method Get

            $allEvents += $response.activityEventEntities

            $continuationUri = $response.continuationUri
        } while ($continuationUri)

        # Check if response contains events
        if ($allEvents.Count -gt 0) {
            Write-Host "Retrieved $($allEvents.Count) events for $DateToExtractLabel"
        }
        else {
            Write-Host "No events found for $DateToExtractLabel"
        }

        # Convert events to JSON
        $jsonContent = $allEvents | ConvertTo-Json 

        Write-Host "Output Blob Path: $myOutputBlob"

        # Write JSON content to the output blob
        Set-Content -Path $myOutputBlob -Value $jsonContent

        Write-Host "File uploaded: $FullExportFileName"
    }

    Write-Host "Extract of Power BI activity events is complete."
}
catch {
    Write-Host "Failed to connect to Power BI Service Account. Error: $_"
    throw
}