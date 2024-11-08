<#
.SYNOPSIS
Connects to Microsoft Graph to retrieve and optionally delete email messages from a specific sender based on the sender's email address in the message headers, optimized to avoid throttling and appending output to a CSV file.

.DESCRIPTION
The script connects to Microsoft Graph Beta using client credentials to access a specified user's inbox. It retrieves messages in batches and includes retry logic with exponential backoff to handle throttling. It filters messages by examining the 'internetMessageHeaders' to find messages where the sender's email address matches the specified address. Optionally, it can delete those messages and appends the output to a CSV file.

.PARAMETER ClientId
Client ID used for the application to authenticate with Microsoft Graph.

.PARAMETER TenantId
Tenant ID corresponding to the Azure AD tenant.

.PARAMETER ClientSecret
Client secret for authenticating the application.

.PARAMETER UserId
The email address or User Principal Name (UPN) of the user whose inbox will be queried.

.PARAMETER SenderEmail
The email address of the sender whose messages will be retrieved and optionally deleted.

.PARAMETER OutputCsvPath
The file path of the CSV file where the output will be appended.

.PARAMETER DaysToSearch
The number of days to look back when searching for messages.

.PARAMETER DeleteMessages
Switch parameter to indicate if messages should be deleted.

.NOTES
Author: Jeffrey "Gig-E" Arsenault, TwistedLogic, LLC
Email: jeff@twistedlogic.io
Date: 18 Sep 2024
Version: 1.6

Before running, ensure you have the latest versions of the following modules installed:
Install-Module Microsoft.Graph -AllowClobber -Force
Install-Module Microsoft.Graph.Beta -AllowClobber -Force

Also requires an Entra App with the following:
- OAuth Token (Client secrets)
- API Permissions:
    - Mail.ReadWrite (Application)
    - User.Read.All (Application)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ClientId,
    [Parameter(Mandatory = $true)]
    [string]$TenantId,
    [Parameter(Mandatory = $true)]
    [string]$ClientSecret,
    [Parameter(Mandatory = $true)]
    [string]$UserId,
    [Parameter(Mandatory = $true)]
    [string]$SenderEmail,
    [Parameter(Mandatory = $true)]
    [string]$OutputCsvPath,
    [Parameter(Mandatory = $true)]
    [int]$DaysToSearch,
    [switch]$DeleteMessages
)

Import-Module Microsoft.Graph.Beta.Mail

# Convert client secret to secure string and create credential object
$ClientSecretPass = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $ClientSecretPass

# Connect to Microsoft Graph
Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential -NoWelcome

# Initialize variables for throttling
$maxRetries = 5
$retryDelay = 2 # Initial delay in seconds

# Function to handle API calls with retry logic
function Invoke-WithRetry {
    param (
        [ScriptBlock]$ScriptBlock
    )
    $retryCount = 0
    while ($true) {
        try {
            return & $ScriptBlock
        } catch {
            if ($_.Exception.Response.StatusCode -eq 429 -or $_.Exception.Response.StatusCode -eq 503) {
                $retryCount++
                if ($retryCount -gt $maxRetries) {
                    throw "Maximum retry attempts exceeded."
                }
                $retryAfter = $_.Exception.Response.Headers['Retry-After']
                if ($retryAfter) {
                    $sleepSeconds = [int]$retryAfter
                } else {
                    $sleepSeconds = $retryDelay * [math]::Pow(2, $retryCount)
                }
                Write-Warning "Throttled by server. Waiting $sleepSeconds seconds before retrying..."
                Start-Sleep -Seconds $sleepSeconds
            } else {
                throw $_
            }
        }
    }
}

# Specify the date range based on the DaysToSearch parameter
$startDateTime = (Get-Date).AddDays(-$DaysToSearch).ToString("o")
$filterDate = "receivedDateTime ge $startDateTime"

# Set batch size (number of messages per request)
$batchSize = 50

# Initialize pagination variables
$morePages = $true
$nextLink = $null

$count = 0
$totalProcessed = 0

# Prepare CSV output
$csvHeaders = @('UserId', 'MessageId', 'SentDateTime', 'Subject', 'SenderEmail', 'Deleted')
if (-not (Test-Path -Path $OutputCsvPath)) {
    # Create CSV file with headers if it doesn't exist
    $null = Out-File -FilePath $OutputCsvPath -Encoding UTF8 -Force
    $csvHeaders | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 | Out-File -FilePath $OutputCsvPath -Append -Encoding UTF8
}

Write-Output "Starting message retrieval and processing..."

while ($morePages) {
    # Build query parameters
    $queryParameters = @{
        'Filter' = $filterDate
        'Top'    = $batchSize
        'Select' = 'id,sentDateTime,subject,sender'
    }

    if ($nextLink) {
        # If there's a nextLink from previous response, use it
        $response = Invoke-WithRetry { Invoke-MgGraphRequest -Uri $nextLink -Method GET }
    } else {
        # Initial request
        $response = Invoke-WithRetry { Get-MgBetaUserMessage -UserId $UserId @queryParameters }
    }

    # Process messages
    foreach ($message in $response.Value) {
        # Retrieve the message with internetMessageHeaders using -Select
        $messageDetail = Invoke-WithRetry {
            Get-MgBetaUserMessage -UserId $UserId -MessageId $message.Id -Select 'internetMessageHeaders,sentDateTime,subject,sender'
        }

        # Initialize flag to check if sender's email is in headers
        $senderInHeaders = $false

        # Loop through headers to find the sender's email address
        foreach ($header in $messageDetail.InternetMessageHeaders) {
            if ($header.Name -match '^(From|Sender|Return-Path)$') {
                if ($header.Value -like "*$SenderEmail*") {
                    $senderInHeaders = $true
                    break
                }
            }
        }

        if ($senderInHeaders) {
            $count++
            Write-Output "$count - $($messageDetail.SentDateTime) - Subject: $($messageDetail.Subject) - Sender in Headers: $SenderEmail"
            $deletedStatus = $false
            if ($DeleteMessages.IsPresent) {
                try {
                    Invoke-WithRetry {
                        Remove-MgBetaUserMessage -UserId $UserId -MessageId $message.Id -Confirm:$false
                    }
                    Write-Output "Message ID $($message.Id) deleted."
                    $deletedStatus = $true
                } catch {
                    Write-Error "Failed to delete message ID $($message.Id): $_"
                }
            }

            # Prepare data for CSV
            $csvData = [PSCustomObject]@{
                UserId        = $UserId
                MessageId     = $message.Id
                SentDateTime  = $messageDetail.SentDateTime
                Subject       = $messageDetail.Subject
                SenderEmail   = $SenderEmail
                Deleted       = $deletedStatus
            }

            # Append data to CSV
            $csvData | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 | Out-File -FilePath $OutputCsvPath -Append -Encoding UTF8
        }

        # Optional: Include a short delay between message processing
        Start-Sleep -Milliseconds 200
    }

    $totalProcessed += $response.Value.Count

    # Check if there's a nextLink for pagination
    if ($response.'@odata.nextLink') {
        $nextLink = $response.'@odata.nextLink'
        Write-Output "Processed $totalProcessed messages so far. Continuing to next batch..."
    } else {
        $morePages = $false
    }
}

Write-Output "$count messages from sender '$SenderEmail' found in message headers and processed."
Write-Output "Total messages processed: $totalProcessed"

# Disconnect from Microsoft Graph
Disconnect-MgGraph