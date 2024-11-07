<#
.SYNOPSIS
Connects to Microsoft Graph to retrieve and optionally delete email messages from a specific sender in a user's inbox.

.DESCRIPTION
The script connects to Microsoft Graph Beta using client credentials to access a specified user's inbox. It retrieves messages directly using OData filters based on the sender's email address and optionally deletes them.

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

.PARAMETER DeleteMessages
Switch parameter to indicate if messages should be deleted.

.NOTES
Author: Jeffrey "Gig-E" Arsenault, TwistedLogic, LLC
Email: jeff@twistedlogic.io
Date: 18 Sep 2024
Version: 1.1

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
    [switch]$DeleteMessages
)

Import-Module Microsoft.Graph.Beta.Mail

# Convert client secret to secure string and create credential object
$ClientSecretPass = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $ClientSecretPass

# Connect to Microsoft Graph
Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential -NoWelcome

# Build the OData filter query
$filter = "from/emailAddress/address eq '$SenderEmail'"

# Get messages directly filtered by sender's email address
$messages = Get-MgBetaUserMessage -UserId $UserId -Filter $filter -All

if ($messages.Count -eq 0) {
    Write-Output "No messages found from sender '$SenderEmail' in user '$UserId' inbox."
} else {
    $count = 0
    foreach ($message in $messages) {
        $count++
        Write-Output "$count - $($message.SentDateTime) - Subject: $($message.Subject) - Sender: $($message.Sender.EmailAddress.Address)"
        if ($DeleteMessages.IsPresent) {
            try {
                Remove-MgBetaUserMessage -UserId $UserId -MessageId $message.Id -Confirm:$false
                Write-Output "Message ID $($message.Id) deleted."
            } catch {
                Write-Error "Failed to delete message ID $($message.Id): $_"
            }
        }
    }
    Write-Output "$count messages processed."
}

# Disconnect from Microsoft Graph
Disconnect-MgGraph
