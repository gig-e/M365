<#
.FILE_NAME
Graph-Email-Delete.ps1

.SYNOPSIS
This script connects to Microsoft Graph to retrieve and optionally delete email messages from a specific sender in a user's inbox.

.DESCRIPTION
The script connects to Microsoft Graph Beta using client credentials to access a specified user's inbox and retrieve messages based on the sender's email address. 
It includes a function to display these messages and optionally delete them from the inbox.

.PARAMETER $ClientId
Client ID used for the application to authenticate with Microsoft Graph.

.PARAMETER $TenantId
Tenant ID corresponding to the Azure AD tenant.

.PARAMETER $ClientSecret
Client secret for authenticating the application.

.PARAMETER $UserId
The email address of the user whose inbox will be queried.

.PARAMETER $SenderEmail
The email address of the sender whose messages will be retrieved and optionally deleted.

.NOTES
Author: Jeffrey "Gig-E" Arsenault, TwistedLogic, LLC
Email: jeff@twistedlogic.io
Date: 18 Sep 2024
Version: 1.0

Before running be sure to have the latest of the following modules installed:
Install-Module Microsoft.Graph -AllowClobber -Force
Install-Module Microsoft.Graph.Beta -AllowClobber -Force

Also requires a Entra App with the following:
- OAUTH Token (Client secrets)
- API Permissions
    - Mail.ReadWrite (Delegated) - Need to test without
    - Mail.ReadWrite (Application)
    - User.Read (Delegated)
    - User.Read.All (Application) - Need to test without
#>

Import-Module Microsoft.Graph.Beta.Mail
$ClientId = ""
$TenantId = ""
$ClientSecret = ""
$ClientSecretPass = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $ClientSecretPass
Connect-MgGraph -TenantId $tenantId -ClientSecretCredential $ClientSecretCredential -NoWelcome
$results = 0
# Inbox you want to clean
$userId = ""
# Sender of the email you want to delete
$senderAddress = ""
$filter = $senderAddress
# Get all messages from $userId that contain the sender's address
$messages = Get-MgBetaUserMessage -All -UserId $userId -Search $filter

function Get-MessagesFromSender {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Messages, # The array of messages to search through
        [Parameter(Mandatory = $true)]
        [string]$SenderEmail, # The email address to match
        [Parameter(Mandatory = $false)]
        [string]$UserId # Optional user ID, required if you want to delete messages
    )
    $count = 0
    foreach ($message in $Messages) {
        if (($message.Sender.EmailAddress.Address) -eq $SenderEmail) {            
            $count++
            Write-Output "$count - $($message.SentDateTime) - Message Subject: $($message.Subject) - Message Sender Email: $($message.Sender.EmailAddress.Address)"
            # Uncomment the line below to remove messages if needed
            # Remove-MgBetaUserMessage -UserId $UserId -MessageId $message.Id
        }
    }
}

# Filter the results to just emails from that sender (optionaly remove them by uncommenting line in funtion)
Get-MessagesFromSender -Messages $messages -SenderEmail $senderAddress -UserId $userId