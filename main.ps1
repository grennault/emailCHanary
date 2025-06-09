# -------------------------------
# emailCHanary Azure Script
# Made by: G. Renault & K. Tamine
# Upload this azure powershell script in your Azure Portal terminal (Upload file) and execute it ./script.ps1
# -------------------------------

# Dot-source other scripts
. "$PSScriptRoot\helper.ps1"
. "$PSScriptRoot\CreateADUser.ps1"
. "$PSScriptRoot\CreateSharedMailbox.ps1"
. "$PSScriptRoot\SetupNotificationMethod.ps1"

# Function to print the email form with the Swiss flag
function Print-WelcomeMessage {
    $welcome = @"
+------------------------------------+
|                                    |
|           +--------+               |
|           |  ####  |               |
|           |  ####  |               |
|       +----------------+           |
|       |    ####  ####  |           |
|       |    ####  ####  |           |
|           |  ####  |               |
|           |  ####  |               |
|           +--------+               |
|                                    |
|         EMAIL FORM                 |
|                                    |
|   Name:                            |
|   Email:                           |
|   Subject:                         |
|                                    |
+------------------------------------+
"@
    Write-Host $welcome -ForegroundColor Red
    $welcome = @"
Welcome to the emailCHanaray script.

IMPORTANT:
This script only supports to be executed as a User in the Azure Portal Powershell terminal in a browser.
You need to have sufficient privilege to create shared mailbox (the canaray mailbox) and to add a forwarding rule to it (to some monitoring email addresses).

"@
    Write-Output $welcome
}

# Function to ask notification method
function Ask-NotificationMethod {
        Param
        (
            [Parameter(Mandatory=$true, Position=0)]
            [string] $SourceMailbox
        )
    
    do {
        Write-Output "How to notify of received emails on email caaray decoy?"
        Write-Output "1. By email (only one supported for now)"
        Write-Output "2. By Teams notification"
        Write-Output "3. By Slack notification"

        $option = Read-Host "Please select an option (1 or 2)"
        switch ($option) {
            "1" { 
                Write-Output "You selected: By email" 
                Setup-NotificationMethod $SourceMailbox
                return
            }
            "2" { 
                Write-Output "You selected: By Teams notification" 
                Write-Output "We apologize but this is not yet supported. Please write a Github Issue to describe your needs."
            }
            "3" { 
                Write-Output "You selected: By Slack notification" 
                Write-Output "We apologize but this is not yet supported. Please write a Github Issue to describe your needs."
            }
            default { Write-Output "Invalid selection. Please choose 1, 2 or 3." }
        }
    } while ($option -ne "1")
}

# Main function
function Main {
    Print-WelcomeMessage
    
    if ("User" -ne ((Get-AzContext).Account).Type) {
        Write-Host "Executed using Managed Identity. Switching as a User... Please complete below steps"
        Write-Host "Trying to login you with Connect-AzAccount command..." -ForegroundColor Yellow
        Connect-AzAccount -UseDeviceAuthentication
    }

    Write-Host "You are logged in as:" -ForegroundColor Green
    az ad signed-in-user show --query "userPrincipalName"
    Write-Host ""

    $UserPrincipalName = az ad signed-in-user show --query "userPrincipalName"
    
    Connect-AzureAD
    Connect-ExchangeOnline

    do {
        Write-Output "Should the email canary be a shared mailbox or an AD User (with Outlook Exchange Licence)"
        Write-Output "1. Canaray as a shared mailbox (could be detected by attacker)"
        Write-Output "2. Canary as an internal AD User (may imply some cost)"

        $option = Read-Host "Please select an option (1 or 2)"
        switch ($option) {
            "1" { 
                Write-Output "You selected: Canaray as a shared mailbox " 
                $SourceMailbox = Create-SharedMailbox
            }
            "2" { 
                Write-Output "You selected: Canary as an internal AD User" 
                Write-Output "Please note that this may induce some cost (cost of one user with Exchange Outlook Licence)"  -ForegroundColor Yellow
                $SourceMailbox = Create-ADUser
            }
            default { Write-Output "Invalid selection. Please choose 1 or 2." }
        }
    } while ($option -ne "1" -and $option -ne "2")

    Ask-NotificationMethod $SourceMailbox
}

try {
    Main
    Write-Host "Script succeeded! Canary mailbox and stealth forwarding rule set up." -ForegroundColor Green
} catch {
    Write-Host "Script failed. Please review the error below or create a Git Issue including this:" -ForegroundColor Red
    Write-Output "An error occurred: $($_.Exception.GetType().FullName)"
    Write-Host $_
    Write-Host $_.ScriptStackTrace
}