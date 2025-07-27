# -------------------------------
# emailCHanary Azure Script
# Made by: G. Renault & K. Tamine
# Upload this azure powershell script in your Azure Portal terminal (Upload file) and execute it ./main.ps1
# -------------------------------

# Dot-source other scripts
. "$PSScriptRoot\common.ps1"
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
Welcome to the emailCHanary script.

IMPORTANT:
This script only supports to be executed as a User in the Azure Portal Powershell terminal in a browser.
You need to have sufficient privilege to create shared mailbox (the canary mailbox) and to add a forwarding rule to it (to some monitoring email addresses).

"@
    Write-Output $welcome
}

# Function to ask notification method
function Ask-NotificationMethod {
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [string] $SourceMailbox
    )
    
    do {
        Write-Log -Message "How to notify of received emails on email canary decoy?" -Level "Info"
        Write-Host "1. By email"
        Write-Host "2. By Teams notification"
        Write-Host "3. By Slack notification"

        $option = Read-Host "Please select an option (1, 2, or 3)"
        switch ($option) {
            "1" {
                Write-Log -Message "You selected: By email" -Level "Success"
                Setup-NotificationMethod $SourceMailbox
                return
            }
            "2" {
                Write-Log -Message "You selected: By Teams notification" -Level "Success"
                # Check if Exchange Online Management module is installed
                if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
                    Write-Log -Message "Exchange Online Management module is required for Teams notifications." -Level "Warning"
                    $installModule = Read-Host "Do you want to install the Exchange Online Management module? (y/n)"
                    if ($installModule -eq "y") {
                        try {
                            Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
                            Write-Log -Message "Exchange Online Management module installed successfully." -Level "Success"
                        } catch {
                            Write-Log -Message "Failed to install Exchange Online Management module: $_" -Level "Error"
                            $tryEmail = Read-Host "Would you like to use email notification instead? (y/n)"
                            if ($tryEmail -eq "y") {
                                Write-Log -Message "Switching to email notification." -Level "Info"
                                Setup-NotificationMethod $SourceMailbox
                                return
                            }
                            return
                        }
                    } else {
                        $tryEmail = Read-Host "Would you like to use email notification instead? (y/n)"
                        if ($tryEmail -eq "y") {
                            Write-Log -Message "Switching to email notification." -Level "Info"
                            Setup-NotificationMethod $SourceMailbox
                            return
                        }
                        return
                    }
                }
                
                # Call the Teams notification setup function
                Setup-TeamsNotification $SourceMailbox
                return
            }
            "3" {
                Write-Log -Message "You selected: By Slack notification" -Level "Success"
                # Check if Exchange Online Management module is installed
                if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
                    Write-Log -Message "Exchange Online Management module is required for Slack notifications." -Level "Warning"
                    $installModule = Read-Host "Do you want to install the Exchange Online Management module? (y/n)"
                    if ($installModule -eq "y") {
                        try {
                            Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
                            Write-Log -Message "Exchange Online Management module installed successfully." -Level "Success"
                        } catch {
                            Write-Log -Message "Failed to install Exchange Online Management module: $_" -Level "Error"
                            $tryEmail = Read-Host "Would you like to use email notification instead? (y/n)"
                            if ($tryEmail -eq "y") {
                                Write-Log -Message "Switching to email notification." -Level "Info"
                                Setup-NotificationMethod $SourceMailbox
                                return
                            }
                            return
                        }
                    } else {
                        $tryEmail = Read-Host "Would you like to use email notification instead? (y/n)"
                        if ($tryEmail -eq "y") {
                            Write-Log -Message "Switching to email notification." -Level "Info"
                            Setup-NotificationMethod $SourceMailbox
                            return
                        }
                        return
                    }
                }
                
                # Call the Slack notification setup function
                Setup-SlackNotification $SourceMailbox
                return
            }
            default { Write-Log -Message "Invalid selection. Please choose 1, 2, or 3." -Level "Warning" }
        }
        
        # If we get here, the user selected an unsupported option and didn't want to use email
        $retry = Read-Host "Do you want to try again? (y/n)"
        if ($retry -ne "y") {
            Write-Log -Message "Notification setup cancelled." -Level "Warning"
            return
        }
    } while ($true)
}

# Main function
function Main {
    Print-WelcomeMessage
    
    # Check and install required modules
    $modulesOk = $true
    $modulesOk = $modulesOk -and (Ensure-ModuleInstalled -ModuleName "Az" -MinimumVersion "9.0.0")
    $modulesOk = $modulesOk -and (Ensure-ModuleInstalled -ModuleName "AzureAD" -MinimumVersion "2.0.0")
    $modulesOk = $modulesOk -and (Ensure-ModuleInstalled -ModuleName "ExchangeOnlineManagement" -MinimumVersion "3.0.0")
    
    if (-not $modulesOk) {
        Write-Log -Message "One or more required modules could not be installed. The script may not function correctly." -Level "Error"
        $continue = Read-Host "Do you want to continue anyway? (y/n)"
        if ($continue -ne "y") {
            Write-Log -Message "Script execution cancelled." -Level "Warning"
            return
        }
    }
    
    # Check authentication
    if (-not (Ensure-ServiceConnected -ServiceName "Azure")) {
        Write-Log -Message "Cannot proceed without Azure authentication." -Level "Error"
        return
    }
    
    # Get and display current user
    $currentUser = Get-CurrentUserPrincipalName
    if ($currentUser) {
        Write-Log -Message "You are logged in as: $currentUser" -Level "Success"
    } else {
        Write-Log -Message "Could not determine current user." -Level "Warning"
    }
    
    # Connect to AzureAD
    if (-not (Ensure-ServiceConnected -ServiceName "AzureAD")) {
        Write-Log -Message "Cannot proceed without AzureAD authentication." -Level "Error"
        return
    }
    
    # Connect to Exchange Online
    if (-not (Ensure-ServiceConnected -ServiceName "ExchangeOnline")) {
        Write-Log -Message "Cannot proceed without Exchange Online authentication." -Level "Error"
        return
    }
    
    # Ask user for canary type
    do {
        Write-Log -Message "Should the email canary be a shared mailbox or an AD User (with Outlook Exchange Licence)" -Level "Info"
        Write-Host "1. Canary as a shared mailbox (could be detected by attacker)"
        Write-Host "2. Canary as an internal AD User (may imply some cost)"

        $option = Read-Host "Please select an option (1 or 2)"
        switch ($option) {
            "1" { 
                Write-Log -Message "You selected: Canary as a shared mailbox" -Level "Success"
                $SourceMailbox = Create-SharedMailbox
                if (-not $SourceMailbox) {
                    Write-Log -Message "Failed to create shared mailbox. Cannot continue." -Level "Error"
                    return
                }
            }
            "2" { 
                Write-Log -Message "You selected: Canary as an internal AD User" -Level "Success"
                Write-Log -Message "Please note that this may induce some cost (cost of one user with Exchange Outlook Licence)" -Level "Warning"
                $SourceMailbox = Create-ADUser
                if (-not $SourceMailbox) {
                    Write-Log -Message "Failed to create AD user. Cannot continue." -Level "Error"
                    return
                }
            }
            default { Write-Log -Message "Invalid selection. Please choose 1 or 2." -Level "Warning" }
        }
    } while ($option -ne "1" -and $option -ne "2")

    Ask-NotificationMethod $SourceMailbox
}

try {
    Main
    Write-Log -Message "Script succeeded! Canary mailbox and stealth forwarding rule set up." -Level "Success"
} catch {
    Write-Log -Message "Script failed. Please review the error below or create a Git Issue including this:" -Level "Error"
    Write-Log -Message "An error occurred: $($_.Exception.GetType().FullName)" -Level "Error"
    Write-Host $_
    Write-Host $_.ScriptStackTrace
}