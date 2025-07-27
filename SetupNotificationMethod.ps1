# -------------------------------
# emailCHanary Notification Setup
# Made by: G. Renault & K. Tamine
# -------------------------------

# Setup the forwarding rule to notify for a received email on the canary email
function Setup-NotificationMethod {
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [string] $SourceMailbox
    )
    
    Write-Log -Message "Starting notification setup process..." -Level "Info"
    
    # Ensure we have the required permissions
    if (-not (Test-RequiredPermissions -PermissionType "CreateDistributionGroup")) {
        Write-Log -Message "You may not have sufficient permissions to create distribution groups." -Level "Warning"
        $continue = Read-Host "Do you want to continue anyway? (y/n)"
        if ($continue -ne "y") {
            Write-Log -Message "Notification setup cancelled." -Level "Warning"
            return
        }
    }

    # Step 1: Verify source mailbox
    Write-Log -Message "=== STEP 1: Source Mailbox ===" -Level "Info"
    
    # Use the provided source mailbox instead of asking again
    Write-Log -Message "Using provided source mailbox: $SourceMailbox" -Level "Info"
    
    # Check if mailbox exists in the organization
    try {
        $mailbox = Get-Mailbox -Identity $SourceMailbox -ErrorAction Stop
        
        Write-Log -Message "Found mailbox:" -Level "Success"
        Write-Host "Name: $($mailbox.DisplayName)"
        Write-Host "Email: $($mailbox.PrimarySmtpAddress)"
        Write-Host "Type: $($mailbox.RecipientTypeDetails)"
        
        $continue = Read-Host "Is this the correct mailbox? (y/n)"
        if ($continue -ne 'y') {
            # If the provided mailbox is not correct, ask for the correct one
            do {
                $SourceMailbox = Get-ValidEmail -PromptMessage "Enter the INTERNAL email address to forward FROM"
                
                try {
                    $mailbox = Get-Mailbox -Identity $SourceMailbox -ErrorAction Stop
                    
                    Write-Log -Message "Found mailbox:" -Level "Success"
                    Write-Host "Name: $($mailbox.DisplayName)"
                    Write-Host "Email: $($mailbox.PrimarySmtpAddress)"
                    Write-Host "Type: $($mailbox.RecipientTypeDetails)"
                    
                    $continue = Read-Host "Is this the correct mailbox? (y/n)"
                    if ($continue -eq 'y') {
                        break
                    }
                } catch {
                    Write-Log -Message "Could not find mailbox '$SourceMailbox'" -Level "Error"
                    Write-Log -Message "Please check the spelling and try again." -Level "Warning"
                }
            } while ($true)
        }
    } catch {
        Write-Log -Message "Error finding mailbox: $_" -Level "Error"
        
        # If the provided mailbox doesn't exist, ask for a valid one
        do {
            $SourceMailbox = Get-ValidEmail -PromptMessage "Enter the INTERNAL email address to forward FROM"
            
            try {
                $mailbox = Get-Mailbox -Identity $SourceMailbox -ErrorAction Stop
                
                Write-Log -Message "Found mailbox:" -Level "Success"
                Write-Host "Name: $($mailbox.DisplayName)"
                Write-Host "Email: $($mailbox.PrimarySmtpAddress)"
                Write-Host "Type: $($mailbox.RecipientTypeDetails)"
                
                $continue = Read-Host "Is this the correct mailbox? (y/n)"
                if ($continue -eq 'y') {
                    break
                }
            } catch {
                Write-Log -Message "Could not find mailbox '$SourceMailbox'" -Level "Error"
                Write-Log -Message "Please check the spelling and try again." -Level "Warning"
            }
        } while ($true)
    }

    # Step 2: Get tenant domain for distribution group
    Write-Log -Message "=== STEP 2: Tenant Domain ===" -Level "Info"
    
    # Get available domains for the distribution group
    $tenantDomains = Get-AzureADDomain | Where-Object { $_.IsDefault -eq $true }
    $defaultDomain = $tenantDomains[0].Name
    
    Write-Log -Message "Default tenant domain: $defaultDomain" -Level "Info"
    $useDifferentDomain = Read-Host "Do you want to use a different domain for the distribution group? (y/n)"
    
    $distributionDomain = $defaultDomain
    if ($useDifferentDomain -eq "y") {
        $selectedDomain = Get-AvailableDomains
        if ($selectedDomain) {
            $distributionDomain = $selectedDomain
        } else {
            Write-Log -Message "Using default domain: $defaultDomain" -Level "Info"
        }
    }

    # Step 3: Create distribution group
    Write-Log -Message "=== STEP 3: Distribution Group ===" -Level "Info"
    do {
        $GroupName = Get-ValidInput -PromptMessage "Enter a name for the distribution group (e.g., ForwardingGroup)" -MaxLength 64 -ErrorMessage "Invalid group name."
        $GroupDisplayName = Get-ValidInput -PromptMessage "Enter a display name for the group (e.g., Forwarding Group)" -MaxLength 64 -ErrorMessage "Invalid display name."
        $GroupEmail = "$GroupName@$distributionDomain"
        
        # Validate the constructed email address
        if ($GroupEmail.Length -gt 256) {
            Write-Log -Message "The email address $GroupEmail exceeds the maximum length of 256 characters." -Level "Error"
            continue
        }
        
        Write-Log -Message "Creating distribution group:" -Level "Info"
        Write-Host "Name: $GroupName"
        Write-Host "Display Name: $GroupDisplayName"
        Write-Host "Email: $GroupEmail"
        
        $continue = Read-Host "Create this distribution group? (y/n)"
        if ($continue -eq 'y') {
            # Check if group already exists
            if (Test-DistributionGroupExists -GroupName $GroupName) {
                Write-Log -Message "A distribution group with the name '$GroupName' already exists." -Level "Warning"
                $useExisting = Read-Host "Do you want to use the existing group? (y/n)"
                if ($useExisting -eq 'y') {
                    Write-Log -Message "Using existing distribution group." -Level "Success"
                    break
                } else {
                    continue
                }
            }
            
            try {
                New-DistributionGroup -Name $GroupName -DisplayName $GroupDisplayName -PrimarySmtpAddress $GroupEmail -ErrorAction Stop
                Write-Log -Message "Distribution group created successfully!" -Level "Success"
                break
            } catch {
                Write-Log -Message "Error creating distribution group: $_" -Level "Error"
                
                # Provide guidance based on common errors
                if ($_.Exception.Message -like "*already exists*") {
                    Write-Log -Message "A group with this name or email already exists. Please try a different name." -Level "Warning"
                } elseif ($_.Exception.Message -like "*permission*") {
                    Write-Log -Message "You may not have sufficient permissions to create distribution groups." -Level "Warning"
                }
                
                Write-Log -Message "Please try again with a different name." -Level "Warning"
            }
        }
    } while ($true)

    # Step 4: Create mail contact
    Write-Log -Message "=== STEP 4: External Contact ===" -Level "Info"
    do {
        $ContactName = Get-ValidInput -PromptMessage "Enter a display name for the contact (e.g., Your Gmail)" -MaxLength 64 -ErrorMessage "Invalid contact name."
        $ExternalEmail = Get-ValidEmail -PromptMessage "Enter the external email address to forward TO"
        
        Write-Log -Message "Creating mail contact:" -Level "Info"
        Write-Host "Name: $ContactName"
        Write-Host "Email: $ExternalEmail"
        
        $continue = Read-Host "Create this mail contact? (y/n)"
        if ($continue -eq 'y') {
            # Check if contact already exists
            if (Test-MailContactExists -ExternalEmailAddress $ExternalEmail) {
                Write-Log -Message "A mail contact with the email '$ExternalEmail' already exists." -Level "Warning"
                $useExisting = Read-Host "Do you want to use the existing contact? (y/n)"
                if ($useExisting -eq 'y') {
                    Write-Log -Message "Using existing mail contact." -Level "Success"
                    break
                } else {
                    continue
                }
            }
            
            try {
                New-MailContact -Name $ContactName -ExternalEmailAddress $ExternalEmail -ErrorAction Stop
                Write-Log -Message "Mail contact created successfully!" -Level "Success"
                break
            } catch {
                Write-Log -Message "Error creating mail contact: $_" -Level "Error"
                
                # Provide guidance based on common errors
                if ($_.Exception.Message -like "*already exists*") {
                    Write-Log -Message "A contact with this name or email already exists. Please try a different name or email." -Level "Warning"
                } elseif ($_.Exception.Message -like "*permission*") {
                    Write-Log -Message "You may not have sufficient permissions to create mail contacts." -Level "Warning"
                }
                
                Write-Log -Message "Please try again with different details." -Level "Warning"
            }
        }
    } while ($true)

    # Step 5: Add contact to group
    Write-Log -Message "=== STEP 5: Add Contact to Group ===" -Level "Info"
    Write-Log -Message "Adding contact '$ContactName' to group '$GroupName'..." -Level "Info"
    
    try {
        Add-DistributionGroupMember -Identity $GroupName -Member $ExternalEmail -ErrorAction Stop
        Write-Log -Message "Contact added to group successfully!" -Level "Success"
    } catch {
        Write-Log -Message "Error adding contact to group: $_" -Level "Error"
        
        # Provide guidance based on common errors
        if ($_.Exception.Message -like "*already a member*") {
            Write-Log -Message "This contact is already a member of the group." -Level "Warning"
        } else {
            $retry = Read-Host "Do you want to retry adding the contact to the group? (y/n)"
            if ($retry -eq 'y') {
                try {
                    Add-DistributionGroupMember -Identity $GroupName -Member $ExternalEmail -ErrorAction Stop
                    Write-Log -Message "Contact added to group successfully on retry!" -Level "Success"
                } catch {
                    Write-Log -Message "Error adding contact to group on retry: $_" -Level "Error"
                    Write-Log -Message "Continuing with setup, but forwarding may not work correctly." -Level "Warning"
                }
            } else {
                Write-Log -Message "Continuing with setup, but forwarding may not work correctly." -Level "Warning"
            }
        }
    }

    # Step 6: Set up forwarding
    Write-Log -Message "=== STEP 6: Set Up Forwarding ===" -Level "Info"
    Write-Log -Message "Final settings:" -Level "Info"
    Write-Host "From: $SourceMailbox"
    Write-Host "To Group: $GroupName ($GroupEmail)"
    Write-Host "Contact: $ContactName ($ExternalEmail)"
    
    $continue = Read-Host "Set up forwarding with these settings? (y/n)"
    if ($continue -eq 'y') {
        try {
            Set-Mailbox -Identity $SourceMailbox -DeliverToMailboxAndForward $true -ForwardingAddress $GroupName -ErrorAction Stop
            
            # Verify forwarding was set up correctly
            $verifyMailbox = Get-Mailbox -Identity $SourceMailbox | Select-Object Name, ForwardingAddress, DeliverToMailboxAndForward
            
            if ($verifyMailbox.ForwardingAddress -eq $GroupName -and $verifyMailbox.DeliverToMailboxAndForward -eq $true) {
                Write-Log -Message "Forwarding set up successfully!" -Level "Success"
                Write-Log -Message "Final forwarding settings:" -Level "Info"
                Write-Host "Name: $($verifyMailbox.Name)"
                Write-Host "Forwarding Address: $($verifyMailbox.ForwardingAddress)"
                Write-Host "Deliver To Mailbox And Forward: $($verifyMailbox.DeliverToMailboxAndForward)"
            } else {
                Write-Log -Message "Forwarding may not have been set up correctly. Please verify manually." -Level "Warning"
            }
        } catch {
            Write-Log -Message "Error setting up forwarding: $_" -Level "Error"
            
            # Provide guidance based on common errors
            if ($_.Exception.Message -like "*permission*") {
                Write-Log -Message "You may not have sufficient permissions to set up forwarding." -Level "Warning"
            }
        }
    }
    
    # Ask if user wants to disconnect from Exchange Online
    $disconnect = Read-Host "Do you want to disconnect from Exchange Online? (y/n)"
    if ($disconnect -eq 'y') {
        try {
            Disconnect-ExchangeOnline -Confirm:$false
            Write-Log -Message "Disconnected from Exchange Online." -Level "Success"
        } catch {
            Write-Log -Message "Error disconnecting from Exchange Online: $_" -Level "Error"
        }
    } else {
        Write-Log -Message "Remaining connected to Exchange Online." -Level "Info"
    }
    
    Write-Log -Message "Notification setup completed." -Level "Success"
}

# Import common functions and helper functions
. "$PSScriptRoot\common.ps1"
. "$PSScriptRoot\helper.ps1"

# Setup Slack notification for a received email on the canary email
function Setup-SlackNotification {
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [string] $SourceMailbox
    )
    
    Write-Log -Message "Starting Slack notification setup process..." -Level "Info"
    
    # Ensure we have the required permissions
    if (-not (Test-RequiredPermissions -PermissionType "CreateDistributionGroup")) {
        Write-Log -Message "You may not have sufficient permissions to create mail flow rules." -Level "Warning"
        $continue = Read-Host "Do you want to continue anyway? (y/n)"
        if ($continue -ne "y") {
            Write-Log -Message "Slack notification setup cancelled." -Level "Warning"
            return
        }
    }

    # Step 1: Verify source mailbox
    Write-Log -Message "=== STEP 1: Source Mailbox ===" -Level "Info"
    
    # Use the provided source mailbox instead of asking again
    Write-Log -Message "Using provided source mailbox: $SourceMailbox" -Level "Info"
    
    # Check if mailbox exists in the organization
    try {
        $mailbox = Get-Mailbox -Identity $SourceMailbox -ErrorAction Stop
        
        Write-Log -Message "Found mailbox:" -Level "Success"
        Write-Host "Name: $($mailbox.DisplayName)"
        Write-Host "Email: $($mailbox.PrimarySmtpAddress)"
        Write-Host "Type: $($mailbox.RecipientTypeDetails)"
        
        $continue = Read-Host "Is this the correct mailbox? (y/n)"
        if ($continue -ne 'y') {
            # If the provided mailbox is not correct, ask for the correct one
            do {
                $SourceMailbox = Get-ValidEmail -PromptMessage "Enter the INTERNAL email address to monitor"
                
                try {
                    $mailbox = Get-Mailbox -Identity $SourceMailbox -ErrorAction Stop
                    
                    Write-Log -Message "Found mailbox:" -Level "Success"
                    Write-Host "Name: $($mailbox.DisplayName)"
                    Write-Host "Email: $($mailbox.PrimarySmtpAddress)"
                    Write-Host "Type: $($mailbox.RecipientTypeDetails)"
                    
                    $continue = Read-Host "Is this the correct mailbox? (y/n)"
                    if ($continue -eq 'y') {
                        break
                    }
                } catch {
                    Write-Log -Message "Could not find mailbox '$SourceMailbox'" -Level "Error"
                    Write-Log -Message "Please check the spelling and try again." -Level "Warning"
                }
            } while ($true)
        }
    } catch {
        Write-Log -Message "Error finding mailbox: $_" -Level "Error"
        
        # If the provided mailbox doesn't exist, ask for a valid one
        do {
            $SourceMailbox = Get-ValidEmail -PromptMessage "Enter the INTERNAL email address to monitor"
            
            try {
                $mailbox = Get-Mailbox -Identity $SourceMailbox -ErrorAction Stop
                
                Write-Log -Message "Found mailbox:" -Level "Success"
                Write-Host "Name: $($mailbox.DisplayName)"
                Write-Host "Email: $($mailbox.PrimarySmtpAddress)"
                Write-Host "Type: $($mailbox.RecipientTypeDetails)"
                
                $continue = Read-Host "Is this the correct mailbox? (y/n)"
                if ($continue -eq 'y') {
                    break
                }
            } catch {
                Write-Log -Message "Could not find mailbox '$SourceMailbox'" -Level "Error"
                Write-Log -Message "Please check the spelling and try again." -Level "Warning"
            }
        } while ($true)
    }

    # Step 2: Get Slack webhook URL
    Write-Log -Message "=== STEP 2: Slack Webhook URL ===" -Level "Info"
    Write-Log -Message "You need to provide a Slack webhook URL to receive notifications." -Level "Info"
    Write-Log -Message "To create a webhook URL in Slack:" -Level "Info"
    Write-Host "1. Go to https://api.slack.com/apps and sign in"
    Write-Host "2. Click 'Create New App' and select 'From scratch'"
    Write-Host "3. Give your app a name (e.g., 'Email Canary Alert') and select your workspace"
    Write-Host "4. Click 'Create App'"
    Write-Host "5. In the left sidebar, click on 'Incoming Webhooks'"
    Write-Host "6. Toggle 'Activate Incoming Webhooks' to On"
    Write-Host "7. Click 'Add New Webhook to Workspace'"
    Write-Host "8. Select the channel where you want to receive notifications and click 'Allow'"
    Write-Host "9. Copy the webhook URL provided"
    
    do {
        $webhookUrl = Read-Host "Enter the Slack webhook URL"
        
        # Validate webhook URL format
        if ($webhookUrl -match "^https://hooks\.slack\.com/services/T[A-Z0-9]+/B[A-Z0-9]+/[a-zA-Z0-9]+$") {
            Write-Log -Message "Webhook URL format appears valid." -Level "Success"
            break
        } else {
            Write-Log -Message "Invalid webhook URL format. Please enter a valid Slack webhook URL." -Level "Warning"
            $retry = Read-Host "Do you want to try again? (y/n)"
            if ($retry -ne "y") {
                Write-Log -Message "Slack notification setup cancelled." -Level "Warning"
                return
            }
        }
    } while ($true)
    
    # Step 3: Create a unique rule name
    $ruleName = "EmailCanary-Slack-Alert-$(Get-Random -Minimum 1000 -Maximum 9999)"
    Write-Log -Message "=== STEP 3: Create Mail Flow Rule ===" -Level "Info"
    Write-Log -Message "Creating mail flow rule with name: $ruleName" -Level "Info"
    
    # Step 4: Create the mail flow rule
    try {
        # Check if a similar rule already exists
        $existingRule = Get-TransportRule | Where-Object { $_.Name -like "EmailCanary-Slack-Alert-*" }
        
        if ($existingRule) {
            Write-Log -Message "An existing email canary alert rule was found: $($existingRule.Name)" -Level "Warning"
            $useExisting = Read-Host "Do you want to update the existing rule? (y/n)"
            
            if ($useExisting -eq "y") {
                # Remove the existing rule
                Remove-TransportRule -Identity $existingRule.Name -Confirm:$false
                Write-Log -Message "Removed existing rule: $($existingRule.Name)" -Level "Success"
            }
        }
        
        # Create a new rule
        $ruleParams = @{
            Name = $ruleName
            SentTo = $SourceMailbox
            SubjectOrBodyContainsWords = "*"  # Match any content
            SetHeaderName = "X-MS-Exchange-Organization-EmailCanaryAlert"
            SetHeaderValue = "True"
            ExceptIfHeaderContainsMessageHeader = "X-MS-Exchange-Organization-EmailCanaryAlert"
            ExceptIfHeaderContainsWords = "True"
            Priority = 0  # High priority
        }
        
        # Create the rule
        New-TransportRule @ruleParams | Out-Null
        
        Write-Log -Message "Mail flow rule created successfully!" -Level "Success"
        
        # Now create a second rule to send the notification to Slack
        $notificationRuleName = "EmailCanary-Slack-Notification-$(Get-Random -Minimum 1000 -Maximum 9999)"
        
        $notificationParams = @{
            Name = $notificationRuleName
            HeaderContainsMessageHeader = "X-MS-Exchange-Organization-EmailCanaryAlert"
            HeaderContainsWords = "True"
            Priority = 0  # High priority
        }
        
        # Create the Slack notification action
        $slackMessage = @"
{
    "blocks": [
        {
            "type": "header",
            "text": {
                "type": "plain_text",
                "text": "⚠️ EMAIL CANARY ALERT ⚠️",
                "emoji": true
            }
        },
        {
            "type": "section",
            "text": {
                "type": "mrkdwn",
                "text": "*An email was received by the canary mailbox*\nThis may indicate that your email canary has been triggered. Please investigate immediately."
            }
        },
        {
            "type": "divider"
        },
        {
            "type": "section",
            "fields": [
                {
                    "type": "mrkdwn",
                    "text": "*Canary Mailbox:*\n$SourceMailbox"
                },
                {
                    "type": "mrkdwn",
                    "text": "*Received:*\n%%{DateTime}%%"
                }
            ]
        },
        {
            "type": "section",
            "fields": [
                {
                    "type": "mrkdwn",
                    "text": "*From:*\n%%{From}%%"
                },
                {
                    "type": "mrkdwn",
                    "text": "*Subject:*\n%%{Subject}%%"
                }
            ]
        },
        {
            "type": "actions",
            "elements": [
                {
                    "type": "button",
                    "text": {
                        "type": "plain_text",
                        "text": "View in Exchange Admin Center",
                        "emoji": true
                    },
                    "url": "https://admin.exchange.microsoft.com/#/mailboxes"
                }
            ]
        }
    ]
}
"@
        
        # Add the webhook action
        $notificationParams.Add("WebhookUrl", $webhookUrl)
        $notificationParams.Add("WebhookPayload", $slackMessage)
        
        # Create the notification rule
        New-TransportRule @notificationParams | Out-Null
        
        Write-Log -Message "Slack notification rule created successfully!" -Level "Success"
        Write-Log -Message "Final settings:" -Level "Info"
        Write-Host "Monitoring: $SourceMailbox"
        Write-Host "Alert Rule: $ruleName"
        Write-Host "Notification Rule: $notificationRuleName"
        Write-Host "Slack Webhook: $webhookUrl"
        
    } catch {
        Write-Log -Message "Error creating mail flow rule: $_" -Level "Error"
        
        # Provide guidance based on common errors
        if ($_.Exception.Message -like "*permission*") {
            Write-Log -Message "You may not have sufficient permissions to create mail flow rules." -Level "Warning"
        }
        
        Write-Log -Message "Slack notification setup failed." -Level "Error"
        return
    }
    
    # Ask if user wants to disconnect from Exchange Online
    $disconnect = Read-Host "Do you want to disconnect from Exchange Online? (y/n)"
    if ($disconnect -eq 'y') {
        try {
            Disconnect-ExchangeOnline -Confirm:$false
            Write-Log -Message "Disconnected from Exchange Online." -Level "Success"
        } catch {
            Write-Log -Message "Error disconnecting from Exchange Online: $_" -Level "Error"
        }
    } else {
        Write-Log -Message "Remaining connected to Exchange Online." -Level "Info"
    }
    
    Write-Log -Message "Slack notification setup completed." -Level "Success"
}

# Setup Teams notification for a received email on the canary email
function Setup-TeamsNotification {
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [string] $SourceMailbox
    )
    
    Write-Log -Message "Starting Teams notification setup process..." -Level "Info"
    
    # Ensure we have the required permissions
    if (-not (Test-RequiredPermissions -PermissionType "CreateDistributionGroup")) {
        Write-Log -Message "You may not have sufficient permissions to create mail flow rules." -Level "Warning"
        $continue = Read-Host "Do you want to continue anyway? (y/n)"
        if ($continue -ne "y") {
            Write-Log -Message "Teams notification setup cancelled." -Level "Warning"
            return
        }
    }

    # Step 1: Verify source mailbox
    Write-Log -Message "=== STEP 1: Source Mailbox ===" -Level "Info"
    
    # Use the provided source mailbox instead of asking again
    Write-Log -Message "Using provided source mailbox: $SourceMailbox" -Level "Info"
    
    # Check if mailbox exists in the organization
    try {
        $mailbox = Get-Mailbox -Identity $SourceMailbox -ErrorAction Stop
        
        Write-Log -Message "Found mailbox:" -Level "Success"
        Write-Host "Name: $($mailbox.DisplayName)"
        Write-Host "Email: $($mailbox.PrimarySmtpAddress)"
        Write-Host "Type: $($mailbox.RecipientTypeDetails)"
        
        $continue = Read-Host "Is this the correct mailbox? (y/n)"
        if ($continue -ne 'y') {
            # If the provided mailbox is not correct, ask for the correct one
            do {
                $SourceMailbox = Get-ValidEmail -PromptMessage "Enter the INTERNAL email address to monitor"
                
                try {
                    $mailbox = Get-Mailbox -Identity $SourceMailbox -ErrorAction Stop
                    
                    Write-Log -Message "Found mailbox:" -Level "Success"
                    Write-Host "Name: $($mailbox.DisplayName)"
                    Write-Host "Email: $($mailbox.PrimarySmtpAddress)"
                    Write-Host "Type: $($mailbox.RecipientTypeDetails)"
                    
                    $continue = Read-Host "Is this the correct mailbox? (y/n)"
                    if ($continue -eq 'y') {
                        break
                    }
                } catch {
                    Write-Log -Message "Could not find mailbox '$SourceMailbox'" -Level "Error"
                    Write-Log -Message "Please check the spelling and try again." -Level "Warning"
                }
            } while ($true)
        }
    } catch {
        Write-Log -Message "Error finding mailbox: $_" -Level "Error"
        
        # If the provided mailbox doesn't exist, ask for a valid one
        do {
            $SourceMailbox = Get-ValidEmail -PromptMessage "Enter the INTERNAL email address to monitor"
            
            try {
                $mailbox = Get-Mailbox -Identity $SourceMailbox -ErrorAction Stop
                
                Write-Log -Message "Found mailbox:" -Level "Success"
                Write-Host "Name: $($mailbox.DisplayName)"
                Write-Host "Email: $($mailbox.PrimarySmtpAddress)"
                Write-Host "Type: $($mailbox.RecipientTypeDetails)"
                
                $continue = Read-Host "Is this the correct mailbox? (y/n)"
                if ($continue -eq 'y') {
                    break
                }
            } catch {
                Write-Log -Message "Could not find mailbox '$SourceMailbox'" -Level "Error"
                Write-Log -Message "Please check the spelling and try again." -Level "Warning"
            }
        } while ($true)
    }

    # Step 2: Get Teams webhook URL
    Write-Log -Message "=== STEP 2: Teams Webhook URL ===" -Level "Info"
    Write-Log -Message "You need to provide a Teams webhook URL to receive notifications." -Level "Info"
    Write-Log -Message "To create a webhook URL in Teams:" -Level "Info"
    Write-Host "1. Go to the Teams channel where you want to receive notifications"
    Write-Host "2. Click the ... (more options) next to the channel name"
    Write-Host "3. Select 'Connectors'"
    Write-Host "4. Find 'Incoming Webhook' and click 'Configure'"
    Write-Host "5. Provide a name (e.g., 'Email Canary Alert')"
    Write-Host "6. Click 'Create'"
    Write-Host "7. Copy the webhook URL provided"
    
    do {
        $webhookUrl = Read-Host "Enter the Teams webhook URL"
        
        # Validate webhook URL format
        if ($webhookUrl -match "^https://.*\.webhook\.office\.com/webhookb2/.*") {
            Write-Log -Message "Webhook URL format appears valid." -Level "Success"
            break
        } else {
            Write-Log -Message "Invalid webhook URL format. Please enter a valid Teams webhook URL." -Level "Warning"
            $retry = Read-Host "Do you want to try again? (y/n)"
            if ($retry -ne "y") {
                Write-Log -Message "Teams notification setup cancelled." -Level "Warning"
                return
            }
        }
    } while ($true)
    
    # Step 3: Create a unique rule name
    $ruleName = "EmailCanary-Teams-Alert-$(Get-Random -Minimum 1000 -Maximum 9999)"
    Write-Log -Message "=== STEP 3: Create Mail Flow Rule ===" -Level "Info"
    Write-Log -Message "Creating mail flow rule with name: $ruleName" -Level "Info"
    
    # Step 4: Create the mail flow rule
    try {
        # Check if a similar rule already exists
        $existingRule = Get-TransportRule | Where-Object { $_.Name -like "EmailCanary-Teams-Alert-*" }
        
        if ($existingRule) {
            Write-Log -Message "An existing email canary alert rule was found: $($existingRule.Name)" -Level "Warning"
            $useExisting = Read-Host "Do you want to update the existing rule? (y/n)"
            
            if ($useExisting -eq "y") {
                # Remove the existing rule
                Remove-TransportRule -Identity $existingRule.Name -Confirm:$false
                Write-Log -Message "Removed existing rule: $($existingRule.Name)" -Level "Success"
            }
        }
        
        # Create a new rule
        $ruleParams = @{
            Name = $ruleName
            SentTo = $SourceMailbox
            SubjectOrBodyContainsWords = "*"  # Match any content
            SetHeaderName = "X-MS-Exchange-Organization-EmailCanaryAlert"
            SetHeaderValue = "True"
            ExceptIfHeaderContainsMessageHeader = "X-MS-Exchange-Organization-EmailCanaryAlert"
            ExceptIfHeaderContainsWords = "True"
            Priority = 0  # High priority
        }
        
        # Create the rule
        New-TransportRule @ruleParams | Out-Null
        
        Write-Log -Message "Mail flow rule created successfully!" -Level "Success"
        
        # Now create a second rule to send the notification to Teams
        $notificationRuleName = "EmailCanary-Teams-Notification-$(Get-Random -Minimum 1000 -Maximum 9999)"
        
        $notificationParams = @{
            Name = $notificationRuleName
            HeaderContainsMessageHeader = "X-MS-Exchange-Organization-EmailCanaryAlert"
            HeaderContainsWords = "True"
            Priority = 0  # High priority
        }
        
        # Create the Teams notification action
        $teamsMessage = @"
{
    "@type": "MessageCard",
    "@context": "http://schema.org/extensions",
    "themeColor": "FF0000",
    "summary": "⚠️ EMAIL CANARY ALERT ⚠️",
    "sections": [
        {
            "activityTitle": "⚠️ EMAIL CANARY ALERT ⚠️",
            "activitySubtitle": "An email was received by the canary mailbox",
            "facts": [
                {
                    "name": "Canary Mailbox:",
                    "value": "$SourceMailbox"
                },
                {
                    "name": "Received:",
                    "value": "%%{DateTime}%%"
                },
                {
                    "name": "From:",
                    "value": "%%{From}%%"
                },
                {
                    "name": "Subject:",
                    "value": "%%{Subject}%%"
                }
            ],
            "text": "This may indicate that your email canary has been triggered. Please investigate immediately."
        }
    ],
    "potentialAction": [
        {
            "@type": "OpenUri",
            "name": "View in Exchange Admin Center",
            "targets": [
                {
                    "os": "default",
                    "uri": "https://admin.exchange.microsoft.com/#/mailboxes"
                }
            ]
        }
    ]
}
"@
        
        # Add the webhook action
        $notificationParams.Add("WebhookUrl", $webhookUrl)
        $notificationParams.Add("WebhookPayload", $teamsMessage)
        
        # Create the notification rule
        New-TransportRule @notificationParams | Out-Null
        
        Write-Log -Message "Teams notification rule created successfully!" -Level "Success"
        Write-Log -Message "Final settings:" -Level "Info"
        Write-Host "Monitoring: $SourceMailbox"
        Write-Host "Alert Rule: $ruleName"
        Write-Host "Notification Rule: $notificationRuleName"
        Write-Host "Teams Webhook: $webhookUrl"
        
    } catch {
        Write-Log -Message "Error creating mail flow rule: $_" -Level "Error"
        
        # Provide guidance based on common errors
        if ($_.Exception.Message -like "*permission*") {
            Write-Log -Message "You may not have sufficient permissions to create mail flow rules." -Level "Warning"
        }
        
        Write-Log -Message "Teams notification setup failed." -Level "Error"
        return
    }
    
    # Ask if user wants to disconnect from Exchange Online
    $disconnect = Read-Host "Do you want to disconnect from Exchange Online? (y/n)"
    if ($disconnect -eq 'y') {
        try {
            Disconnect-ExchangeOnline -Confirm:$false
            Write-Log -Message "Disconnected from Exchange Online." -Level "Success"
        } catch {
            Write-Log -Message "Error disconnecting from Exchange Online: $_" -Level "Error"
        }
    } else {
        Write-Log -Message "Remaining connected to Exchange Online." -Level "Info"
    }
    
    Write-Log -Message "Teams notification setup completed." -Level "Success"
}