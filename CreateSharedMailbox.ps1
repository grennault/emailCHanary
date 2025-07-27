# -------------------------------
# emailCHanary Shared Mailbox Creation
# Made by: G. Renault & K. Tamine
# -------------------------------

# Function to create a shared mailbox with error handling
function Create-SharedMailbox {
    Write-Log -Message "Starting shared mailbox creation process..." -Level "Info"
    
    # Ensure we have the required permissions
    if (-not (Test-RequiredPermissions -PermissionType "CreateMailbox")) {
        Write-Log -Message "You may not have sufficient permissions to create a shared mailbox." -Level "Warning"
        $continue = Read-Host "Do you want to continue anyway? (y/n)"
        if ($continue -ne "y") {
            Write-Log -Message "Shared mailbox creation cancelled." -Level "Warning"
            return $null
        }
    }
    
    # Get available domains
    $selectedDomain = Get-AvailableDomains
    if (-not $selectedDomain) {
        Write-Log -Message "Failed to get a valid domain. Cannot create shared mailbox." -Level "Error"
        return $null
    }

    # Get Shared Mailbox Info from User
    $SharedMailboxName = Get-ValidInput -PromptMessage "Enter the name for the shared mailbox (e.g., Security Canary):" -MaxLength 64 -ErrorMessage "Invalid mailbox name."
    $SharedMailboxAlias = Get-ValidAlias -PromptMessage "Enter the alias for the shared mailbox (e.g., securitycanary):"
    $SharedMailboxEmail = "$SharedMailboxAlias@$selectedDomain"
    
    # Validate the constructed email address
    if ($SharedMailboxEmail.Length -gt 256) {
        Write-Log -Message "The email address $SharedMailboxEmail exceeds the maximum length of 256 characters." -Level "Error"
        return $null
    }
    
    # Check if mailbox already exists
    if (Test-MailboxExists -EmailAddress $SharedMailboxEmail) {
        Write-Log -Message "A mailbox with the address $SharedMailboxEmail already exists." -Level "Warning"
        $continue = Read-Host "Do you want to use this existing mailbox? (y/n)"
        if ($continue -eq "y") {
            Write-Log -Message "Using existing mailbox: $SharedMailboxEmail" -Level "Success"
            return $SharedMailboxEmail
        } else {
            Write-Log -Message "Please try again with a different alias." -Level "Info"
            return Create-SharedMailbox  # Recursive call to try again
        }
    }
    
    Write-Log -Message "Creating $SharedMailboxEmail as a shared mailbox..." -Level "Info"
    
    # Create the Shared Mailbox with error handling
    try {
        New-Mailbox -Shared `
            -Name $SharedMailboxName `
            -Alias $SharedMailboxAlias `
            -PrimarySmtpAddress $SharedMailboxEmail -ErrorAction Stop
            
        Write-Log -Message "Shared mailbox created successfully." -Level "Success"
        
        # Ensure GAL visibility and set email address
        try {
            Set-Mailbox -Identity $SharedMailboxEmail `
                -DisplayName $SharedMailboxName `
                -HiddenFromAddressListsEnabled $false `
                -WindowsEmailAddress $SharedMailboxEmail -ErrorAction Stop
                
            Write-Log -Message "Mailbox properties configured successfully." -Level "Success"
        } catch {
            Write-Log -Message "Error configuring mailbox properties: $_" -Level "Error"
            Write-Log -Message "The mailbox was created but may not be properly configured." -Level "Warning"
        }
        
        Write-Log -Message "Shared mailbox $SharedMailboxEmail has been created successfully" -Level "Success"
        return $SharedMailboxEmail
    } catch {
        Write-Log -Message "Error creating shared mailbox: $_" -Level "Error"
        
        # Provide guidance based on common errors
        if ($_.Exception.Message -like "*already exists*") {
            Write-Log -Message "A mailbox with this name or alias already exists. Please try a different name or alias." -Level "Warning"
        } elseif ($_.Exception.Message -like "*permission*") {
            Write-Log -Message "You may not have sufficient permissions to create mailboxes." -Level "Warning"
        }
        
        return $null
    }
}

# Import common functions and helper functions
. "$PSScriptRoot\common.ps1"
. "$PSScriptRoot\helper.ps1"