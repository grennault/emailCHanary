# Setup the forwarding rule to notify for a received email on the canary email
function Setup-NotificationMethod {
        Param
        (
            [Parameter(Mandatory=$true, Position=0)]
            [string] $SourceMailbox,
        )

    # Step 1: Get and verify source mailbox
    Write-Host "`n=== STEP 1: Source Mailbox ===" -ForegroundColor Cyan
    do {
        $SourceMailbox = Get-ValidEmail -PromptMessage "Enter the INTERNAL email address to forward FROM"
        
        # Check if mailbox exists in the organization
        $mailbox = Get-Mailbox -Identity $SourceMailbox -ErrorAction SilentlyContinue
        
        if ($mailbox) {
            Write-Host "`nFound mailbox:" -ForegroundColor Green
            Write-Host "Name: $($mailbox.DisplayName)"
            Write-Host "Email: $($mailbox.PrimarySmtpAddress)"
            Write-Host "Type: $($mailbox.RecipientTypeDetails)"
            
            $continue = Read-Host "`nIs this the correct mailbox? (y/n)"
            if ($continue -eq 'y') {
                break
            }
        } else {
            Write-Host "`nError: Could not find mailbox '$SourceMailbox'" -ForegroundColor Red
            Write-Host "Please check the spelling and try again." -ForegroundColor Yellow
        }
    } while ($true)

    # Step 2: Create distribution group
    Write-Host "`n=== STEP 2: Distribution Group ===" -ForegroundColor Cyan
    do {
        $GroupName = Read-Host "Enter a name for the distribution group (e.g., ForwardingGroup)"
        $GroupDisplayName = Read-Host "Enter a display name for the group (e.g., Forwarding Group)"
        $GroupEmail = "$($GroupName)@mailchanaryprotonmail.onmicrosoft.com"

        Write-Host "`nCreating distribution group:" -ForegroundColor Yellow
        Write-Host "Name: $GroupName"
        Write-Host "Display Name: $GroupDisplayName"
        Write-Host "Email: $GroupEmail"

        $continue = Read-Host "`nCreate this distribution group? (y/n)"
        if ($continue -eq 'y') {
            try {
                New-DistributionGroup -Name $GroupName -DisplayName $GroupDisplayName -PrimarySmtpAddress $GroupEmail
                Write-Host "Distribution group created successfully!" -ForegroundColor Green
                break
            } catch {
                Write-Host "Error creating distribution group:" -ForegroundColor Red
                Write-Host $_
                Write-Host "`nPlease try again with a different name." -ForegroundColor Yellow
            }
        }
    } while ($true)

    # Step 3: Create mail contact
    Write-Host "`n=== STEP 3: External Contact ===" -ForegroundColor Cyan
    do {
        $ContactName = Read-Host "Enter a display name for the contact (e.g., Your Gmail)"
        $ExternalEmail = Get-ValidEmail -PromptMessage "Enter the external email address to forward TO"

        Write-Host "`nCreating mail contact:" -ForegroundColor Yellow
        Write-Host "Name: $ContactName"
        Write-Host "Email: $ExternalEmail"

        $continue = Read-Host "`nCreate this mail contact? (y/n)"
        if ($continue -eq 'y') {
            try {
                New-MailContact -Name $ContactName -ExternalEmailAddress $ExternalEmail
                Write-Host "Mail contact created successfully!" -ForegroundColor Green
                break
            } catch {
                Write-Host "Error creating mail contact:" -ForegroundColor Red
                Write-Host $_
                Write-Host "`nPlease try again with different details." -ForegroundColor Yellow
            }
        }
    } while ($true)

    # Step 4: Add contact to group
    Write-Host "`n=== STEP 4: Add Contact to Group ===" -ForegroundColor Cyan
    Write-Host "Adding contact '$ContactName' to group '$GroupName'..." -ForegroundColor Yellow
    
    try {
        Add-DistributionGroupMember -Identity $GroupName -Member $ExternalEmail
        Write-Host "Contact added to group successfully!" -ForegroundColor Green
    } catch {
        Write-Host "Error adding contact to group:" -ForegroundColor Red
        Write-Host $_
        return
    }

    # Step 5: Set up forwarding
    Write-Host "`n=== STEP 5: Set Up Forwarding ===" -ForegroundColor Cyan
    Write-Host "`nFinal settings:" -ForegroundColor Yellow
    Write-Host "From: $SourceMailbox"
    Write-Host "To Group: $GroupName ($GroupEmail)"
    Write-Host "Contact: $ContactName ($ExternalEmail)"

    $continue = Read-Host "`nSet up forwarding with these settings? (y/n)"
    if ($continue -eq 'y') {
        try {
            Set-Mailbox -Identity $SourceMailbox -DeliverToMailboxAndForward $true -ForwardingAddress $GroupName
            Write-Host "`nForwarding set up successfully!" -ForegroundColor Green
            Write-Host "`nFinal forwarding settings:"
            Get-Mailbox -Identity $SourceMailbox | Select-Object Name, ForwardingAddress, DeliverToMailboxAndForward
        } catch {
            Write-Host "Error setting up forwarding:" -ForegroundColor Red
            Write-Host $_
        }
    }

    Disconnect-ExchangeOnline -Confirm:$false
}