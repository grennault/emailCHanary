# Function to create a shared mailbox 
function Create-SharedMailbox {
    $selectedDomain = Get-AvailableDomains

    # Get Shared Mailbox Info from User
    $SharedMailboxName = (Read-Host "Enter the name for the shared mailbox (e.g., Security Canary):").Trim()
    $SharedMailboxAlias = Get-ValidAlias -PromptMessage "Enter the alias for the shared mailbox (e.g., securitycanary):"
    $SharedMailboxEmail = $SharedMailboxAlias + "@" + $selectedDomain

    Write-Output "Creating $SharedMailboxAlias@$selectedDomain as a shared email..."

    # Create the Shared Mailbox
    New-Mailbox -Shared `
        -Name $SharedMailboxName `
        -Alias $SharedMailboxAlias `
        -PrimarySmtpAddress $SharedMailboxEmail

    # Ensure GAL visibility and set email address
    Set-Mailbox -Identity $SharedMailboxEmail `
        -DisplayName $SharedMailboxName `
        -HiddenFromAddressListsEnabled $false `
        -WindowsEmailAddress $SharedMailboxEmail

    Write-Output "Shared email $SharedMailboxAlias@$selectedDomain has been created" -ForegroundColor Green

    return $SharedMailboxEmail
}