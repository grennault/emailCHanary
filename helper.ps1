# Validate alias format
function Get-ValidAlias {
    param (
        [string]$PromptMessage
    )
    do {
        $alias = Read-Host ($PromptMessage).Trim()
        $isValid = $alias -match "^[a-zA-Z0-9!#\$%&'\*\+\-/=\?\^_`{\|}~\.]+$" -and ($alias -notmatch "^\." -and $alias -notmatch "\.$")
        if (-not $isValid) {
            Write-Host "Invalid alias. Use only letters, numbers, and allowed special characters. No spaces." -ForegroundColor Yellow
        }
    } while (-not $isValid)
    return $alias
}

# Function to validate email input
function Get-ValidEmail {
    param (
        [string]$PromptMessage
    )
    do {
        $email = (Read-Host $PromptMessage).Trim()
        $isValid = $email -match '^[^\s@]+@[^\s@]+\.[^\s@]+$'
        if (-not $isValid) {
            Write-Host "Invalid email format. Please enter in the format: user@domain.com" -ForegroundColor Yellow
        }
    } while (-not $isValid)
    return $email
}

function Get-AvailableDomains{
    Connect-AzureAD

    # Get all verified domains for the organization
    $domains = Get-AzureADDomain

    # Display the domains
    Write-Output "The following email domains are managed by your organization:"
    $domainNames = @()

    foreach ($domain in $domains) {
        $domainNames += $domain.Name
        Write-Output $domain.Name
    }

    # Check the number of managed domains
    $selectedDomain = $null

    if ($domainNames.Count -eq 1) {
        # Automatically select the domain if only one is available
        $selectedDomain = $domainNames[0]
        Write-Output "Only one domain available. Automatically selected: $selectedDomain"
    } else {
        # Force user to select a domain if more than one domain is available
        do {
            $selectedDomain = Read-Host "Please enter a domain from the above list"
        } until ($domainNames -contains $selectedDomain)

        Write-Output "Valid domain selected: $selectedDomain"
    }
    return $selectedDomain
}