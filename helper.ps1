# -------------------------------
# emailCHanary Helper Functions
# Made by: G. Renault & K. Tamine
# -------------------------------

# Validate alias format
function Get-ValidAlias {
    param (
        [string]$PromptMessage
    )
    
    do {
        $alias = Read-Host ($PromptMessage).Trim()
        
        # Improved regex pattern for alias validation
        # Allows letters, numbers, and specific special characters
        # Doesn't allow starting or ending with a period
        $isValid = $alias -match "^[a-zA-Z0-9!#$%&'*+\-/=?^_{}|~.]+$" -and 
                  ($alias -notmatch "^\." -and $alias -notmatch "\.$") -and
                  ($alias.Length -le 64)
                  
        if (-not $isValid) {
            Write-Log -Message "Invalid alias. Use only letters, numbers, and allowed special characters. No spaces. Cannot start or end with a period. Maximum length is 64 characters." -Level "Warning"
        }
        
        # Allow user to exit
        if ($alias -eq "exit" -or $alias -eq "quit") {
            $confirm = Read-Host "Are you sure you want to exit? (y/n)"
            if ($confirm -eq "y") {
                Write-Log -Message "Operation cancelled by user." -Level "Warning"
                exit
            }
        }
    } while (-not $isValid)
    
    return $alias
}

# Function to validate email input with improved regex
function Get-ValidEmail {
    param (
        [string]$PromptMessage
    )
    
    do {
        $email = (Read-Host $PromptMessage).Trim()
        
        # More comprehensive email validation regex
        # Checks for proper format, valid characters, and length limits
        $isValid = $email -match '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$' -and
                   ($email.Length -le 254)
                   
        if (-not $isValid) {
            Write-Log -Message "Invalid email format. Please enter in the format: user@domain.com" -Level "Warning"
        }
        
        # Allow user to exit
        if ($email -eq "exit" -or $email -eq "quit") {
            $confirm = Read-Host "Are you sure you want to exit? (y/n)"
            if ($confirm -eq "y") {
                Write-Log -Message "Operation cancelled by user." -Level "Warning"
                exit
            }
        }
    } while (-not $isValid)
    
    return $email
}

# Function to get available domains with error handling
function Get-AvailableDomains {
    try {
        # Ensure we're connected to AzureAD
        if (-not (Ensure-ServiceConnected -ServiceName "AzureAD")) {
            Write-Log -Message "Cannot retrieve domains without AzureAD connection." -Level "Error"
            return $null
        }
        
        # Get all verified domains for the organization
        $domains = Get-AzureADDomain -ErrorAction Stop
        
        if (-not $domains -or $domains.Count -eq 0) {
            Write-Log -Message "No domains found for your organization." -Level "Error"
            return $null
        }
        
        # Display the domains
        Write-Log -Message "The following email domains are managed by your organization:" -Level "Info"
        $domainNames = @()
        
        foreach ($domain in $domains) {
            $domainNames += $domain.Name
            Write-Host $domain.Name
        }
        
        # Check the number of managed domains
        $selectedDomain = $null
        
        if ($domainNames.Count -eq 1) {
            # Automatically select the domain if only one is available
            $selectedDomain = $domainNames[0]
            Write-Log -Message "Only one domain available. Automatically selected: $selectedDomain" -Level "Info"
            $confirm = Read-Host "Is this domain correct? (y/n)"
            if ($confirm -ne "y") {
                Write-Log -Message "Domain selection cancelled by user." -Level "Warning"
                return $null
            }
        } else {
            # Force user to select a domain if more than one domain is available
            do {
                $selectedDomain = Read-Host "Please enter a domain from the above list (or 'exit' to cancel)"
                
                if ($selectedDomain -eq "exit" -or $selectedDomain -eq "quit") {
                    Write-Log -Message "Domain selection cancelled by user." -Level "Warning"
                    return $null
                }
            } until ($domainNames -contains $selectedDomain)
            
            Write-Log -Message "Valid domain selected: $selectedDomain" -Level "Success"
        }
        
        return $selectedDomain
    }
    catch {
        Write-Log -Message "Error retrieving domains: $_" -Level "Error"
        return $null
    }
}

# Function to validate string input with optional regex pattern
function Get-ValidInput {
    param (
        [string]$PromptMessage,
        [string]$RegexPattern = ".*",
        [string]$ErrorMessage = "Invalid input. Please try again.",
        [int]$MaxLength = 0
    )
    
    do {
        $input = (Read-Host $PromptMessage).Trim()
        
        $isValid = $input -match $RegexPattern
        
        if ($MaxLength -gt 0) {
            $isValid = $isValid -and ($input.Length -le $MaxLength)
        }
        
        if (-not $isValid) {
            if ($MaxLength -gt 0) {
                Write-Log -Message "$ErrorMessage Maximum length is $MaxLength characters." -Level "Warning"
            } else {
                Write-Log -Message $ErrorMessage -Level "Warning"
            }
        }
        
        # Allow user to exit
        if ($input -eq "exit" -or $input -eq "quit") {
            $confirm = Read-Host "Are you sure you want to exit? (y/n)"
            if ($confirm -eq "y") {
                Write-Log -Message "Operation cancelled by user." -Level "Warning"
                exit
            }
        }
    } while (-not $isValid)
    
    return $input
}

# Function to check if a mailbox exists
function Test-MailboxExists {
    param (
        [string]$EmailAddress
    )
    
    try {
        $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction SilentlyContinue
        return ($mailbox -ne $null)
    }
    catch {
        return $false
    }
}

# Function to check if a user exists
function Test-UserExists {
    param (
        [string]$UserPrincipalName
    )
    
    try {
        $user = Get-AzureADUser -Filter "userPrincipalName eq '$UserPrincipalName'" -ErrorAction SilentlyContinue
        return ($user -ne $null)
    }
    catch {
        return $false
    }
}

# Function to check if a distribution group exists
function Test-DistributionGroupExists {
    param (
        [string]$GroupName
    )
    
    try {
        $group = Get-DistributionGroup -Identity $GroupName -ErrorAction SilentlyContinue
        return ($group -ne $null)
    }
    catch {
        return $false
    }
}

# Function to check if a mail contact exists
function Test-MailContactExists {
    param (
        [string]$ExternalEmailAddress
    )
    
    try {
        $contact = Get-MailContact -Filter "ExternalEmailAddress -eq 'SMTP:$ExternalEmailAddress'" -ErrorAction SilentlyContinue
        return ($contact -ne $null)
    }
    catch {
        return $false
    }
}

# Import common functions
. "$PSScriptRoot\common.ps1"