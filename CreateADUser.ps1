# -------------------------------
# emailCHanary AD User Creation
# Made by: G. Renault & K. Tamine
# -------------------------------

# Function to create an AD User with error handling
function Create-ADUser {
    Write-Log -Message "Starting AD user creation process..." -Level "Info"
    
    # Ensure we have the required permissions
    if (-not (Test-RequiredPermissions -PermissionType "CreateUser")) {
        Write-Log -Message "You may not have sufficient permissions to create AD users." -Level "Warning"
        $continue = Read-Host "Do you want to continue anyway? (y/n)"
        if ($continue -ne "y") {
            Write-Log -Message "AD user creation cancelled." -Level "Warning"
            return $null
        }
    }
    
    # Check license permissions
    if (-not (Test-RequiredPermissions -PermissionType "AssignLicense")) {
        Write-Log -Message "You may not have sufficient permissions to assign licenses." -Level "Warning"
        $continue = Read-Host "Do you want to continue anyway? (y/n)"
        if ($continue -ne "y") {
            Write-Log -Message "AD user creation cancelled." -Level "Warning"
            return $null
        }
    }
    
    # Get available domains
    $selectedDomain = Get-AvailableDomains
    if (-not $selectedDomain) {
        Write-Log -Message "Failed to get a valid domain. Cannot create AD user." -Level "Error"
        return $null
    }

    # Get AD User Info with validation
    $DisplayName = Get-ValidInput -PromptMessage "Enter the AD User DisplayName:" -MaxLength 64 -ErrorMessage "Invalid display name."
    $UserPrincipalName = Get-ValidInput -PromptMessage "Enter the AD User UserPrincipalName:" -RegexPattern "^[a-zA-Z0-9._%+-]+$" -MaxLength 64 -ErrorMessage "Invalid user principal name. Use only letters, numbers, and allowed special characters."
    $MailNickName = Get-ValidInput -PromptMessage "Enter the AD User MailNickName:" -RegexPattern "^[a-zA-Z0-9._%+-]+$" -MaxLength 64 -ErrorMessage "Invalid mail nickname. Use only letters, numbers, and allowed special characters."
    
    $createdEmail = "$UserPrincipalName@$selectedDomain"
    
    # Validate the constructed email address
    if ($createdEmail.Length -gt 256) {
        Write-Log -Message "The email address $createdEmail exceeds the maximum length of 256 characters." -Level "Error"
        return $null
    }
    
    # Check if user already exists
    if (Test-UserExists -UserPrincipalName $createdEmail) {
        Write-Log -Message "A user with the address $createdEmail already exists." -Level "Warning"
        $continue = Read-Host "Do you want to use this existing user? (y/n)"
        if ($continue -eq "y") {
            Write-Log -Message "Using existing user: $createdEmail" -Level "Success"
            return $createdEmail
        } else {
            Write-Log -Message "Please try again with a different user principal name." -Level "Info"
            return Create-ADUser  # Recursive call to try again
        }
    }
    
    Write-Log -Message "Creating AD user with email $createdEmail..." -Level "Info"
    
    # Prompt for a secure password
    $SecurePassword = Read-Host -AsSecureString "Enter password for the new user"
    
    # Security best practice: Force password change at next sign-in
    $forceChange = Read-Host "Force password change at next sign-in? (Recommended for security) (y/n)"
    $ForceChangePasswordNextSignIn = ($forceChange -eq "y")
    
    if (-not $ForceChangePasswordNextSignIn) {
        Write-Log -Message "Warning: Not forcing password change at next sign-in is a security risk." -Level "Warning"
    }
    
    # Create a new Azure AD user with error handling
    try {
        New-AzureADUser -DisplayName $DisplayName `
            -UserPrincipalName $createdEmail `
            -AccountEnabled $true `
            -MailNickName $MailNickName `
            -PasswordProfile (New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile -Property @{
                Password = $SecurePassword
                ForceChangePasswordNextSignIn = $ForceChangePasswordNextSignIn
            }) -ErrorAction Stop
            
        Write-Log -Message "AD user created successfully." -Level "Success"
        
        # Check available licenses
        try {
            $licenses = Get-AzureADSubscribedSku | Where-Object { $_.CapabilityStatus -eq "Enabled" }
            
            if (-not $licenses -or $licenses.Count -eq 0) {
                Write-Log -Message "No enabled licenses found in the tenant." -Level "Error"
                Write-Log -Message "User created but no license assigned." -Level "Warning"
                return $createdEmail
            }
            
            # Display available licenses
            Write-Log -Message "Available licenses:" -Level "Info"
            $licenseOptions = @{}
            $i = 1
            
            foreach ($license in $licenses) {
                $licenseOptions[$i] = $license
                Write-Host "$i. $($license.SkuPartNumber) (Available: $($license.PrepaidUnits.Enabled - $license.ConsumedUnits))"
                $i++
            }
            
            # Let user select a license
            $validSelection = $false
            do {
                $selection = Read-Host "Select a license to assign (1-$($licenseOptions.Count)), or 'skip' to skip license assignment"
                
                if ($selection -eq "skip") {
                    Write-Log -Message "License assignment skipped." -Level "Warning"
                    return $createdEmail
                }
                
                if ([int]::TryParse($selection, [ref]$null)) {
                    $selectionNum = [int]$selection
                    if ($selectionNum -ge 1 -and $selectionNum -le $licenseOptions.Count) {
                        $selectedLicense = $licenseOptions[$selectionNum]
                        $validSelection = $true
                    }
                }
                
                if (-not $validSelection) {
                    Write-Log -Message "Invalid selection. Please try again." -Level "Warning"
                }
            } while (-not $validSelection)
            
            # Assign the license to the user
            try {
                Set-AzureADUserLicense -ObjectId $createdEmail -AddLicenses $selectedLicense.SkuId -ErrorAction Stop
                Write-Log -Message "License $($selectedLicense.SkuPartNumber) assigned successfully." -Level "Success"
            } catch {
                Write-Log -Message "Error assigning license: $_" -Level "Error"
                Write-Log -Message "User created but license assignment failed." -Level "Warning"
            }
        } catch {
            Write-Log -Message "Error retrieving licenses: $_" -Level "Error"
            Write-Log -Message "User created but license assignment failed." -Level "Warning"
        }
        
        Write-Log -Message "AD user $createdEmail has been created successfully" -Level "Success"
        return $createdEmail
    } catch {
        Write-Log -Message "Error creating AD user: $_" -Level "Error"
        
        # Provide guidance based on common errors
        if ($_.Exception.Message -like "*already exists*") {
            Write-Log -Message "A user with this name or email already exists. Please try a different name or email." -Level "Warning"
        } elseif ($_.Exception.Message -like "*permission*") {
            Write-Log -Message "You may not have sufficient permissions to create users." -Level "Warning"
        }
        
        return $null
    }
}

# Import common functions and helper functions
. "$PSScriptRoot\common.ps1"
. "$PSScriptRoot\helper.ps1"