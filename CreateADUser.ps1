# Function to create a AD User 
function Create-ADUser {
    $selectedDomain = Get-AvailableDomains

    $DisplayName = (Read-Host "Enter the AD User DisplayName:").Trim()
    $UserPrincipalName = (Read-Host "Enter the AD User UserPrincipalName:").Trim()
    $MailNickName = (Read-Host "Enter the AD User MailNickName:").Trim()
    
    $createdEmail = $UserPrincipalName + "@" + $selectedDomain

    # Prompt for a secure password
    $SecurePassword = Read-Host -AsSecureString "Enter password for the new user"

    # Create a new Azure AD user without specifying the password directly in the script
    New-AzureADUser -DisplayName $DisplayName -UserPrincipalName $createdEmail -AccountEnabled $true -MailNickName $MailNickName -PasswordProfile (New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile -Property @{Password = $SecurePassword; ForceChangePasswordNextSignIn = $false})

    # Get the license details
    $License = Get-AzureADSubscribedSku | Where-Object { $_.SkuPartNumber -eq "EXCHANGE_S_STANDARD" }

    # Assign the license to the user
    Set-AzureADUserLicense -ObjectId $createdEmail -AddLicenses $License.SkuId

    return $createdEmail
}