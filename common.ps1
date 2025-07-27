# -------------------------------
# emailCHanary Common Functions
# Made by: G. Renault & K. Tamine
# -------------------------------

# Function to check and install required modules
function Ensure-ModuleInstalled {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ModuleName,
        
        [Parameter(Mandatory=$false)]
        [string]$MinimumVersion = $null
    )
    
    $moduleParams = @{
        Name = $ModuleName
        ListAvailable = $true
    }
    
    if ($MinimumVersion) {
        $moduleParams.Add("MinimumVersion", $MinimumVersion)
    }
    
    $module = Get-Module @moduleParams
    
    if (-not $module) {
        Write-Host "Module $ModuleName is not installed." -ForegroundColor Yellow
        $install = Read-Host "Do you want to install it? (y/n)"
        
        if ($install -eq 'y') {
            try {
                if ($MinimumVersion) {
                    Install-Module -Name $ModuleName -MinimumVersion $MinimumVersion -Force -AllowClobber -Scope CurrentUser
                } else {
                    Install-Module -Name $ModuleName -Force -AllowClobber -Scope CurrentUser
                }
                Write-Host "Module $ModuleName installed successfully." -ForegroundColor Green
                return $true
            } catch {
                Write-Host "Failed to install module $ModuleName. Error: $_" -ForegroundColor Red
                return $false
            }
        } else {
            Write-Host "Module $ModuleName is required for this script to function properly." -ForegroundColor Red
            return $false
        }
    } else {
        Write-Host "Module $ModuleName is already installed." -ForegroundColor Green
        return $true
    }
}

# Function to check if user is connected to a service
function Ensure-ServiceConnected {
    param (
        [Parameter(Mandatory=$true)]
        [ValidateSet("Azure", "AzureAD", "ExchangeOnline")]
        [string]$ServiceName
    )
    
    switch ($ServiceName) {
        "Azure" {
            try {
                $context = Get-AzContext -ErrorAction Stop
                if (-not $context.Account) {
                    Write-Host "Not connected to Azure. Connecting..." -ForegroundColor Yellow
                    Connect-AzAccount -UseDeviceAuthentication
                    return $true
                } else {
                    Write-Host "Already connected to Azure as: $($context.Account.Id)" -ForegroundColor Green
                    return $true
                }
            } catch {
                Write-Host "Error checking Azure connection: $_" -ForegroundColor Red
                return $false
            }
        }
        "AzureAD" {
            try {
                Get-AzureADTenantDetail -ErrorAction Stop | Out-Null
                Write-Host "Already connected to Azure AD." -ForegroundColor Green
                return $true
            } catch {
                Write-Host "Not connected to Azure AD. Connecting..." -ForegroundColor Yellow
                try {
                    Connect-AzureAD
                    return $true
                } catch {
                    Write-Host "Error connecting to Azure AD: $_" -ForegroundColor Red
                    return $false
                }
            }
        }
        "ExchangeOnline" {
            try {
                Get-EXOMailbox -ResultSize 1 -ErrorAction Stop | Out-Null
                Write-Host "Already connected to Exchange Online." -ForegroundColor Green
                return $true
            } catch {
                Write-Host "Not connected to Exchange Online. Connecting..." -ForegroundColor Yellow
                try {
                    Connect-ExchangeOnline
                    return $true
                } catch {
                    Write-Host "Error connecting to Exchange Online: $_" -ForegroundColor Red
                    return $false
                }
            }
        }
    }
}

# Function to check if user has required permissions
function Test-RequiredPermissions {
    param (
        [Parameter(Mandatory=$true)]
        [ValidateSet("CreateMailbox", "CreateUser", "AssignLicense", "CreateDistributionGroup")]
        [string]$PermissionType
    )
    
    switch ($PermissionType) {
        "CreateMailbox" {
            try {
                # Check if user can get mailbox information
                Get-EXOMailbox -ResultSize 1 -ErrorAction Stop | Out-Null
                Write-Host "You have permissions to manage mailboxes." -ForegroundColor Green
                return $true
            } catch {
                Write-Host "You may not have permissions to create mailboxes. Error: $_" -ForegroundColor Red
                return $false
            }
        }
        "CreateUser" {
            try {
                # Check if user can get user information
                Get-AzureADUser -Top 1 -ErrorAction Stop | Out-Null
                Write-Host "You have permissions to manage users." -ForegroundColor Green
                return $true
            } catch {
                Write-Host "You may not have permissions to create users. Error: $_" -ForegroundColor Red
                return $false
            }
        }
        "AssignLicense" {
            try {
                # Check if user can get license information
                Get-AzureADSubscribedSku -ErrorAction Stop | Out-Null
                Write-Host "You have permissions to manage licenses." -ForegroundColor Green
                return $true
            } catch {
                Write-Host "You may not have permissions to assign licenses. Error: $_" -ForegroundColor Red
                return $false
            }
        }
        "CreateDistributionGroup" {
            try {
                # Check if user can get distribution group information
                Get-DistributionGroup -ResultSize 1 -ErrorAction Stop | Out-Null
                Write-Host "You have permissions to manage distribution groups." -ForegroundColor Green
                return $true
            } catch {
                Write-Host "You may not have permissions to create distribution groups. Error: $_" -ForegroundColor Red
                return $false
            }
        }
    }
}

# Function to get user principal name using PowerShell instead of Azure CLI
function Get-CurrentUserPrincipalName {
    try {
        $context = Get-AzContext
        if ($context -and $context.Account) {
            return $context.Account.Id
        } else {
            Write-Host "Not connected to Azure. Cannot determine user principal name." -ForegroundColor Red
            return $null
        }
    } catch {
        Write-Host "Error getting user principal name: $_" -ForegroundColor Red
        return $null
    }
}

# Function for enhanced logging
function Write-Log {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("Info", "Warning", "Error", "Success")]
        [string]$Level = "Info"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    switch ($Level) {
        "Info" { Write-Host $logMessage -ForegroundColor White }
        "Warning" { Write-Host $logMessage -ForegroundColor Yellow }
        "Error" { Write-Host $logMessage -ForegroundColor Red }
        "Success" { Write-Host $logMessage -ForegroundColor Green }
    }
}