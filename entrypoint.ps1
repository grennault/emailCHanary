# -------------------------------
# emailCHanary Azure Script - Entry Point
# Made by: G. Renault & K. Tamine
# Upload this azure powershell script in your Azure Portal terminal (Upload file) and execute it ./entrypoint.ps1
# -------------------------------

# Set execution policy for the current process
try {
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
    Write-Host "Execution policy set to Bypass for the current process." -ForegroundColor Green
} catch {
    Write-Host "Warning: Could not set execution policy. You may need to run this script with elevated privileges." -ForegroundColor Yellow
    Write-Host "Error: $_" -ForegroundColor Red
}

# Check if running in Azure Cloud Shell
$isCloudShell = $env:ACC_LOCATION -ne $null
if ($isCloudShell) {
    Write-Host "Running in Azure Cloud Shell environment." -ForegroundColor Green
} else {
    Write-Host "Not running in Azure Cloud Shell. Some features may not work as expected." -ForegroundColor Yellow
    Write-Host "For best results, please run this script in Azure Cloud Shell." -ForegroundColor Yellow
}

# Get the directory of the current script
$scriptPath = $PSScriptRoot
if (-not $scriptPath) {
    $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
}

# Check if all required script files exist
$requiredFiles = @(
    "main.ps1",
    "common.ps1",
    "helper.ps1",
    "CreateADUser.ps1",
    "CreateSharedMailbox.ps1",
    "SetupNotificationMethod.ps1"
)

$missingFiles = @()
foreach ($file in $requiredFiles) {
    $filePath = Join-Path -Path $scriptPath -ChildPath $file
    if (-not (Test-Path -Path $filePath)) {
        $missingFiles += $file
    }
}

if ($missingFiles.Count -gt 0) {
    Write-Host "Error: The following required files are missing:" -ForegroundColor Red
    foreach ($file in $missingFiles) {
        Write-Host "  - $file" -ForegroundColor Red
    }
    Write-Host "Please ensure all required files are in the same directory as this script." -ForegroundColor Red
    exit 1
}

# Execute the main script
try {
    Write-Host "Starting emailCHanary script..." -ForegroundColor Cyan
    & "$scriptPath\main.ps1"
} catch {
    Write-Host "An error occurred while executing the main script:" -ForegroundColor Red
    Write-Host $_ -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    exit 1
}