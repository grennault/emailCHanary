# This is script downloads all ps1 from emailCHanary github and excute the main.ps1 script. 

# Define the directory and GitHub URL
$directory = "emailCHanary"
$baseUrl = "https://raw.githubusercontent.com/grennault/$directory/blob/main/"

# Create the directory if it doesn't exist
if (-Not (Test-Path -Path $directory)) {
    New-Item -ItemType Directory -Path $directory
}

# List of files to download
$files = @("helper.ps1", "main.ps1", "CreateADUser.ps1", "CreateSharedMailbox.ps1", "SetupNotificationMethod.ps1")

# Download each file
foreach ($file in $files) {
    Invoke-WebRequest -Uri ($baseUrl + $file) -OutFile (Join-Path -Path $directory -ChildPath $file)
}

# Dot-source main.ps1
. "$PSScriptRoot\$directory\main.ps1"