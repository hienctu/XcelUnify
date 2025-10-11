param(
    [string]$InstallPath
)
# Get the current user's profile folder
$userProfile = $env:USERPROFILE

$filePath = Join-Path $InstallPath "appSettings.json"

if (Test-Path $filePath) {

    # Read the entire file as text
    $content = Get-Content $filePath -Raw

    # Replace all occurrences of <<user-name>> with user profile path
    $content = $content -replace '<<user-name>>', [regex]::Escape($userProfile)

    # Save back to the JSON file
    Set-Content $filePath $content -Encoding UTF8

    Write-Host "Replaced <<user-name>> with $userProfile in $filePath"
}
else {
    Write-Host "File not found: $filePath"
}