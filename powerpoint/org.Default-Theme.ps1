# This script renames a custom default PowerPoint theme named Default Theme.potx
# to Corp.Default Theme.potx, creating a backup before PowerPoint generates a new default.

# Step 1: Define the file path.
# The $env:APPDATA variable automatically points to the user's AppData\Roaming folder.
$templatesPath = "$env:APPDATA\Microsoft\Templates\Document Themes"
$oldFileName = "Default Theme.potx"
$newFileName = "Corp.Default Theme.potx"
$oldFilePath = Join-Path -Path $templatesPath -ChildPath $oldFileName

# Step 2: Check if the directory and the file exist.
Write-Host "Checking for file '$oldFileName' in '$templatesPath'..." -ForegroundColor Cyan

if (Test-Path -Path $oldFilePath -PathType Leaf) {
    # If the file exists, proceed with the rename operation.
    try {
        # Rename the file using the Rename-Item cmdlet.
        Rename-Item -Path $oldFilePath -NewName $newFileName -Force

        # Provide a success message to the user.
        Write-Host "Success! The PowerPoint theme file has been renamed from '$oldFileName' to '$newFileName'." -ForegroundColor Green
        Write-Host "You can now open PowerPoint, and it will use its default Office theme." -ForegroundColor Green
    }
    catch {
        # Catch any errors during the rename process and display them.
        Write-Host "An error occurred while trying to rename the file:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
} else {
    # If the file does not exist, inform the user with a warning.
    Write-Host "WARNING: The file '$oldFileName' was not found at '$templatesPath'." -ForegroundColor Yellow
    Write-Host "This is normal if a custom default theme has not been saved in this location." -ForegroundColor Yellow
    Write-Host "You can create one by opening a new PowerPoint presentation, saving it as a template named 'Default Theme' in this folder." -ForegroundColor Cyan
}
