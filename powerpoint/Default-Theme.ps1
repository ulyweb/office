# This script renames the default Microsoft PowerPoint theme, Default Theme.thmx,
# to Corp.Default Theme.thmx. This creates a backup and forces PowerPoint to
# use its built-in default theme for new presentations.

# Step 1: Define the file path.
# The $env:APPDATA variable automatically points to the user's AppData\Roaming folder.
$templatesPath = "$env:APPDATA\Microsoft\Templates\Document Themes"
$oldFileName = "Default Theme.thmx"
$newFileName = "Corp.Default Theme.thmx"
$oldFilePath = Join-Path -Path $templatesPath -ChildPath $oldFileName
$newFilePath = Join-Path -Path $templatesPath -ChildPath $newFileName

# Step 2: Check if the original file exists.
if (Test-Path -Path $oldFilePath) {
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
    # If the file does not exist, inform the user.
    Write-Host "The file '$oldFileName' was not found at '$templatesPath'." -ForegroundColor Yellow
    Write-Host "This is normal if a custom default theme has not been saved." -ForegroundColor Yellow
    Write-Host "You can create one by opening a new PowerPoint presentation, saving it as a theme, and naming it 'Default Theme' in this folder." -ForegroundColor Cyan
}
