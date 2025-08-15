# This script renames the custom default Word template, Normal.dotm,
# to Corp.Normal.dotm, creating a backup before Word generates a new default.

# Step 1: Define the file path.
# The $env:APPDATA variable automatically points to the user's AppData\Roaming folder.
$templatesPath = "$env:APPDATA\Microsoft\Templates"
$oldFileName = "Normal.dotm"
$newFileName = "Corp.Normal.dotm"
$oldFilePath = Join-Path -Path $templatesPath -ChildPath $oldFileName

# Step 2: Check if the directory and the file exist.
Write-Host "Checking for file '$oldFileName' in '$templatesPath'..." -ForegroundColor Cyan

if (Test-Path -Path $oldFilePath -PathType Leaf) {
    # If the file exists, proceed with the rename operation.
    try {
        # Rename the file using the Rename-Item cmdlet.
        Rename-Item -Path $oldFilePath -NewName $newFileName -Force

        # Provide a success message to the user.
        Write-Host "Success! The Word template file has been renamed from '$oldFileName' to '$newFileName'." -ForegroundColor Green
        Write-Host "You can now open Word, and it will use its default template." -ForegroundColor Green
    }
    catch {
        # Catch any errors during the rename process and display them.
        Write-Host "An error occurred while trying to rename the file:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
} else {
    # If the file does not exist, inform the user with a warning.
    Write-Host "WARNING: The file '$oldFileName' was not found at '$templatesPath'." -ForegroundColor Yellow
    Write-Host "This is unexpected, as Word usually creates this file by default." -ForegroundColor Yellow
    Write-Host "It may have already been renamed, deleted, or is in a different location." -ForegroundColor Cyan
}
