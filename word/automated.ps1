# This script renames the default Microsoft Word template, Normal.dotm,
# to Corp.Normal.dotm. This is useful for corporate environments or
# for creating a backup of a customized template before a new one is created.

# Step 1: Define the file path.
# The $env:APPDATA variable automatically points to the user's AppData\Roaming folder.
$templatesPath = "$env:APPDATA\Microsoft\Templates"
$oldFileName = "Normal.dotm"
$newFileName = "Corp.Normal.dotm"
$oldFilePath = Join-Path -Path $templatesPath -ChildPath $oldFileName
$newFilePath = Join-Path -Path $templatesPath -ChildPath $newFileName

# Step 2: Check if the original file exists.
if (Test-Path -Path $oldFilePath) {
    # If the file exists, proceed with the rename operation.
    try {
        # Rename the file using the Rename-Item cmdlet.
        Rename-Item -Path $oldFilePath -NewName $newFileName -Force

        # Provide a success message to the user.
        Write-Host "Success! The file has been renamed from '$oldFileName' to '$newFileName'." -ForegroundColor Green
        Write-Host "You can now open Word, and it will automatically generate a new default Normal.dotm template." -ForegroundColor Green
    }
    catch {
        # Catch any errors during the rename process and display them.
        Write-Host "An error occurred while trying to rename the file:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
} else {
    # If the file does not exist, inform the user.
    Write-Host "The file '$oldFileName' was not found at '$templatesPath'." -ForegroundColor Yellow
    Write-Host "It may have already been renamed or deleted." -ForegroundColor Yellow
}
