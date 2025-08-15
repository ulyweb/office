# This script safely renames the custom default Word template, Normal.dotm,
# to Corp.Normal.dotm, creating a backup before Word generates a new default.
# It waits for the user to close Microsoft Word to prevent data loss and then
# automatically re-opens the application.

# Step 1: Check if Word is running and wait for it to close.
$wordProcess = Get-Process -Name "winword" -ErrorAction SilentlyContinue

while ($wordProcess) {
    # If the Word process is found, display a warning and wait.
    Write-Host "WARNING: Microsoft Word is currently running." -ForegroundColor Yellow
    Write-Host "Please save all your active documents and close Word to proceed." -ForegroundColor Yellow
    Write-Host "The script is waiting for Word to be closed..." -ForegroundColor Cyan
    
    # Wait for 5 seconds before checking again.
    Start-Sleep -Seconds 5
    
    # Check for the Word process again.
    $wordProcess = Get-Process -Name "winword" -ErrorAction SilentlyContinue
}

# Step 2: Proceed with the file operation once Word is closed.
Write-Host "Microsoft Word is no longer running. Proceeding with the operation." -ForegroundColor Green

# Define the file path.
# The $env:APPDATA variable automatically points to the user's AppData\Roaming folder.
$templatesPath = "$env:APPDATA\Microsoft\Templates"
$oldFileName = "Normal.dotm"
$newFileName = "Corp.Normal.dotm"
$oldFilePath = Join-Path -Path $templatesPath -ChildPath $oldFileName

# Step 3: Check if the file exists.
Write-Host "Checking for file '$oldFileName' in '$templatesPath'..." -ForegroundColor Cyan

if (Test-Path -Path $oldFilePath -PathType Leaf) {
    # If the file exists, proceed with the rename operation.
    try {
        # Rename the file using the Rename-Item cmdlet.
        Rename-Item -Path $oldFilePath -NewName $newFileName -Force

        # Provide a success message to the user.
        Write-Host "Success! The Word template file has been renamed from '$oldFileName' to '$newFileName'." -ForegroundColor Green
        
        # Re-open Microsoft Word for the user.
        Write-Host "Reopening Microsoft Word now..." -ForegroundColor Green
        Start-Process -FilePath "winword.exe"

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
