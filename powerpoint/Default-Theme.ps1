# This script safely renames a custom default PowerPoint theme named Default Theme.potx
# to Corp.Default Theme.potx, creating a backup before PowerPoint generates a new default.
# It waits for the user to close PowerPoint to prevent data loss and then
# automatically re-opens the application.

# Step 1: Check if PowerPoint is running and wait for it to close.
$powerPointProcess = Get-Process -Name "powerpnt" -ErrorAction SilentlyContinue

while ($powerPointProcess) {
    # If the PowerPoint process is found, display a warning and wait.
    Write-Host "WARNING: Microsoft PowerPoint is currently running." -ForegroundColor Yellow
    Write-Host "Please save all your active presentations and close PowerPoint to proceed." -ForegroundColor Yellow
    Write-Host "The script is waiting for PowerPoint to be closed..." -ForegroundColor Cyan
    
    # Wait for 5 seconds before checking again.
    Start-Sleep -Seconds 5
    
    # Check for the PowerPoint process again.
    $powerPointProcess = Get-Process -Name "powerpnt" -ErrorAction SilentlyContinue
}

# Step 2: Proceed with the file operation once PowerPoint is closed.
Write-Host "Microsoft PowerPoint is no longer running. Proceeding with the operation." -ForegroundColor Green

# Define the file path.
# The $env:APPDATA variable automatically points to the user's AppData\Roaming folder.
$templatesPath = "$env:APPDATA\Microsoft\Templates\Document Themes"
$oldFileName = "Default Theme.potx"
$newFileName = "Corp.Default Theme.potx"
$oldFilePath = Join-Path -Path $templatesPath -ChildPath $oldFileName

# Step 3: Check if the file exists.
Write-Host "Checking for file '$oldFileName' in '$templatesPath'..." -ForegroundColor Cyan

if (Test-Path -Path $oldFilePath -PathType Leaf) {
    # If the file exists, proceed with the rename operation.
    try {
        # Rename the file using the Rename-Item cmdlet.
        Rename-Item -Path $oldFilePath -NewName $newFileName -Force

        # Provide a success message to the user.
        Write-Host "Success! The PowerPoint theme file has been renamed from '$oldFileName' to '$newFileName'." -ForegroundColor Green
        
        # Re-open Microsoft PowerPoint for the user.
        Write-Host "Reopening Microsoft PowerPoint now..." -ForegroundColor Green
        Start-Process -FilePath "powerpnt.exe"

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
