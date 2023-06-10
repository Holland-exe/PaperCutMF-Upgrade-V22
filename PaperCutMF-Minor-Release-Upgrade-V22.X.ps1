# Get the path of the current script
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

# Check if the PaperCut installation folder exists in the C drive
$pcInstallationLocationC = "C:\Program Files\PaperCut MF"
if (Test-Path -Path $pcInstallationLocationC -PathType Container) {
    $pcInstallationLocation = $pcInstallationLocationC
}
else {
    # Check if the PaperCut installation folder exists in the E drive
    $pcInstallationLocationE = "E:\Program Files\PaperCut MF"
    if (Test-Path -Path $pcInstallationLocationE -PathType Container) {
        $pcInstallationLocation = $pcInstallationLocationE
    }
    else {
        Write-Host "Error: PaperCut installation folder not found."
        return
    }
}

# Define the source and destination paths
$sourcePath = $pcInstallationLocation
$destinationPath = "C:\EDUIT\PaperCut-Backup\22.1.1"

# Define the log file path
$logFilePath = Join-Path -Path $scriptPath -ChildPath "papercut-upgrade-log.dat"

# Function to write log entries to the log file
function Write-LogEntry {
    param(
        [string]$Message
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[${timestamp}] $Message"

    $logEntry | Out-File -FilePath $logFilePath -Append
}

# Echo step: Stop PaperCut Server
Write-Host "Stopping PaperCut Server..."
Write-LogEntry -Message "Stopping PaperCut Server"
Stop-Service -Name "PaperCut Application Server"

# Echo step: Create Backup Folder
$backupFolder = Join-Path -Path $destinationPath -ChildPath "PaperCut MF"
if (!(Test-Path -Path $backupFolder -PathType Container)) {
    Write-Host "Creating Backup Folder..."
    Write-LogEntry -Message "Creating Backup Folder: $backupFolder"
    New-Item -Path $backupFolder -ItemType Directory
} else {
    Write-Host "Backup folder already exists. Skipping backup creation."
    Write-LogEntry -Message "Backup folder already exists. Skipping backup creation: $backupFolder"
}

# Copy the installation folder to the backup folder
Write-Host "Copying installation folder to backup..."
Write-LogEntry -Message "Copying installation folder to backup: $sourcePath to $backupFolder"
Copy-Item -Path $sourcePath -Destination $backupFolder -Recurse -Force

# Echo step: Backup Complete
Write-Host "Backup complete."
Write-LogEntry -Message "Backup complete: $backupFolder"

# Specify the URL of the exe file
$downloadUrl = "https://cdn.papercut.com/web/products/ng-mf/installers/mf/22.x/pcmf-setup-22.1.1.66714.exe"

# Specify the filename for the downloaded file
$fileName = "pcmf-setup-22.1.1.66714.exe"

# Specify the path to save the downloaded file
$exePath = Join-Path -Path $scriptPath -ChildPath $fileName

# Download the file
Write-Host "Downloading the upgrade file..."
Write-LogEntry -Message "Downloading the upgrade file"
Invoke-WebRequest -Uri $downloadUrl -OutFile $exePath

# Check if the file was downloaded successfully
if (Test-Path -Path $exePath -PathType Leaf) {
    # Echo step: Run PaperCut Upgrade
    Write-Host "Running PaperCut Upgrade..."
    Write-LogEntry -Message "Running PaperCut Upgrade"
    Start-Process -FilePath $exePath -ArgumentList "/VERYSILENT"
} else {
    Write-Host "Error: Failed to download the upgrade file."
    Write-LogEntry -Message "Error: Failed to download the upgrade file."
}

# Echo step: Start PaperCut Server
Write-Host "Starting PaperCut Server..."
Write-LogEntry -Message "Starting PaperCut Server, please wait."
Start-Service -Name "PaperCut Application Server"


# Wait for 10 minutes
# Start-Sleep -Seconds (10 * 60)

# Get the current version number from version.txt
$versionFilePath = Join-Path -Path $pcInstallationLocation -ChildPath "server\version.txt"
if (Test-Path -Path $versionFilePath -PathType Leaf) {
    $currentVersion = Get-Content -Path $versionFilePath -Raw
    Write-LogEntry -Message "Current version: name-with-version=$currentVersion"
} else {
    Write-LogEntry -Message "Error: Failed to retrieve the current version number."
}

# Echo: Upgrade completed
Write-Host "Upgrade completed."
Write-LogEntry -Message "Upgrade completed."
