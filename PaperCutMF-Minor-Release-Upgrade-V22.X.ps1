# Get the path of the current script
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

# Get PCMF Server install path from registry
$pcInstallationLocation = Get-ItemProperty -Path 'HKLM:\SOFTWARE\PaperCut MF' -Name 'InstallPath' | Select-Object -ExpandProperty 'InstallPath'

# Define the source and destination paths
$sourcePath = $pcInstallationLocation
$destinationPath = "C:\EDUIT\PaperCut-Backup\22.1.1"

# Define the log file path
$logFilePath = Join-Path -Path $scriptPath -ChildPath "papercut-upgrade-log.dat"

# Function to write log entries to the log file
function Write-LogEntry {
    param(
        [string]$Data
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[${timestamp}] $data"
    Write-Output -InputObject "$($data)"
    $logEntry | Out-File -FilePath $logFilePath -Append
}

# Echo step: Stop PaperCut Server
Write-LogEntry -Message "Stopping PaperCut Server"
Stop-Service -Name "PaperCut Application Server"

# Echo step: Create Backup Folder
$backupFolder = Join-Path -Path $destinationPath -ChildPath "PaperCut MF"
if (!(Test-Path -Path $backupFolder -PathType Container)) {
    Write-LogEntry -Message "Creating Backup Folder: $backupFolder"
    New-Item -Path $backupFolder -ItemType Directory
} else {
    Write-LogEntry -Message "Backup folder already exists. Skipping backup creation: $backupFolder"
}

# Copy the installation folder to the backup folder
Write-LogEntry -Message "Copying installation folder to backup: $sourcePath to $backupFolder"
Copy-Item -Path $sourcePath -Destination $backupFolder -Recurse -Force

# Echo step: Backup Complete
Write-LogEntry -Message "Backup complete: $backupFolder"

# Specify the URL of the exe file
$downloadUrl = "https://cdn.papercut.com/web/products/ng-mf/installers/mf/22.x/pcmf-setup-22.1.1.66714.exe"

# Specify the filename for the downloaded file
$fileName = "pcmf-setup-22.1.1.66714.exe"

# Specify the path to save the downloaded file
$exePath = Join-Path -Path $scriptPath -ChildPath $fileName

# Download the file
Write-LogEntry -Message "Downloading the upgrade file"
Invoke-WebRequest -Uri $downloadUrl -OutFile $exePath

# Check if the file was downloaded successfully
if (Test-Path -Path $exePath -PathType Leaf) {
    # Echo step: Run PaperCut Upgrade
    Write-LogEntry -Message "Running PaperCut Upgrade"
    Start-Process -FilePath $exePath -ArgumentList "/VERYSILENT"
} else {
    Write-LogEntry -Message "Error: Failed to download the upgrade file."
}

# Echo step: Start PaperCut Server
Write-LogEntry -Message "Starting PaperCut Server, please wait."
Start-Service -Name "PaperCut Application Server"

# Get the current version number from version.txt
Write-LogEntry -Message "Current version: $(Get-ItemProperty -Path 'HKLM:\SOFTWARE\PaperCut MF' -Name 'Version' | Select-Object -ExpandProperty 'Version')"

# Echo: Upgrade completed
Write-LogEntry -Message "Upgrade completed."
