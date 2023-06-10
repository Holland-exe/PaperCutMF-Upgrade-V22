# PaperCut MF 22.X Backup and Upgrade Script 

## Description

This script automates the backup and upgrade process for PaperCut MF, a print management software. It stops the PaperCut server, creates a backup of the installation folder, downloads the latest upgrade file, runs the upgrade silently, starts the PaperCut server, and logs the entire process.

## Prerequisites

- PowerShell 3.0 or later.
- Access to the PaperCut installation folder on the local machine.
- Internet connectivity to download the upgrade file.

## Usage

1. Clone or download this repository to your local machine.
2. Open PowerShell and navigate to the directory containing the script.
3. Run the script using the following command:

   ```.\papercut-upgrade.ps1```

**Note You may need to set your execution policy to unrestricted, alternatively run the following command within the working directory of the papercut-upgrade.ps1 script.**
   ```Set-ExecutionPolicy Bypass -Scope Process; & ".\PaperCutMF-Minor-Release-Upgrade-V22.X.ps1"```

4. Follow the on-screen instructions and monitor the progress.

## Configuration

Before running the script, you may need to adjust the following configuration options:

- **$pcInstallationLocationC**: The path to the PaperCut installation folder on the C drive. Modify this variable if your installation is located elsewhere.
- **$pcInstallationLocationE**: The path to the PaperCut installation folder on the E drive. Modify this variable if your installation is located elsewhere.
- **$destinationPath**: The path where the backup folder will be created. Modify this variable to specify a different destination.
- **$downloadUrl**: The URL of the PaperCut upgrade file. Modify this variable if you need to download a different version.
- **$fileName**: The filename for the downloaded upgrade file. Modify this variable if you want to use a different name.
- **$exePath**: The path to save the downloaded upgrade file. Modify this variable to specify a different location.
- **$logFilePath**: The path to the log file. Modify this variable to specify a different location or filename.

## Logging

The script logs all the major steps and events to a log file for reference. The default log file is named "papercut-upgrade-log.dat" and is stored in the same directory as the script. Each log entry includes a timestamp and a description of the event or step.

## Disclaimer

This script is provided as-is without any warranty. Use it at your own risk. Review the script and customise it according to your specific requirements before running it in a production environment.
