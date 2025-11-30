<#
.SYNOPSIS
    Downloads and installs OneDrive for all users on the machine.

.DESCRIPTION
    This script downloads the OneDrive installer from Microsoft and installs it
    for all users on the local machine. It requires administrator privileges.
    Optimized for unattended execution in AVD Custom Image Template process.

.PARAMETER MaxRetries
    Maximum number of retry attempts for download failures. Default is 3.

.PARAMETER TimeoutSeconds
    Timeout in seconds for download operations. Default is 300 (5 minutes).

.EXAMPLE
    .\OneDriveperMachine.ps1 -Verbose

.NOTES
    Requires: Administrator privileges
    Author: AVD Automation
#>

[CmdletBinding()]
param(
    [int]$MaxRetries = 3,
    [int]$TimeoutSeconds = 300
)

# Configure error handling for unattended execution
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# Setup logging to file for unattended diagnostics
$LogPath = "$env:ProgramData\AVD\Logs"
$LogFile = "$LogPath\OneDriveInstall_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

if (-not (Test-Path $LogPath)) {
    New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
}

# Transcription for complete audit trail in unattended environments
Start-Transcript -Path $LogFile -Append | Out-Null

function Write-LogMessage {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )
    
    $Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $LogEntry = "[$Timestamp] [$Level] $Message"
    
    switch ($Level) {
        'Info' { Write-Output $LogEntry }
        'Warning' { Write-Warning $LogEntry }
        'Error' { Write-Error $LogEntry }
    }
}

# Validate administrator privileges
$Principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $Principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-LogMessage "Script requires administrator privileges" 'Error'
    throw "This script requires administrator privileges. Please run as Administrator."
}

try {
    $OneDriveLink = "https://go.microsoft.com/fwlink/?linkid=844652"
    $OneDriveInstaller = "$env:TEMP\OneDriveSetup.exe"
    $RetryCount = 0

    # Download with retry logic and timeout
    Write-LogMessage "Attempting to download OneDrive installer from: $OneDriveLink"
    
    while ($RetryCount -lt $MaxRetries) {
        try {
            Invoke-WebRequest -Uri $OneDriveLink `
                -OutFile $OneDriveInstaller `
                -TimeoutSec $TimeoutSeconds `
                -UseBasicParsing `
                -ErrorAction Stop
            
            Write-LogMessage "OneDrive installer downloaded successfully"
            break
        }
        catch {
            $RetryCount++
            if ($RetryCount -lt $MaxRetries) {
                $WaitTime = [Math]::Pow(2, $RetryCount - 1)
                Write-LogMessage "Download failed (attempt $RetryCount of $MaxRetries). Retrying in $WaitTime seconds..." 'Warning'
                Start-Sleep -Seconds $WaitTime
            }
            else {
                throw "Download failed after $MaxRetries attempts: $_"
            }
        }
    }

    # Validate that the file was downloaded successfully
    if (-not (Test-Path $OneDriveInstaller)) {
        throw "OneDrive installer download failed - file not found at $OneDriveInstaller"
    }

    $FileSize = (Get-Item $OneDriveInstaller).Length
    Write-LogMessage "Installer file size: $($FileSize / 1MB) MB"

    # Silent installation for unattended execution
    Write-LogMessage "Starting OneDrive installation with /allusers /silent parameters"
    $Process = Start-Process -FilePath $OneDriveInstaller `
        -ArgumentList "/allusers /silent" `
        -Wait `
        -PassThru `
        -NoNewWindow `
        -ErrorAction Stop

    # Check if installation was successful (exit code 0)
    if ($Process.ExitCode -eq 0) {
        Write-LogMessage "OneDrive installation completed successfully (exit code: 0)"

        # Clean up installer
        Write-LogMessage "Removing installer from $OneDriveInstaller"
        if (Test-Path $OneDriveInstaller) {
            Remove-Item $OneDriveInstaller -Force -ErrorAction Stop
            Write-LogMessage "Installer removed successfully"
        }
    }
    else {
        throw "OneDrive installation failed with exit code: $($Process.ExitCode)"
    }
}
catch {
    Write-LogMessage "CRITICAL: An error occurred during OneDrive installation: $_" 'Error'
    Stop-Transcript | Out-Null
    throw $_
}
finally {
    Stop-Transcript | Out-Null
    Write-Output "Installation process complete. Log file: $LogFile"
}