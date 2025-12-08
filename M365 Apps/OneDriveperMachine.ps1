<#
.SYNOPSIS
    Downloads and installs OneDrive for all users on the machine.

.DESCRIPTION
    This script downloads the OneDrive installer from Microsoft and installs it
    for all users on the local machine. It requires administrator privileges.
    Optimized for unattended execution in AVD Custom Image Template process.
    Detects and removes existing OneDrive installations before deploying the new version.

.PARAMETER MaxRetries
    Maximum number of retry attempts for download failures. Default is 3.

.PARAMETER TimeoutSeconds
    Timeout in seconds for download operations. Default is 300 (5 minutes).

.PARAMETER ForceUninstall
    Forces uninstall of existing OneDrive before installing new version. Default is $true.

.EXAMPLE
    .\OneDriveperMachine.ps1 -MaxRetries 5 -TimeoutSeconds 600

.NOTES
    Requires: Administrator privileges
    Author: AVD Automation
    Version: 2.0
#>

[CmdletBinding()]
param(
    [ValidateRange(1, 10)]
    [int]$MaxRetries = 3,
    
    [ValidateRange(60, 3600)]
    [int]$TimeoutSeconds = 300,
    
    [bool]$ForceUninstall = $true
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$WarningPreference = 'Continue'

# Inline constants for better performance
$script:OneDriveLink = "https://go.microsoft.com/fwlink/?linkid=844652"
$script:OneDriveInstaller = "$env:TEMP\OneDriveSetup.exe"
$script:LogPath = "$env:ProgramData\AVD\Logs"
$script:LogFile = "$script:LogPath\OneDriveInstall_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Ensure log directory exists
if (-not (Test-Path $script:LogPath)) {
    $null = New-Item -ItemType Directory -Path $script:LogPath -Force
}

# Start transcript
$null = Start-Transcript -Path $script:LogFile -Append

function Write-LogMessage {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )
    
    $LogEntry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
    
    switch ($Level) {
        'Info'    { Write-Output $LogEntry }
        'Warning' { Write-Warning $Message }
        'Error'   { Write-Host $LogEntry -ForegroundColor Red }
    }
}

function Test-AdminPrivileges {
    try {
        ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch { $false }
}

function Uninstall-OneDrive {
    $OneDriveFound = $false
    Write-LogMessage "Starting OneDrive uninstall detection"
    
    # Check per-machine installations (Program Files) - Combined path check
    $MachinePaths = @(
        "$env:ProgramFiles\Microsoft OneDrive",
        "${env:ProgramFiles(x86)}\Microsoft OneDrive"
    )
    
    foreach ($Path in $MachinePaths) {
        if (-not (Test-Path $Path)) { continue }
        
        Write-LogMessage "Found per-machine OneDrive installation at: $Path"
        $OneDriveSetup = Join-Path -Path $Path -ChildPath "OneDriveSetup.exe"
        
        if (Test-Path $OneDriveSetup) {
            try {
                Write-LogMessage "Uninstalling per-machine OneDrive from: $Path"
                $UninstallProcess = Start-Process -FilePath $OneDriveSetup `
                    -ArgumentList "/uninstall /allusers /silent" `
                    -Wait `
                    -PassThru `
                    -NoNewWindow `
                    -ErrorAction Stop
                
                $ExitCode = $UninstallProcess.ExitCode
                if ($ExitCode -eq 0) {
                    Write-LogMessage "Per-machine OneDrive uninstalled successfully"
                } else {
                    Write-LogMessage "Per-machine OneDrive uninstall exit code: $ExitCode" 'Warning'
                }
                $OneDriveFound = $true
            }
            catch {
                Write-LogMessage "Error uninstalling per-machine OneDrive: $_" 'Warning'
            }
        }
    }
    
    # Check per-user installations in AppData
    Get-ChildItem -Path "$env:SystemDrive\Users" -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $OneDrivePath = Join-Path $_.FullName "AppData\Local\Microsoft\OneDrive"
        if (Test-Path $OneDrivePath) {
            Write-LogMessage "Found per-user OneDrive installation at: $OneDrivePath"
            $OneDriveFound = $true
        }
    }
    
    # Check Registry - Combined path check
    $RegistryPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\OneDriveSetup.exe",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\OneDriveSetup.exe"
    )
    
    foreach ($RegPath in $RegistryPaths) {
        if (-not (Test-Path $RegPath -ErrorAction SilentlyContinue)) { continue }
        
        Write-LogMessage "Found OneDrive registry entry: $RegPath"
        $OneDriveFound = $true
        
        try {
            $UninstallString = (Get-ItemProperty -Path $RegPath -ErrorAction SilentlyContinue).UninstallString
            if ($UninstallString) {
                Write-LogMessage "Executing registry uninstall: $UninstallString"
                $null = Invoke-Expression $UninstallString -ErrorAction SilentlyContinue
                Write-LogMessage "Registry-based uninstall initiated"
            }
        }
        catch {
            Write-LogMessage "Error executing registry uninstall: $_" 'Warning'
        }
    }
    
    if ($OneDriveFound) {
        Write-LogMessage "Waiting 5 seconds for OneDrive uninstall processes to complete"
        Start-Sleep -Seconds 5
        return $true
    }
    
    Write-LogMessage "No existing OneDrive installation detected"
    return $false
}

function Download-OneDriveInstaller {
    param(
        [string]$Uri,
        [string]$OutFile,
        [int]$MaxRetries,
        [int]$TimeoutSeconds
    )
    
    for ($Attempt = 1; $Attempt -le $MaxRetries; $Attempt++) {
        try {
            Write-LogMessage "Downloading OneDrive installer from: $Uri (attempt $Attempt/$MaxRetries)"
            
            Invoke-WebRequest -Uri $Uri -OutFile $OutFile -TimeoutSec $TimeoutSeconds -UseBasicParsing -ErrorAction Stop
            
            # Validate file exists and has content
            $FileInfo = Get-Item $OutFile -ErrorAction Stop
            if ($FileInfo.Length -lt 1MB) {
                throw "Downloaded file size ($($FileInfo.Length) bytes) is suspiciously small"
            }
            
            Write-LogMessage "OneDrive installer downloaded successfully - Size: $('{0:F2}' -f ($FileInfo.Length / 1MB)) MB"
            return $true
        }
        catch {
            if ($Attempt -lt $MaxRetries) {
                $WaitTime = [Math]::Min([Math]::Pow(2, $Attempt - 1), 30)
                Write-LogMessage "Download failed (attempt $Attempt): $_ - Retrying in $WaitTime seconds" 'Warning'
                Start-Sleep -Seconds $WaitTime
            }
            else {
                throw "Failed to download OneDrive installer after $MaxRetries attempts: $_"
            }
        }
    }
}

function Install-OneDrive {
    param([string]$InstallerPath)
    
    if (-not (Test-Path $InstallerPath)) {
        throw "Installer not found at: $InstallerPath"
    }
    
    Write-LogMessage "Starting OneDrive installation (silent, all users)"
    
    $Process = Start-Process -FilePath $InstallerPath -ArgumentList "/allusers /silent" -Wait -PassThru -NoNewWindow
    
    if ($Process.ExitCode -ne 0) {
        throw "OneDrive installation failed with exit code: $($Process.ExitCode)"
    }
    
    Write-LogMessage "OneDrive installation completed successfully"
}

function Remove-InstallerFile {
    param([string]$FilePath)
    
    if (Test-Path $FilePath) {
        Remove-Item -Path $FilePath -Force -ErrorAction SilentlyContinue
        Write-LogMessage "Installer file cleanup completed"
    }
}

# Main execution
try {
    # Validate prerequisites
    if (-not (Test-AdminPrivileges)) {
        throw "Script requires administrator privileges"
    }
    
    Write-LogMessage "========== OneDrive Installation Script Started =========="
    Write-LogMessage "MaxRetries: $MaxRetries, TimeoutSeconds: $TimeoutSeconds, ForceUninstall: $ForceUninstall"
    
    # Uninstall existing OneDrive if enabled
    if ($ForceUninstall) {
        Uninstall-OneDrive
    }
    
    # Download installer with retry logic
    Download-OneDriveInstaller -Uri $script:OneDriveLink `
        -OutFile $script:OneDriveInstaller `
        -MaxRetries $MaxRetries `
        -TimeoutSeconds $TimeoutSeconds
    
    # Install OneDrive
    Install-OneDrive -InstallerPath $script:OneDriveInstaller
    
    # Clean up
    Remove-InstallerFile -FilePath $script:OneDriveInstaller
    
    Write-LogMessage "========== OneDrive Installation Script Completed Successfully =========="
}
catch {
    Write-LogMessage "CRITICAL ERROR: $_" 'Error'
    throw $_
}
finally {
    $null = Stop-Transcript
    Write-Output "Installation process complete. Log file: $script:LogFile"
}