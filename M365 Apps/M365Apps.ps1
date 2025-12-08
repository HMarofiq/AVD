#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Installs Microsoft 365 Apps for Enterprise with Shared Device Activation for Azure Virtual Desktop.

.DESCRIPTION
    This script performs the following:
    1. Uninstalls existing Office installations (MSI-based and Click-to-Run)
    2. Downloads and installs Office Deployment Tool (ODT)
    3. Installs Microsoft 365 Apps with Shared Device Activation enabled
    4. Excludes: OneDrive, Lync, Groove, Access, Publisher, and Teams
    
    Note: OneDrive should be installed separately using per-machine mode
    Note: Teams should be installed separately using Teams VDI installation

.NOTES
    Author: Azure Virtual Desktop Deployment
    Version: 1.0
    Requirements: 
        - Run as Administrator
        - Internet connectivity to download ODT and Office files
        - Valid Microsoft 365 Apps license with Shared Computer Activation support
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$LogPath = "$env:TEMP\M365AppsInstall",
    
    [Parameter(Mandatory = $false)]
    [string]$ODTDownloadPath = "$env:TEMP\ODT"
)

#region Functions

function Write-Log {
    <#
    .SYNOPSIS
        Writes log entries to console and log file.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Create log directory if it doesn't exist
    if (-not (Test-Path -Path $LogPath)) {
        New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
    }
    
    $logFile = Join-Path -Path $LogPath -ChildPath "M365AppsInstall_$(Get-Date -Format 'yyyyMMdd').log"
    Add-Content -Path $logFile -Value $logMessage
    
    switch ($Level) {
        'Info'    { Write-Host $logMessage -ForegroundColor Cyan }
        'Warning' { Write-Host $logMessage -ForegroundColor Yellow }
        'Error'   { Write-Host $logMessage -ForegroundColor Red }
    }
}

function Test-AdminPrivileges {
    <#
    .SYNOPSIS
        Verifies the script is running with administrator privileges.
    #>
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Uninstall-ExistingOffice {
    <#
    .SYNOPSIS
        Removes existing Office installations including MSI and Click-to-Run versions.
    #>
    Write-Log -Message "Starting removal of existing Office installations..." -Level Info
    
    # Check for Click-to-Run Office installations
    $c2rPath = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe"
    if (Test-Path -Path $c2rPath) {
        Write-Log -Message "Found Click-to-Run Office installation. Initiating uninstall..." -Level Info
        try {
            $uninstallArgs = "scenario=install scenariosubtype=ARP sourcetype=None productstoremove=AllProducts"
            Start-Process -FilePath $c2rPath -ArgumentList $uninstallArgs -Wait -NoNewWindow
            Write-Log -Message "Click-to-Run Office uninstallation completed." -Level Info
        }
        catch {
            Write-Log -Message "Error uninstalling Click-to-Run Office: $_" -Level Warning
        }
    }
    
    # Check for MSI-based Office installations
    $msiOfficeProducts = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like "*Microsoft Office*" -or $_.Name -like "*Microsoft 365*" }
    
    foreach ($product in $msiOfficeProducts) {
        Write-Log -Message "Uninstalling MSI product: $($product.Name)" -Level Info
        try {
            $product.Uninstall() | Out-Null
            Write-Log -Message "Successfully uninstalled: $($product.Name)" -Level Info
        }
        catch {
            Write-Log -Message "Error uninstalling $($product.Name): $_" -Level Warning
        }
    }
    
    Write-Log -Message "Office uninstallation process completed." -Level Info
}

function Get-OfficeDeploymentTool {
    <#
    .SYNOPSIS
        Downloads and extracts the Office Deployment Tool.
    #>
    Write-Log -Message "Downloading Office Deployment Tool..." -Level Info
    
    # Create ODT directory
    if (-not (Test-Path -Path $ODTDownloadPath)) {
        New-Item -ItemType Directory -Path $ODTDownloadPath -Force | Out-Null
    }
    
    # Download ODT from Office CDN
    $odtExeUrl = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
    $setupPath = Join-Path -Path $ODTDownloadPath -ChildPath "setup.exe"
    
    try {
        Write-Log -Message "Downloading ODT setup.exe from Office CDN..." -Level Info
        
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $odtExeUrl -OutFile $setupPath -UseBasicParsing
        
        if (Test-Path -Path $setupPath) {
            Write-Log -Message "Office Deployment Tool downloaded successfully to: $setupPath" -Level Info
            return $setupPath
        }
        else {
            throw "Failed to download Office Deployment Tool."
        }
    }
    catch {
        Write-Log -Message "Error downloading ODT: $_" -Level Error
        throw
    }
}

function New-M365AppsConfiguration {
    <#
    .SYNOPSIS
        Creates the configuration XML file for Office installation with Shared Device Activation.
        Excludes: OneDrive, Lync, Groove, Access, Publisher, and Teams
    #>
    Write-Log -Message "Creating M365 Apps configuration file..." -Level Info
    
    $configContent = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="MonthlyEnterprise">
    <Product ID="O365ProPlusRetail">
      <Language ID="en-us" />
      <Language ID="MatchOS" />
      <ExcludeApp ID="Access" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="OneDrive" />
      <ExcludeApp ID="Publisher" />
      <ExcludeApp ID="Teams" />
    </Product>
  </Add>
  <Property Name="SharedComputerLicensing" Value="1" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="DeviceBasedLicensing" Value="0" />
  <Property Name="SCLCacheOverride" Value="0" />
  <Property Name="PinIconsToTaskbar" Value="FALSE" />
  <Updates Enabled="FALSE" />
  <RemoveMSI />
  <Display Level="None" AcceptEULA="TRUE" />
  <Logging Level="Standard" Path="$LogPath" />
</Configuration>
"@
    
    $configPath = Join-Path -Path $ODTDownloadPath -ChildPath "configuration.xml"
    $configContent | Out-File -FilePath $configPath -Encoding UTF8 -Force
    
    Write-Log -Message "Configuration file created at: $configPath" -Level Info
    return $configPath
}

function Install-M365Apps {
    <#
    .SYNOPSIS
        Installs Microsoft 365 Apps using the Office Deployment Tool.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$SetupPath,
        
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )
    
    Write-Log -Message "Starting Microsoft 365 Apps installation..." -Level Info
    Write-Log -Message "Setup path: $SetupPath" -Level Info
    Write-Log -Message "Configuration path: $ConfigPath" -Level Info
    
    try {
        $installArgs = "/configure `"$ConfigPath`""
        Write-Log -Message "Running: $SetupPath $installArgs" -Level Info
        
        $process = Start-Process -FilePath $SetupPath -ArgumentList $installArgs -Wait -PassThru -NoNewWindow
        
        if ($process.ExitCode -eq 0) {
            Write-Log -Message "Microsoft 365 Apps installation completed successfully." -Level Info
        }
        else {
            Write-Log -Message "Installation completed with exit code: $($process.ExitCode)" -Level Warning
        }
    }
    catch {
        Write-Log -Message "Error during installation: $_" -Level Error
        throw
    }
}

function Set-AVDOptimizations {
    <#
    .SYNOPSIS
        Applies AVD-specific registry optimizations for Office.
    #>
    Write-Log -Message "Applying AVD optimizations for Office..." -Level Info
    
    try {
        # Ensure IsWVDEnvironment is set for Teams VDI optimization (when Teams is installed separately)
        $teamsRegPath = "HKLM:\SOFTWARE\Microsoft\Teams"
        if (-not (Test-Path -Path $teamsRegPath)) {
            New-Item -Path $teamsRegPath -Force | Out-Null
        }
        New-ItemProperty -Path $teamsRegPath -Name "IsWVDEnvironment" -PropertyType DWORD -Value 1 -Force | Out-Null
        Write-Log -Message "Set IsWVDEnvironment registry key for Teams VDI optimization." -Level Info
        
        # Disable Office automatic updates (updates should be managed through image updates)
        $officeUpdatePath = "HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate"
        if (-not (Test-Path -Path $officeUpdatePath)) {
            New-Item -Path $officeUpdatePath -Force | Out-Null
        }
        New-ItemProperty -Path $officeUpdatePath -Name "enableautomaticupdates" -PropertyType DWORD -Value 0 -Force | Out-Null
        Write-Log -Message "Disabled Office automatic updates." -Level Info
        
        # Configure Office telemetry settings
        $officeTelemetryPath = "HKLM:\SOFTWARE\Policies\Microsoft\office\16.0\common\clienttelemetry"
        if (-not (Test-Path -Path $officeTelemetryPath)) {
            New-Item -Path $officeTelemetryPath -Force | Out-Null
        }
        New-ItemProperty -Path $officeTelemetryPath -Name "sendtelemetry" -PropertyType DWORD -Value 3 -Force | Out-Null
        Write-Log -Message "Configured Office telemetry settings." -Level Info
        
        Write-Log -Message "AVD optimizations applied successfully." -Level Info
    }
    catch {
        Write-Log -Message "Error applying AVD optimizations: $_" -Level Warning
    }
}

function Test-M365AppsInstallation {
    <#
    .SYNOPSIS
        Verifies that Microsoft 365 Apps was installed successfully.
    #>
    Write-Log -Message "Verifying Microsoft 365 Apps installation..." -Level Info
    
    $officePath = "C:\Program Files\Microsoft Office\root\Office16"
    $expectedApps = @("WINWORD.EXE", "EXCEL.EXE", "POWERPNT.EXE", "OUTLOOK.EXE")
    $installedApps = @()
    $missingApps = @()
    
    foreach ($app in $expectedApps) {
        $appPath = Join-Path -Path $officePath -ChildPath $app
        if (Test-Path -Path $appPath) {
            $installedApps += $app
            Write-Log -Message "Verified: $app is installed." -Level Info
        }
        else {
            $missingApps += $app
            Write-Log -Message "Missing: $app not found." -Level Warning
        }
    }
    
    # Check Shared Computer Activation is enabled
    $sclPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    if (Test-Path -Path $sclPath) {
        $sclValue = Get-ItemProperty -Path $sclPath -Name "SharedComputerLicensing" -ErrorAction SilentlyContinue
        if ($sclValue.SharedComputerLicensing -eq "1") {
            Write-Log -Message "Shared Computer Activation is enabled." -Level Info
        }
        else {
            Write-Log -Message "Shared Computer Activation may not be properly configured." -Level Warning
        }
    }
    
    if ($missingApps.Count -eq 0) {
        Write-Log -Message "All expected Microsoft 365 Apps are installed successfully." -Level Info
        return $true
    }
    else {
        Write-Log -Message "Some apps are missing: $($missingApps -join ', ')" -Level Warning
        return $false
    }
}

#endregion Functions

#region Main Script

try {
    Write-Log -Message "========================================" -Level Info
    Write-Log -Message "Microsoft 365 Apps Installation Script" -Level Info
    Write-Log -Message "For Azure Virtual Desktop with Shared Device Activation" -Level Info
    Write-Log -Message "========================================" -Level Info
    
    # Verify admin privileges
    if (-not (Test-AdminPrivileges)) {
        throw "This script must be run with Administrator privileges."
    }
    Write-Log -Message "Administrator privileges verified." -Level Info
    
    # Step 1: Uninstall existing Office
    Write-Log -Message "Step 1: Removing existing Office installations..." -Level Info
    Uninstall-ExistingOffice
    
    # Step 2: Download Office Deployment Tool
    Write-Log -Message "Step 2: Downloading Office Deployment Tool..." -Level Info
    $setupPath = Get-OfficeDeploymentTool
    
    # Step 3: Create configuration file
    Write-Log -Message "Step 3: Creating configuration file..." -Level Info
    $configPath = New-M365AppsConfiguration
    
    # Step 4: Install Microsoft 365 Apps
    Write-Log -Message "Step 4: Installing Microsoft 365 Apps..." -Level Info
    Install-M365Apps -SetupPath $setupPath -ConfigPath $configPath
    
    # Step 5: Apply AVD optimizations
    Write-Log -Message "Step 5: Applying AVD optimizations..." -Level Info
    Set-AVDOptimizations
    
    # Step 6: Verify installation
    Write-Log -Message "Step 6: Verifying installation..." -Level Info
    $installSuccess = Test-M365AppsInstallation
    
    Write-Log -Message "========================================" -Level Info
    Write-Log -Message "Installation process completed." -Level Info
    Write-Log -Message "========================================" -Level Info
    Write-Log -Message "" -Level Info
    Write-Log -Message "IMPORTANT NEXT STEPS:" -Level Info
    Write-Log -Message "1. Install OneDrive using per-machine mode - Run OneDriveperMachine.ps1" -Level Info
    Write-Log -Message "2. Install Microsoft Teams using Teams VDI installation method" -Level Info
    Write-Log -Message "   - Download: https://go.microsoft.com/fwlink/?linkid=2243204" -Level Info
    Write-Log -Message "   - Run: teamsbootstrapper.exe -p" -Level Info
    Write-Log -Message "========================================" -Level Info
}
catch {
    Write-Log -Message "Script failed with error: $_" -Level Error
    Write-Log -Message $_.ScriptStackTrace -Level Error
    exit 1
}

#endregion Main Script
