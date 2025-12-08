<#
.SYNOPSIS
    Check and install Windows Updates on Azure VMs.

.DESCRIPTION
    This script connects to Azure, retrieves available Windows Updates on a specified VM,
    and allows you to select and install specific KB updates.

.PARAMETER SubscriptionId
    The Azure Subscription ID where the VM resides.

.PARAMETER ResourceGroupName
    The Resource Group name containing the VM.

.PARAMETER VMName
    The name of the Azure VM to check for updates.

.EXAMPLE
    .\Check-AzureVMWindowsUpdate.ps1 -SubscriptionId "your-sub-id" -ResourceGroupName "your-rg" -VMName "your-vm"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$SubscriptionId,

    [Parameter(Mandatory = $false)]
    [string]$ResourceGroupName,

    [Parameter(Mandatory = $false)]
    [string]$VMName
)

#region Functions

function Connect-ToAzure {
    <#
    .SYNOPSIS
        Connects to Azure and sets the subscription context.
    #>
    param(
        [string]$SubscriptionId
    )

    try {
        # Check if already connected
        $context = Get-AzContext -ErrorAction SilentlyContinue
        if (-not $context) {
            Write-Host "Connecting to Azure..." -ForegroundColor Yellow
            Connect-AzAccount -ErrorAction Stop
        }
        else {
            Write-Host "Already connected to Azure as: $($context.Account.Id)" -ForegroundColor Green
        }

        # Set subscription if provided
        if ($SubscriptionId) {
            Set-AzContext -SubscriptionId $SubscriptionId -ErrorAction Stop | Out-Null
            Write-Host "Subscription set to: $SubscriptionId" -ForegroundColor Green
        }
    }
    catch {
        Write-Error "Failed to connect to Azure: $_"
        exit 1
    }
}

function Get-AzureVMList {
    <#
    .SYNOPSIS
        Retrieves list of Azure VMs for selection.
    #>
    Write-Host "`nRetrieving Azure VMs..." -ForegroundColor Yellow
    $vms = Get-AzVM -Status | Where-Object { $_.StorageProfile.OSDisk.OSType -eq 'Windows' }
    
    if ($vms.Count -eq 0) {
        Write-Warning "No Windows VMs found in the current subscription."
        return $null
    }

    Write-Host "`nAvailable Windows VMs:" -ForegroundColor Cyan
    Write-Host "=" * 80

    $index = 1
    $vmList = @()
    foreach ($vm in $vms) {
        $powerState = ($vm.PowerState -split ' ')[-1]
        $status = switch ($powerState) {
            'running' { "Running" }
            'deallocated' { "Deallocated" }
            'stopped' { "Stopped" }
            default { $powerState }
        }
        
        Write-Host ("{0,3}. {1,-30} | RG: {2,-25} | Status: {3}" -f $index, $vm.Name, $vm.ResourceGroupName, $status)
        $vmList += [PSCustomObject]@{
            Index             = $index
            Name              = $vm.Name
            ResourceGroupName = $vm.ResourceGroupName
            Status            = $status
        }
        $index++
    }

    Write-Host "=" * 80
    return $vmList
}

function Get-WindowsUpdatesScript {
    <#
    .SYNOPSIS
        Returns the script to run on the VM to check for Windows Updates.
    #>
    return @'
# Create Windows Update Session
$UpdateSession = New-Object -ComObject Microsoft.Update.Session
$UpdateSearcher = $UpdateSession.CreateUpdateSearcher()

Write-Output "Searching for available Windows Updates..."

try {
    # Search for updates that are not installed
    $SearchResult = $UpdateSearcher.Search("IsInstalled=0 and Type='Software'")
    
    if ($SearchResult.Updates.Count -eq 0) {
        Write-Output "NO_UPDATES_AVAILABLE"
    }
    else {
        Write-Output "UPDATES_FOUND"
        Write-Output "COUNT:$($SearchResult.Updates.Count)"
        
        foreach ($Update in $SearchResult.Updates) {
            $KBArticles = ($Update.KBArticleIDs | ForEach-Object { "KB$_" }) -join ","
            if ([string]::IsNullOrEmpty($KBArticles)) { $KBArticles = "N/A" }
            
            $Size = [math]::Round($Update.MaxDownloadSize / 1MB, 2)
            $Severity = if ($Update.MsrcSeverity) { $Update.MsrcSeverity } else { "Unspecified" }
            
            # Output format: KB|Title|Size(MB)|Severity|UpdateID
            Write-Output "UPDATE|$KBArticles|$($Update.Title)|$Size|$Severity|$($Update.Identity.UpdateID)"
        }
    }
}
catch {
    Write-Output "ERROR:$($_.Exception.Message)"
}
'@
}

function Get-InstallUpdatesScript {
    <#
    .SYNOPSIS
        Returns the script to install selected updates on the VM.
    #>
    param(
        [string[]]$UpdateIDs
    )

    $updateIDsString = $UpdateIDs -join ','
    
    return @"
`$UpdateIDs = '$updateIDsString' -split ','

# Create Windows Update Session
`$UpdateSession = New-Object -ComObject Microsoft.Update.Session
`$UpdateSearcher = `$UpdateSession.CreateUpdateSearcher()

Write-Output "Searching for updates to install..."

try {
    `$SearchResult = `$UpdateSearcher.Search("IsInstalled=0 and Type='Software'")
    
    # Create update collection for selected updates
    `$UpdatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
    
    foreach (`$Update in `$SearchResult.Updates) {
        if (`$UpdateIDs -contains `$Update.Identity.UpdateID) {
            # Accept EULA if required
            if (`$Update.EulaAccepted -eq `$false) {
                `$Update.AcceptEula()
            }
            `$UpdatesToInstall.Add(`$Update) | Out-Null
            Write-Output "Selected: `$(`$Update.Title)"
        }
    }
    
    if (`$UpdatesToInstall.Count -eq 0) {
        Write-Output "NO_UPDATES_SELECTED"
    }
    else {
        Write-Output "DOWNLOADING:`$(`$UpdatesToInstall.Count) updates"
        
        # Download updates
        `$Downloader = `$UpdateSession.CreateUpdateDownloader()
        `$Downloader.Updates = `$UpdatesToInstall
        `$DownloadResult = `$Downloader.Download()
        
        Write-Output "Download Result: `$(`$DownloadResult.ResultCode)"
        
        # Install updates
        Write-Output "INSTALLING:`$(`$UpdatesToInstall.Count) updates"
        
        `$Installer = `$UpdateSession.CreateUpdateInstaller()
        `$Installer.Updates = `$UpdatesToInstall
        `$InstallResult = `$Installer.Install()
        
        Write-Output "Installation Result: `$(`$InstallResult.ResultCode)"
        Write-Output "Reboot Required: `$(`$InstallResult.RebootRequired)"
        
        for (`$i = 0; `$i -lt `$UpdatesToInstall.Count; `$i++) {
            `$UpdateResult = `$InstallResult.GetUpdateResult(`$i)
            Write-Output "RESULT|`$(`$UpdatesToInstall.Item(`$i).Title)|`$(`$UpdateResult.ResultCode)"
        }
        
        if (`$InstallResult.RebootRequired) {
            Write-Output "REBOOT_REQUIRED"
        }
    }
}
catch {
    Write-Output "ERROR:`$(`$_.Exception.Message)"
}
"@
}

function Invoke-VMScript {
    <#
    .SYNOPSIS
        Runs a PowerShell script on the Azure VM using Run Command.
    #>
    param(
        [string]$ResourceGroupName,
        [string]$VMName,
        [string]$Script
    )

    try {
        $result = Invoke-AzVMRunCommand `
            -ResourceGroupName $ResourceGroupName `
            -VMName $VMName `
            -CommandId 'RunPowerShellScript' `
            -ScriptString $Script `
            -ErrorAction Stop

        return $result.Value[0].Message
    }
    catch {
        Write-Error "Failed to execute script on VM: $_"
        return $null
    }
}

function Show-AvailableUpdates {
    <#
    .SYNOPSIS
        Displays available updates in a formatted table.
    #>
    param(
        [array]$Updates
    )

    Write-Host "`nAvailable Windows Updates:" -ForegroundColor Cyan
    Write-Host "=" * 120

    $format = "{0,4} | {1,-15} | {2,-60} | {3,10} | {4,-12}"
    Write-Host ($format -f "#", "KB", "Title", "Size (MB)", "Severity") -ForegroundColor Yellow
    Write-Host "-" * 120

    $index = 1
    foreach ($update in $Updates) {
        $title = if ($update.Title.Length -gt 57) { $update.Title.Substring(0, 57) + "..." } else { $update.Title }
        Write-Host ($format -f $index, $update.KB, $title, $update.Size, $update.Severity)
        $index++
    }
    Write-Host "=" * 120
}

function Select-Updates {
    <#
    .SYNOPSIS
        Prompts user to select updates for installation.
    #>
    param(
        [array]$Updates
    )

    Write-Host "`nSelect updates to install:" -ForegroundColor Yellow
    Write-Host "  - Enter update numbers separated by commas (e.g., 1,3,5)"
    Write-Host "  - Enter 'all' to install all updates"
    Write-Host "  - Enter 'q' to quit without installing"
    Write-Host ""

    $selection = Read-Host "Your selection"

    if ($selection -eq 'q') {
        return $null
    }

    if ($selection -eq 'all') {
        return $Updates
    }

    $selectedIndices = $selection -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' }
    $selectedUpdates = @()

    foreach ($idx in $selectedIndices) {
        $i = [int]$idx - 1
        if ($i -ge 0 -and $i -lt $Updates.Count) {
            $selectedUpdates += $Updates[$i]
        }
        else {
            Write-Warning "Invalid selection: $idx (skipped)"
        }
    }

    return $selectedUpdates
}

#endregion Functions

#region Main Script

# Ensure Az module is available
if (-not (Get-Module -ListAvailable -Name Az.Compute)) {
    Write-Error "Az.Compute module is not installed. Please install it using: Install-Module -Name Az -Scope CurrentUser"
    exit 1
}

Import-Module Az.Compute -ErrorAction SilentlyContinue

# Connect to Azure
Connect-ToAzure -SubscriptionId $SubscriptionId

# If VM details not provided, show selection menu
if (-not $VMName -or -not $ResourceGroupName) {
    $vmList = Get-AzureVMList
    
    if (-not $vmList) {
        exit 1
    }

    Write-Host "`nEnter the number of the VM to check for updates: " -ForegroundColor Yellow -NoNewline
    $vmSelection = Read-Host

    if ($vmSelection -match '^\d+$') {
        $selectedVM = $vmList | Where-Object { $_.Index -eq [int]$vmSelection }
        if ($selectedVM) {
            $VMName = $selectedVM.Name
            $ResourceGroupName = $selectedVM.ResourceGroupName
            
            if ($selectedVM.Status -ne 'Running') {
                Write-Warning "VM '$VMName' is not running (Status: $($selectedVM.Status)). Please start the VM first."
                exit 1
            }
        }
        else {
            Write-Error "Invalid selection."
            exit 1
        }
    }
    else {
        Write-Error "Invalid input."
        exit 1
    }
}

Write-Host "`nChecking Windows Updates on VM: $VMName" -ForegroundColor Cyan
Write-Host "This may take a few minutes..." -ForegroundColor Yellow

# Get available updates
$checkScript = Get-WindowsUpdatesScript
$output = Invoke-VMScript -ResourceGroupName $ResourceGroupName -VMName $VMName -Script $checkScript

if (-not $output) {
    Write-Error "Failed to retrieve update information from VM."
    exit 1
}

# Parse the output
$lines = $output -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ }

if ($lines -contains "NO_UPDATES_AVAILABLE") {
    Write-Host "`nNo updates available. The VM is up to date!" -ForegroundColor Green
    exit 0
}

if ($lines | Where-Object { $_ -match "^ERROR:" }) {
    $errorLine = $lines | Where-Object { $_ -match "^ERROR:" }
    Write-Error "Error checking updates: $($errorLine -replace '^ERROR:', '')"
    exit 1
}

# Parse updates
$updates = @()
foreach ($line in $lines) {
    if ($line -match "^UPDATE\|") {
        $parts = $line -split '\|'
        if ($parts.Count -ge 6) {
            $updates += [PSCustomObject]@{
                KB       = $parts[1]
                Title    = $parts[2]
                Size     = $parts[3]
                Severity = $parts[4]
                UpdateID = $parts[5]
            }
        }
    }
}

if ($updates.Count -eq 0) {
    Write-Host "`nNo updates found or unable to parse update information." -ForegroundColor Yellow
    exit 0
}

# Display available updates
Show-AvailableUpdates -Updates $updates

# Get user selection
$selectedUpdates = Select-Updates -Updates $updates

if (-not $selectedUpdates -or $selectedUpdates.Count -eq 0) {
    Write-Host "`nNo updates selected. Exiting." -ForegroundColor Yellow
    exit 0
}

Write-Host "`nSelected updates for installation:" -ForegroundColor Cyan
foreach ($update in $selectedUpdates) {
    Write-Host "  - $($update.KB): $($update.Title)" -ForegroundColor White
}

# Confirm installation
Write-Host "`nDo you want to proceed with the installation? (y/n): " -ForegroundColor Yellow -NoNewline
$confirm = Read-Host

if ($confirm -ne 'y') {
    Write-Host "Installation cancelled." -ForegroundColor Yellow
    exit 0
}

Write-Host "`nInstalling selected updates on VM: $VMName" -ForegroundColor Cyan
Write-Host "This may take a while depending on the updates..." -ForegroundColor Yellow

# Install updates
$updateIDs = $selectedUpdates | ForEach-Object { $_.UpdateID }
$installScript = Get-InstallUpdatesScript -UpdateIDs $updateIDs
$installOutput = Invoke-VMScript -ResourceGroupName $ResourceGroupName -VMName $VMName -Script $installScript

if (-not $installOutput) {
    Write-Error "Failed to install updates on VM."
    exit 1
}

# Display installation results
Write-Host "`nInstallation Results:" -ForegroundColor Cyan
Write-Host "=" * 80

$installLines = $installOutput -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ }

foreach ($line in $installLines) {
    if ($line -match "^RESULT\|") {
        $parts = $line -split '\|'
        $resultCode = switch ($parts[2]) {
            "2" { "Succeeded" }
            "3" { "Succeeded with errors" }
            "4" { "Failed" }
            "5" { "Aborted" }
            default { "Unknown ($($parts[2]))" }
        }
        Write-Host "  $($parts[1]): $resultCode" -ForegroundColor $(if ($resultCode -eq "Succeeded") { "Green" } else { "Red" })
    }
    elseif ($line -eq "REBOOT_REQUIRED") {
        Write-Host "`n*** REBOOT REQUIRED ***" -ForegroundColor Yellow
        Write-Host "The VM needs to be restarted to complete the update installation." -ForegroundColor Yellow
        
        Write-Host "`nDo you want to restart the VM now? (y/n): " -ForegroundColor Yellow -NoNewline
        $rebootConfirm = Read-Host
        
        if ($rebootConfirm -eq 'y') {
            Write-Host "Restarting VM: $VMName..." -ForegroundColor Cyan
            Restart-AzVM -ResourceGroupName $ResourceGroupName -Name $VMName -NoWait
            Write-Host "VM restart initiated." -ForegroundColor Green
        }
    }
}

Write-Host "`nUpdate installation completed." -ForegroundColor Green

#endregion Main Script
