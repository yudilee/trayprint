<#
.SYNOPSIS
    Registers TrayPrint as a Windows Scheduled Task that runs at system startup
    with SYSTEM privileges (highest available).

.DESCRIPTION
    Creates a Scheduled Task named "TrayPrintAgent" that launches
    trayprint.exe --silent at system startup. The task runs as the SYSTEM
    account with highest run level, ensuring the print agent has full
    administrator privileges for printer management.

    The task is configured to:
    - Start at system boot (before any user logs in)
    - Run as NT AUTHORITY\SYSTEM
    - Restart up to 3 times on failure (1 minute apart)
    - Start even if on battery
    - Not stop if switching to battery

.PARAMETER InstallPath
    The full path to the directory containing trayprint.exe.

.PARAMETER TaskName
    The name of the scheduled task (default: TrayPrintAgent).

.PARAMETER Unregister
    If specified, removes the scheduled task instead of creating it.

.EXAMPLE
    .\Register-TrayPrintTask.ps1 -InstallPath "C:\Program Files\PrintHub\TrayPrint"

.EXAMPLE
    .\Register-TrayPrintTask.ps1 -Unregister
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$InstallPath,

    [Parameter(Mandatory = $false)]
    [string]$TaskName = "TrayPrintAgent",

    [switch]$Unregister
)

# ── Unregister mode ──
if ($Unregister) {
    Write-Host "Removing scheduled task '$TaskName'..."
    try {
        $existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        if ($existing) {
            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
            Write-Host "[SUCCESS] Scheduled task '$TaskName' removed."
        } else {
            Write-Host "[INFO] Scheduled task '$TaskName' does not exist."
        }
    } catch {
        Write-Host "[WARNING] Failed to remove scheduled task: $_"
    }
    return
}

# ── Validate parameters ──
if (-not $InstallPath) {
    Write-Host "[ERROR] -InstallPath is required when not using -Unregister."
    exit 1
}

$exePath = Join-Path $InstallPath "trayprint.exe"
if (-not (Test-Path $exePath)) {
    Write-Host "[ERROR] trayprint.exe not found at: $exePath"
    exit 1
}

# ── Check if task already exists ──
$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existing) {
    Write-Host "[INFO] Task '$TaskName' already exists. Removing and recreating..."
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

# ── Create the scheduled task ──
try {
    Write-Host "Creating scheduled task '$TaskName'..."
    Write-Host "  Executable: $exePath"
    Write-Host "  Arguments:  --silent"
    Write-Host "  User:       SYSTEM"
    Write-Host "  Trigger:    AtStartup"

    $action = New-ScheduledTaskAction `
        -Execute $exePath `
        -Argument "--silent" `
        -WorkingDirectory $InstallPath

    $trigger = New-ScheduledTaskTrigger -AtStartup

    $settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -RestartCount 3 `
        -RestartInterval (New-TimeSpan -Minutes 1) `
        -Compatibility Win8

    $principal = New-ScheduledTaskPrincipal `
        -UserId "SYSTEM" `
        -LogonType ServiceAccount `
        -RunLevel Highest

    Register-ScheduledTask `
        -TaskName $TaskName `
        -Action $action `
        -Trigger $trigger `
        -Settings $settings `
        -Principal $principal `
        -Force

    Write-Host "[SUCCESS] Scheduled task '$TaskName' created."
    Write-Host "         The task will run at every system startup as SYSTEM."

    # Show the task details
    Get-ScheduledTask -TaskName $TaskName | Format-List TaskName, State, Actions, Triggers, Principal, Settings

} catch {
    Write-Host "[ERROR] Failed to create scheduled task: $_"
    exit 1
}
