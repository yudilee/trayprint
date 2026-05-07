<#
.SYNOPSIS
    Removes the TrayPrint Scheduled Task.

.DESCRIPTION
    Unregisters the "TrayPrintAgent" scheduled task that was created by
    Register-TrayPrintTask.ps1. This is called during MSI uninstall.

.PARAMETER TaskName
    The name of the scheduled task to remove (default: TrayPrintAgent).

.EXAMPLE
    .\Unregister-TrayPrintTask.ps1

.EXAMPLE
    .\Unregister-TrayPrintTask.ps1 -TaskName "TrayPrintAgent"
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$TaskName = "TrayPrintAgent"
)

Write-Host "Removing scheduled task '$TaskName'..."

try {
    $existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($existing) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Host "[SUCCESS] Scheduled task '$TaskName' removed."
    } else {
        Write-Host "[INFO] Scheduled task '$TaskName' does not exist — nothing to remove."
    }
} catch {
    Write-Host "[WARNING] Failed to remove scheduled task: $_"
    # Non-zero exit to signal the installer that cleanup had an issue
    exit 1
}
