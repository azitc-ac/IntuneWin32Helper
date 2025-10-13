<#
.SYNOPSIS
	App (un-)installation script
					 
.DESCRIPTION
	This script installs or uninstalls the app.
					 
.NOTES
	Version:        2.0
	LastMod Date:   2024-12-15
	Purpose/Change: added support for ARM/x86, optimized logging
    # inspired by https://github.com/FlorianSLZ/Intune-Win32-Deployer
    # adapted by AZ - https://blog.zarenko.net
#>

Param
  (
    [parameter(Mandatory=$false)]
    [String[]]
    $param
  )

$Action = "Uninstall"
$PackageID = "WINGETPROGRAMID"
$logFile = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\$($PackageID)_$Action.log"

function Write-Log {
    param (
        [string]$message,
        [string]$logFilePath=$logfile
    )
    Add-Content -Path $logFilePath -Value "$(Get-Date): $message" -Force
}
  
Write-Log "------------------------------------"
Write-Log "$Action $PackageID"
Write-Log "Resolving winget_exe"
$wingetPaths = Resolve-Path "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller*\winget.exe"

if ($wingetPaths.Count -gt 1) {
    # Sortiere die Pfade nach dem Installationsdatum und wähle den neuesten aus
    $wingetPaths = $wingetPaths | Sort-Object { (Get-Item $_.Path).CreationTime } -Descending
    $wingetPath = $wingetPaths[0].Path
    Write-Log "Path: $wingetPath"
} elseif ($wingetPaths.Count -eq 1) {
    $wingetPath = $wingetPaths[0].Path
    Write-Log "Path: $wingetPath"
} else {
    Write-Log "Winget not installed."
    exit
}

if($Action -eq "Install"){$accpackagree = "--accept-package-agreements"}

Start-Transcript $logfile -Append -NoClobber -Force 
& $wingetPath $Action --exact --id $PackageID --silent --accept-source-agreements --scope=machine $accpackagree $param 
Stop-Transcript
Write-Log "$Action finished."