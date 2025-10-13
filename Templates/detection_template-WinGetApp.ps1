<#
.SYNOPSIS
	App detection script
					 
.DESCRIPTION
	This script detects if an app is installed.
					 
.NOTES
	Version:        2.2
	LastMod Date:   2024-01-31
	Purpose/Change: added support for ARM/x86, optimized logging, fixes
    # inspired by https://github.com/FlorianSLZ/Intune-Win32-Deployer
    # adapted by AZ - https://blog.zarenko.net
#>

Param
  (
    [parameter(Mandatory=$false)]
    [String[]]
    $param
  )

$Action = "Detect"
$PackageID = "WINGETPROGRAMID"
$logFile = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\$($PackageID)_$Action.log"

function Write-Log {
    param (
        [string]$message,
        [string]$logFilePath=$logfile
    )
    Add-Content -Path $logFilePath -Value "$(Get-Date): $message" -Force
    Write-Output $message
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
    Write-Log "Winget found."
} elseif ($wingetPaths.Count -eq 1) {
    $wingetPath = $wingetPaths[0].Path
    Write-Log "Path: $wingetPath"
    Write-Log "Winget found."

} else {
    Write-Log "Winget NOT installed, exiting."
    exit 1
}

$wingetPrg_Existing = & $wingetPath list --id $PackageID --exact --accept-source-agreements
if ($wingetPrg_Existing -like "*$PackageID*"){
    Write-Log "App $PackageID found!"
    Write-Log "$Action finished."
    exit 0
}
else{
    Write-Log "App $PackageID NOT found!"
    Write-Log "$Action finished."
    exit 1
}