<#
.SYNOPSIS
	App detection script
					 
.DESCRIPTION
	This script detects if an app is installed.
					 
.NOTES
	Version:        3.1
	LastMod Date:   2025-07-08
	Purpose/Change: Initial version for detection of specific app versions
    by AZ - https://blog.zarenko.net based on a script by thomas.froitzheim@iteracon.de
#>

Param
  (
    [parameter(Mandatory=$false)]
    [String[]]
    $param
  )

$Action = "Detect"
$PackageID = "#DN#"
$TargetVersion = "#VER#"
$logFile = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\$($PackageID)-$($TargetVersion)_$Action.log"

function Write-Log {
    param (
        [string]$message,
        [string]$logFilePath=$logfile
    )
    Add-Content -Path $logFilePath -Value "$(Get-Date): $message" -Force
    Write-Output $message
}

Function Test-AppInstallation {
    param (
        [string]$AppName,
        [string]$TargetVersion
    )

    $searchPattern = $AppName + "*"
    $targetVersion = [System.Version]::Parse($TargetVersion)
    $registryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    )

    Try {
        $applicationData = @()
        foreach ($registryPath in $registryPaths) {
            try {
                $applicationData += Get-ChildItem -Path $registryPath -ErrorAction SilentlyContinue |
                    Get-ItemProperty |
                    Where-Object { $_.DisplayName -like $searchPattern } |
                    Select-Object -Property DisplayName, DisplayVersion
            } catch {
                continue
            }
        }

        if ($applicationData.Count -eq 0) {
            Write-Log "App $PackageID NOT found!"
            Write-Log "$Action finished." 
            exit 1
        }

        $installedVersion = [System.Version]::Parse($applicationData[0].DisplayVersion)
        if ($installedVersion -ge $targetVersion) {
            Write-Log "App $PackageID found! (Exact or newer version found.)"
            Write-Log "$Action finished."
            exit 0
        } else {
            Write-Log "App $PackageID NOT found! (Older version found.)"
            Write-Log "$Action finished."
            exit 1
        }
    } catch {
            Write-Log "App $PackageID NOT found! (Error during detection.)"
            Write-Log "$Action finished."            
        exit 2
    }
}

Write-Log "------------------------------------"
Write-Log "$Action $PackageID - $TargetVersion"

Test-AppInstallation -AppName $PackageID -TargetVersion $TargetVersion
