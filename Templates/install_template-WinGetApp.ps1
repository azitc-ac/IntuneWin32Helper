Param (
    [parameter(Mandatory=$false)]
    [String[]] $param
)

$Action = "Install"
$PackageID = "WINGETPROGRAMID"
$logFile = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\$($PackageID)_$Action.log"

function Write-Log {
    param (
        [string]$message,
        [string]$logFilePath = $logFile
    )
    Add-Content -Path $logFilePath -Value "$(Get-Date): $message" -Force
}

Write-Log "------------------------------------"
Write-Log "$Action $PackageID"
Write-Log "Resolving winget_exe"

try {
    $wingetPaths = Resolve-Path "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller*\winget.exe" -ErrorAction Stop
} catch {
    Write-Log "Winget not installed or path resolution failed: $_"
    exit 1
}

if ($wingetPaths.Count -gt 1) {
    $wingetPaths = $wingetPaths | Sort-Object { (Get-Item $_.Path).CreationTime } -Descending
    $wingetPath = $wingetPaths[0].Path
} elseif ($wingetPaths.Count -eq 1) {
    $wingetPath = $wingetPaths[0].Path
} else {
    Write-Log "Winget executable not found."
    exit 1
}

Write-Log "Using winget path: $wingetPath"

$accpackagree = "--accept-package-agreements"
$arguments = @(
    $Action
    "--exact"
    "--id", $PackageID
    "--silent"
    "--accept-source-agreements"
    $accpackagree
    "--scope=machine"
) + $param

Start-Transcript -Path $logFile -Append -NoClobber -Force
try {
    & $wingetPath @arguments
    $exitCode = $LASTEXITCODE
    Stop-Transcript    
} catch {
    $errmsg = $_
    Stop-Transcript
    $exitCode = 1
}
Write-Log "Winget exited with code $exitCode"
if($errmsg){Write-Log "Exception during winget execution: $errmsg"}
