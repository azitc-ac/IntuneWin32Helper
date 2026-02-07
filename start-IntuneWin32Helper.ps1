$toolVersion = "1.3"
$rootDir = $PSScriptRoot
if (-not $PSScriptRoot) { $rootDir = "C:\Users\alex\OneDrive - AZITC\Tools\Administration\IntuneWin32Helper" }
$cleandatetime = get-date -uformat "%Y-%m-%d_%H-%M-%S"
$log = $rootDir + "\Logs\" + $cleandatetime + ".log"
Start-Transcript $log

$config = Get-Content -Raw -Path "$rootDir\config\config.json" -Encoding UTF8 | ConvertFrom-Json

$cloudName = $config.cloudName
$ApiKey = $config.apiKey
$ApiSecret = $config.apiSecret
$packetRoot = $config.packetRoot

if (-not (Test-Path $packetRoot)) { md $packetRoot }
. "$rootDir\functions\functions.ps1"

check-prereqs


Add-Type -AssemblyName PresentationFramework | Out-Null
Add-Type -AssemblyName System.Windows.Forms    | Out-Null


$continue = $true
while ($continue) {
    $choice = Show-StartDialog
    switch ($choice) {
        'CreateNew'          { "-> Start packaging assistant"; createApps }
        'CreateNewAndDeploy' { "-> Start packaging + deployment"; createApps -createAndDeploy }
        'DeployExisting'     { "-> Start deployment of existing app"; deployApps }
        'Cancel'             { "-> Cancelled"; $continue = $false }
        'Closed'             { "-> Closed with [X]"; $continue = $false }
        default              { "-> Unexpected: $choice"; $continue = $false }
    }
}
Stop-Transcript
