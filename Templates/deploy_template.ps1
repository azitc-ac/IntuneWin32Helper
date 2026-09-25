param(
    [switch]$bulk,

    # Ziel-Tenant, uebergeben vom aufrufenden Lauf (deployApps / createApps).
    # Ist er gesetzt, erscheint in diesem Skript kein Auswahldialog mehr.
    $Tenant
)
$rootDir = "#ROOT#"

# Funktionen zuerst laden: die Konfiguration kommt ueber Get-ToolConfig.
. $rootDir\functions\functions.ps1

$config = Get-ToolConfig -RootDir $rootDir

$packetRoot = $config.packetRoot
if(-not (Test-Path $packetRoot)){md $packetRoot}

check-prereqs

# Tenant-Auswahl und Anmeldung laufen ueber EINEN Pfad (Initialize-IntuneConnection).
# Wurde der Tenant vom aufrufenden Lauf uebergeben, erscheint hier kein Dialog -
# auch dann nicht, wenn der Token abgelaufen ist und neu geholt werden muss.
# Fehlt er (direkter Aufruf dieses Skripts), wird genau einmal gefragt.
$Tenant = Initialize-IntuneConnection -Tenant $Tenant -Tenants $config.tenants
Write-Host "Tenant: $($Tenant.name)"

# Names Application, description and publisher info
$PackageName = "#PN#"
$Displayname = "#DN#"
$Description = "#DESC#"
$Publisher = "#PUB#"
$AppVersion = "#VER#"

# Create working direcotry for the Application, set download location, and download installer
$appname=$PackageName
$apppath=$PSScriptRoot
$inpath=$apppath + "\in"
$outpath=$apppath + "\out"

# create a temporary short source path to prevent problems with long fullnames
$drive = Get-FirstFreeDriveLetter
subst $drive $inpath
$shortsourcepath = $drive + "\"
# Create the intunewin file from source and destination variables
$installer="Invoke-AppDeployToolkit.ps1"
$SetupFile = $installer
$Destination = $outpath
$CreateAppPackage = New-IntuneWin32AppPackage -SourceFolder $shortsourcepath -SetupFile $SetupFile -OutputFolder $Destination -Force -Verbose
# Get intunewin file Meta data and assign intunewin file location variable
$IntuneWinFile = $CreateAppPackage.Path
$IntuneWinMetaData = Get-IntuneWin32AppMetaData -FilePath $IntuneWinFile
# remove the temporary short source path again
subst $drive /d 

# Create Detection Rule
$DetectionRule = New-IntuneWin32AppDetectionRuleScript -ScriptFile ($apppath + "\detection.ps1")

# Create Requirement Rule
$RequirementRule = New-IntuneWin32AppRequirementRule -Architecture x64 -MinimumSupportedOperatingSystem W10_20H2

# Create a Icon from an image file
if(Test-Path "$apppath\$appname.png"){$ImageFile = "$apppath\$appname.png"}else{$ImageFile = "$apppath\defaultLogo.png"}
$Icon = New-IntuneWin32AppIcon -FilePath $ImageFile

#Install and Uninstall Commands
$InstallCommandLine = "ServiceUi.exe -Process:Explorer.exe Invoke-AppDeployToolkit.exe -DeploymentType Install -DeployMode Silent"
$UninstallCommandLine = "ServiceUi.exe -Process:Explorer.exe Invoke-AppDeployToolkit.exe -DeploymentType Uninstall -DeployMode Silent"

# check if there is an app with the same name already which could be updated
$existingapps = $null
$existingapps = Get-IntuneWin32App -DisplayName $Displayname

if($existingapps){
    Add-Type -AssemblyName Microsoft.VisualBasic
    # if the parameter bulk is not set, ask if a new app should be created
    if($bulk -ne $true){
    
        $result = [Microsoft.VisualBasic.Interaction]::MsgBox('An existing application with the same name has been detected. Create a new application? Select "No" to update an existing one. ','YesNoCancel,SystemModal,Information', 'Create or update an application')
        if($result -eq "Yes"){
            #BULK IS NOT SET
            #ANSWER WAS "YES, CREATE A NEW APP"
            #Builds the App and Uploads to Intune
            Add-IntuneWin32App -FilePath $IntuneWinFile -DisplayName $DisplayName -Description $Description -Publisher $Publisher -AppVersion $AppVersion -InstallExperience "system" -RestartBehavior "suppress" -DetectionRule $DetectionRule -RequirementRule $RequirementRule -InstallCommandLine $InstallCommandLine -UninstallCommandLine $UninstallCommandLine -Icon $Icon -Notes "Created by IntuneWin32Helper #TOOLVER#" -Verbose
        }
        if($result -eq "No"){
            #BULK IS NOT SET
            #ANSWER WAS "NO, UPDATE an existing APP"
            #Builds the App and Uploads to Intune
            #Updates the App 
            # Auswahl ueber den Dialog des Tools statt Out-GridView (ogv):
            # ogv braucht einen STA-Host und fehlt in PowerShell 7 ohne Zusatzmodul.
            $updateCandidates = Get-IntuneWin32App -DisplayName $Displayname |
                Select-Object id, displayName, displayVersion, createdDateTime
            $app = Get-SingleDialogSelection -Value (Open-SelectDialog -data @($updateCandidates) -title "Select the app to update" -size medium)
            if (-not $app) { throw "No application selected for update - aborting." }
            Update-IntuneWin32AppPackageFile -ID $app.id -FilePath $IntuneWinFile
        }
    }
    else{
        #BULK IS SET, always build a NEW App without asking
        Add-IntuneWin32App -FilePath $IntuneWinFile -DisplayName $DisplayName -Description $Description -Publisher $Publisher -AppVersion $AppVersion -InstallExperience "system" -RestartBehavior "suppress" -DetectionRule $DetectionRule -RequirementRule $RequirementRule -InstallCommandLine $InstallCommandLine -UninstallCommandLine $UninstallCommandLine -Icon $Icon -Notes "Created by IntuneWin32Helper #TOOLVER#" -Verbose
    }
}
else{    
    Add-IntuneWin32App -FilePath $IntuneWinFile -DisplayName $DisplayName -Description $Description -Publisher $Publisher -AppVersion $AppVersion -InstallExperience "system" -RestartBehavior "suppress" -DetectionRule $DetectionRule -RequirementRule $RequirementRule -InstallCommandLine $InstallCommandLine -UninstallCommandLine $UninstallCommandLine -Icon $Icon -Verbose
}
Write-Host "Finished."
#if($bulk -ne $true){pause}else{Start-Sleep -Seconds 3}
