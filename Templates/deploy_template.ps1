param(
    [switch]$bulk
)
$rootDir = $PSScriptRoot
if(-not $scriptpath){$rootDir = "C:\Users\alex\OneDrive - AZITC\Tools\Administration\IntuneWin32Helper"}
$config = Get-Content -Raw -Path "$rootDir\config\config.json" -Encoding UTF8 | ConvertFrom-Json

$packetRoot = $config.packetRoot
if(-not (Test-Path $packetRoot)){$packetRoot = "C:\Users\alex.HOME\AZITC\work - Pakete"}
. $rootDir\functions\functions.ps1

check-prereqs
if($bulk -ne $true){
    if(-not $Tenant){$Tenant = Open-SelectDialogWithEdit -data $config.tenants -title "Select Tenant" -size small}
}
$Tenant = $Tenant | Where-Object { $_ -is [System.Management.Automation.PSCustomObject] } # bug mit Dialog und Rückgabe Collections, sonst auch int werte enthalten

if(Test-Path $rootDir\tokeninfo.json){
    Write-Host "Token detected, checking validity."
    $lastinfo = Get-Content -Raw -Path "$rootDir\tokeninfo.json" -Encoding UTF8 | ConvertFrom-Json
    if($lastinfo.ExpiresOn.ToUniversalTime() -lt (get-date).ToUniversalTime()){
        # token ist expired, neu holen
        if(-not (Test-AccessToken)){
            Write-Host "Token expired, re-auth required."
            if(-not $Tenant){$Tenant = Open-SelectDialog -data $config.tenants -title "Select Tenant" -size small}
            $tokeninfo = Connect-MSIntuneGraph -TenantID $Tenant.name -Interactive -ClientId $Tenant.AppId
        }
    }
    else{        
        # token ist ggf. noch gut
        if(-not (Test-AccessToken)){
            if(-not $Tenant){$Tenant = Open-SelectDialog -data $config.tenants -title "Select Tenant" -size small}
            $tokeninfo = Connect-MSIntuneGraph -TenantID $Tenant.name -Interactive -ClientId $Tenant.AppId
        }
        else{
            Write-Host "Token still valid."
        }
    }
}
else{
    Write-Host "No Token detected, authentication required."
    # no token info, neues Token erforderlich
    if(-not $Tenant){$Tenant = Open-SelectDialog -data $config.tenants -title "Select Tenant" -size small}
    $tokeninfo = Connect-MSIntuneGraph -TenantID $Tenant.name -Interactive -ClientId $Tenant.AppId
}

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
$CreateAppPackage = New-IntuneWin32AppPackage -SourceFolder $shortsourcepath -SetupFile $SetupFile -OutputFolder $Destination -Force #-Verbose
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
if(Test-Path "$apppath\$appname.png"){$ImageFile = "$apppath\$appname.png"}else{"$apppath\defaultLogo.png"}
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
            Add-IntuneWin32App -FilePath $IntuneWinFile -DisplayName $DisplayName -Description $Description -Publisher $Publisher -AppVersion $AppVersion -InstallExperience "system" -RestartBehavior "suppress" -DetectionRule $DetectionRule -RequirementRule $RequirementRule -InstallCommandLine $InstallCommandLine -UninstallCommandLine $UninstallCommandLine -Icon $Icon -Notes "Created by IntuneWin32Helper #TOOLVER#" #-Verbose
        }
        if($result -eq "No"){
            #BULK IS NOT SET
            #ANSWER WAS "NO, UPDATE an existing APP"
            #Builds the App and Uploads to Intune
            #Updates the App 
            $app = Get-IntuneWin32App -DisplayName $Displayname | ogv -PassThru -Title "Select the app to update"
            Update-IntuneWin32AppPackageFile -ID $app.id -FilePath $IntuneWinFile
        }
    }
    else{
        #BULK IS SET, always build a NEW App without asking
        Add-IntuneWin32App -FilePath $IntuneWinFile -DisplayName $DisplayName -Description $Description -Publisher $Publisher -AppVersion $AppVersion -InstallExperience "system" -RestartBehavior "suppress" -DetectionRule $DetectionRule -RequirementRule $RequirementRule -InstallCommandLine $InstallCommandLine -UninstallCommandLine $UninstallCommandLine -Icon $Icon -Notes "Created by IntuneWin32Helper #TOOLVER#" #-Verbose
    }
}
else{    
    Add-IntuneWin32App -FilePath $IntuneWinFile -DisplayName $DisplayName -Description $Description -Publisher $Publisher -AppVersion $AppVersion -InstallExperience "system" -RestartBehavior "suppress" -DetectionRule $DetectionRule -RequirementRule $RequirementRule -InstallCommandLine $InstallCommandLine -UninstallCommandLine $UninstallCommandLine -Icon $Icon #-Verbose
}