param(
    [switch]$bulk
)
$rootDir = "#ROOT#"
$config = Get-Content -Raw -Path "$rootDir\config\config.json" -Encoding UTF8 | ConvertFrom-Json

$packetRoot = $config.packetRoot
if(-not (Test-Path $packetRoot)){md $packetRoot}
. $rootDir\functions\functions.ps1

check-prereqs

if($bulk -ne $true){
	# always ask for tenant in single app deployments
    Write-Host "Show tenant selection dialog"
    if(-not $Tenant){$Tenant = Open-SelectDialog -data $config.tenants -title "Select Tenant" -size small}
	$Tenant = $Tenant | Where-Object { $_ -isnot [int]] } # bug mit Dialog und Rückgabe Collections, sonst auch int werte enthalten
	$tokeninfo = Connect-MSIntuneGraph -TenantID $Tenant.name -ClientId $Tenant.AppId -ClientSecret $Tenant.clientSecret -Verbose
	Write-Host "Tenant: $($Tenant.name)"
}

if(-not (Test-AccessToken)){
    Write-Host "No access token detected, authentication required."
    Write-Host "Show tenant selection dialog"
	if(-not $Tenant){$Tenant = Open-SelectDialog -data $config.tenants -title "Select Tenant" -size small}
	$Tenant = $Tenant | Where-Object { $_ -isnot [int]] } # bug mit Dialog und Rückgabe Collections, sonst auch int werte enthalten
	$tokeninfo = Connect-MSIntuneGraph -TenantID $Tenant.name -ClientId $Tenant.AppId -ClientSecret $Tenant.clientSecret -Verbose
	Write-Host "Tenant: $($Tenant.name)"
}
else{
	Write-Host "Access Token still valid."
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
            $app = Get-IntuneWin32App -DisplayName $Displayname | ogv -PassThru -Title "Select the app to update"
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
