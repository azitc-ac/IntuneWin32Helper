$toolVersion = "1.1"
$rootDir = $PSScriptRoot
if(-not $scriptpath){$rootDir = "C:\Users\alex\OneDrive - AZITC\Tools\Administration\IntuneWin32Helper"}
$config = Get-Content -Raw -Path "$rootDir\config\config.json" -Encoding UTF8 | ConvertFrom-Json

$cloudName = $config.cloudName
$ApiKey = $config.apiKey
$ApiSecret = $config.apiSecret
$packetRoot = $config.packetRoot
if(-not (Test-Path $packetRoot)){$packetRoot = "C:\Users\alex.HOME\AZITC\work - Pakete"}
. $rootDir\functions\functions.ps1

check-prereqs
$csv = "$rootDir\apps.csv"
<#$csv = Import-Csv "$rootDir\apps.csv" -Delimiter ";"
# neusortieren, falls etwas hinzugekommen ist
$csv | sort DisplayName | export-csv -Delimiter ";" "$rootDir\apps.csv" -NoTypeInformation
#>
$apps = Open-SelectDialogWithEdit $csv -title "Welche Applications sollen erstellt werden?" -size large
$apps = $apps | Where-Object { $_ -is [System.Management.Automation.PSCustomObject] } # bug mit Dialog und Rückgabe Collections, sonst auch int werte enthalten

foreach($app in $apps){    
    # Erstellen der Anwendung
    $AppName = $app.DisplayName
    $AppVersion = $app.Version
    $AppPublisher = $app.Publisher
    $AppDescription = $AppName
    $AppNameCombined = $AppName + " - " + $AppVersion
    $SourcePath = "$packetRoot\$AppNameCombined"
    $ProgramId = $app.ProgramID
    $InstallCmd = "ServiceUi.exe -Process:Explorer.exe Invoke-AppDeployToolkit.exe -DeploymentType Install -DeployMode Silent"
    $UninstallCmd = "ServiceUi.exe -Process:Explorer.exe Invoke-AppDeployToolkit.exe -DeploymentType Uninstall -DeployMode Silent"
    $InstallCmdInternal = $app.InstallCmd
    $UninstallCmdInternal = $app.UninstallCmd
    $winGetParams = $app.WinGetParams -replace "`"",""

    Write-Host "Erstelle Application: $AppNameCombined" -ForegroundColor Cyan
    if((Test-Path $SourcePath) -and ($config.removeExistingPacketDirOnEachRun -eq $true)){del $SourcePath -Recurse -Force}
    New-ADTTemplate -Destination $packetRoot -Name $AppNameCombined

    Write-Host "Warte 5 Sekunden..."
    Start-Sleep -Seconds 5
    
    # erstellen von \in und \out, move eine ebene tiefer nach \in, 
    Write-Host "Verschiebe Daten nach .\in"
    $psadtdirs = Get-ChildItem $SourcePath
    md $SourcePath\in
    md $SourcePath\out
    $psadtdirs | %{mv $_.fullname $SourcePath\in}
    
    # log path ändern
    $psadtconfigfilepath = "$SourcePath\in\Config\config.psd1" 
    $psadtconfig= get-content $psadtconfigfilepath
    $psadtconfig= $psadtconfig.Replace("envWinDir\Logs\Software", "envProgramData\Microsoft\IntuneManagementExtension\Logs")
    $psadtconfigfolderPath = Get-Item "$SourcePath\in\Config"
    (Get-Item $psadtconfigfolderPath).Attributes = ((Get-Item $psadtconfigfolderPath).Attributes -band -bnot [System.IO.FileAttributes]::ReadOnly)
    $psadtconfig | Out-File $psadtconfigfilepath -Encoding utf8 -Force

    # für normale Pakete
    # kopieren von detect.ps1 und anpassen
    Write-Host "Kopieren und anpassen: detection_template.ps1"
    if($AppVersion -ne "LatestAvailable"){
        $DetectionScript = get-content "$rootDir\Templates\detection_template.ps1" 
        $DetectionScript -replace "#DN#", $AppName -replace "#VER#",$AppVersion | Out-File "$SourcePath\detection.ps1" -Encoding utf8 -Force
    }
    else{
        # für WinGet Pakete
        # kopieren von detect.ps1 und anpassen
        Write-Host "Kopieren und anpassen: detection_template-WinGetApp.ps1"
        $DetectionScript = get-content "$rootDir\Templates\detection_template-WinGetApp.ps1"
        $DetectionScript -replace "WINGETPROGRAMID", $ProgramId | Out-File "$SourcePath\detection.ps1" -Encoding utf8 -Force
    }
    # kopieren von serviceui.exe ins neu erstellte dir
    Write-Host "Kopieren: ServiceUI.exe"
    cp $rootDir\ServiceUI.exe $SourcePath\in

    # einpflegen von publisher, appname, version ins invoke-AppDeployToolkit.ps1
    # Update der Invoke-AppDeployToolkit.ps1, außer im Fall des Zero-Config Deployment mit 1 single MSI, dann darf hier nichts angepasst werden
    if(-not $app.SingleMSI){
        Write-Host "Kopieren und anpassen: Invoke-AppDeployToolkit.ps1"    
        $creationdate = get-date -Format "yyyy-MM-dd"
        $psadtscript = get-content "$SourcePath\in\Invoke-AppDeployToolkit.ps1"
        $psadtscript -replace "AppVendor = ''","AppVendor = '$AppPublisher'" -replace "AppName = ''", "AppName = '$AppName'" -replace "AppVersion = ''", "AppVersion = '$AppVersion'" `
            -replace "AppScriptDate = '2000-12-31'", "AppScriptDate = '$creationdate'" -replace "AppScriptAuthor = '<author name>'", "AppScriptAuthor = 'alexander@zarenko.net'" `
            | Out-File "$SourcePath\in\Invoke-AppDeployToolkit.ps1" -Encoding utf8 -Force
    }
    if($InstallCmdInternal){
        Insert-Commands -Install $InstallCmdInternal -FilePath "$SourcePath\in\Invoke-AppDeployToolkit.ps1"
    }
    if($UninstallCmdInternal){
        Insert-Commands -Uninstall $UninstallCmdInternal -FilePath "$SourcePath\in\Invoke-AppDeployToolkit.ps1"
    }
    if($AppVersion -eq "LatestAvailable"){        
        $InstallCmdInternal = get-WinGetCommands -type Install -id $ProgramId -wgparams $winGetParams
        $UninstallCmdInternal= get-WinGetCommands -type Uninstall -id $ProgramId -wgparams $winGetParams
        Insert-Commands -Install $InstallCmdInternal -FilePath "$SourcePath\in\Invoke-AppDeployToolkit.ps1"
        Insert-Commands -Uninstall $UninstallCmdInternal -FilePath "$SourcePath\in\Invoke-AppDeployToolkit.ps1"        
    }

    #App version setzen
    if($AppVersion -eq "LatestAvailable"){
        $desc = "Installed using PSADT and WinGet"
    }
    else{
        $desc = "Installed using PSADT"
    }

    #Logo download from Appstore
    $logoURL = $app.logoURL
    $appfolder = $SourcePath
    #schaue nach, ob es schon ein logo gibt und falls ja, nimm es, kein DL
    $existingLogo = "$rootDir\Logos\$AppName.png"
    if(test-path $existingLogo){
        write-host "Verwende existierendes Logo [$existingLogo]"
        cp $existingLogo "$appfolder\"
    }
    else{
        if(-not $logoURL){
            #copy default logo
            Write-Host "Keine Logo URL angegeben. Nehme Default logo."
            cp "$rootDir\Logos\defaultlogo.png" $appfolder
        }
        else{
            #try DL
            if($logoURL -like "*.png"){$LogoFileName = "$AppName.png"}else{$LogoFileName = "$AppName.webp"}
            Write-Host "Versuche Logo Download..."
            Invoke-WebRequest -Uri $logoURL -D -OutFile $appfolder\$LogoFileName
            if(Test-Path $appfolder\$LogoFileName){
                Write-Host "Logo Download erfolgreich."
                if($LogoFileName -eq "$AppName.webp"){
                    # convert webp to png
                    Write-Host "Konvertiere Logo von .webp zu .png"
                    Convert-WebPToPngCloudinary "$appfolder\$LogoFileName" -CloudName $cloudName -ApiKey $ApiKey -ApiSecret $ApiSecret
                    $LogoFileName = "$AppName.png"
                }
                cp "$appfolder\$LogoFileName" "$rootDir\Logos\" -Force
            }
            else{
                #copy default logo
                Write-Host "Fehler beim Logo Download. Nehme Default logo."
                cp "$rootDir\Logos\defaultlogo.png" $appfolder
            }
        }
    }

    #deploy template an App anpassen und kopieren
    $go = Get-Content "$rootDir\Templates\deploy_template.ps1"
    $go -replace "#DN#",$AppName -replace "#PN#",$AppName -replace "#PUB#",$AppPublisher `
         -replace "#DM#", "DetectionScript" -replace "#VER#", $AppVersion -replace "#DESC#", $desc`
         -replace "#TOOLVER#", $toolVersion | Out-File $SourcePath\deploy.ps1 -Encoding utf8 -Force
  
    # manuell: files reinpacken, install u uninstall routine einpflegen
    if($AppVersion -ne "LatestAvailable"){
        Write-Host "ToDo: Files reinkopieren, dann ENTER" -ForegroundColor Cyan
        explorer "$SourcePath\in\Files"
        pause
        Write-Host "ToDo: Install & Uninstall Section mit Leben füllen, dann ENTER, dann wird nach Intune deployed" -ForegroundColor Cyan
        powershell_ise $SourcePath\in\Invoke-AppDeployToolkit.ps1
        pause
    }

    #deploy.ps1 aufrufen
    if($apps.count -gt 1){
        write-host "Parameter -bulk is set."
        & $SourcePath\deploy.ps1 -bulk
    }
    else{
        write-host "Parameter -bulk is NOT set."
        & $SourcePath\deploy.ps1    
    }
    # und weiter gehts mit der nächsten App
}