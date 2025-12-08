function deployApps{    
    
    function Parse-AppFolderName {
        param(
            [Parameter(Mandatory=$true)]
            [string]$FolderName,
            [Parameter(Mandatory=$false)]
            [string]$Delimiter = ' - '
        )

        # Standardwerte
        $appName = $FolderName
        $appVersion = ''

        if ($FolderName -and $Delimiter -and $FolderName.Contains($Delimiter)) {
            $lastIndex = $FolderName.LastIndexOf($Delimiter)
            if ($lastIndex -ge 0) {
                $left  = $FolderName.Substring(0, $lastIndex)
                $right = $FolderName.Substring($lastIndex + $Delimiter.Length)
                if ($left)  { $appName    = $left.Trim() }
                if ($right) { $appVersion = $right.Trim() }
            }
        }

        # Rückgabe als Objekt
        return [pscustomobject]@{
            AppName    = $appName
            AppVersion = $appVersion
        }
    }
        
    function Get-DeployScripts {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$PacketRoot,

            [Parameter(Mandatory = $false)]
            [string]$Delimiter = ' - ',

            [Parameter(Mandatory = $false)]
            [switch]$Recurse,

            [Parameter(Mandatory = $false)]
            [string]$ExportCsvPath
        )

        if (-not (Test-Path -LiteralPath $PacketRoot)) {
            throw "PacketRoot does not exist: $PacketRoot"
        }

        # Unterverzeichnisse holen (optional rekursiv)
        if ($Recurse) {
            $dirs = Get-ChildItem -LiteralPath $PacketRoot -Directory -Recurse
        } else {
            $dirs = Get-ChildItem -LiteralPath $PacketRoot -Directory
        }

        $results = @()

        foreach ($dir in $dirs) {
            $deployPath = Join-Path -Path $dir.FullName -ChildPath 'deploy.ps1'

            if (Test-Path -LiteralPath $deployPath) {
                $parsed = Parse-AppFolderName -FolderName $dir.Name -Delimiter $Delimiter

                $fi = Get-Item -LiteralPath $deployPath
                $obj = [pscustomobject]@{
                    AppName      = $parsed.AppName
                    AppVersion   = $parsed.AppVersion
                    LastModified = $fi.LastWriteTime
                    FullPath     = $fi.FullName
                }

                $results += $obj
            }
        }

        # Optional: CSV exportieren
        if ($ExportCsvPath -and $ExportCsvPath.Trim().Length -gt 0) {
            try {
                $results | Export-Csv -Path $ExportCsvPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
            } catch {
                Write-Warning ("Konnte CSV nicht schreiben: {0}" -f $_.Exception.Message)
            }
        }

        return $results
    }

    $deployableApps = Get-DeployScripts -PacketRoot $packetRoot #-Recurse
    $appsToDeploy = Open-SelectDialog -data $deployableApps -title "Select Apps to deploy" -large
    
    # Rückgabe bereinigen (bekannter Workaround gegen int-Werte in Collections)
    if ($appsToDeploy -ne $null) {
        $appsToDeploy = $appsToDeploy | Where-Object { $_ -is [System.Management.Automation.PSCustomObject] }
    }

    # Abbruchbedingung: wenn Nutzer "Cancel" klickt oder Fenster schließt → keine gültigen Items
    if ($appsToDeploy -eq $null -or ($appsToDeploy | Measure-Object).Count -eq 0) {
        break
    }

    # Verarbeitung der ausgewählten Apps
    foreach($app in $appsToDeploy){
        Write-Host "Deploy Application: $($app.AppName) - $($app.Version)" -ForegroundColor Cyan
        #deploy.ps1 aufrufen
        if($appsToDeploy.count -gt 1){
            write-host "Parameter -bulk is set."
            & $app.FullPath -bulk
        }
        else{
            write-host "Parameter -bulk is NOT set."
            & $app.FullPath    
        }
    }
}

function createApps{
    param(
        [switch]$createAndDeploy,
        [string]$csvPath
    )
    $csvPath = "$rootDir\apps.csv"

    # --- Wiederholte Auswahl + Verarbeitung, bis Nutzer abbricht ---
    while ($true) {
        # Dialog öffnen (gibt nur bei OK ein Ergebnis zurück)
        if($createAndDeploy){$title = "Select Applications to create and deploy"}else{$title = "Select Applications to create"}
        $apps = Open-SelectDialogWithEdit -CsvPath $csvPath -title $title -size large

        # Rückgabe bereinigen (bekannter Workaround gegen int-Werte in Collections)
        if ($apps -ne $null) {
            $apps = $apps | Where-Object { $_ -is [System.Management.Automation.PSCustomObject] }
        }

        # Abbruchbedingung: wenn Nutzer "Cancel" klickt oder Fenster schließt → keine gültigen Items
        if ($apps -eq $null -or ($apps | Measure-Object).Count -eq 0) {
            break
        }

        # Verarbeitung der ausgewählten Apps

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

        Write-Host "Creating PSADT application: $AppNameCombined" -ForegroundColor Cyan
        if((Test-Path $SourcePath) -and ($config.removeExistingPacketDirOnEachRun -eq $true)){del $SourcePath -Recurse -Force}
        New-ADTTemplate -Destination $packetRoot -Name $AppNameCombined

        Write-Host "Wait 5 seconds..."
        Start-Sleep -Seconds 5
    
        # erstellen von \in und \out, move eine ebene tiefer nach \in, 
        Write-Host "Moving data to .\in"
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
        Write-Host "Copying and customizing: detection_template.ps1"
        if($AppVersion -ne "LatestAvailable"){
            $DetectionScript = get-content "$rootDir\Templates\detection_template.ps1" 
            $DetectionScript -replace "#DN#", $AppName -replace "#VER#",$AppVersion | Out-File "$SourcePath\detection.ps1" -Encoding utf8 -Force
        }
        else{
            # für WinGet Pakete
            # kopieren von detect.ps1 und anpassen
            Write-Host "Copying and customizing: detection_template-WinGetApp.ps1"
            $DetectionScript = get-content "$rootDir\Templates\detection_template-WinGetApp.ps1"
            $DetectionScript -replace "WINGETPROGRAMID", $ProgramId | Out-File "$SourcePath\detection.ps1" -Encoding utf8 -Force
        }
        # kopieren von serviceui.exe ins neu erstellte dir
        Write-Host "Copying: ServiceUI.exe"
        cp $rootDir\ServiceUI.exe $SourcePath\in

        # einpflegen von publisher, appname, version ins invoke-AppDeployToolkit.ps1
        # Update der Invoke-AppDeployToolkit.ps1, außer im Fall des Zero-Config Deployment mit 1 single MSI, dann darf hier nichts angepasst werden
        if(-not $app.SingleMSI){
            Write-Host "Copying and customizing: Invoke-AppDeployToolkit.ps1"    
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
            write-host "Using existing logo [$existingLogo]"
            cp $existingLogo "$appfolder\"
        }
        else{
            if(-not $logoURL){
                #copy default logo
                Write-Host "No logo URL specified. Taking default logo."
                cp "$rootDir\Logos\defaultlogo.png" $appfolder
            }
            else{
                #try DL
                if($logoURL -like "*.png"){$LogoFileName = "$AppName.png"}else{$LogoFileName = "$AppName.webp"}
                Write-Host "Trying logo download..."
                Invoke-WebRequest -Uri $logoURL -D -OutFile $appfolder\$LogoFileName
                if(Test-Path $appfolder\$LogoFileName){
                    Write-Host "Logo download successful."
                    if($LogoFileName -eq "$AppName.webp"){
                        # convert webp to png
                        Write-Host "Converting logo from .webp to .png"
                        Convert-WebPToPngCloudinary "$appfolder\$LogoFileName" -CloudName $cloudName -ApiKey $ApiKey -ApiSecret $ApiSecret
                        $LogoFileName = "$AppName.png"
                    }
                    cp "$appfolder\$LogoFileName" "$rootDir\Logos\" -Force
                }
                else{
                    #copy default logo
                    Write-Host "Error during logo download. Taking default logo."
                    cp "$rootDir\Logos\defaultlogo.png" $appfolder
                }
            }
        }

        #deploy template an App anpassen und kopieren
        $go = Get-Content "$rootDir\Templates\deploy_template.ps1"
        $go -replace "#ROOT#",$rootDir -replace "#DN#",$AppName -replace "#PN#",$AppName -replace "#PUB#",$AppPublisher `
             -replace "#DM#", "DetectionScript" -replace "#VER#", $AppVersion -replace "#DESC#", $desc`
             -replace "#TOOLVER#", $toolVersion | Out-File $SourcePath\deploy.ps1 -Encoding utf8 -Force
  
        # manuell: files reinpacken, install u uninstall routine einpflegen
        if($AppVersion -ne "LatestAvailable"){
            Write-Host "ToDo: now add/copy all required files for setup, then press ENTER" -ForegroundColor Cyan
            explorer "$SourcePath\in\Files"
            pause
            Write-Host "ToDo: fill Install & Uninstall sections with life, then press ENTER. Deployment to Intune will begin if previously selected." -ForegroundColor Cyan
            powershell_ise $SourcePath\in\Invoke-AppDeployToolkit.ps1
            pause
        }
        if($createAndDeploy){
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
    }

        # Nach der Verarbeitung geht die Schleife automatisch weiter
        # → Der Dialog wird erneut geöffnet, bis der Nutzer abbricht.
    }
}

function Show-StartDialog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [System.Windows.Window]$Owner,
        [string]$Title = "IntuneWin32Helper – https://blog.zarenko.net/",
        [int]$TileWidth = 220 # Breite je Kachel zur sauberen Textumbruch-Steuerung
    )

    Add-Type -AssemblyName PresentationCore           | Out-Null
    Add-Type -AssemblyName PresentationFramework      | Out-Null
    Add-Type -AssemblyName System.Windows.Forms       | Out-Null

    # Fenster
    $dlg = New-Object Windows.Window
    $dlg.Title = $Title
    $dlg.Width = 740
    $dlg.Height = 450
    if ($Owner -ne $null) {
        $dlg.Owner = $Owner
        $dlg.WindowStartupLocation = "CenterOwner"
    } else {
        $dlg.WindowStartupLocation = "CenterScreen"
    }
    $dlg.ResizeMode = 'NoResize'

    # Root-Grid
    $root = New-Object Windows.Controls.Grid
    $root.Margin = "16"
    $rowTitle   = New-Object Windows.Controls.RowDefinition; $rowTitle.Height   = [Windows.GridLength]::Auto
    $rowContent = New-Object Windows.Controls.RowDefinition; $rowContent.Height = New-Object Windows.GridLength -ArgumentList 1, ([Windows.GridUnitType]::Star)
    $rowFooter  = New-Object Windows.Controls.RowDefinition; $rowFooter.Height  = [Windows.GridLength]::Auto
    $null = $root.RowDefinitions.Add($rowTitle)
    $null = $root.RowDefinitions.Add($rowContent)
    $null = $root.RowDefinitions.Add($rowFooter)

    # Titel
    $txtTitle = New-Object Windows.Controls.TextBlock
    $txtTitle.Text = "What would you like to do?"
    $txtTitle.FontSize = 18
    $txtTitle.FontWeight = 'Bold'
    $txtTitle.Margin = "0,0,0,12"
    [Windows.Controls.Grid]::SetRow($txtTitle, 0)
    $null = $root.Children.Add($txtTitle)

    # Inhalte: 3 Kacheln in UniformGrid
    $uniform = New-Object Windows.Controls.Primitives.UniformGrid
    $uniform.Rows = 1
    $uniform.Columns = 3
    $uniform.Margin = "0,8,0,0"
    [Windows.Controls.Grid]::SetRow($uniform, 1)
    $null = $root.Children.Add($uniform)

    # Hilfsfunktion: Kachel erstellen (Border + StackPanel + Icon + Texte)
    function New-OptionTile {
        param(
            [string]$Caption,
            [string]$Description,
            [System.Windows.UIElement]$IconElement,
            [string]$ReturnValue,
            [int]$Width = 220
        )

        # Umrandete, hoverbare Kachel
        $border = New-Object Windows.Controls.Border
        $border.BorderBrush = [System.Windows.Media.Brushes]::LightGray
        $border.BorderThickness = '1'
        $border.CornerRadius = '6'
        $border.Margin = '6'
        $border.Padding = '12'
        $border.Background = [System.Windows.Media.Brushes]::White
        $border.SnapsToDevicePixels = $true
        $border.Width = $Width
        $border.Cursor = 'Hand'

        # >>> stabiler Rückgabewert direkt an der Kachel speichern
        $border.Tag = $ReturnValue

        # Hover-Effekt
        $border.Add_MouseEnter({ param($s,$e) $s.BorderBrush = [System.Windows.Media.Brushes]::DodgerBlue })
        $border.Add_MouseLeave({ param($s,$e) $s.BorderBrush = [System.Windows.Media.Brushes]::LightGray })

        # Inhalt
        $stack = New-Object Windows.Controls.StackPanel
        $stack.Orientation = 'Vertical'
        $stack.VerticalAlignment = 'Center'
        $stack.HorizontalAlignment = 'Center'
        $stack.Width = $Width - 24

        if ($IconElement -ne $null) {
            $iconHost = New-Object Windows.Controls.ContentControl
            $iconHost.Content = $IconElement
            $iconHost.HorizontalAlignment = 'Center'
            $iconHost.Margin = '0,6,0,8'
            $null = $stack.Children.Add($iconHost)
        }

        $lbl = New-Object Windows.Controls.TextBlock
        $lbl.Text = $Caption
        $lbl.FontWeight = 'Bold'
        $lbl.FontSize = 14
        $lbl.HorizontalAlignment = 'Center'
        $lbl.TextAlignment = 'Center'
        $lbl.TextWrapping = 'Wrap'
        $lbl.Margin = '0,0,0,4'
        $lbl.MaxWidth = $Width - 24

        $desc = New-Object Windows.Controls.TextBlock
        $desc.Text = $Description
        $desc.TextAlignment = 'Center'
        $desc.Foreground = [System.Windows.Media.Brushes]::DimGray
        $desc.Margin = '0,0,0,6'
        $desc.TextWrapping = 'Wrap'
        $desc.MaxWidth = $Width - 24

        $null = $stack.Children.Add($lbl)
        $null = $stack.Children.Add($desc)

        $border.Child = $stack

        return $border
    }

    # ---------- ICONS ----------
    $glyphAdd         = [char]0xE710  # Add
    $glyphCloudUpload = [char]0xE898  # CloudUpload
    $glyphSettings    = [char]0xE713  # Settings

    # Links: Add-Icon (einzeln)
    $iconCreate = New-Object Windows.Controls.TextBlock
    $iconCreate.Text = $glyphAdd
    $iconCreate.FontFamily = New-Object System.Windows.Media.FontFamily 'Segoe MDL2 Assets'
    $iconCreate.FontSize = 42
    $iconCreate.Foreground = [System.Windows.Media.Brushes]::DodgerBlue
    $iconCreate.HorizontalAlignment = 'Center'
    $iconCreate.Margin = '0,6,0,8'

    # Mitte: Add + CloudUpload nebeneinander
    $iconCreateDeployWrap = New-Object Windows.Controls.StackPanel
    $iconCreateDeployWrap.Orientation = 'Horizontal'
    $iconCreateDeployWrap.HorizontalAlignment = 'Center'
    $iconCreateDeployWrap.Margin = '0,6,0,8'

    $iconCreateDeployAdd = New-Object Windows.Controls.TextBlock
    $iconCreateDeployAdd.Text = $glyphAdd
    $iconCreateDeployAdd.FontFamily = New-Object System.Windows.Media.FontFamily 'Segoe MDL2 Assets'
    $iconCreateDeployAdd.FontSize = 36
    $iconCreateDeployAdd.Foreground = [System.Windows.Media.Brushes]::DodgerBlue
    $iconCreateDeployAdd.HorizontalAlignment = 'Center'
    $iconCreateDeployAdd.Margin = '0,0,8,0'

    $iconCreateDeployUpload = New-Object Windows.Controls.TextBlock
    $iconCreateDeployUpload.Text = $glyphCloudUpload
    $iconCreateDeployUpload.FontFamily = New-Object System.Windows.Media.FontFamily 'Segoe MDL2 Assets'
    $iconCreateDeployUpload.FontSize = 36
    $iconCreateDeployUpload.Foreground = [System.Windows.Media.Brushes]::DodgerBlue
    $iconCreateDeployUpload.HorizontalAlignment = 'Center'

    $null = $iconCreateDeployWrap.Children.Add($iconCreateDeployAdd)
    $null = $iconCreateDeployWrap.Children.Add($iconCreateDeployUpload)

    # Rechts: CloudUpload-Icon (einzeln)
    $iconDeploy = New-Object Windows.Controls.TextBlock
    $iconDeploy.Text = $glyphCloudUpload
    $iconDeploy.FontFamily = New-Object System.Windows.Media.FontFamily 'Segoe MDL2 Assets'
    $iconDeploy.FontSize = 42
    $iconDeploy.Foreground = [System.Windows.Media.Brushes]::DodgerBlue
    $iconDeploy.HorizontalAlignment = 'Center'
    $iconDeploy.Margin = '0,6,0,8'
    # ---------- ENDE ICONS ----------

    # Kacheln erstellen
    $tileCreate = New-OptionTile -Caption 'Create apps' `
        -Description 'Create app packages in the file system.' `
        -IconElement $iconCreate `
        -ReturnValue 'CreateNew' `
        -Width $TileWidth

    $tileCreateDeploy = New-OptionTile -Caption 'Create and deploy apps' `
        -Description 'Create app packages in the file system and deploy them to Intune.' `
        -IconElement $iconCreateDeployWrap `
        -ReturnValue 'CreateNewAndDeploy' `
        -Width $TileWidth

    $tileDeploy = New-OptionTile -Caption 'Deploy existing apps' `
        -Description 'Deploy existing app packages to Intune.' `
        -IconElement $iconDeploy `
        -ReturnValue 'DeployExisting' `
        -Width $TileWidth

    # Rückgabewert direkt an den Kacheln speichern
    $tileCreate.Tag        = 'CreateNew'
    $tileCreateDeploy.Tag  = 'CreateNewAndDeploy'
    $tileDeploy.Tag        = 'DeployExisting'

    $null = $uniform.Children.Add($tileCreate)
    $null = $uniform.Children.Add($tileCreateDeploy)
    $null = $uniform.Children.Add($tileDeploy)

    # --- ZENTRALER Klick-Handler: PreviewMouseLeftButtonUp am UniformGrid ---
    # Greift, egal ob auf Icon, Text, leere Fläche oder Stack geklickt wird.
    $uniform.Add_PreviewMouseLeftButtonUp({
        param($s,$e)

        # Ursprüngliches Ziel der Maus
        $src = $e.OriginalSource

        # Zum nächsten Border (Tile) nach oben laufen
        $elem = $src
        $borderFound = $null

        while ($elem -ne $null -and $borderFound -eq $null) {
            if ($elem -is [Windows.Controls.Border]) {
                $borderFound = $elem
            } else {
                # Parent über FrameworkElement/VisualTree ermitteln
                if ($elem -is [System.Windows.FrameworkElement] -and $elem.Parent -ne $null) {
                    $elem = $elem.Parent
                } else {
                    $elem = [System.Windows.Media.VisualTreeHelper]::GetParent($elem)
                }
            }
        }

        if ($borderFound -ne $null -and $borderFound.Tag -ne $null -and [string]::IsNullOrWhiteSpace([string]$borderFound.Tag) -eq $false) {
            $dlg.Tag = [string]$borderFound.Tag
            $dlg.Close()
        }
    })

    # --- Footer mit Zahnrad links & Cancel rechts ---
    $footer = New-Object Windows.Controls.Grid
    $footer.Margin = "0,14,0,0"
    $colLeft = New-Object Windows.Controls.ColumnDefinition; $colLeft.Width = "Auto"
    $colFill = New-Object Windows.Controls.ColumnDefinition; $colFill.Width = "*"
    $colRight = New-Object Windows.Controls.ColumnDefinition; $colRight.Width = "Auto"
    $null = $footer.ColumnDefinitions.Add($colLeft)
    $null = $footer.ColumnDefinitions.Add($colFill)
    $null = $footer.ColumnDefinitions.Add($colRight)

    # Einstellungen links
    $btnSettings = New-Object Windows.Controls.Button
    $btnSettings.ToolTip = "Settings"
    $btnSettings.Padding = "10,6"
    $btnSettings.MinWidth = 40
    $btnSettings.HorizontalAlignment = "Left"
    $btnSettings.VerticalAlignment = "Center"
    $settingsIcon = New-Object Windows.Controls.TextBlock
    $settingsIcon.Text = $glyphSettings
    $settingsIcon.FontFamily = New-Object System.Windows.Media.FontFamily 'Segoe MDL2 Assets'
    $settingsIcon.FontSize = 18
    $settingsIcon.Foreground = [System.Windows.Media.Brushes]::Gray
    $btnSettings.Content = $settingsIcon
    [Windows.Controls.Grid]::SetColumn($btnSettings, 0)
    $null = $footer.Children.Add($btnSettings)

    # Cancel rechts
    $spClose = New-Object Windows.Controls.StackPanel
    $spClose.Orientation = 'Horizontal'
    $spClose.HorizontalAlignment = 'Right'
    $btnClose = New-Object Windows.Controls.Button
    $btnClose.Content = 'Cancel'
    $btnClose.Padding = '14,6'
    $btnClose.Margin = '0,0,0,0'
    $btnClose.Add_Click({
        $dlg.Tag = 'Cancel'
        $dlg.Close()
    })
    $null = $spClose.Children.Add($btnClose)
    [Windows.Controls.Grid]::SetColumn($spClose, 2)
    $null = $footer.Children.Add($spClose)
    [Windows.Controls.Grid]::SetRow($footer, 2)
    $null = $root.Children.Add($footer)

    # Click: Einstellungen öffnen (modal, zentriert)
    $btnSettings.Add_Click({
        try {
            $null = Edit-SettingsDialog -Owner $dlg -PreferredPaths @(                
                (Join-Path $rootDir "Config\config.json")
            )
        } catch {
            [System.Windows.MessageBox]::Show(("Error while opening the settings: {0}" -f $_.Exception.Message), "Settings", "OK", "Error") | Out-Null
        }
    })

    # Fensterinhalt setzen
    $dlg.Content = $root

    # Initiales Tag
    $dlg.Tag = $null

    # [X]-Schließen → 'Closed' setzen, falls noch kein Wert
    $dlg.Add_Closing({
        if ($dlg.Tag -eq $null -or [string]::IsNullOrWhiteSpace([string]$dlg.Tag)) {
            $dlg.Tag = 'Closed'
        }
    })

    # Anzeigen (ohne DialogResult) und Rückgabe
    $null = $dlg.ShowDialog()
    return $dlg.Tag
}

function Open-EditDialog {
    param (
        [hashtable]$item,
        [string]$title,
        [string[]]$PropertyOrder
    )

    Add-Type -AssemblyName PresentationFramework

    $window = New-Object Windows.Window
    $window.Title = $title
    $window.Width = 900                     # breiter
    $window.Height = 800
    $window.SizeToContent = 'Height'        # Breite bleibt fix, Höhe passt sich an
    $window.WindowStartupLocation = 'CenterScreen'

    $scrollViewer = New-Object Windows.Controls.ScrollViewer
    $scrollViewer.VerticalScrollBarVisibility = 'Auto'
    $scrollViewer.HorizontalScrollBarVisibility = 'Disabled' # wir wollen Inhalte strecken statt horizontal scrollen
    $scrollViewer.HorizontalAlignment = 'Stretch'

    $stackPanel = New-Object Windows.Controls.StackPanel
    $stackPanel.Margin = "10"
    $stackPanel.Orientation = 'Vertical'
    $stackPanel.HorizontalAlignment = 'Stretch'              # wichtig für Breitenübernahme

    $textBoxes = [ordered]@{}

    $keys = if ($PropertyOrder) { $PropertyOrder } else { $item.Keys }

    foreach ($key in $keys) {
        $label = New-Object Windows.Controls.Label
        $label.Content = $key
        $label.Margin = "0,0,0,2"
        $label.HorizontalAlignment = 'Left'
        $stackPanel.Children.Add($label)

        $value = $item[$key]
        $isMultiline = ($value -is [string]) -and ($value -match "`n")
        if ($key -match "cmd") { $isMultiline = $true }

        $textBox = New-Object Windows.Controls.TextBox
        $textBox.Text = $value
        $textBox.Margin = "0,0,0,8"
        $textBox.AcceptsReturn = $isMultiline
        $textBox.TextWrapping = 'Wrap'
        $textBox.HorizontalAlignment = 'Stretch'             # << streckt die TextBox in Fensterbreite
        $textBox.MinWidth = 800                               # << sorgt für breite Felder
        if ($isMultiline) {
            $textBox.Height = 100
            $textBox.VerticalScrollBarVisibility = 'Auto'
        } else {
            $textBox.Height = 30
        }

        $stackPanel.Children.Add($textBox)
        $textBoxes[$key] = $textBox
    }

    # --- Neue Buttons: WinGet und MSI ---
    $wingetButton = New-Object Windows.Controls.Button
    $wingetButton.Content = "WinGet"
    $wingetButton.Width = 100
    $wingetButton.Margin = "5"

    $msiButton = New-Object Windows.Controls.Button
    $msiButton.Content = "MSI"
    $msiButton.Width = 100
    $msiButton.Margin = "5"

    $wingetButton.Add_Click({
        $dlg = Show-WinGetBrowserDialog
        if ($dlg -and $dlg.Result -eq 'OK' -and $dlg.Url) {
            $info = Get-WinGetInfoFromUri -Uri $dlg.Url

            # Felder befüllen – flexible Zuordnung
            Set-IfPresent -TextBoxes $textBoxes -CandidateKeys @('Name','App','AppName','ProductName') -Value $info.Name
            Set-IfPresent -TextBoxes $textBoxes -CandidateKeys @('Version','ProductVersion') -Value $info.Version
            Set-IfPresent -TextBoxes $textBoxes -CandidateKeys @('Publisher','Hersteller','Vendor','Company') -Value $info.Publisher
            Set-IfPresent -TextBoxes $textBoxes -CandidateKeys @('WingetId','Id','PackageIdentifier') -Value $info.WingetId
            Set-IfPresent -TextBoxes $textBoxes -CandidateKeys @('cmd','InstallCmd','Command') -Value $info.InstallCmd
        }
    })

    $msiButton.Add_Click({
        try {
            Add-Type -AssemblyName PresentationFramework
            $ofd = New-Object Microsoft.Win32.OpenFileDialog
            $ofd.Title  = "Select MSI"
            $ofd.Filter = "MSI files (*.msi)|*.msi|All files (*.*)|*.*"
            $ofd.Multiselect = $false
            $ok = $ofd.ShowDialog()
            if ($ok -eq $true -and $ofd.FileName) {
                $props = Get-MsiProperties -Path $ofd.FileName

                # Name / Display
                Set-IfPresent -TextBoxes $textBoxes -CandidateKeys @('Display','DisplayName','Name','App','AppName','ProductName') -Value $props.ProductName
                # Version / Publisher
                Set-IfPresent -TextBoxes $textBoxes -CandidateKeys @('Version','ProductVersion') -Value $props.ProductVersion
                Set-IfPresent -TextBoxes $textBoxes -CandidateKeys @('Publisher','Hersteller','Vendor','Company') -Value $props.Manufacturer
                # ProductCode (eigene Spalte, wenn vorhanden)
                Set-IfPresent -TextBoxes $textBoxes -CandidateKeys @('ProductCode','MsiProductCode') -Value $props.ProductCode

                # Install- & Uninstall-Command
                $msiInstall = 'msiexec /i "{0}" /qn' -f $ofd.FileName
                Set-IfPresent -TextBoxes $textBoxes -CandidateKeys @('cmd','InstallCmd','Command') -Value $msiInstall

                if ($props.ProductCode) {
                    $msiUninstall = 'msiexec /x {0} /qn' -f $props.ProductCode  # /x = Uninstall, /qn = quiet
                    Set-IfPresent -TextBoxes $textBoxes -CandidateKeys @('Uninstall','UninstallCmd','RemoveCmd') -Value $msiUninstall
                }
            }
        } catch { }
    })

    # OK/Cancel
    $okButton = New-Object Windows.Controls.Button
    $okButton.Content = "OK"
    $okButton.Width = 100
    $okButton.Margin = "5"
    $okButton.Add_Click({ $window.DialogResult = $true })

    $cancelButton = New-Object Windows.Controls.Button
    $cancelButton.Content = "Cancel"
    $cancelButton.Width = 100
    $cancelButton.Margin = "5"
    $cancelButton.Add_Click({ $window.DialogResult = $false })

    $buttonPanel = New-Object Windows.Controls.StackPanel
    $buttonPanel.Orientation = 'Horizontal'
    $buttonPanel.HorizontalAlignment = 'Right'

    # Reihenfolge: WinGet | MSI | OK | Cancel  (WinGet/MSI links neben OK)
    $buttonPanel.Children.Add($wingetButton)
    $buttonPanel.Children.Add($msiButton)
    $buttonPanel.Children.Add($okButton)
    $buttonPanel.Children.Add($cancelButton)

    $stackPanel.Children.Add($buttonPanel)
    $scrollViewer.Content = $stackPanel
    $window.Content = $scrollViewer

    $result = $window.ShowDialog()

    if ($result -eq $true) {
        $newItem = [ordered]@{}
        foreach ($key in $keys) { $newItem[$key] = $textBoxes[$key].Text }
        return $newItem
    }
}

# --- Hilfsfunktion: IE11-Emulation für WPF WebBrowser aktivieren (HKCU) ---
function Ensure-WebBrowserIE11 {
    try {
        $exeName = [System.Diagnostics.Process]::GetCurrentProcess().ProcessName + ".exe"
        $regPath = 'HKCU:\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION'

        if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
        $current = (Get-ItemProperty -Path $regPath -Name $exeName -ErrorAction SilentlyContinue).$exeName

        # 11001 = IE11 Edge Mode
        if (-not $current) {
            New-ItemProperty -Path $regPath -Name $exeName -Value 11001 -PropertyType DWord -Force | Out-Null
        }
    } catch { }
}

# --- Hilfsfunktion: MSI-Eigenschaften lesen (ProductName/Version/Manufacturer) ---
function Get-MsiProperties {
    param([Parameter(Mandatory=$true)][string]$Path)

    $props = @{}
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $database  = $installer.GetType().InvokeMember('OpenDatabase','InvokeMethod',$null,$installer,@($Path,0))

        foreach ($p in 'ProductName','ProductVersion','Manufacturer','ProductCode') {
            $view   = $database.GetType().InvokeMember('OpenView','InvokeMethod',$null,$database,@("SELECT `Value` FROM `Property` WHERE `Property`='$p'"))
            $null   = $view.GetType().InvokeMember('Execute','InvokeMethod',$null,$view,$null)
            $record = $view.GetType().InvokeMember('Fetch','InvokeMethod',$null,$view,$null)
            if ($record) {
                $val = $record.GetType().InvokeMember('StringData','GetProperty',$null,$record,1)
                if ($val) { $props[$p] = $val }
            }
            $null = $view.GetType().InvokeMember('Close','InvokeMethod',$null,$view,$null)
        }
    } catch { }
    return $props
}

# --- Hilfsfunktion: Mini-Browser für winget.run anzeigen ---
function Show-WinGetBrowserDialog {
    Add-Type -AssemblyName PresentationFramework

    Ensure-WebBrowserIE11

    $dlg = New-Object Windows.Window
    $dlg.Title = "WinGet – Paket suchen"
    $dlg.Width = 900
    $dlg.Height = 600
    $dlg.WindowStartupLocation = 'CenterScreen'

    $root = New-Object Windows.Controls.DockPanel

    $web = New-Object Windows.Controls.WebBrowser
    $web.Source = [Uri] "https://winget.run"

    $btnPanel = New-Object Windows.Controls.StackPanel
    $btnPanel.Orientation = 'Horizontal'
    $btnPanel.HorizontalAlignment = 'Right'
    $btnPanel.Margin = "5"

    $ok = New-Object Windows.Controls.Button
    $ok.Content = "OK"
    $ok.Width = 100
    $ok.Margin = "5"

    $cancel = New-Object Windows.Controls.Button
    $cancel.Content = "Cancel"
    $cancel.Width = 100
    $cancel.Margin = "5"

    $ok.Add_Click({
        # Rückgabe über Tag: aktuelle URL
        try {
            $uri = $web.Source
            if ($uri) {
                $dlg.Tag = @{ Result = 'OK'; Url = $uri.AbsoluteUri }
            } else {
                $dlg.Tag = @{ Result = 'OK'; Url = $null }
            }
        } catch {
            $dlg.Tag = @{ Result = 'OK'; Url = $null }
        }
        $dlg.Close()
    })
    $cancel.Add_Click({
        $dlg.Tag = @{ Result = 'Cancel' }
        $dlg.Close()
    })

    $btnPanel.Children.Add($ok)
    $btnPanel.Children.Add($cancel)

    [Windows.Controls.DockPanel]::SetDock($btnPanel, 'Bottom')
    $root.Children.Add($btnPanel)
    $root.Children.Add($web)
    $dlg.Content = $root

    $null = $dlg.ShowDialog()
    return $dlg.Tag
}

# --- Hilfsfunktion: Infos aus einer winget.run/winstall-Seite ableiten ---
function Get-WinGetInfoFromUri {
    param([string]$Uri)

    $out = @{}
    if (-not $Uri) { return $out }

    try {
        # Server-HTML abrufen
        $resp = Invoke-WebRequest -Uri $Uri -UseBasicParsing -ErrorAction Stop
        $html = $resp.Content

        # 1) Winget-ID aus "winget install ... --id ..." herausziehen
        $id = $null
        $m = [regex]::Match($html,'winget\s+install.*?(?:--id[=\s]+)([A-Za-z0-9\.\-]+)', 'IgnoreCase')
        if ($m.Success) { $id = $m.Groups[1].Value }

        # Fallback: /package/<ID> in der URL
        if (-not $id) {
            $m2 = [regex]::Match($Uri,'/package/([A-Za-z0-9\.\-]+)')
            if ($m2.Success) { $id = $m2.Groups[1].Value }
        }
        if ($id) { $out.WingetId = $id }

        # Name grob aus <h1> extrahieren (Seiten-Titel)
        $m3 = [regex]::Match($html,'<h1[^>]*>(.*?)</h1>','IgnoreCase')
        if ($m3.Success) { $out.Name = ($m3.Groups[1].Value -replace '<.*?>','').Trim() }

        # Publisher/Version heuristisch
        $m4 = [regex]::Match($html,'Publisher[^:<]*:\s*([A-Za-z0-9\.\-\s&,]+)','IgnoreCase')
        if ($m4.Success) { $out.Publisher = $m4.Groups[1].Value.Trim() }

        $m5 = [regex]::Match($html,'Version[^:<]*:\s*([0-9][A-Za-z0-9\.\-\+]+)','IgnoreCase')
        if ($m5.Success) { $out.Version = $m5.Groups[1].Value.Trim() }

        # 2) Wenn Winget vorhanden: autoritative Details via "winget show --id"
        $wg = Get-Command winget -ErrorAction SilentlyContinue
        if ($wg -and $id) {
            $show = winget show --id $id -e --accept-source-agreements 2>&1
            foreach ($line in $show) {
                if ($line -match '^\s*Name\s*:\s*(.+)$')      { $out.Name      = $matches[1].Trim() }
                elseif ($line -match '^\s*Publisher\s*:\s*(.+)$') { $out.Publisher = $matches[1].Trim() }
                elseif ($line -match '^\s*Version\s*:\s*(.+)$')   { $out.Version   = $matches[1].Trim() }
            }
        }

        if ($id) { $out.InstallCmd = "winget install -e --id $id" }
    } catch { }
    return $out
}

# --- Hilfsfunktion: Wert in vorhandene Textbox schreiben (wenn Key existiert) ---
function Set-IfPresent {
    param(
        [Parameter(Mandatory=$true)][hashtable]$TextBoxes,
        [Parameter(Mandatory=$true)][string[]]$CandidateKeys,
        [Parameter()][string]$Value
    )
    if ([string]::IsNullOrWhiteSpace($Value)) { return }
    foreach ($k in $CandidateKeys) {
        if ($TextBoxes.Contains($k)) {
            $TextBoxes[$k].Text = $Value
            break
        }
    }
}

function Open-SelectDialogWithEdit {
    param (
        [string]$CsvPath,
        [string]$title = "Selection",
        [ValidateSet("small", "medium", "large")]
        [string]$size = "medium"
    )

    Add-Type -AssemblyName PresentationFramework | Out-Null

    function Save-Data {
        param (
            [System.Collections.IEnumerable]$data,
            [string[]]$columnOrder
        )
        $exportData = $data | Select-Object -Property $columnOrder
        $exportData | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    }

    # Initiales Laden + Sortieren + Speichern
    $data = Import-Csv -Path $CsvPath -Delimiter ";" |
    ForEach-Object {
        $obj = $_ | Select-Object *
        $obj | Add-Member -MemberType NoteProperty -Name __InternalId -Value ([guid]::NewGuid().ToString())
        $obj
    }

    # Sortieren nach DisplayName, falls vorhanden
    if ("DisplayName" -in $data[0].PSObject.Properties.Name) {
        $data = $data | Sort-Object DisplayName
    }

    # CSV sofort neu schreiben (ohne __InternalId)
    $exportData = $data | Select-Object -Property ($data[0].PSObject.Properties.Name | Where-Object { $_ -ne "__InternalId" })
    $exportData | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"

    # Grid vollständig neu laden
    function Load-Data {
        Import-Csv -Path $CsvPath -Delimiter ";" |
        ForEach-Object {
            $obj = $_ | Select-Object *
            $obj | Add-Member -MemberType NoteProperty -Name __InternalId -Value ([guid]::NewGuid().ToString())
            $obj
        }
    }

    $data = Load-Data

    # Fenster
    $window = New-Object Windows.Window
    $window.Title = $title
    switch ($size) {
        "small"  { $window.Width = 500;  $window.Height = 400 }
        "medium" { $window.Width = 800;  $window.Height = 600 }
        "large"  { $window.Width = 2048; $window.Height = 768 }
    }
    $window.WindowStartupLocation = "CenterScreen"
    $window.ResizeMode = 'CanResize'

    # DataGrid
    $dataGrid = New-Object Windows.Controls.DataGrid
    $dataGrid.AutoGenerateColumns = $false
    $dataGrid.CanUserSortColumns = $true
    $dataGrid.SelectionMode = 'Extended'
    $dataGrid.SelectionUnit = 'FullRow'

    $firstItem   = $data | Select-Object -First 1
    $columnOrder = $firstItem.PSObject.Properties.Name | Where-Object { $_ -ne "__InternalId" }

    foreach ($property in $columnOrder) {
        $column = New-Object Windows.Controls.DataGridTextColumn
        $column.Header  = $property
        $column.Binding = New-Object Windows.Data.Binding($property)
        $column.CanUserSort = $true
        [void]$dataGrid.Columns.Add($column)
    }

    function Refresh-Grid {
        $data = Load-Data
        $dataGrid.ItemsSource = $null
        $dataGrid.ItemsSource = $data
    }
    Refresh-Grid

    # Buttons
    $okButton = New-Object Windows.Controls.Button
    $okButton.Content = "OK"
    $okButton.Width   = 100
    $okButton.Margin  = "5"
    $okButton.IsDefault = $true

    $cancelButton = New-Object Windows.Controls.Button
    $cancelButton.Content = "Cancel"
    $cancelButton.Width   = 100
    $cancelButton.Margin  = "5"
    $cancelButton.IsCancel = $true

    $editButton = New-Object Windows.Controls.Button
    $editButton.Content = "Edit"
    $editButton.Width   = 100
    $editButton.Margin  = "5"

    $newButton = New-Object Windows.Controls.Button
    $newButton.Content = "New"
    $newButton.Width   = 100
    $newButton.Margin  = "5"

    $duplicateButton = New-Object Windows.Controls.Button
    $duplicateButton.Content = "Duplicate"
    $duplicateButton.Width   = 100
    $duplicateButton.Margin  = "5"

    $deleteButton = New-Object Windows.Controls.Button
    $deleteButton.Content = "Delete"
    $deleteButton.Width   = 100
    $deleteButton.Margin  = "5"

    # --- Ereignisse (ohne DialogResult) ---
    $okButton.Add_Click({
        # Auswahl einsammeln
        $selection = @()
        foreach ($item in $dataGrid.SelectedItems) {
            if ($item -isnot [int]) {
                $selection += $item
            }
        }
        # Ergebnis im Tag ablegen
        $window.Tag = [pscustomobject]@{
            Result    = 'Ok'
            Selection = $selection
        }
        $window.Close()
    })

    $cancelButton.Add_Click({
        $window.Tag = [pscustomobject]@{
            Result    = 'Cancel'
            Selection = @()
        }
        $window.Close()
    })

    $editButton.Add_Click({
        if ($dataGrid.SelectedItem) {
            $selected = $dataGrid.SelectedItem
            $hash = @{}
            foreach ($prop in $columnOrder) { $hash[$prop] = $selected.$prop }

            $edited = Open-EditDialog -item $hash -title "Edit entry" -PropertyOrder $columnOrder
            $edited = $edited | Where-Object { $_ -isnot [int] }

            if ($edited) {
                foreach ($key in $edited.Keys) {
                    $selected.$key = $edited[$key]
                }
                Save-Data -data $dataGrid.ItemsSource -columnOrder $columnOrder
                Refresh-Grid
            }
        }
    })

    $newButton.Add_Click({
        $template = [ordered]@{}
        foreach ($property in $columnOrder) { $template[$property] = "" }

        $newItem = Open-EditDialog -item $template -title "Add new entry" -PropertyOrder $columnOrder
        $newItem = $newItem | Where-Object { $_ -isnot [int] }

        if ($newItem) {
            $newObject = New-Object PSObject
            foreach ($key in $newItem.Keys) {
                $newObject | Add-Member -MemberType NoteProperty -Name $key -Value $newItem[$key]
            }
            $newObject | Add-Member -MemberType NoteProperty -Name __InternalId -Value ([guid]::NewGuid().ToString())

            $data = @($dataGrid.ItemsSource) + $newObject
            Save-Data -data $data -columnOrder $columnOrder
            Refresh-Grid
        }
    })

    $duplicateButton.Add_Click({
        if ($dataGrid.SelectedItem) {
            $selected = $dataGrid.SelectedItem
            $hash = @{}
            foreach ($prop in $columnOrder) { $hash[$prop] = $selected.$prop }

            $duplicated = Open-EditDialog -item $hash -title "Edit duplicate entry" -PropertyOrder $columnOrder
            $duplicated = $duplicated | Where-Object { $_ -isnot [int] }

            if ($duplicated) {
                $newObject = New-Object PSObject
                foreach ($key in $duplicated.Keys) {
                    $newObject | Add-Member -MemberType NoteProperty -Name $key -Value $duplicated[$key]
                }
                $newObject | Add-Member -MemberType NoteProperty -Name __InternalId -Value ([guid]::NewGuid().ToString())

                $data = @($dataGrid.ItemsSource) + $newObject
                Save-Data -data $data -columnOrder $columnOrder
                Refresh-Grid
            }
        }
    })

    $deleteButton.Add_Click({
        $selectedItems = $dataGrid.SelectedItems
        if ($selectedItems.Count -gt 0) {
            $confirm = [System.Windows.MessageBox]::Show("Möchtest du die ausgewählten Einträge wirklich löschen?", "Löschen bestätigen", "YesNo", "Warning")
            if ($confirm -eq "Yes") {
                $idsToDelete = $selectedItems | ForEach-Object { $_.__InternalId }
                $data = @($dataGrid.ItemsSource) | Where-Object { $_.__InternalId -notin $idsToDelete }
                Save-Data -data $data -columnOrder $columnOrder
                Refresh-Grid
            } else {
                Write-Host "Löschvorgang abgebrochen."
            }
        } else {
            [System.Windows.MessageBox]::Show("Keine Einträge ausgewählt zum Löschen.", "Hinweis", "OK", "Information") | Out-Null
        }
    })

    # Button-Panel
    $buttonPanel = New-Object Windows.Controls.StackPanel
    $buttonPanel.Orientation = 'Horizontal'
    $buttonPanel.HorizontalAlignment = 'Right'
    $buttonPanel.Margin = "10"
    [void]$buttonPanel.Children.Add($newButton)
    [void]$buttonPanel.Children.Add($duplicateButton)
    [void]$buttonPanel.Children.Add($editButton)
    [void]$buttonPanel.Children.Add($deleteButton)
    [void]$buttonPanel.Children.Add($okButton)
    [void]$buttonPanel.Children.Add($cancelButton)

    # Grid-Layout
    $grid = New-Object Windows.Controls.Grid
    [void]$grid.RowDefinitions.Add((New-Object Windows.Controls.RowDefinition)) # Filter (nicht genutzt)
    [void]$grid.RowDefinitions.Add((New-Object Windows.Controls.RowDefinition)) # Grid
    [void]$grid.RowDefinitions.Add((New-Object Windows.Controls.RowDefinition)) # Buttons
    $grid.RowDefinitions[0].Height = [Windows.GridLength]::Auto
    $grid.RowDefinitions[2].Height = [Windows.GridLength]::Auto

    [void]$grid.Children.Add($dataGrid)
    [Windows.Controls.Grid]::SetRow($dataGrid, 1)

    $footerGrid = New-Object Windows.Controls.Grid
    $footerGrid.Margin = "10"
    $colLeft = New-Object Windows.Controls.ColumnDefinition; $colLeft.Width = "Auto"
    $colFill = New-Object Windows.Controls.ColumnDefinition; $colFill.Width = "*"
    $colRight = New-Object Windows.Controls.ColumnDefinition; $colRight.Width = "Auto"
    [void]$footerGrid.ColumnDefinitions.Add($colLeft)
    [void]$footerGrid.ColumnDefinitions.Add($colFill)
    [void]$footerGrid.ColumnDefinitions.Add($colRight)

#    [Windows.Controls.Grid]::SetColumn($settingsButton, 0)
#    [void]$footerGrid.Children.Add($settingsButton)

    [Windows.Controls.Grid]::SetColumn($buttonPanel, 2)
    [void]$footerGrid.Children.Add($buttonPanel)

    [void]$grid.Children.Add($footerGrid)
    [Windows.Controls.Grid]::SetRow($footerGrid, 2)

    $window.Content = $grid

    # [X]-Schließen: falls nichts gesetzt, auf Closed stellen (leere Auswahl)
    $window.Add_Closing({
        if ($window.Tag -eq $null) {
            $window.Tag = [pscustomobject]@{
                Result    = 'Closed'
                Selection = @()
            }
        }
    })

    # --- WICHTIG: Dialog anzeigen ---
    $null = $window.ShowDialog()

    # Rückgabe: Bei OK → Auswahl; bei Cancel/Closed → leeres Array
    if ($window.Tag -ne $null -and $window.Tag.Result -eq 'Ok') {
        return $window.Tag.Selection
    } else {
        return @()
    }
}

function Open-SelectDialog {
    param (
        $data,
        [string]$title,
        [switch]$large
    )

    Add-Type -AssemblyName PresentationFramework

    # Fenster erstellen
    $window = New-Object Windows.Window
    $window.Title = $title
    if($large){$window.Width = 1024; $window.Height = 768}
    else{$window.Width = 800; $window.Height = 600}

    # DataGrid erstellen
    $dataGrid = New-Object Windows.Controls.DataGrid
    $dataGrid.CanUserSortColumns = $true
    $dataGrid.SelectionMode = 'Extended'
    $dataGrid.SelectionUnit = 'FullRow'
    $dataGrid.AutoGenerateColumns = $false

    # Spalten manuell erzeugen
    $firstItem = $data | Select-Object -First 1
    foreach ($property in $firstItem.PSObject.Properties.Name) {
        $column = New-Object Windows.Controls.DataGridTextColumn
        $column.Header = $property
        $column.Binding = New-Object Windows.Data.Binding($property)
        $column.CanUserSort = $true
        $dataGrid.Columns.Add($column)
    }

    # ItemsSource setzen
    $dataGrid.ItemsSource = $data

    # OK-Button
    $okButton = New-Object Windows.Controls.Button
    $okButton.Height = 40
    $okButton.Width = 100
    $okButton.Content = "OK"
    $okButton.Margin = "5"
    $okButton.Add_Click({
        $window.DialogResult = $true
    })

    # Cancel-Button
    $cancelButton = New-Object Windows.Controls.Button
    $cancelButton.Height = 40
    $cancelButton.Width = 100
    $cancelButton.Content = "Cancel"
    $cancelButton.Margin = "5"
    $cancelButton.Add_Click({
        $window.DialogResult = $false
    })

    # Button-Panel
    $buttonPanel = New-Object Windows.Controls.StackPanel
    $buttonPanel.Orientation = 'Horizontal'
    $buttonPanel.HorizontalAlignment = 'Right'
    $buttonPanel.Margin = "10"
    $buttonPanel.Children.Add($okButton)
    $buttonPanel.Children.Add($cancelButton)

    # Layout-Grid
    $grid = New-Object Windows.Controls.Grid
    $grid.RowDefinitions.Add((New-Object Windows.Controls.RowDefinition))
    $grid.RowDefinitions.Add((New-Object Windows.Controls.RowDefinition))
    $grid.RowDefinitions[1].Height = [Windows.GridLength]::Auto

    $grid.Children.Add($dataGrid)
    [Windows.Controls.Grid]::SetRow($dataGrid, 0)
    $grid.Children.Add($buttonPanel)
    [Windows.Controls.Grid]::SetRow($buttonPanel, 1)

    $window.Content = $grid

    # Dialog anzeigen
    $window.WindowStartupLocation = 'CenterScreen'
    $result = $window.ShowDialog()

    if ($result -eq $true) {
        return $dataGrid.SelectedItems
    }
}

function Edit-SettingsDialog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [System.Windows.Window]$Owner,

        [string[]]$PreferredPaths = @(            
            (Join-Path $rootPath "Config\config.json")
        )
    )

    Add-Type -AssemblyName PresentationCore | Out-Null
    Add-Type -AssemblyName PresentationFramework | Out-Null
    Add-Type -AssemblyName System.Windows.Forms | Out-Null

    function Resolve-ConfigPath {
        param([string[]]$Paths)
        $resolved = $null
        foreach ($p in $Paths) { if ($p -and (Test-Path -LiteralPath $p)) { $resolved = $p; break } }
        if ($resolved -eq $null) {
            foreach ($p in $Paths) {
                if ($p) {
                    $dir = Split-Path -Path $p -Parent
                    if ($dir -and (Test-Path -LiteralPath $dir)) { $resolved = $p; break }
                }
            }
        }
        return $resolved
    }

    function Load-ConfigObject {
        param([string]$Path)
        $obj = $null
        if ($Path -and (Test-Path -LiteralPath $Path)) {
            try {
                $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
                $obj = $raw | ConvertFrom-Json -ErrorAction Stop
            } catch {
                [System.Windows.MessageBox]::Show(("JSON could not be loaded: {0}" -f $_.Exception.Message), "Error", "OK", "Error") | Out-Null
            }
        }
        if ($obj -eq $null) {
            $obj = [PSCustomObject]@{
                packetRoot = "$env:TEMP\IntuneWin32Helper\out"
                removeExistingPacketDirOnEachRun = $false
                cloudName = ""
                apiKey    = ""
                apiSecret = ""
                tenants   = @()
            }
        }
        if ($obj.tenants -eq $null) { $obj.tenants = @() }
        return $obj
    }

    function Save-ConfigObject {
        param([hashtable]$Data, [string]$Path)
        try {
            $dir = Split-Path -Path $Path -Parent
            if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            $json = ConvertTo-Json -InputObject $Data -Depth 8
            Set-Content -LiteralPath $Path -Value $json -Encoding UTF8
            return $true
        } catch {
            [System.Windows.MessageBox]::Show(("JSON could not be saved: {0}" -f $_.Exception.Message), "Error", "OK", "Error") | Out-Null
            return $false
        }
    }

    function Test-GuidString {
        param([string]$Text)
        $isGuid = $false
        try { $null = [Guid]::Parse($Text); $isGuid = $true } catch { $isGuid = $false }
        return $isGuid
    }

    # Pfad & Laden
    $configPath = Resolve-ConfigPath -Paths $PreferredPaths
    $cfg = Load-ConfigObject -Path $configPath

    # Fenster
    $dlg = New-Object Windows.Window
    $dlg.Title = "Edit settings"
    $dlg.Width = 900
    $dlg.Height = 640
    $dlg.WindowStartupLocation = "CenterOwner"
    if ($Owner -ne $null) { $dlg.Owner = $Owner }

    # Root-Grid
    $root = New-Object Windows.Controls.Grid
    $root.Margin = "12"
    $rowPath = New-Object Windows.Controls.RowDefinition; $rowPath.Height = [Windows.GridLength]::Auto
    $rowTabs = New-Object Windows.Controls.RowDefinition; $rowTabs.Height = New-Object Windows.GridLength -ArgumentList 1, ([Windows.GridUnitType]::Star)
    $rowBtns = New-Object Windows.Controls.RowDefinition; $rowBtns.Height = [Windows.GridLength]::Auto
    $null = $root.RowDefinitions.Add($rowPath); $null = $root.RowDefinitions.Add($rowTabs); $null = $root.RowDefinitions.Add($rowBtns)

    # Pfadzeile
    $pathGrid = New-Object Windows.Controls.Grid
    $c1 = New-Object Windows.Controls.ColumnDefinition; $c1.Width = "*"
    $c2 = New-Object Windows.Controls.ColumnDefinition; $c2.Width = "Auto"
    $null = $pathGrid.ColumnDefinitions.Add($c1); $null = $pathGrid.ColumnDefinitions.Add($c2)

    $tbPath = New-Object Windows.Controls.TextBox
    $tbPath.Text = $configPath
    $tbPath.Margin = "0,0,8,0"
    [Windows.Controls.Grid]::SetColumn($tbPath, 0)

    $btnBrowse = New-Object Windows.Controls.Button
    $btnBrowse.Content = "Browse…"
    [Windows.Controls.Grid]::SetColumn($btnBrowse, 1)

    $null = $pathGrid.Children.Add($tbPath); $null = $pathGrid.Children.Add($btnBrowse)
    [Windows.Controls.Grid]::SetRow($pathGrid, 0); $null = $root.Children.Add($pathGrid)

    # Tabs
    $tabs = New-Object Windows.Controls.TabControl
    [Windows.Controls.Grid]::SetRow($tabs, 1); $null = $root.Children.Add($tabs)

    # --- Tab: Allgemein ---
    $tabGeneral = New-Object Windows.Controls.TabItem
    $tabGeneral.Header = "Common"
    $generalGrid = New-Object Windows.Controls.Grid
    $generalGrid.Margin = "10"

    # 2 Spalten
    $gCol1 = New-Object Windows.Controls.ColumnDefinition; $gCol1.Width = "Auto"
    $gCol2 = New-Object Windows.Controls.ColumnDefinition; $gCol2.Width = "*"
    $null = $generalGrid.ColumnDefinitions.Add($gCol1); $null = $generalGrid.ColumnDefinitions.Add($gCol2)

    # ZEILEN:
    # 0: packetRoot
    # 1: removeExistingPacketDirOnEachRun
    # 2: Überschrift "Logo image conversion"
    # 3: cloudName
    # 4: apiKey
    # 5: apiSecret
    # 6: Hyperlink (Cloudinary)
    for ($i=0; $i -lt 7; $i++) { $r = New-Object Windows.Controls.RowDefinition; $r.Height = [Windows.GridLength]::Auto; $null = $generalGrid.RowDefinitions.Add($r) }

    # packetRoot (mit Browse)
    $lblPacketRoot = New-Object Windows.Controls.TextBlock; $lblPacketRoot.Text = "packageRoot:"; $lblPacketRoot.VerticalAlignment = "Center"
    [Windows.Controls.Grid]::SetRow($lblPacketRoot,0); [Windows.Controls.Grid]::SetColumn($lblPacketRoot,0)

    $tbPacketRoot = New-Object Windows.Controls.TextBox; $tbPacketRoot.Margin = "6,0,0,0"; $tbPacketRoot.Text = $cfg.packetRoot
    $btnPacketBrowse = New-Object Windows.Controls.Button; $btnPacketBrowse.Content = "…"; $btnPacketBrowse.Width = 28; $btnPacketBrowse.Margin = "6,0,0,0"
    $cellGrid = New-Object Windows.Controls.Grid
    $cellCol1 = New-Object Windows.Controls.ColumnDefinition; $cellCol1.Width = "*"
    $cellCol2 = New-Object Windows.Controls.ColumnDefinition; $cellCol2.Width = "Auto"
    $null = $cellGrid.ColumnDefinitions.Add($cellCol1); $null = $cellGrid.ColumnDefinitions.Add($cellCol2)
    [Windows.Controls.Grid]::SetColumn($tbPacketRoot,0); [Windows.Controls.Grid]::SetColumn($btnPacketBrowse,1)
    $null = $cellGrid.Children.Add($tbPacketRoot); $null = $cellGrid.Children.Add($btnPacketBrowse)
    [Windows.Controls.Grid]::SetRow($cellGrid,0); [Windows.Controls.Grid]::SetColumn($cellGrid,1)

    # removeExistingPacketDirOnEachRun
    $lblRemove = New-Object Windows.Controls.TextBlock; $lblRemove.Text = "removeExistingPackageDirOnEachRun:"; $lblRemove.VerticalAlignment = "Center"
    [Windows.Controls.Grid]::SetRow($lblRemove,1); [Windows.Controls.Grid]::SetColumn($lblRemove,0)
    $cbRemove = New-Object Windows.Controls.CheckBox; $cbRemove.Margin = "6,0,0,0"; $cbRemove.IsChecked = $false
    if ($cfg.removeExistingPacketDirOnEachRun -is [bool]) { $cbRemove.IsChecked = $cfg.removeExistingPacketDirOnEachRun }
    elseif ($cfg.removeExistingPacketDirOnEachRun -is [string]) {
        $valLower = $cfg.removeExistingPacketDirOnEachRun.ToLower()
        if ($valLower -eq "true") { $cbRemove.IsChecked = $true }
        if ($valLower -eq "false") { $cbRemove.IsChecked = $false }
    }
    [Windows.Controls.Grid]::SetRow($cbRemove,1); [Windows.Controls.Grid]::SetColumn($cbRemove,1)

    # --- Überschrift-Gruppe: Logo image conversion ---
    $lblLogoHeader = New-Object Windows.Controls.TextBlock
    $lblLogoHeader.Text = "Logo image conversion"
    $lblLogoHeader.FontWeight = "Bold"
    $lblLogoHeader.Margin = "0,12,0,4"
    [Windows.Controls.Grid]::SetRow($lblLogoHeader,2); [Windows.Controls.Grid]::SetColumnSpan($lblLogoHeader,2)

    # cloudName
    $lblCloud = New-Object Windows.Controls.TextBlock; $lblCloud.Text = "cloudName:"; $lblCloud.VerticalAlignment = "Center"
    [Windows.Controls.Grid]::SetRow($lblCloud,3); [Windows.Controls.Grid]::SetColumn($lblCloud,0)
    $tbCloud = New-Object Windows.Controls.TextBox; $tbCloud.Margin = "6,0,0,0"; $tbCloud.Text = $cfg.cloudName
    [Windows.Controls.Grid]::SetRow($tbCloud,3); [Windows.Controls.Grid]::SetColumn($tbCloud,1)

    # apiKey
    $lblApiKey = New-Object Windows.Controls.TextBlock; $lblApiKey.Text = "apiKey:"; $lblApiKey.VerticalAlignment = "Center"
    [Windows.Controls.Grid]::SetRow($lblApiKey,4); [Windows.Controls.Grid]::SetColumn($lblApiKey,0)
    $tbApiKey = New-Object Windows.Controls.TextBox; $tbApiKey.Margin = "6,0,0,0"; $tbApiKey.Text = $cfg.apiKey
    [Windows.Controls.Grid]::SetRow($tbApiKey,4); [Windows.Controls.Grid]::SetColumn($tbApiKey,1)

    # apiSecret
    $lblApiSecret = New-Object Windows.Controls.TextBlock; $lblApiSecret.Text = "apiSecret:"; $lblApiSecret.VerticalAlignment = "Center"
    [Windows.Controls.Grid]::SetRow($lblApiSecret,5); [Windows.Controls.Grid]::SetColumn($lblApiSecret,0)
    $tbApiSecret = New-Object Windows.Controls.TextBox; $tbApiSecret.Margin = "6,0,0,0"; $tbApiSecret.Text = $cfg.apiSecret
    [Windows.Controls.Grid]::SetRow($tbApiSecret,5); [Windows.Controls.Grid]::SetColumn($tbApiSecret,1)

    # Hyperlink (Cloudinary)
    $linkTextBlock = New-Object Windows.Controls.TextBlock
    $linkTextBlock.Margin = "0,6,0,0"
    $hyperlink = New-Object System.Windows.Documents.Hyperlink
    $hyperlink.NavigateUri = [Uri]::new("https://cloudinary.com/users/register_free")
    $run = New-Object System.Windows.Documents.Run
    $run.Text = "https://cloudinary.com/users/register_free"
    $null = $hyperlink.Inlines.Add($run)
    $hyperlink.Foreground = [System.Windows.Media.Brushes]::Blue
    $hyperlink.Cursor = [System.Windows.Input.Cursors]::Hand
    # Klick öffnet Systembrowser
    $hyperlink.Add_Click({
        param($sender, $e)
        try {
            $url = $sender.NavigateUri.AbsoluteUri
            Start-Process $url
        } catch {
            [System.Windows.MessageBox]::Show(("Konnte den Link nicht öffnen: {0}" -f $_.Exception.Message), "Error", "OK", "Error") | Out-Null
        }
    })
    $null = $linkTextBlock.Inlines.Add($hyperlink)
    [Windows.Controls.Grid]::SetRow($linkTextBlock,6); [Windows.Controls.Grid]::SetColumnSpan($linkTextBlock,2)

    # Controls in Tab "Allgemein" einfügen
    $null = $generalGrid.Children.Add($lblPacketRoot)
    $null = $generalGrid.Children.Add($cellGrid)
    $null = $generalGrid.Children.Add($lblRemove)
    $null = $generalGrid.Children.Add($cbRemove)
    $null = $generalGrid.Children.Add($lblLogoHeader)
    $null = $generalGrid.Children.Add($lblCloud)
    $null = $generalGrid.Children.Add($tbCloud)
    $null = $generalGrid.Children.Add($lblApiKey)
    $null = $generalGrid.Children.Add($tbApiKey)
    $null = $generalGrid.Children.Add($lblApiSecret)
    $null = $generalGrid.Children.Add($tbApiSecret)
    $null = $generalGrid.Children.Add($linkTextBlock)

    $tabGeneral.Content = $generalGrid
    $null = $tabs.Items.Add($tabGeneral)

    # --- Tab: Tenants (inkl. clientSecret) ---
    $tabTenants = New-Object Windows.Controls.TabItem
    $tabTenants.Header = "Tenants"
    $tenantsGrid = New-Object Windows.Controls.Grid
    $tenantsGrid.Margin = "10"
    $tRow1 = New-Object Windows.Controls.RowDefinition; $tRow1.Height = New-Object Windows.GridLength -ArgumentList 1, ([Windows.GridUnitType]::Star)
    $tRow2 = New-Object Windows.Controls.RowDefinition; $tRow2.Height = [Windows.GridLength]::Auto
    $null = $tenantsGrid.RowDefinitions.Add($tRow1); $null = $tenantsGrid.RowDefinitions.Add($tRow2)

    $dgTenants = New-Object Windows.Controls.DataGrid
    $dgTenants.AutoGenerateColumns = $false
    $dgTenants.CanUserAddRows = $false
    $dgTenants.CanUserDeleteRows = $false
    $dgTenants.IsReadOnly = $false
    $dgTenants.SelectionMode = 'Extended'
    $dgTenants.SelectionUnit = 'FullRow'

    $colTName   = New-Object Windows.Controls.DataGridTextColumn; $colTName.Header   = "name";         $colTName.Binding   = New-Object Windows.Data.Binding("name")
    $colTAppId  = New-Object Windows.Controls.DataGridTextColumn; $colTAppId.Header  = "appid";        $colTAppId.Binding  = New-Object Windows.Data.Binding("appid")
    $colTSecret = New-Object Windows.Controls.DataGridTextColumn; $colTSecret.Header = "clientSecret"; $colTSecret.Binding = New-Object Windows.Data.Binding("clientSecret"); $colTSecret.Width = 260

    $null = $dgTenants.Columns.Add($colTName)
    $null = $dgTenants.Columns.Add($colTAppId)
    $null = $dgTenants.Columns.Add($colTSecret)

    $tenantItems = @()
    foreach ($t in $cfg.tenants) {
        $secret = ""
        if ($t.PSObject.Properties.Name -contains "clientSecret") { $secret = $t.clientSecret }
        $tenantItems += [PSCustomObject]@{ name = $t.name; appid = $t.appid; clientSecret = $secret }
    }
    $dgTenants.ItemsSource = $tenantItems

    [Windows.Controls.Grid]::SetRow($dgTenants, 0); $null = $tenantsGrid.Children.Add($dgTenants)

    $spTenantBtns = New-Object Windows.Controls.StackPanel
    $spTenantBtns.Orientation = "Horizontal"
    $spTenantBtns.HorizontalAlignment = "Right"
    $btnTenantAdd    = New-Object Windows.Controls.Button; $btnTenantAdd.Content    = "Add"; $btnTenantAdd.Margin    = "0,10,8,0"; $btnTenantAdd.Padding    = "14,6"
    $btnTenantEdit   = New-Object Windows.Controls.Button; $btnTenantEdit.Content   = "Edit"; $btnTenantEdit.Margin   = "0,10,8,0"; $btnTenantEdit.Padding   = "14,6"
    $btnTenantDelete = New-Object Windows.Controls.Button; $btnTenantDelete.Content = "Delete";    $btnTenantDelete.Margin = "0,10,0,0";  $btnTenantDelete.Padding = "14,6"
    $null = $spTenantBtns.Children.Add($btnTenantAdd)
    $null = $spTenantBtns.Children.Add($btnTenantEdit)
    $null = $spTenantBtns.Children.Add($btnTenantDelete)
    [Windows.Controls.Grid]::SetRow($spTenantBtns, 1); $null = $tenantsGrid.Children.Add($spTenantBtns)

    $tabTenants.Content = $tenantsGrid
    $null = $tabs.Items.Add($tabTenants)

    # --- Untere Buttons ---
    $spBtns = New-Object Windows.Controls.StackPanel
    $spBtns.Orientation = "Horizontal"
    $spBtns.HorizontalAlignment = "Right"
    $btnReload = New-Object Windows.Controls.Button; $btnReload.Content = "Reload"; $btnReload.Margin = "0,10,8,0"; $btnReload.Padding = "14,6"
    $btnSave   = New-Object Windows.Controls.Button; $btnSave.Content   = "Save"; $btnSave.Margin  = "0,10,8,0"; $btnSave.Padding  = "14,6"
    $btnClose  = New-Object Windows.Controls.Button; $btnClose.Content  = "Close"; $btnClose.Margin = "0,10,0,0";  $btnClose.Padding = "14,6"
    $null = $spBtns.Children.Add($btnReload); $null = $spBtns.Children.Add($btnSave); $null = $spBtns.Children.Add($btnClose)
    [Windows.Controls.Grid]::SetRow($spBtns, 2); $null = $root.Children.Add($spBtns)

    $dlg.Content = $root

    # --- Events ---

    # Datei durchsuchen
    $btnBrowse.Add_Click({
        try {
            $ofd = New-Object System.Windows.Forms.OpenFileDialog
            $ofd.Filter = "JSON file (*.json)|*.json|All files (*.*)|*.*"
            $ofd.Multiselect = $false
            $ofd.CheckFileExists = $false
            $ofd.FileName = "config.json"
            $res = $ofd.ShowDialog()
            if ($res -eq [System.Windows.Forms.DialogResult]::OK) {
                if ($ofd.FileName) { $tbPath.Text = $ofd.FileName }
            }
        } catch {
            [System.Windows.MessageBox]::Show(("File selection falied: {0}" -f $_.Exception.Message), "Error", "OK", "Error") | Out-Null
        }
    })

    # packetRoot Ordnerauswahl
    $btnPacketBrowse.Add_Click({
        try {
            $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
            $fbd.Description = "Select PackageRoot"
            $fbd.ShowNewFolderButton = $true
            $res = $fbd.ShowDialog()
            if ($res -eq [System.Windows.Forms.DialogResult]::OK) {
                if ($fbd.SelectedPath) { $tbPacketRoot.Text = $fbd.SelectedPath }
            }
        } catch {
            [System.Windows.MessageBox]::Show(("Folder selection failed: {0}" -f $_.Exception.Message), "Error", "OK", "Error") | Out-Null
        }
    })

    # Tenants: Hinzufügen
    $btnTenantAdd.Add_Click({
        try {
            $newTenant = Edit-TenantDialog -Owner $dlg
            if ($newTenant -ne $null) {
                $list = @($dgTenants.ItemsSource)
                $list += [PSCustomObject]@{ name = $newTenant.name; appid = $newTenant.appid; clientSecret = $newTenant.clientSecret }
                $dgTenants.ItemsSource = $null
                $dgTenants.ItemsSource = $list
            }
        } catch {
            [System.Windows.MessageBox]::Show(("Tenant edit failed: {0}" -f $_.Exception.Message), "Error", "OK", "Error") | Out-Null
        }
    })

    # Tenants: Bearbeiten
    $btnTenantEdit.Add_Click({
        try {
            $sel = $dgTenants.SelectedItem
            if ($sel -eq $null) {
                [System.Windows.MessageBox]::Show("Please select a tenant.", "Hint", "OK", "Information") | Out-Null
                return
            }
            $initialSecret = ""
            if ($sel.PSObject.Properties.Name -contains "clientSecret") { $initialSecret = $sel.clientSecret }
            $edited = Edit-TenantDialog -Owner $dlg -InitialName $sel.name -InitialAppId $sel.appid -InitialClientSecret $initialSecret
            if ($edited -ne $null) {
                $sel.name         = $edited.name
                $sel.appid        = $edited.appid
                $sel.clientSecret = $edited.clientSecret
                $dgTenants.Items.Refresh()
            }
        } catch {
            [System.Windows.MessageBox]::Show(("Tenant edit failed: {0}" -f $_.Exception.Message), "Error", "OK", "Error") | Out-Null
        }
    })

    # Tenants: Löschen
    $btnTenantDelete.Add_Click({
        try {
            $selected = $dgTenants.SelectedItems
            if ($selected -eq $null -or $selected.Count -eq 0) {
                [System.Windows.MessageBox]::Show("No tenants selected for deletion.", "Hint", "OK", "Information") | Out-Null
                return
            }
            $confirm = [System.Windows.MessageBox]::Show("Delete selected Tenants ?", "Confirm", "YesNo", "Warning")
            if ($confirm -eq "Yes") {
                $remaining = @()
                $current = @($dgTenants.ItemsSource)
                foreach ($item in $current) {
                    $isSelected = $false
                    foreach ($s in $selected) { if ($item -eq $s) { $isSelected = $true; break } }
                    if (-not $isSelected) { $remaining += $item }
                }
                $dgTenants.ItemsSource = $null
                $dgTenants.ItemsSource = $remaining
            }
        } catch {
            [System.Windows.MessageBox]::Show(("Tenant edit failed: {0}" -f $_.Exception.Message), "Error", "OK", "Error") | Out-Null
        }
    })

    # Neu laden
    $btnReload.Add_Click({
        try {
            $cfg = Load-ConfigObject -Path $tbPath.Text
            # Allgemein
            $tbPacketRoot.Text = $cfg.packetRoot
            $cbRemove.IsChecked = $false
            if ($cfg.removeExistingPacketDirOnEachRun -is [bool]) { $cbRemove.IsChecked = $cfg.removeExistingPacketDirOnEachRun }
            elseif ($cfg.removeExistingPacketDirOnEachRun -is [string]) {
                $valLower = $cfg.removeExistingPacketDirOnEachRun.ToLower()
                if ($valLower -eq "true") { $cbRemove.IsChecked = $true }
                if ($valLower -eq "false") { $cbRemove.IsChecked = $false }
            }
            $tbCloud.Text = $cfg.cloudName
            $tbApiKey.Text = $cfg.apiKey
            $tbApiSecret.Text = $cfg.apiSecret
            # Tenants
            $tenantItems = @()
            foreach ($t in $cfg.tenants) {
                $secret = ""
                if ($t.PSObject.Properties.Name -contains "clientSecret") { $secret = $t.clientSecret }
                $tenantItems += [PSCustomObject]@{ name = $t.name; appid = $t.appid; clientSecret = $secret }
            }
            $dgTenants.ItemsSource = $null
            $dgTenants.ItemsSource = $tenantItems
        } catch {
            [System.Windows.MessageBox]::Show(("Reloading failed: {0}" -f $_.Exception.Message), "Error", "OK", "Error") | Out-Null
        }
    })

    # Speichern
    $btnSave.Add_Click({
        try {
            $targetPath = $tbPath.Text
            if (-not $targetPath) {
                [System.Windows.MessageBox]::Show("No path given.", "Hint", "OK", "Warning") | Out-Null
                return
            }

            # Tenants validieren/sammeln
            $tenantArray = @()
            foreach ($row in @($dgTenants.ItemsSource)) {
                $n = $row.name
                $a = $row.appid
                $s = ""
                if ($row.PSObject.Properties.Name -contains "clientSecret") { $s = $row.clientSecret }
                if (-not $n -or $n.Trim().Length -eq 0) {
                    [System.Windows.MessageBox]::Show("Tenant name must not be empty.", "Validation", "OK", "Warning") | Out-Null
                    return
                }
                if (-not (Test-GuidString -Text $a)) {
                    [System.Windows.MessageBox]::Show(("Invalid AppId (GUID) for tenant '{0}'." -f $n), "Validation", "OK", "Warning") | Out-Null
                    return
                }
                $tenantArray += @{ name = $n; appid = $a; clientSecret = $s }
            }

            # Allgemein sammeln
            $removeBool = $false
            if ($cbRemove.IsChecked -eq $true) { $removeBool = $true }

            $data = @{
                packetRoot = $tbPacketRoot.Text
                removeExistingPacketDirOnEachRun = $removeBool
                cloudName = $tbCloud.Text
                apiKey    = $tbApiKey.Text
                apiSecret = $tbApiSecret.Text
                tenants   = $tenantArray
            }

            $ok = Save-ConfigObject -Data $data -Path $targetPath
            if ($ok) { [System.Windows.MessageBox]::Show(("Saved: {0}" -f $targetPath), "Success", "OK", "Information") | Out-Null }
        } catch {
            [System.Windows.MessageBox]::Show(("Save failed: {0}" -f $_.Exception.Message), "Error", "OK", "Error") | Out-Null
        }
    })

    # Schließen
    $btnClose.Add_Click({
        $dlg.DialogResult = $false
        $dlg.Close()
    })

    $null = $dlg.ShowDialog()
    return $true
}
#
function Edit-TenantDialog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [System.Windows.Window]$Owner,
        [string]$InitialName = "",
        [string]$InitialAppId = "",
        [string]$InitialClientSecret = ""
    )

    Add-Type -AssemblyName PresentationCore | Out-Null
    Add-Type -AssemblyName PresentationFramework | Out-Null

    function Test-GuidString {
        param([string]$Text)
        $isGuid = $false
        try {
            $null = [Guid]::Parse($Text)
            $isGuid = $true
        } catch {
            $isGuid = $false
        }
        return $isGuid
    }

    $dlg = New-Object Windows.Window
    $dlg.Title = "Tenant Edit"
    $dlg.Width = 560
    $dlg.Height = 280
    $dlg.WindowStartupLocation = "CenterOwner"
    if ($Owner -ne $null) { $dlg.Owner = $Owner }

    $grid = New-Object Windows.Controls.Grid
    $grid.Margin = "12"

    # Zeilen: name, appid, clientSecret, Buttons
    $row1 = New-Object Windows.Controls.RowDefinition; $row1.Height = [Windows.GridLength]::Auto
    $row2 = New-Object Windows.Controls.RowDefinition; $row2.Height = [Windows.GridLength]::Auto
    $row3 = New-Object Windows.Controls.RowDefinition; $row3.Height = [Windows.GridLength]::Auto
    $row4 = New-Object Windows.Controls.RowDefinition; $row4.Height = [Windows.GridLength]::Auto
    $null = $grid.RowDefinitions.Add($row1)
    $null = $grid.RowDefinitions.Add($row2)
    $null = $grid.RowDefinitions.Add($row3)
    $null = $grid.RowDefinitions.Add($row4)

    # Spalten: Label + Control
    $col1 = New-Object Windows.Controls.ColumnDefinition; $col1.Width = "Auto"
    $col2 = New-Object Windows.Controls.ColumnDefinition; $col2.Width = "*"
    $null = $grid.ColumnDefinitions.Add($col1)
    $null = $grid.ColumnDefinitions.Add($col2)

    # --- name ---
    $lblName = New-Object Windows.Controls.TextBlock
    $lblName.Text = "name:"
    $lblName.VerticalAlignment = "Center"
    [Windows.Controls.Grid]::SetRow($lblName,0); [Windows.Controls.Grid]::SetColumn($lblName,0)

    $tbName = New-Object Windows.Controls.TextBox
    $tbName.Margin = "6,0,0,0"
    $tbName.Text = $InitialName
    [Windows.Controls.Grid]::SetRow($tbName,0); [Windows.Controls.Grid]::SetColumn($tbName,1)

    # --- appid ---
    $lblAppId = New-Object Windows.Controls.TextBlock
    $lblAppId.Text = "appid (GUID):"
    $lblAppId.VerticalAlignment = "Center"
    [Windows.Controls.Grid]::SetRow($lblAppId,1); [Windows.Controls.Grid]::SetColumn($lblAppId,0)

    $tbAppId = New-Object Windows.Controls.TextBox
    $tbAppId.Margin = "6,0,0,0"
    $tbAppId.Text = $InitialAppId
    [Windows.Controls.Grid]::SetRow($tbAppId,1); [Windows.Controls.Grid]::SetColumn($tbAppId,1)

    # --- clientSecret (PasswordBox + Anzeigen-Checkbox) ---
    $lblSecret = New-Object Windows.Controls.TextBlock
    $lblSecret.Text = "clientSecret (optional):"
    $lblSecret.VerticalAlignment = "Center"
    [Windows.Controls.Grid]::SetRow($lblSecret,2); [Windows.Controls.Grid]::SetColumn($lblSecret,0)

    # Wir bauen rechts eine kleine Zelle mit PasswordBox + "anzeigen" Checkbox
    $secretCell = New-Object Windows.Controls.Grid
    $scCol1 = New-Object Windows.Controls.ColumnDefinition; $scCol1.Width = "*"
    $scCol2 = New-Object Windows.Controls.ColumnDefinition; $scCol2.Width = "Auto"
    $null = $secretCell.ColumnDefinitions.Add($scCol1)
    $null = $secretCell.ColumnDefinitions.Add($scCol2)

    $pbSecret = New-Object Windows.Controls.PasswordBox
    $pbSecret.Margin = "6,0,0,0"
    # PasswordBox kann nicht direkt vorbefüllt werden mit Klartext in sicheren Szenarien;
    # in WPF geht Set-Password nur über .Password:
    if ($InitialClientSecret) { $pbSecret.Password = $InitialClientSecret }
    [Windows.Controls.Grid]::SetColumn($pbSecret,0)

    $cbShow = New-Object Windows.Controls.CheckBox
    $cbShow.Content = "anzeigen"
    $cbShow.Margin = "6,0,0,0"
    [Windows.Controls.Grid]::SetColumn($cbShow,1)

    $null = $secretCell.Children.Add($pbSecret)
    $null = $secretCell.Children.Add($cbShow)

    [Windows.Controls.Grid]::SetRow($secretCell,2); [Windows.Controls.Grid]::SetColumn($secretCell,1)

    # Optional: bei "anzeigen" den Secret-Wert temporär in einem TextBox anzeigen
    # (Wir tauschen visuell zwischen PasswordBox und TextBox)
    $tbSecretPlain = New-Object Windows.Controls.TextBox
    $tbSecretPlain.Margin = "6,0,0,0"
    $tbSecretPlain.Visibility = 'Collapsed'
    if ($InitialClientSecret) { $tbSecretPlain.Text = $InitialClientSecret }
    # Wir legen die TextBox oben auf Column 0 derselben Zelle
    [Windows.Controls.Grid]::SetColumn($tbSecretPlain,0)
    $null = $secretCell.Children.Add($tbSecretPlain)

    # Toggle-Logik für Anzeigen
    $cbShow.Add_Checked({
        # Plain sichtbar, PasswordBox ausblenden; Inhalt synchronisieren
        $tbSecretPlain.Text = $pbSecret.Password
        $tbSecretPlain.Visibility = 'Visible'
        $pbSecret.Visibility = 'Collapsed'
    })
    $cbShow.Add_Unchecked({
        # PasswordBox sichtbar, Plain ausblenden; Inhalt synchronisieren
        $pbSecret.Password = $tbSecretPlain.Text
        $pbSecret.Visibility = 'Visible'
        $tbSecretPlain.Visibility = 'Collapsed'
    })

    # --- Buttons ---
    $spBtns = New-Object Windows.Controls.StackPanel
    $spBtns.Orientation = "Horizontal"
    $spBtns.HorizontalAlignment = "Right"
    [Windows.Controls.Grid]::SetRow($spBtns, 3); [Windows.Controls.Grid]::SetColumnSpan($spBtns, 2)

    $btnOK = New-Object Windows.Controls.Button
    $btnOK.Content = "OK"
    $btnOK.Margin = "0,10,8,0"
    $btnOK.Padding = "14,6"

    $btnCancel = New-Object Windows.Controls.Button
    $btnCancel.Content = "Cancel"
    $btnCancel.Margin = "0,10,0,0"
    $btnCancel.Padding = "14,6"

    $null = $spBtns.Children.Add($btnOK)
    $null = $spBtns.Children.Add($btnCancel)

    # --- Layout zusammenfügen ---
    $null = $grid.Children.Add($lblName)
    $null = $grid.Children.Add($tbName)
    $null = $grid.Children.Add($lblAppId)
    $null = $grid.Children.Add($tbAppId)
    $null = $grid.Children.Add($lblSecret)
    $null = $grid.Children.Add($secretCell)
    $null = $grid.Children.Add($spBtns)

    $dlg.Content = $grid

    $script:TenantResult = $null

    $btnCancel.Add_Click({
        $dlg.DialogResult = $false
        $dlg.Close()
    })

    $btnOK.Add_Click({
        try {
            $n = $tbName.Text
            $a = $tbAppId.Text
            # Secret: je nach Sichtbarkeit aus PasswordBox oder Plain-TextBox lesen
            $s = $pbSecret.Password
            if ($tbSecretPlain.Visibility -eq 'Visible') { $s = $tbSecretPlain.Text }

            if (-not $n -or $n.Trim().Length -eq 0) {
                [System.Windows.MessageBox]::Show("Name must not be empty.", "Validation", "OK", "Warning") | Out-Null
                return
            }
            if (-not (Test-GuidString -Text $a)) {
                [System.Windows.MessageBox]::Show("AppId is not a valid GUID.", "Validation", "OK", "Warning") | Out-Null
                return
            }
            # clientSecret darf leer sein; keine weitere Validierung erforderlich
            $script:TenantResult = [PSCustomObject]@{
                name         = $n
                appid        = $a
                clientSecret = $s
            }
            $dlg.DialogResult = $true
            $dlg.Close()
        } catch {
            [System.Windows.MessageBox]::Show(("Error: {0}" -f $_.Exception.Message), "Error", "OK", "Error") | Out-Null
        }
    })

    $null = $dlg.ShowDialog()
    return $script:TenantResult
}

function check-prereqs{
    #Pre-reqs 
    Write-Host "Checking required PowerShell modules"
    $installedmodules=(Get-InstalledModule -ErrorAction SilentlyContinue).Name
    $requiredmodules=@(
        "IntuneWin32App"
    )
    foreach($requiredmodule in $requiredmodules){ 
        if ($installedmodules -notcontains $requiredmodule){
            Write-Host "Required module [$requiredmodule] not detected - installing..." -ForegroundColor Yellow
            Install-Module $requiredmodule -Force -Scope CurrentUser
        }
        else{
            Write-Host "Required module [$requiredmodule] detected." -ForegroundColor Green
        }
    }
    #end function
}

function Convert-WebPToPngCloudinary {
    param (
        [Parameter(Mandatory = $true)]
        [string]$InputFile,

        [Parameter(Mandatory = $true)]
        [string]$CloudName,

        [Parameter(Mandatory = $true)]
        [string]$ApiKey,

        [Parameter(Mandatory = $true)]
        [string]$ApiSecret
    )

    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)
    $dirname = [System.IO.Path]::GetDirectoryName($InputFile)
    $timestamp = [int][double]::Parse((Get-Date -UFormat %s))
    $paramsToSign = "public_id=$fileName&timestamp=$timestamp$ApiSecret"

    $sha1 = New-Object System.Security.Cryptography.SHA1Managed
    $signatureBytes = $sha1.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($paramsToSign))
    $signature = [BitConverter]::ToString($signatureBytes) -replace "-", ""

    $uploadUrl = "https://api.cloudinary.com/v1_1/$CloudName/image/upload"

    # Datei als Base64 kodieren
    $bytes = [System.IO.File]::ReadAllBytes($InputFile)
    $base64 = [Convert]::ToBase64String($bytes)
    $base64File = "data:image/webp;base64,$base64"

    $body = @{
        file       = $base64File
        api_key    = $ApiKey
        timestamp  = $timestamp
        public_id  = $fileName
        signature  = $signature
    }

    try {
        $response = Invoke-RestMethod -Uri $uploadUrl -Method Post -Body $body
        $pngUrl = "https://res.cloudinary.com/$CloudName/image/upload/f_png/$fileName"
        Write-Host "PNG-URL: $pngUrl"

        $outputFile = "$fileName.png"
        Invoke-WebRequest -Uri $pngUrl -OutFile "$dirname\$outputFile"
        Write-Host "PNG-Datei erfolgreich heruntergeladen: $outputFile"
    } catch {
        Write-Error "Error: $_"
    }
}

function Get-FirstFreeDriveLetter {
    $used = (Get-PSDrive -PSProvider 'FileSystem').Name
    $all = [char[]]([byte][char]'C'..[byte][char]'Z')
    foreach ($letter in $all) {
        if ($letter -notin $used) {
            return "$letter" + ":"
        }
    }
}

function Insert-Commands {
    param (
        [string]$FilePath,
        [string[]]$Install,
        [string[]]$Uninstall
    )

    if (!(Test-Path $FilePath)) {
        Write-Error "File '$FilePath' was not found."
        return
    }

    $content = Get-Content $FilePath
    $newContent = @()
    $insertedMarkers = @{}

    if($Install){
        $markers = @{
                '## <Perform Installation tasks here>' = $Install
            }
    }
    elseif($Uninstall){
        $markers = @{
                '## <Perform Uninstallation tasks here>' = $Uninstall
            }
    }
    for ($i = 0; $i -lt $content.Count; $i++) {
        $line = $content[$i]
        $newContent += $line

        foreach ($marker in $markers.Keys) {
            if ($line -like "*$marker*" -and -not $insertedMarkers.ContainsKey($marker)) {
                $newContent += $markers[$marker]
                $insertedMarkers[$marker] = $true
            }
        }
    }

    Set-Content -Path $FilePath -Value $newContent
    Write-Host "Code successfully added after: $($insertedMarkers.Keys -join ', ')"
}

function get-WinGetCommands{
    param(
        [ValidateSet("Install", "Uninstall")]
        [string]$type,
        [string]$id,
        [string]$wgparams
    )

$WingetDefaultCmdsPart1= @'
$logFile = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\<WINGETPROGRAMID>_<ACT>.log"
Write-ADTLogEntry "Action: [<ACT>], PackageID: [<WINGETPROGRAMID>]"
Write-ADTLogEntry "Resolving winget_exe"
try {
    $wingetPaths = Resolve-Path "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller*\winget.exe" -ErrorAction Stop
} catch {
    Write-ADTLogEntry "Winget not installed or path resolution failed: $_"
    exit 1
}
if ($wingetPaths.Count -gt 1) {
    $wingetPaths = $wingetPaths | Sort-Object { (Get-Item $_.Path).CreationTime } -Descending
    $wingetPath = $wingetPaths[0].Path
} elseif ($wingetPaths.Count -eq 1) {
    $wingetPath = $wingetPaths[0].Path
} else {
    Write-ADTLogEntry "Winget executable not found."
    exit 1
}
Write-ADTLogEntry "Using winget path: $wingetPath"
$accpackagree = "--accept-package-agreements"
'@
$WinGetInstallArgs = @'
$arguments = @(
    "<ACT>"
    "--exact"
    "--id", "<WINGETPROGRAMID>"
    "--silent"
    "--accept-source-agreements"
    $accpackagree    
) + "<WINGETPARAMS>"
'@
$WinGetUninstallArgs = @'
$arguments = @(
    "<ACT>"
    "--exact"
    "--id", "<WINGETPROGRAMID>"
    "--silent"
    "--accept-source-agreements"    
) + "<WINGETPARAMS>"
'@
$WingetDefaultCmdsPart3= @'
$ArgumentList = $($arguments -join ' ').trim()
$result=Start-ADTProcess -FilePath $wingetPath -ArgumentList "$ArgumentList" -PassThru
$result = $result -replace 'Ôûê', '░' -replace 'ÔûÆ', '█'
Write-ADTLogEntry "WinGet output:"
Write-ADTLogEntry $result
'@

    $WingetDefaultCmdsPart1 = $WingetDefaultCmdsPart1 -replace "<ACT>", $type -replace "<WINGETPROGRAMID>", $id
    $WinGetInstallArgs = $WinGetInstallArgs -replace "<WINGETPARAMS>", $wgparams -replace "<WINGETPROGRAMID>", $id -replace "<ACT>", $type
    $WinGetUninstallArgs = $WinGetUninstallArgs -replace "<WINGETPARAMS>", $wgparams -replace "<WINGETPROGRAMID>", $id -replace "<ACT>", $type

    switch ($type) {
        "install" {
            $WingetDefaultCmdsPart2 = $WinGetInstallArgs
        }
        "uninstall" {
            $WingetDefaultCmdsPart2 = $WinGetUninstallArgs        
        }
    }
    
    $WingetDefaultCmds = @(
        $WingetDefaultCmdsPart1
        $WingetDefaultCmdsPart2
        $WingetDefaultCmdsPart3
    ) -join "`n"

    return $WingetDefaultCmds

}
