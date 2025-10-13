function Open-EditDialog {
    param (
        [hashtable]$item,
        [string]$title,
        [string[]]$PropertyOrder
    )

    Add-Type -AssemblyName PresentationFramework

    $window = New-Object Windows.Window
    $window.Title = $title
    $window.Width = 450
    $window.Height = 800
    $window.WindowStartupLocation = 'CenterScreen'

    $scrollViewer = New-Object Windows.Controls.ScrollViewer
    $scrollViewer.VerticalScrollBarVisibility = 'Auto'

    $stackPanel = New-Object Windows.Controls.StackPanel
    $stackPanel.Margin = "10"
    $stackPanel.Orientation = 'Vertical'

    $textBoxes = [ordered]@{}

    # Verwende PropertyOrder, wenn vorhanden, sonst item.Keys
    $keys = if ($PropertyOrder) { $PropertyOrder } else { $item.Keys }

    foreach ($key in $keys) {
        $label = New-Object Windows.Controls.Label
        $label.Content = $key
        $label.Margin = "0,0,0,2"
        $stackPanel.Children.Add($label)

        $value = $item[$key]
        $isMultiline = ($value -is [string]) -and ($value -match "`n")
        if($key -match "cmd"){$isMultiline = $true}
        $textBox = New-Object Windows.Controls.TextBox
        $textBox.Text = $value
        $textBox.Margin = "0,0,0,8"
        $textBox.AcceptsReturn = $isMultiline
        $textBox.TextWrapping = 'Wrap'
        if ($isMultiline) {
            $textBox.Height = 60
        } else {
            $textBox.Height = 25
        }

        $stackPanel.Children.Add($textBox)
        $textBoxes[$key] = $textBox
    }

    $okButton = New-Object Windows.Controls.Button
    $okButton.Content = "OK"
    $okButton.Width = 100
    $okButton.Margin = "5"
    $okButton.Add_Click({ $window.DialogResult = $true })

    $cancelButton = New-Object Windows.Controls.Button
    $cancelButton.Content = "Abbrechen"
    $cancelButton.Width = 100
    $cancelButton.Margin = "5"
    $cancelButton.Add_Click({ $window.DialogResult = $false })

    $buttonPanel = New-Object Windows.Controls.StackPanel
    $buttonPanel.Orientation = 'Horizontal'
    $buttonPanel.HorizontalAlignment = 'Right'
    $buttonPanel.Children.Add($okButton)
    $buttonPanel.Children.Add($cancelButton)

    $stackPanel.Children.Add($buttonPanel)
    $scrollViewer.Content = $stackPanel
    $window.Content = $scrollViewer

    $result = $window.ShowDialog()

    if ($result -eq $true) {
        $newItem = [ordered]@{}
        foreach ($key in $keys) {
            $newItem[$key] = $textBoxes[$key].Text
        }
        return $newItem
    }
}

function Open-SelectDialogWithEdit {
    param (
        [string]$CsvPath,
        [string]$title = "Auswahl",
        [ValidateSet("small", "medium", "large")]
        [string]$size = "medium"
    )

    Add-Type -AssemblyName PresentationFramework

    function Save-Data {
        param (
            [System.Collections.IEnumerable]$data,
            [string[]]$columnOrder
        )
        $exportData = $data | Select-Object -Property $columnOrder
        $exportData | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    }
    
    # Initiales Laden + Sortieren + Speichern
    $data = Import-Csv -Path $CsvPath -Delimiter ";" | ForEach-Object {
        $obj = $_ | Select-Object *
        $obj | Add-Member -MemberType NoteProperty -Name __InternalId -Value ([guid]::NewGuid().ToString())
        $obj
    }

    # Sortieren nach DisplayName, falls vorhanden
    if ("DisplayName" -in $data[0].PSObject.Properties.Name) {
        $data = $data | Sort-Object DisplayName
    }

    # CSV sofort neu schreiben
    $exportData = $data | Select-Object -Property ($data[0].PSObject.Properties.Name | Where-Object { $_ -ne "__InternalId" })
    $exportData | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"

    # Grid vollständig neu laden
    function Load-Data {
        Import-Csv -Path $CsvPath -Delimiter ";" | ForEach-Object {
            $obj = $_ | Select-Object *
            $obj | Add-Member -MemberType NoteProperty -Name __InternalId -Value ([guid]::NewGuid().ToString())
            $obj
        }
    }
    $data = Load-Data

    $window = New-Object Windows.Window
    $window.Title = $title
    switch ($size) {
        "small" { $window.Width = 500; $window.Height = 400 }
        "medium" { $window.Width = 800; $window.Height = 600 }
        "large" { $window.Width = 2048; $window.Height = 768 }
    }

    $dataGrid = New-Object Windows.Controls.DataGrid
    $dataGrid.AutoGenerateColumns = $false
    $dataGrid.CanUserSortColumns = $true
    $dataGrid.SelectionMode = 'Extended'
    $dataGrid.SelectionUnit = 'FullRow'

    $firstItem = $data | Select-Object -First 1
    $columnOrder = $firstItem.PSObject.Properties.Name | Where-Object { $_ -ne "__InternalId" }

    foreach ($property in $columnOrder) {
        $column = New-Object Windows.Controls.DataGridTextColumn
        $column.Header = $property
        $column.Binding = New-Object Windows.Data.Binding($property)
        $column.CanUserSort = $true
        $dataGrid.Columns.Add($column)
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
    $okButton.Width = 100
    $okButton.Margin = "5"
    $okButton.Add_Click({ $window.DialogResult = $true })

    $cancelButton = New-Object Windows.Controls.Button
    $cancelButton.Content = "Abbrechen"
    $cancelButton.Width = 100
    $cancelButton.Margin = "5"
    $cancelButton.Add_Click({ $window.DialogResult = $false })

    $editButton = New-Object Windows.Controls.Button
    $editButton.Content = "Bearbeiten"
    $editButton.Width = 100
    $editButton.Margin = "5"
    $editButton.Add_Click({
        if ($dataGrid.SelectedItem) {
            $selected = $dataGrid.SelectedItem
            $hash = @{}
            foreach ($prop in $columnOrder) {
                $hash[$prop] = $selected.$prop
            }
            $edited = Open-EditDialog -item $hash -title "Eintrag bearbeiten" -PropertyOrder $columnOrder
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

    $newButton = New-Object Windows.Controls.Button
    $newButton.Content = "Neu"
    $newButton.Width = 100
    $newButton.Margin = "5"
    $newButton.Add_Click({
        $template = [ordered]@{}
        foreach ($property in $columnOrder) {
            $template[$property] = ""
        }
        $newItem = Open-EditDialog -item $template -title "Neuen Eintrag hinzufügen" -PropertyOrder $columnOrder
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

    $duplicateButton = New-Object Windows.Controls.Button
    $duplicateButton.Content = "Duplizieren"
    $duplicateButton.Width = 100
    $duplicateButton.Margin = "5"
    $duplicateButton.Add_Click({
        if ($dataGrid.SelectedItem) {
            $selected = $dataGrid.SelectedItem
            $hash = @{ }
            foreach ($prop in $columnOrder) {
                $hash[$prop] = $selected.$prop
            }
            $duplicated = Open-EditDialog -item $hash -title "Duplizierten Eintrag bearbeiten" -PropertyOrder $columnOrder
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
    $deleteButton = New-Object Windows.Controls.Button
    $deleteButton.Content = "Löschen"
    $deleteButton.Width = 100
    $deleteButton.Margin = "5"
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
            [System.Windows.MessageBox]::Show("Keine Einträge ausgewählt zum Löschen.", "Hinweis", "OK", "Information")
        }
    })

    $buttonPanel = New-Object Windows.Controls.StackPanel
    $buttonPanel.Orientation = 'Horizontal'
    $buttonPanel.HorizontalAlignment = 'Right'
    $buttonPanel.Margin = "10"
    $buttonPanel.Children.Add($newButton)
    $buttonPanel.Children.Add($duplicateButton)
    $buttonPanel.Children.Add($editButton)
    $buttonPanel.Children.Add($deleteButton)
    $buttonPanel.Children.Add($okButton)
    $buttonPanel.Children.Add($cancelButton)

    $grid = New-Object Windows.Controls.Grid
    $grid.RowDefinitions.Add((New-Object Windows.Controls.RowDefinition)) # Filter (nicht genutzt)
    $grid.RowDefinitions.Add((New-Object Windows.Controls.RowDefinition)) # Grid
    $grid.RowDefinitions.Add((New-Object Windows.Controls.RowDefinition)) # Buttons
    $grid.RowDefinitions[0].Height = [Windows.GridLength]::Auto
    $grid.RowDefinitions[2].Height = [Windows.GridLength]::Auto

    $grid.Children.Add($dataGrid)
    [Windows.Controls.Grid]::SetRow($dataGrid, 1)
    $grid.Children.Add($buttonPanel)
    [Windows.Controls.Grid]::SetRow($buttonPanel, 2)

    $window.Content = $grid
    $result = $window.ShowDialog()

    if ($result -eq $true) {
        return $dataGrid.SelectedItems | Where-Object { $_ -isnot [int] }
    }
}

function Open-SelectDialog {
    param (
        $data,
        [string]$title
    )

    Add-Type -AssemblyName PresentationFramework

    # Fenster erstellen
    $window = New-Object Windows.Window
    $window.Title = $title
    $window.Width = 800
    $window.Height = 600

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
    $result = $window.ShowDialog()

    if ($result -eq $true) {
        return $dataGrid.SelectedItems
    }
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
            Install-Module $requiredmodule -Force
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
        Write-Error "Fehler: $_"
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
        Write-Error "Datei '$FilePath' wurde nicht gefunden."
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
    Write-Host "Code erfolgreich eingefügt nach: $($insertedMarkers.Keys -join ', ')"
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
