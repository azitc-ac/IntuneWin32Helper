<#
    .SYNOPSIS
    Struktur- und Syntaxpruefungen fuer IntuneWin32Helper.

    .DESCRIPTION
    Jede Pruefung hier gehoert zu einem Fehler, der dieses Repo schon einmal
    getroffen hat. Sie ist absichtlich ausfuehrbar und nicht nur dokumentiert:
    ein Merksatz haelt nicht, ein fehlschlagender Check schon.

    Baut man einen der Fehler zurueck, MUSS dieses Skript fehlschlagen.

    .EXAMPLE
    .\Tests\Invoke-RepoChecks.ps1
    Prueft das Repo und beendet mit Exit-Code 1, wenn etwas gefunden wurde.
#>
[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot)
)

$ErrorActionPreference = "Stop"
$failures = New-Object System.Collections.ArrayList
$checked  = 0

function Add-Failure {
    param([string]$Check, [string]$Detail)
    $null = $failures.Add([PSCustomObject]@{ Check = $Check; Detail = $Detail })
}

$psFiles = Get-ChildItem -Path $RepoRoot -Filter *.ps1 -Recurse -File |
    Where-Object { $_.FullName -notmatch '\\\.git\\' }

if (-not $psFiles) { throw "No .ps1 files found under $RepoRoot" }

# Einmal parsen, Ergebnisse fuer alle Pruefungen wiederverwenden.
$parsed = @{}
foreach ($file in $psFiles) {
    $tokens = $null; $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors)
    $parsed[$file.FullName] = [PSCustomObject]@{
        File = $file; Ast = $ast; Tokens = $tokens; Errors = $errors
    }
}

# ---------------------------------------------------------------------------
# 1) Syntax. Faengt z.B. "$_ -isnot [int]]" - eine doppelte Klammer machte das
#    komplette deploy.ps1 unausfuehrbar, ohne dass es beim Lesen auffiel.
# ---------------------------------------------------------------------------
foreach ($p in $parsed.Values) {
    $checked++
    if ($p.Errors -and $p.Errors.Count -gt 0) {
        foreach ($e in $p.Errors) {
            Add-Failure "Syntax" ("{0}:{1} {2}" -f $p.File.Name, $e.Extent.StartLineNumber, $e.Message)
        }
    }
}

# ---------------------------------------------------------------------------
# 2) BOM. Alle Skripte dieses Repos sind UTF-8 MIT BOM; ohne BOM liest
#    Windows PowerShell 5.1 Umlaute als Mojibake.
# ---------------------------------------------------------------------------
foreach ($file in $psFiles) {
    $checked++
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
    if ($bytes.Length -lt 3 -or $bytes[0] -ne 0xEF -or $bytes[1] -ne 0xBB -or $bytes[2] -ne 0xBF) {
        Add-Failure "BOM" ("{0} hat keinen UTF-8-BOM" -f $file.Name)
    }
}

# ---------------------------------------------------------------------------
# 3) Unterdrueckte .Add()-Rueckgaben. UIElementCollection.Add() gibt den
#    Einfuege-Index zurueck. Unterdrueckt man ihn in einer Funktion nicht,
#    landen int-Werte im Rueckgabewert und vermischen sich mit der Auswahl -
#    genau der Grund fuer die frueheren "-isnot [int]"-Filter.
#    Geprueft wird ueber die AST, nicht per Text: Kommentare zaehlen nicht.
# ---------------------------------------------------------------------------
foreach ($p in $parsed.Values) {
    $checked++
    $pipelines = $p.Ast.FindAll({
        param($n) $n -is [System.Management.Automation.Language.PipelineAst]
    }, $true)

    foreach ($pipe in $pipelines) {
        # Nur Pipelines, deren Wert verworfen wird (direkt als Statement).
        if ($pipe.Parent -isnot [System.Management.Automation.Language.StatementBlockAst] -and
            $pipe.Parent -isnot [System.Management.Automation.Language.NamedBlockAst]) { continue }
        if ($pipe.PipelineElements.Count -ne 1) { continue }

        $el = $pipe.PipelineElements[0]
        if ($el -isnot [System.Management.Automation.Language.CommandExpressionAst]) { continue }

        $expr = $el.Expression
        if ($expr -isnot [System.Management.Automation.Language.InvokeMemberExpressionAst]) { continue }
        if ($expr.Member -isnot [System.Management.Automation.Language.StringConstantExpressionAst]) { continue }
        if ($expr.Member.Value -ne "Add") { continue }

        Add-Failure "AddNotSuppressed" ("{0}:{1} {2} - Rueckgabe nicht unterdrueckt ([void] davor setzen)" -f `
            $p.File.Name, $expr.Extent.StartLineNumber, $expr.Extent.Text)
    }
}

# ---------------------------------------------------------------------------
# 4) Kein Out-GridView. ogv braucht einen STA-Host, fehlt in PowerShell 7 ohne
#    Zusatzmodul und passt nicht zu den WPF-Dialogen des Tools. Geprueft wird
#    ueber die Token, damit erklaerende Kommentare nicht anschlagen.
# ---------------------------------------------------------------------------
foreach ($p in $parsed.Values) {
    $checked++
    foreach ($t in $p.Tokens) {
        if ($t.Kind -ne [System.Management.Automation.Language.TokenKind]::Generic -and
            $t.Kind -ne [System.Management.Automation.Language.TokenKind]::Identifier) { continue }
        if ($t.Text -in @("ogv", "Out-GridView")) {
            Add-Failure "OutGridView" ("{0}:{1} verwendet {2}" -f $p.File.Name, $t.Extent.StartLineNumber, $t.Text)
        }
    }
}

# ---------------------------------------------------------------------------
# 5) Aufrufe passen zu den Signaturen. "Open-SelectDialog -size small" lief ins
#    Leere, weil nur Open-SelectDialogWithEdit ein -size hatte. So etwas faellt
#    erst zur Laufzeit auf - hier faellt es sofort auf.
# ---------------------------------------------------------------------------
$funcParams = @{}
foreach ($p in $parsed.Values) {
    $defs = $p.Ast.FindAll({
        param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]
    }, $true)
    foreach ($d in $defs) {
        $names = @()
        $pb = $d.Body.ParamBlock
        if ($pb) { foreach ($par in $pb.Parameters) { $names += $par.Name.VariablePath.UserPath } }
        elseif ($d.Parameters) { foreach ($par in $d.Parameters) { $names += $par.Name.VariablePath.UserPath } }
        $funcParams[$d.Name] = $names
    }
}

$commonParams = @(
    "Verbose","Debug","ErrorAction","WarningAction","InformationAction","ErrorVariable",
    "WarningVariable","InformationVariable","OutVariable","OutBuffer","PipelineVariable",
    "WhatIf","Confirm"
)

foreach ($p in $parsed.Values) {
    $checked++
    $commands = $p.Ast.FindAll({
        param($n) $n -is [System.Management.Automation.Language.CommandAst]
    }, $true)

    foreach ($cmd in $commands) {
        $name = $cmd.GetCommandName()
        if (-not $name) { continue }
        if (-not $funcParams.ContainsKey($name)) { continue }

        $known = @($funcParams[$name]) + $commonParams

        foreach ($element in $cmd.CommandElements) {
            if ($element -isnot [System.Management.Automation.Language.CommandParameterAst]) { continue }
            $used = $element.ParameterName

            # PowerShell erlaubt eindeutige Abkuerzungen - Praefix genuegt.
            $match = $known | Where-Object { $_ -like ($used + "*") }
            if (-not $match) {
                Add-Failure "UnknownParameter" ("{0}:{1} {2} -{3} - kein solcher Parameter" -f `
                    $p.File.Name, $element.Extent.StartLineNumber, $name, $used)
            }
        }
    }
}

# ---------------------------------------------------------------------------
# 6) Der Tenant laeuft ueber EINEN Pfad. Frueher fragte jedes erzeugte
#    deploy.ps1 selbst - im Bulk-Lauf also pro App erneut.
# ---------------------------------------------------------------------------
$checked++
$templatePath = Join-Path $RepoRoot "Templates\deploy_template.ps1"
if (-not (Test-Path -LiteralPath $templatePath)) {
    Add-Failure "TenantOnePath" "Templates\deploy_template.ps1 fehlt"
}
else {
    $tpl = $parsed[(Get-Item -LiteralPath $templatePath).FullName]

    $hasTenantParam = $false
    if ($tpl.Ast.ParamBlock) {
        $hasTenantParam = [bool]($tpl.Ast.ParamBlock.Parameters |
            Where-Object { $_.Name.VariablePath.UserPath -eq "Tenant" })
    }
    if (-not $hasTenantParam) {
        Add-Failure "TenantOnePath" "deploy_template.ps1 hat keinen Parameter -Tenant - ein Bulk-Lauf fragt sonst pro App"
    }

    $callsInit = $tpl.Ast.FindAll({
        param($n)
        $n -is [System.Management.Automation.Language.CommandAst] -and
        $n.GetCommandName() -eq "Initialize-IntuneConnection"
    }, $true)
    if (-not $callsInit) {
        Add-Failure "TenantOnePath" "deploy_template.ps1 ruft Initialize-IntuneConnection nicht auf"
    }

    $connects = $tpl.Ast.FindAll({
        param($n)
        $n -is [System.Management.Automation.Language.CommandAst] -and
        $n.GetCommandName() -eq "Connect-MSIntuneGraph"
    }, $true)
    if ($connects) {
        Add-Failure "TenantOnePath" ("deploy_template.ps1 ruft Connect-MSIntuneGraph direkt auf (Zeile {0}) - das gehoert in Initialize-IntuneConnection" -f `
            $connects[0].Extent.StartLineNumber)
    }
}

# ---------------------------------------------------------------------------
# 7) break/continue nur innerhalb einer Schleife (oder eines switch) derselben
#    Funktion. Ein break ohne eigene Schleife bricht die Schleife des AUFRUFERS
#    ab: in deployApps schloss "Cancel" dadurch das komplette Tool.
# ---------------------------------------------------------------------------
$loopTypes = @(
    [System.Management.Automation.Language.ForEachStatementAst],
    [System.Management.Automation.Language.ForStatementAst],
    [System.Management.Automation.Language.WhileStatementAst],
    [System.Management.Automation.Language.DoWhileStatementAst],
    [System.Management.Automation.Language.DoUntilStatementAst],
    [System.Management.Automation.Language.SwitchStatementAst]
)

foreach ($p in $parsed.Values) {
    $checked++
    $funcs = $p.Ast.FindAll({
        param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]
    }, $true)

    foreach ($fn in $funcs) {
        $jumps = $fn.Body.FindAll({
            param($n)
            ($n -is [System.Management.Automation.Language.BreakStatementAst]) -or
            ($n -is [System.Management.Automation.Language.ContinueStatementAst])
        }, $true)

        foreach ($jump in $jumps) {
            # Ein Label bezeichnet eine Schleife ausdruecklich - das ist Absicht.
            if ($jump.Label) { continue }

            $node = $jump.Parent
            $inLoop = $false
            while ($node -ne $null -and $node -ne $fn) {
                foreach ($lt in $loopTypes) {
                    if ($node -is $lt) { $inLoop = $true; break }
                }
                if ($inLoop) { break }
                # Verschachtelte Funktion: ab hier zaehlt deren Rahmen, nicht der aeussere.
                if ($node -is [System.Management.Automation.Language.FunctionDefinitionAst]) { break }
                $node = $node.Parent
            }

            if (-not $inLoop) {
                Add-Failure "BreakOutsideLoop" ("{0}:{1} {2} enthaelt '{3}' ohne eigene Schleife - bricht die Schleife des Aufrufers ab (return verwenden)" -f `
                    $p.File.Name, $jump.Extent.StartLineNumber, $fn.Name, $jump.Extent.Text.Trim())
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Ergebnis
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host ("Dateien: {0}   Pruefungen: {1}" -f $psFiles.Count, $checked)

if ($failures.Count -eq 0) {
    Write-Host "Alle Pruefungen bestanden." -ForegroundColor Green
    exit 0
}

Write-Host ("{0} Befund(e):" -f $failures.Count) -ForegroundColor Red
$failures | Group-Object Check | ForEach-Object {
    Write-Host ("  [{0}]" -f $_.Name) -ForegroundColor Yellow
    foreach ($f in $_.Group) { Write-Host ("    - " + $f.Detail) }
}
exit 1
