<#
.SYNOPSIS
    Erstellt einen Bericht über die Berechtigungen von Exchange-Postfächern.
.DESCRIPTION
    Dieses PowerShell-Skript sammelt Postfachberechtigungen (Postfachzugriffsrechte, "Send-As" und "Send on Behalf") für alle Benutzer- und freigegebene Postfächer in einem Exchange-Server und erstellt einen HTML-Report.
    
.EXAMPLE
    PS> .\ExchangePermissionsReport.ps1
    (Erstellt einen Bericht über die Berechtigungen und speichert ihn als HTML-Datei.)

.LINK
    https://github.com/chris-20/ExchangePermissionsReport

.NOTES
    Lizenz: MIT
    
# Exchange Public Folder Permission Analysis Script
# Requires Exchange Management Shell

# Hole das Verzeichnis, in dem das Skript liegt
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$reportPath = Join-Path $scriptPath "PublicFolderPermissions_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
$progressPreference = 'Continue'

function Write-ProgressHelper {
    param(
        [int]$ProgressCounter,
        [int]$TotalFolders,
        [string]$CurrentFolder
    )
    $percentComplete = [Math]::Min(($ProgressCounter / [Math]::Max($TotalFolders, 1)) * 100, 100)
    Write-Progress -Activity "Analysiere Öffentliche Ordner" -Status "Verarbeite: $CurrentFolder" `
        -PercentComplete $percentComplete -CurrentOperation "$ProgressCounter von $TotalFolders Ordnern"
}

function Get-AllPublicFolders {
    try {
        # Verwende Get-PublicFolder ohne Identity-Parameter für Root-Level
        $folders = Get-PublicFolder -Recurse -ErrorAction Stop
        Write-Host "Gefundene Ordner: $($folders.Count)"
        return $folders
    }
    catch {
        Write-Warning "Fehler beim Abrufen der öffentlichen Ordner: $_"
        return @()
    }
}

function Get-FolderPermissions {
    param(
        [Parameter(Mandatory=$true)]
        $Folder,
        [int]$CurrentCount,
        [int]$TotalCount
    )
    
    try {
        Write-ProgressHelper -ProgressCounter $CurrentCount -TotalFolders $TotalCount -CurrentFolder $Folder.Name
        
        # Hole die Berechtigungen
    $permissions = Get-PublicFolderClientPermission -Identity $Folder.EntryID -ErrorAction Stop | 
    Where-Object { $_.User.DisplayName -ne "Standard" -and $_.User.DisplayName -ne "Anonymous" } |
    ForEach-Object {
        @{
            User = $_.User.DisplayName
            AccessRights = $_.AccessRights -join ", "
        }
    }
        
        $result = @{
            Identity = $Folder.Identity
            EntryID = $Folder.EntryID
            Name = $Folder.Name
            FolderPath = $Folder.Identity
            ParentPath = if ($Folder.ParentPath) { $Folder.ParentPath } else { "\" }
            Permissions = $permissions
        }
        
        return $result
    }
    catch {
        Write-Warning "Fehler beim Abrufen der Berechtigungen für $($Folder.Name): $_"
        return $null
    }
}

function Generate-HTMLReport {
    param([array]$FolderData)
    
    $html = @"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <title>Exchange Public Folder Berechtigungen</title>
    <style>
        body {
            font-family: 'Segoe UI', Arial, sans-serif;
            line-height: 1.6;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background-color: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        h1, h2 {
            color: #2c3e50;
            padding-bottom: 10px;
        }
        h1 {
            border-bottom: 2px solid #3498db;
        }
        .summary {
            background-color: #e8f4f8;
            padding: 15px;
            border-radius: 5px;
            margin: 20px 0;
        }
        .folder {
            margin: 20px 0;
            padding: 15px;
            background-color: #f8f9fa;
            border-radius: 5px;
            border-left: 4px solid #3498db;
        }
        .folder-header {
            font-size: 1.2em;
            font-weight: bold;
            color: #2c3e50;
            margin-bottom: 10px;
            padding-bottom: 5px;
            border-bottom: 1px solid #ddd;
        }
        .folder-path {
            font-size: 0.9em;
            color: #666;
            margin-bottom: 10px;
            word-break: break-all;
        }
        .permissions-table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 10px;
            background-color: white;
        }
        .permissions-table th {
            background-color: #3498db;
            color: white;
            padding: 8px;
            text-align: left;
        }
        .permissions-table td {
            padding: 8px;
            border-bottom: 1px solid #ddd;
        }
        .permissions-table tr:nth-child(even) {
            background-color: #f2f2f2;
        }
        .no-permissions {
            color: #666;
            font-style: italic;
            padding: 10px;
            background-color: #f8f9fa;
            border-radius: 4px;
        }
        .timestamp {
            color: #666;
            font-size: 0.9em;
            margin-top: 20px;
            text-align: right;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>Exchange Public Folder Berechtigungsübersicht</h1>
        <div class="summary">
            <h2>Zusammenfassung</h2>
            <p>Gesamtanzahl der analysierten Ordner: $($FolderData.Count)</p>
            <p>Analysezeitpunkt: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')</p>
        </div>
"@

    # Sortiere die Ordner nach Pfad
    $sortedFolders = $FolderData | Sort-Object -Property FolderPath

    foreach ($folder in $sortedFolders) {
        if ($null -ne $folder) {
            $html += @"
            <div class="folder">
                <div class="folder-header">$([System.Web.HttpUtility]::HtmlEncode($folder.Name))</div>
                <div class="folder-path">Pfad: $([System.Web.HttpUtility]::HtmlEncode($folder.FolderPath))</div>
"@
            if ($folder.Permissions -and @($folder.Permissions).Count -gt 0) {
                $html += @"
                <table class="permissions-table">
                    <tr>
                        <th>Benutzer/Gruppe</th>
                        <th>Berechtigungen</th>
                    </tr>
"@
                foreach ($perm in $folder.Permissions) {
                    $html += @"
                    <tr>
                        <td>$([System.Web.HttpUtility]::HtmlEncode($perm.User))</td>
                        <td>$([System.Web.HttpUtility]::HtmlEncode($perm.AccessRights))</td>
                    </tr>
"@
                }
                $html += "</table>"
            }
            else {
                $html += @"
                <div class="no-permissions">Keine spezifischen Berechtigungen gefunden</div>
"@
            }
            $html += "</div>"
        }
    }

    $html += @"
        <div class="timestamp">Report erstellt: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')</div>
    </div>
</body>
</html>
"@

    return $html
}

# Hauptskript
try {
    # Initialisiere Arrays und Zähler
    $counter = 0
    
    # Hole zunächst alle Ordner
    Write-Host "Erfasse öffentliche Ordner..."
    $allFolders = Get-AllPublicFolders
    $totalFolders = $allFolders.Count
    
    if ($totalFolders -eq 0) {
        Write-Warning "Keine öffentlichen Ordner gefunden!"
        return
    }
    
    Write-Host "Starte Analyse von $totalFolders öffentlichen Ordnern..."
    
    # Analysiere jeden Ordner
    $folderPermissions = @()
    
    foreach ($folder in $allFolders) {
        $counter++
        $folderInfo = Get-FolderPermissions -Folder $folder -CurrentCount $counter -TotalCount $totalFolders
        if ($null -ne $folderInfo) {
            $folderPermissions += $folderInfo
        }
    }
    
    # Generiere und speichere HTML-Report
    Write-Host "Erstelle HTML-Report..."
    $htmlReport = Generate-HTMLReport -FolderData $folderPermissions
    $htmlReport | Out-File -FilePath $reportPath -Encoding UTF8
    
    Write-Host "Report wurde erfolgreich erstellt: $reportPath"
}
catch {
    Write-Error "Ein Fehler ist aufgetreten: $_"
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}
