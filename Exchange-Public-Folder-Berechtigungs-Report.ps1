<#
.SYNOPSIS
    Erstellt einen Bericht über die Berechtigungen von Exchange-Postfächern.
.DESCRIPTION
    Dieses PowerShell-Skript sammelt Postfachberechtigungen (Postfachzugriffsrechte, "Send-As" und "Send on Behalf") für alle Benutzer- und freigegebene Postfächer in einem Exchange-Server und erstellt einen HTML-Report.
    
.EXAMPLE
    PS> .\ExchangePermissionsReport.ps1
    (Erstellt einen Bericht über die Berechtigungen und speichert ihn als HTML-Datei.)

.LINK
    https://github.com/chris-20/Exchange-On-Premise-Public-Folder-Berechtigungs-Report

.NOTES
    Lizenz: MIT
    Version: 1.0
#>

# Exchange Public Folder Permission Analysis Script
[CmdletBinding()]
param(
    [string]$OutputPath = "",
    [switch]$IncludeDefaultPermissions
)

# Setze Encoding für korrekte Umlaut-Darstellung
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'

# BOM (Byte Order Mark) für UTF-8 hinzufügen
$BOM = [System.Text.UTF8Encoding]::new($true)

# Setze Ausgabepfad
if ([string]::IsNullOrEmpty($OutputPath)) {
    $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
    $timestamp = Get-Date -Format "dd.MM.yyyy HH-mm"
    $reportPath = Join-Path $scriptPath "Public_Folder_Permissions_Report_$timestamp.html"
} else {
    $reportPath = $OutputPath
}

$progressPreference = 'Continue'

# Prüfe und lade Exchange Management Shell
function Initialize-ExchangeEnvironment {
    try {
        $exchangeSnapin = Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue
        
        if ($null -eq $exchangeSnapin) {
            Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
            Write-Host "Exchange Management Shell erfolgreich geladen." -ForegroundColor Green
        }
        
        $null = Get-ExchangeServer -ErrorAction Stop
        return $true
    }
    catch {
        Write-Error "Fehler beim Laden der Exchange Management Shell: $_"
        Write-Host "Bitte stellen Sie sicher, dass:"
        Write-Host "1. Sie das Skript auf einem Exchange-Server ausführen"
        Write-Host "2. Sie als Exchange-Administrator berechtigt sind"
        Write-Host "3. Sie PowerShell als Administrator ausführen"
        return $false
    }
}

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
        Write-Verbose "Hole alle öffentlichen Ordner..."
        $folders = Get-PublicFolder -Recurse -ErrorAction Stop
        Write-Host "Gefundene Ordner: $($folders.Count)" -ForegroundColor Green
        return $folders
    }
    catch {
        Write-Warning "Fehler beim Abrufen der öffentlichen Ordner: $_"
        throw
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
        
        $permissionFilter = {
            if ($IncludeDefaultPermissions) {
                return $true
            }
            return $_.User.DisplayName -ne "Standard" -and $_.User.DisplayName -ne "Anonymous"
        }
        
        $permissions = Get-PublicFolderClientPermission -Identity $Folder.EntryID -ErrorAction Stop | 
            Where-Object $permissionFilter |
            ForEach-Object {
                $userType = "Unbekannt"
                try {
                    $adObject = Get-User $_.User.DisplayName -ErrorAction SilentlyContinue
                    if ($adObject) {
                        $userType = "Benutzer"
                    } else {
                        $adGroup = Get-Group $_.User.DisplayName -ErrorAction SilentlyContinue
                        if ($adGroup) {
                            $userType = "Gruppe"
                        }
                    }
                } catch {
                    Write-Verbose "Konnte Typ für $($_.User.DisplayName) nicht ermitteln: $_"
                }

                @{
                    User = $_.User.DisplayName
                    AccessRights = ($_.AccessRights | Sort-Object) -join ", "
                    UserType = $userType
                }
            }
        
        return @{
            Identity = $Folder.Identity
            EntryID = $Folder.EntryID
            Name = $Folder.Name
            FolderPath = $Folder.Identity
            ParentPath = if ($Folder.ParentPath) { $Folder.ParentPath } else { "\" }
            FolderSize = Get-PublicFolderStatistics -Identity $Folder.EntryID | Select-Object -ExpandProperty TotalItemSize
            Permissions = $permissions
        }
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
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Exchange Public Folder Berechtigungen</title>
    <link rel="icon" type="image/svg+xml" href="data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNDAgMjQwIj48ZGVmcz48bGluZWFyR3JhZGllbnQgaWQ9InByaW1hcnlHcmFkaWVudCIgeDE9IjAlIiB5MT0iMCUiIHgyPSIxMDAlIiB5Mj0iMTAwJSI+PHN0b3Agb2Zmc2V0PSIwJSIgc3R5bGU9InN0b3AtY29sb3I6IzJCNTg3NiIvPjxzdG9wIG9mZnNldD0iMTAwJSIgc3R5bGU9InN0b3AtY29sb3I6IzRFNDM3NiIvPjwvbGluZWFyR3JhZGllbnQ+PGxpbmVhckdyYWRpZW50IGlkPSJhY2NlbnRHcmFkaWVudCIgeDE9IjAlIiB5MT0iMCUiIHgyPSIxMDAlIiB5Mj0iMCUiPjxzdG9wIG9mZnNldD0iMCUiIHN0eWxlPSJzdG9wLWNvbG9yOiMwMGM2ZmYiLz48c3RvcCBvZmZzZXQ9IjEwMCUiIHN0eWxlPSJzdG9wLWNvbG9yOiMwMDcyZmYiLz48L2xpbmVhckdyYWRpZW50PjwvZGVmcz48Y2lyY2xlIGN4PSIxMjAiIGN5PSIxMjAiIHI9IjExMCIgZmlsbD0idXJsKCNwcmltYXJ5R3JhZGllbnQpIi8+PHBhdGggZD0iTSA2MCwxMjAgTCA5MCwxMjAgTCAxMDUsNzAgTCAxMzUsMTcwIEwgMTUwLDEyMCBMIDE4MCwxMjAiIGZpbGw9Im5vbmUiIHN0cm9rZT0idXJsKCNhY2NlbnRHcmFkaWVudCkiIHN0cm9rZS13aWR0aD0iMTQiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCIvPjwvc3ZnPg==">
    <style>
        :root {
            --primary-gradient-start: #2B5876;
            --primary-gradient-end: #4E4376;
            --accent-gradient-start: #00c6ff;
            --accent-gradient-end: #0072ff;
            --background-color: #f8fafc;
            --card-background: #ffffff;
            --text-primary: #1e293b;
            --text-secondary: #64748b;
            --border-radius: 12px;
            --transition-smooth: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            --shadow-sm: 0 1px 3px rgba(0,0,0,0.12), 0 1px 2px rgba(0,0,0,0.24);
            --shadow-md: 0 4px 6px -1px rgba(0,0,0,0.1), 0 2px 4px -2px rgba(0,0,0,0.1);
            --shadow-lg: 0 10px 15px -3px rgba(0,0,0,0.1), 0 4px 6px -4px rgba(0,0,0,0.1);
        }

        body { 
            font-family: 'Segoe UI', system-ui, sans-serif;
            line-height: 1.6; 
            margin: 0; 
            padding: 32px; 
            background: linear-gradient(135deg, var(--background-color), #ffffff);
            color: var(--text-primary);
            min-height: 100vh;
        }

        .container { 
            max-width: 1200px; 
            margin: 0 auto; 
            background-color: var(--card-background); 
            padding: 32px;
            border-radius: var(--border-radius);
            box-shadow: var(--shadow-lg);
            border: 1px solid rgba(43, 88, 118, 0.08);
            position: relative;
            overflow: hidden;
        }

        .container::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 4px;
            background: linear-gradient(90deg, 
                var(--accent-gradient-start), 
                var(--accent-gradient-end), 
                var(--primary-gradient-start), 
                var(--primary-gradient-end));
            background-size: 300% 100%;
            animation: gradientMove 8s ease infinite;
        }

        @keyframes gradientMove {
            0% { background-position: 0% 50%; }
            50% { background-position: 100% 50%; }
            100% { background-position: 0% 50%; }
        }

        h1 { 
            color: var(--text-primary);
            font-size: 2.25rem;
            font-weight: 700;
            letter-spacing: -0.025em;
            margin-bottom: 2rem;
            padding-bottom: 1rem;
            background: linear-gradient(135deg, var(--primary-gradient-start), var(--primary-gradient-end));
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            position: relative;
        }

        h1::after {
            content: '';
            position: absolute;
            bottom: 0;
            left: 0;
            right: 0;
            height: 2px;
            background: linear-gradient(90deg, 
                var(--accent-gradient-start), 
                var(--accent-gradient-end));
            border-radius: 2px;
            opacity: 0.8;
        }

        .summary {
            margin: 2rem 0;
            padding: 1.5rem;
            background: linear-gradient(135deg, 
                rgba(43, 88, 118, 0.03), 
                rgba(78, 67, 118, 0.03));
            border-radius: var(--border-radius);
            border: 1px solid rgba(43, 88, 118, 0.08);
        }

        .folder-group { 
            margin: 2rem 0;
            border-radius: var(--border-radius);
            overflow: hidden;
            box-shadow: var(--shadow-sm);
            border: 1px solid rgba(43, 88, 118, 0.08);
            transition: var(--transition-smooth);
        }

        .folder-group:hover {
            box-shadow: var(--shadow-md);
        }

        .folder-group-header { 
            background: linear-gradient(135deg, 
                var(--primary-gradient-start), 
                var(--primary-gradient-end));
            color: white;
            padding: 1rem;
            cursor: pointer;
            user-select: none;
            font-weight: 500;
        }

        .folder-group-content { 
            display: none;
            padding: 1rem;
            background: linear-gradient(135deg, 
                var(--card-background), 
                rgba(248, 250, 252, 0.5));
        }

        .folder-group-content.active { 
            display: block;
        }

        .folder { 
            margin: 1rem 0;
            padding: 1.5rem;
            background: var(--card-background);
            border-radius: var(--border-radius);
            box-shadow: var(--shadow-sm);
            border: 1px solid rgba(43, 88, 118, 0.08);
            transition: var(--transition-smooth);
        }

        .folder:hover {
            box-shadow: var(--shadow-md);
            transform: translateX(4px);
        }

        .folder-header { 
            font-size: 1.2em;
            font-weight: 700;
            color: var(--text-primary);
            margin-bottom: 1rem;
        }

        .folder-info { 
            font-size: 0.9em;
            color: var(--text-secondary);
            margin-bottom: 1rem;
        }

        .permissions-table {
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
            margin: 1rem 0;
            border-radius: var(--border-radius);
            overflow: hidden;
            box-shadow: var(--shadow-sm);
        }

        .permissions-table th {
            background: linear-gradient(135deg, 
                var(--primary-gradient-start), 
                var(--primary-gradient-end));
            color: white;
            padding: 1rem;
            text-align: left;
            font-weight: 500;
        }

        .permissions-table td {
            padding: 0.75rem 1rem;
            border-bottom: 1px solid rgba(43, 88, 118, 0.08);
        }

        .permissions-table tr:last-child td {
            border-bottom: none;
        }

        .permissions-table tr:nth-child(even) {
            background: linear-gradient(135deg, 
                rgba(43, 88, 118, 0.02), 
                rgba(78, 67, 118, 0.02));
        }

        .permissions-table tr:hover {
            background: linear-gradient(135deg, 
                rgba(0, 198, 255, 0.05), 
                rgba(0, 114, 255, 0.05));
        }

        .search-box {
            margin: 2rem 0;
            padding: 1rem;
            background: linear-gradient(135deg, 
                rgba(43, 88, 118, 0.02), 
                rgba(78, 67, 118, 0.02));
            border-radius: var(--border-radius);
            border: 1px solid rgba(43, 88, 118, 0.08);
        }

        .search-box input {
            width: 100%;
            padding: 0.75rem 1rem;
            border: 1px solid rgba(43, 88, 118, 0.16);
            border-radius: var(--border-radius);
            font-size: 1rem;
            transition: var(--transition-smooth);
        }

        .search-box input:focus {
            outline: none;
            border-color: var(--accent-gradient-start);
            box-shadow: 0 0 0 3px rgba(0, 198, 255, 0.1);
        }

        .timestamp {
            color: var(--text-secondary);
            font-size: 0.875rem;
            margin-top: 2rem;
            padding: 1rem;
            background: linear-gradient(135deg, 
                rgba(43, 88, 118, 0.02), 
                rgba(78, 67, 118, 0.02));
            border-radius: var(--border-radius);
            text-align: right;
            border: 1px solid rgba(43, 88, 118, 0.06);
        }

        @media (max-width: 768px) {
            body {
                padding: 16px;
            }

            .container {
                padding: 16px;
            }
        }
    </style>
    <script>
    function toggleFolderGroup(element) {
        const content = element.nextElementSibling;
        content.classList.toggle('active');
    }

    function searchFolders() {
        const input = document.getElementById('searchInput').value.toLowerCase();
        const folders = document.getElementsByClassName('folder');
        const folderGroups = document.getElementsByClassName('folder-group');
        
        for (let folder of folders) {
            const text = folder.textContent.toLowerCase();
            const shouldShow = text.includes(input);
            folder.style.display = shouldShow ? '' : 'none';
            
            if (shouldShow) {
                let parent = folder.closest('.folder-group-content');
                if (parent) {
                    parent.classList.add('active');
                }
            }
        }
        
        for (let group of folderGroups) {
            const visibleFolders = group.querySelectorAll('.folder[style=""]').length;
            group.style.display = visibleFolders > 0 ? '' : 'none';
        }
    }
    </script>
</head>
<body>
    <div class="container">
        <h1>Exchange Public Folder Berechtigungsübersicht</h1><div class="summary">
            <h2>Zusammenfassung</h2>
            <p>Gesamtanzahl der analysierten Ordner: $($FolderData.Count)</p>
            <p>Analysezeitpunkt: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')</p>
        </div>

        <div class="search-box">
            <input type="text" id="searchInput" onkeyup="searchFolders()" placeholder="Suche nach Ordnern oder Berechtigungen...">
        </div>
        
        <div class="folder-tree">
"@

    # Gruppiere Ordner nach ihrem Hauptpfad
    $groupedFolders = $FolderData | Group-Object { 
        $pathParts = $_.FolderPath -split '\\'
        if ($pathParts.Count -gt 2) {
            $pathParts[1]  # Nimm den ersten Teil nach dem Root
        } else {
            "Root"
        }
    }

    foreach ($group in ($groupedFolders | Sort-Object Name)) {
        $groupName = if ($group.Name -eq "Root") { "Root-Verzeichnis" } else { $group.Name }
        
        $html += @"
            <div class="folder-group">
                <div class="folder-group-header" onclick="toggleFolderGroup(this)">
                    $([System.Web.HttpUtility]::HtmlEncode($groupName)) ($($group.Count) Ordner)
                </div>
                <div class="folder-group-content">
"@

        $sortedFolders = $group.Group | Sort-Object FolderPath
        
        foreach ($folder in $sortedFolders) {
            if ($null -ne $folder) {
                $html += @"
                    <div class="folder">
                        <div class="folder-header">$([System.Web.HttpUtility]::HtmlEncode($folder.Name))</div>
                        <div class="folder-info">
                            Pfad: $([System.Web.HttpUtility]::HtmlEncode($folder.FolderPath))<br>
                            Größe: $($folder.FolderSize)
                        </div>
"@
                if ($folder.Permissions -and @($folder.Permissions).Count -gt 0) {
                    $html += @"
                        <table class="permissions-table">
                            <tr>
                                <th>Benutzer/Gruppe</th>
                                <th>Typ</th>
                                <th>Berechtigungen</th>
                            </tr>
"@
                    foreach ($perm in $folder.Permissions) {
                        $html += @"
                            <tr>
                                <td>$([System.Web.HttpUtility]::HtmlEncode($perm.User))</td>
                                <td>$([System.Web.HttpUtility]::HtmlEncode($perm.UserType))</td>
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
        
        $html += "</div></div>"
    }

    $html += @"
        </div>
        <div class="timestamp">Report erstellt: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')</div>
    </div>
</body>
</html>
"@

    return $html
}

# Hauptskript
try {
    if (-not (Initialize-ExchangeEnvironment)) {
        throw "Exchange-Umgebung konnte nicht initialisiert werden."
    }
    
    $counter = 0
    
    Write-Host "Erfasse öffentliche Ordner..."
    $allFolders = Get-AllPublicFolders
    $totalFolders = $allFolders.Count
    
    if ($totalFolders -eq 0) {
        throw "Keine öffentlichen Ordner gefunden!"
    }
    
    Write-Host "Starte Analyse von $totalFolders öffentlichen Ordnern..."
    
    $folderPermissions = @()
    
    foreach ($folder in $allFolders) {
        $counter++
        $folderInfo = Get-FolderPermissions -Folder $folder -CurrentCount $counter -TotalCount $totalFolders
        if ($null -ne $folderInfo) {
            $folderPermissions += $folderInfo
        }
    }
    
    Write-Host "Erstelle HTML-Report..."
    $htmlReport = Generate-HTMLReport -FolderData $folderPermissions
    
    # Schreibe HTML mit BOM
    $htmlBytes = $BOM.GetBytes($htmlReport)
    [System.IO.File]::WriteAllBytes($reportPath, $htmlBytes)
    
    Write-Host "Report wurde erfolgreich erstellt: $reportPath" -ForegroundColor Green
}
catch {
    Write-Error "Ein Fehler ist aufgetreten: $_"
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}
