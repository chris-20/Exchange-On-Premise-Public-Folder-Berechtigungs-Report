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
# Requires Exchange Management Shell
# Hole das Verzeichnis, in dem das Skript liegt
[CmdletBinding()]
param(
    [string]$OutputPath = "",
    [switch]$IncludeDefaultPermissions
)

# Setze Ausgabepfad
if ([string]::IsNullOrEmpty($OutputPath)) {
    $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
    $timestamp = Get-Date -Format "dd.MM.yyyy HH-mm"
    $reportPath = Join-Path $scriptPath "Public_Folder_Permissions_Report_$timestamp.html"
} else {
    $reportPath = $OutputPath
}

# Setze Encoding für korrekte Umlaut-Darstellung
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'

$progressPreference = 'Continue'

# Prüfe und lade Exchange Management Shell
function Initialize-ExchangeEnvironment {
    try {
        # Prüfe, ob Exchange-Cmdlets verfügbar sind
        $exchangeSnapin = Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue
        
        if ($null -eq $exchangeSnapin) {
            # Versuche Exchange-Snapin zu laden
            Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
            Write-Host "Exchange Management Shell erfolgreich geladen." -ForegroundColor Green
        }
        
        # Teste ob Exchange-Cmdlets verfügbar sind
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
        
        # Hole die Berechtigungen
        $permissions = Get-PublicFolderClientPermission -Identity $Folder.EntryID -ErrorAction Stop | 
            Where-Object $permissionFilter |
            ForEach-Object {
                # Verbesserte Benutzertyp-Erkennung
                $userType = "Unbekannt"
                try {
                    # Versuche den Benutzer oder die Gruppe zu finden
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
    
    $css = @"
    body { font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; margin: 0; padding: 20px; background-color: #f5f5f5; }
    .container { max-width: 1200px; margin: 0 auto; background-color: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
    h1, h2 { color: #2c3e50; padding-bottom: 10px; }
    h1 { border-bottom: 2px solid #3498db; }
    .summary { background-color: #e8f4f8; padding: 15px; border-radius: 5px; margin: 20px 0; }
    .folder { margin: 20px 0; padding: 15px; background-color: #f8f9fa; border-radius: 5px; border-left: 4px solid #3498db; }
    .folder-header { font-size: 1.2em; font-weight: bold; color: #2c3e50; margin-bottom: 10px; }
    .folder-info { font-size: 0.9em; color: #666; margin-bottom: 10px; }
    .permissions-table { width: 100%; border-collapse: collapse; margin-top: 10px; }
    .permissions-table th { background-color: #3498db; color: white; padding: 8px; text-align: left; }
    .permissions-table td { padding: 8px; border-bottom: 1px solid #ddd; }
    .permissions-table tr:nth-child(even) { background-color: #f2f2f2; }
    .no-permissions { color: #666; font-style: italic; padding: 10px; }
    .timestamp { color: #666; font-size: 0.9em; margin-top: 20px; text-align: right; }
    .search-box { margin: 20px 0; padding: 10px; }
    .search-box input { width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; }
"@

    $html = @"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <title>Exchange Public Folder Berechtigungen</title>
    <style>$css</style>
    <script>
    function searchFolders() {
        const input = document.getElementById('searchInput').value.toLowerCase();
        const folders = document.getElementsByClassName('folder');
        
        for (let folder of folders) {
            const text = folder.textContent.toLowerCase();
            folder.style.display = text.includes(input) ? '' : 'none';
        }
    }
    </script>
</head>
<body>
    <div class="container">
        <h1>Exchange Public Folder Berechtigungsübersicht</h1>
        
        <div class="summary">
            <h2>Zusammenfassung</h2>
            <p>Gesamtanzahl der analysierten Ordner: $($FolderData.Count)</p>
            <p>Analysezeitpunkt: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')</p>
        </div>

        <div class="search-box">
            <input type="text" id="searchInput" onkeyup="searchFolders()" placeholder="Suche nach Ordnern oder Berechtigungen...">
        </div>
"@

    $sortedFolders = $FolderData | Sort-Object -Property FolderPath

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
    # Prüfe Exchange-Umgebung
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
    $htmlReport | Out-File -FilePath $reportPath -Encoding UTF8
    
    Write-Host "Report wurde erfolgreich erstellt: $reportPath" -ForegroundColor Green
}
catch {
    Write-Error "Ein Fehler ist aufgetreten: $_"
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}
