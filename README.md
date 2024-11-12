📊 **Exchange Public Folder Berechtigungs-Report**

Dieses PowerShell-Skript erstellt einen detaillierten Berechtigungsbericht für öffentliche Ordner auf einem Exchange-Server. Es erfasst die Zugriffsrechte der Benutzer auf die öffentlichen Ordner und stellt sie in einem übersichtlichen HTML-Bericht dar. Der Report hilft, Zugriffsberechtigungen zu überprüfen, Sicherheitsrichtlinien einzuhalten und eine klare Übersicht zu behalten.

✨ **Funktionen**

- 🔎 **Umfassende Berechtigungsanalyse**: Zeigt alle relevanten Berechtigungen für jeden öffentlichen Ordner, einschließlich der Berechtigungen für Benutzer und Gruppen.
- 📄 **Stilvolle HTML-Darstellung**: Der Bericht wird in einer modernen und benutzerfreundlichen HTML-Datei dargestellt, die die Analyse der Ordnersicherheit vereinfacht.
- 📅 **Automatisierter Zeitstempel**: Jeder Bericht enthält das Erstellungsdatum für eine einfache Nachverfolgung.
- 👥 **Öffentliche Ordner**: Der Bericht listet detaillierte Berechtigungen für alle öffentlichen Ordner im Exchange-Server auf.

📋 **Voraussetzungen**

- PowerShell-Zugriff auf den Exchange Server
- **Exchange Management Shell** zur Ausführung des Skripts

🚀 **Verwendung**

1. Lade das Skript herunter.
2. Öffne die **Exchange Management Shell**.
3. Navigiere in das Verzeichnis, in dem sich das Skript befindet.
4. Gib folgenden Befehl ein, um das Skript auszuführen:  
   `.\Exchange-Public-Folder-Berechtigungs-Report.ps1`
5. Der Bericht wird als HTML-Datei im gleichen Verzeichnis gespeichert und enthält einen Zeitstempel im Dateinamen.

📘 **Beispielausgabe**

Der HTML-Bericht zeigt:

- 🛠 **Berechtigungstyp** (z.B. Ordnerschutz, Lesezugriff, Vollzugriff)
- 🧾 **Spezifische Berechtigungen** für Benutzer und Gruppen
- 👤 **Zugewiesene Benutzer** und deren Zugriffsrechte auf die öffentlichen Ordner
