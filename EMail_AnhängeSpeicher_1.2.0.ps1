# 11:43 20.02.2025 
# Aktualisiertes Skript: Überprüfung, ob USB verfügbar ist
# Speichert E-Mail-Anhänge von beliebige Email
# Überprüft, ob das USB-Laufwerk verfügbar ist Falls nicht, wird das Skript gestoppt.
# Überprüft vor dem Speichern, ob die Datei bereits existiert. 
# Speichert Anhänge im „beliebige Verzeichnisse”

Add-Type -AssemblyName System.Windows.Forms

# USB-Laufwerk prüfen
$UsbPath = „beliebige Verzeichnisse”

if (!(Test-Path $UsbPath)) {
    Write-Output "USB-Laufwerk nicht verfügbar. Skript wird beendet."
    exit
}

Write-Output "USB-Laufwerk erkannt. Starte das Skript..."

# Outlook-Anwendung starten
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")
$Inbox = $Namespace.GetDefaultFolder(6)  # 6 = Posteingang

# Ordner ggf. neu anlegen
if (!(Test-Path $UsbPath)) {
    New-Item -ItemType Directory -Path $UsbPath | Out-Null
}

# Filterparameter
$SenderFilter = „beliebige E-Mail”
$NeuesGespeichert = 0

# E-Mails durchsuchen
foreach ($Mail in $Inbox.Items) {
    if ($Mail.SenderEmailAddress -match $SenderFilter) {
        Write-Output "Bearbeite E-Mail von Susanne Vogt: $($Mail.Subject)"

        if ($Mail.Attachments.Count -gt 0) {
            foreach ($Attachment in $Mail.Attachments) {
                $FilePath = Join-Path -Path $UsbPath -ChildPath $Attachment.FileName

                if (Test-Path $FilePath) {
                    Write-Output "Datei existiert bereits und wird nicht erneut gespeichert: $FilePath"
                } else {
                    $Attachment.SaveAsFile($FilePath)
                    Write-Output "Anhang gespeichert: $FilePath"
                    $NeuesGespeichert++
                }
            }
        }
    }
}

# Benutzerhinweis anzeigen
if ($NeuesGespeichert -gt 0) {
    [System.Windows.Forms.MessageBox]::Show("Hat sehr gut funktioniert.", "Erfolg", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
} else {
    $heute = Get-Date -Format "dd.MM.yyyy"
    [System.Windows.Forms.MessageBox]::Show("Die Anhänge wurden schon am $heute gespeichert.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}
