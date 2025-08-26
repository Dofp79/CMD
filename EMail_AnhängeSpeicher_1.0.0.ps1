# 11:43 20.02.2025 
# Aktualisiertes Skript: Überprüfung, ob USB verfügbar ist
# Speichert E-Mail-Anhänge von susanne.vogt@gsthaden.de
# Überprüft, ob das USB-Laufwerk verfügbar ist → Falls nicht, wird das Skript gestoppt.
# Überprüft vor dem Speichern, ob die Datei bereits existiert.
# Speichert Anhänge im Ordner:

# USB-Laufwerk prüfen
$UsbPath = „Beliebige Speicher”

if (!(Test-Path $UsbPath)) {
    Write-Output "USB-Laufwerk nicht verfügbar. Skript wird beendet."
    exit
}

Write-Output "USB-Laufwerk erkannt. Starte das Skript..."

# Outlook-Anwendung starten
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")

# Posteingang abrufen
$Inbox = $Namespace.GetDefaultFolder(6)  # 6 = Posteingang

# Falls der Ordner nicht existiert, erstelle ihn (falls z. B. der USB-Stick neu formatiert wurde)
if (!(Test-Path $UsbPath)) {
    New-Item -ItemType Directory -Path $UsbPath | Out-Null
}

# Relevante E-Mail-Adresse (nur E-Mails von Susanne Vogt verarbeiten)
$SenderFilter = „beliebige E-Mail”

# E-Mails durchsuchen
foreach ($Mail in $Inbox.Items) {
    # Prüfen, ob die E-Mail von Susanne Vogt ist
    if ($Mail.SenderEmailAddress -match $SenderFilter) {
        Write-Output "Bearbeite E-Mail von Susanne Vogt: $($Mail.Subject)"

        # Falls die E-Mail Anhänge hat
        if ($Mail.Attachments.Count -gt 0) {
            foreach ($Attachment in $Mail.Attachments) {
                # Speicherpfad für den Anhang
                $FilePath = Join-Path -Path $UsbPath -ChildPath $Attachment.FileName

                # Prüfen, ob die Datei bereits existiert
                if (Test-Path $FilePath) {
                    Write-Output "Datei existiert bereits und wird nicht erneut gespeichert: $FilePath"
                } else {
                    # Datei speichern, falls sie noch nicht existiert
                    $Attachment.SaveAsFile($FilePath)
                    Write-Output "Anhang gespeichert: $FilePath"
                }
            }
        }
    }
}
