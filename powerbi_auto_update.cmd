@echo off
setlocal

echo ----------------------------------------
echo Starte Power BI Bericht mit SharePoint-Daten
echo ----------------------------------------

:: Pfad zur PBIX-Datei (lokal gespeichert)
set "PBIX_PATH=C:\Users\%USERNAME%\PowerBI\Reports\SalesDashboard.pbix"

:: Power BI Desktop starten und Bericht öffnen
start "" "C:\Program Files\Microsoft Power BI Desktop\bin\PBIDesktop.exe" "%PBIX_PATH%"

:: Warten, bis Daten geladen sind (z. B. 2 Minuten)
timeout /t 120 >nul

:: OPTIONAL: Power BI nach Ladezeit schließen
:: Achtung: Macht nur Sinn, wenn Power BI so eingestellt ist, dass es beim Öffnen Daten lädt!
:: taskkill /f /im PBIDesktop.exe

echo Bericht wurde geöffnet und aktualisiert (manuell oder automatisch je nach Modell)
pause
