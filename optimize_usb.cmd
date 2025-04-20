@echo off 
::  optimize_usb.cmd  -  PC- und USB-Optimierung  (18 Apr 2025)
::  ----------------------------------------------------------
::  - prüft & repariert C: plus alle USB-Laufwerke (DriveType 2)
::  - reserviert maximale MFT-Zone, defragmentiert, bereinigt
::  - Skript in UTF‑8, OEM-Tools werden via :RunOEM in CP 850 ausgeführt

REM ---------- 1) UTF‑8-Codepage aktivieren + Farbe ----------------
chcp 65001 >nul
color 0A

REM ---------- 2) Kopfzeile ----------------------------------------
cls
echo ============================================================
echo   COMPUTER- UND USB-OPTIMIERUNG  –  ALLES WIRD GEORDNET
echo ============================================================

REM ---------- 3) MFT-Reserve auf Maximum setzen -------------------
echo.
echo [1] MFT-Reserve auf Maximum setzen ...
call :RunOEM fsutil behavior set mftzone 4

REM ---------- 4) CHKDSK auf Systemlaufwerk C: ---------------------
echo.
echo [2] Laufwerk C: auf Fehler prüfen ...
call :RunOEM chkdsk C: /F /R /X
if errorlevel 1 (
    echo     -> C: wird beim nächsten Neustart repariert.
) else (
    echo     -> C: ist fehlerfrei.
)

REM ---------- 5) CHKDSK auf allen USB-Laufwerken ------------------
echo.
echo [3] USB-/Wechseldatenträger erkennen und prüfen ...

set "USB_FOUND=0"

:: 5a) WMIC-Methode
for /f %%D in ('wmic logicaldisk where "DriveType=2" get deviceid ^| find ":"') do (
    set "USB_FOUND=1"
    call :RunChk %%D
)

:: 5b) PowerShell-Fallback, falls WMIC nicht verfügbar
if "%USB_FOUND%"=="0" (
    for /f %%D in ('powershell -NoLogo -NoProfile -Command ^
        "Get-CimInstance Win32_LogicalDisk | where {$_.DriveType -eq 2} | ForEach-Object {$_.DeviceID}"') do (
        set "USB_FOUND=1"
        call :RunChk %%D
    )
)

if "%USB_FOUND%"=="0" (
    echo     -> Keine USB-Laufwerke gefunden.
)

REM ---------- 6) TEMP-Ordner leeren -------------------------------
echo.
echo [4] Temporäre Dateien entfernen ...
del /s /q /f "%temp%\*" >nul 2>&1
echo     -> Temp-Ordner geleert.

REM ---------- 7) Defragmentierung & Konsolidierung von C: ---------
echo.
echo [5] Festplatte C: defragmentieren und freien Platz zusammenführen ...
call :RunOEM defrag C: /H /X /U /V

REM ---------- 8) Datenträgerbereinigung ---------------------------
echo.
echo [6] System bereinigen (Datenträgerbereinigung) ...
call :RunOEM cleanmgr /sageset:99
call :RunOEM cleanmgr /sagerun:99

REM ---------- 9) Systemdateien prüfen -----------------------------
echo.
echo [7] Systemdateien prüfen (SFC) ...
call :RunOEM sfc /scannow

REM ----------10) Abschluss ----------------------------------------
echo.
echo ============================================================
echo   Optimierung abgeschlossen – Rechner und USB sind top!
echo ============================================================
pause
exit /b


:: ==============================================================
::  Unterprogramm :RunChk - einzelnes Laufwerk prüfen
::  Aufruf:  call :RunChk E:
:: ==============================================================
:RunChk
echo     -> Laufwerk %1 auf Fehler prüfen ...
call :RunOEM chkdsk %1 /F /R /X
if errorlevel 1 (
    echo     -> %1 wird beim nächsten Neustart repariert.
) else (
    echo     -> %1 ist fehlerfrei.
)
exit /b


:: ==============================================================
::  Unterprogramm :RunOEM
::  Wechselt temporär auf Codepage 850, führt den Befehl aus
::  und stellt anschließend UTF‑8 (65001) wieder her.
::  Aufruf:  call :RunOEM <Befehl mit Parametern>
:: ==============================================================
:RunOEM
for /f "tokens=2 delims=:." %%C in ('chcp') do set "SAVE_CP=%%C"
chcp 850 >nul
%*
chcp %SAVE_CP% >nul
exit /b %errorlevel%
