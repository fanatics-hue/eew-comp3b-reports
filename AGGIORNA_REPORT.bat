@echo off
call "%~dp0..\_GUARD.bat" "EEW COMP3B Report"
if "%PROTEZIONE_OK%"=="0" ( pause & exit /b 0 )

setlocal enabledelayedexpansion
chcp 65001 >nul
title EEW COMP3B - Aggiornamento Report Settimanale
cd /d "%~dp0"

echo.
echo ===========================================================
echo   EEW COMP3B  ^|  Aggiornamento Report Settimanale
echo ===========================================================
echo.

:: Verifica Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERRORE: Python non trovato.
    echo Installa Python da https://www.python.org/downloads/
    pause & exit /b 1
)

:: Verifica python-docx
python -c "import docx" >nul 2>&1
if errorlevel 1 (
    echo Installazione python-docx...
    pip install python-docx --quiet
)

:: Trova il docx piu recente
set DOCX=
for /f "delims=" %%f in ('dir /b /o-d "*.docx" 2^>nul') do (
    if not defined DOCX set DOCX=%%f
)
if not defined DOCX (
    echo Nessun file .docx trovato. Inserisci il nome del file:
    set /p DOCX="> "
)
if not exist "!DOCX!" (
    echo ERRORE: File non trovato - !DOCX!
    pause & exit /b 1
)

echo   Docx: !DOCX!
echo.

:: Genera HTML
echo [1/3] Generazione dashboard...
set PYTHONUTF8=1
python auto_report.py "!DOCX!"
if errorlevel 1 (
    echo ERRORE: Generazione fallita.
    pause & exit /b 1
)

:: Trova HTML generato (il piu recente)
set HTML=
for /f "delims=" %%f in ('dir /b /o-d "IR-0*_redesign.html" 2^>nul') do (
    if not defined HTML set HTML=%%f
)

:: Aggiorna index.html
if defined HTML (
    (
        echo ^<!DOCTYPE html^>
        echo ^<html lang="en"^>
        echo ^<head^>
        echo ^<meta charset="UTF-8"^>
        echo ^<meta http-equiv="refresh" content="0; url=!HTML!"^>
        echo ^<title^>EEW COMP3B - Weekly Reports^</title^>
        echo ^<style^>body{font-family:-apple-system,'Segoe UI',sans-serif;background:#1c1c1a;color:#f0efe8;display:flex;align-items:center;justify-content:center;height:100vh;margin:0;flex-direction:column;gap:16px}a{color:#1D9E75}^</style^>
        echo ^</head^>
        echo ^<body^>
        echo ^<div style="font-size:13px;color:#888"^>Redirecting...^</div^>
        echo ^<a href="!HTML!"^>Latest report: !HTML!^</a^>
        echo ^</body^>
        echo ^</html^>
    ) > index.html
)

:: Git commit e push
git --version >nul 2>&1
if errorlevel 1 goto :skip_git
if not exist ".git" goto :skip_git

echo.
echo [2/3] Commit e push su GitHub...
git add "!HTML!" index.html >nul 2>&1
git commit -m "Weekly report !HTML:.html=! - %DATE%" >nul 2>&1
git push origin main
if errorlevel 1 (
    echo ATTENZIONE: Push fallito. Controlla connessione o credenziali GitHub.
    goto :open_local
)

:: Ricava GitHub Pages URL
set GH_USER=
set GH_REPO=
for /f "delims=" %%u in ('git remote get-url origin 2^>nul') do set ORIGIN=%%u
set ORIGIN=!ORIGIN:https://github.com/=!
set ORIGIN=!ORIGIN:.git=!
for /f "tokens=1,2 delims=/" %%a in ("!ORIGIN!") do (
    set GH_USER=%%a
    set GH_REPO=%%b
)

if defined GH_USER (
    set PAGES_URL=https://!GH_USER!.github.io/!GH_REPO!/!HTML!
    echo.
    echo [3/3] Pubblicato!
    echo.
    echo   Dashboard online:
    echo   !PAGES_URL!
    echo.
    start "" "!PAGES_URL!"
    goto :done
)

:open_local
if defined HTML (
    echo   Apertura in locale: !HTML!
    start "" "!HTML!"
)

:done
echo.
echo ===========================================================
echo   Completato!
echo ===========================================================
echo.
pause
exit /b 0

:skip_git
echo [SKIP] Git non configurato - apro il file in locale.
if defined HTML start "" "!HTML!"
pause