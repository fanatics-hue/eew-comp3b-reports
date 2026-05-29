@echo off
chcp 65001 >nul
title EEW COMP3B — Aggiornamento Report Settimanale
cd /d "%~dp0"

echo.
echo ==========================================================
echo   EEW COMP3B  ^|  Aggiornamento Report Settimanale
echo ==========================================================
echo.

:: ── Verifica dipendenze ───────────────────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo ERRORE: Python non trovato.
    echo Installa Python da https://www.python.org/downloads/
    pause & exit /b 1
)
python -c "import docx" >nul 2>&1
if errorlevel 1 (
    echo Installazione python-docx...
    pip install python-docx --quiet
)

:: ── Trova il docx più recente ─────────────────────────────────────────────────
set DOCX=
for /f "delims=" %%f in ('dir /b /o-d "IR-0*.docx" 2^>nul') do (
    if not defined DOCX set DOCX=%%f
)
if not defined DOCX (
    echo Nessun file .docx trovato. Inserisci il nome del file:
    set /p DOCX="> "
)
if not exist "%DOCX%" (
    echo ERRORE: File non trovato — %DOCX%
    pause & exit /b 1
)

echo   Docx:  %DOCX%
echo.

:: ── Genera HTML ───────────────────────────────────────────────────────────────
echo [1/3] Generazione dashboard...
set PYTHONUTF8=1
python auto_report.py "%DOCX%"
if errorlevel 1 (
    echo ERRORE: Generazione fallita.
    pause & exit /b 1
)

:: ── Trova l'HTML appena generato ─────────────────────────────────────────────
set HTML=
for /f "delims=" %%f in ('dir /b /o-d "IR-0*_redesign.html" 2^>nul') do (
    if not defined HTML set HTML=%%f
)

:: ── Aggiorna index.html con redirect al nuovo report ─────────────────────────
if defined HTML (
    :: Estrai IR number e date dal nome file per il link in index.html
    set IR_FILE=%HTML%
    powershell -NoProfile -Command ^
        "$html='%HTML%'; $stem=$html -replace '_redesign\.html$',''; " ^
        "$ir=if($stem -match 'IR-(\d+)'){$Matches[1]}else{'???'}; " ^
        "$range=if($stem -match '(\d{2}_\d{2}_\d{4})_-_(\d{2}_\d{2}_\d{4})'){($Matches[1]+' to '+$Matches[2]) -replace '_','.' }else{''}; " ^
        "$content=[System.IO.File]::ReadAllText('index.html'); " ^
        "$content=$content -replace 'url=IR-\d+[^\"'']+','url='+$html; " ^
        "$content=$content -replace 'href=\"IR-\d+[^\"]+\"','href=\"'+$html+'\"'; " ^
        "$content=$content -replace 'IR-\d+ — [^<]+','IR-'+$ir+' — '+$range; " ^
        "[System.IO.File]::WriteAllText('index.html',$content,[System.Text.Encoding]::UTF8)"
)

:: ── Git commit e push ────────────────────────────────────────────────────────
git --version >nul 2>&1
if errorlevel 1 goto :skip_git
if not exist ".git" goto :skip_git

echo.
echo [2/3] Commit e push su GitHub...

git add "%HTML%" index.html 2>nul
git commit -m "Weekly report %HTML:.html=% — %DATE%" 2>nul

git push origin main
if errorlevel 1 (
    echo ATTENZIONE: Push fallito. Controlla la connessione o le credenziali GitHub.
    echo Il file HTML e' stato generato correttamente in locale.
    goto :open_local
)

:: ── Ricava GitHub Pages URL ───────────────────────────────────────────────────
set PAGES_URL=
for /f "delims=" %%u in ('git remote get-url origin 2^>nul') do set ORIGIN=%%u
:: Converti https://github.com/user/repo.git -> https://user.github.io/repo/file.html
set ORIGIN=%ORIGIN:https://github.com/=%
set ORIGIN=%ORIGIN:.git=%
for /f "tokens=1,2 delims=/" %%a in ("%ORIGIN%") do (
    set GH_USER=%%a
    set GH_REPO=%%b
)
if defined GH_USER if defined GH_REPO (
    set PAGES_URL=https://%GH_USER%.github.io/%GH_REPO%/%HTML%
)

echo.
echo [3/3] Pubblicato!
echo.
if defined PAGES_URL (
    echo   Dashboard online:
    echo   %PAGES_URL%
    echo.
    echo   Apertura nel browser...
    start "" "%PAGES_URL%"
) else (
    goto :open_local
)
goto :done

:open_local
if defined HTML (
    echo   Apertura in locale: %HTML%
    start "" "%HTML%"
)

:done
echo.
echo ==========================================================
echo   Completato! Premi un tasto per chiudere.
echo ==========================================================
pause >nul

:skip_git
echo.
echo [SKIP] Git non configurato — apro il file in locale.
if defined HTML start "" "%HTML%"
pause >nul
