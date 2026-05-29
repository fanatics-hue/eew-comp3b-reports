@echo off
chcp 65001 >nul
title EEW COMP3B — Setup GitHub (una tantum)
cd /d "%~dp0"

echo.
echo ==========================================================
echo   EEW COMP3B  ^|  Setup GitHub (eseguire UNA SOLA VOLTA)
echo ==========================================================
echo.

:: ── Verifica git ──────────────────────────────────────────────────────────────
git --version >nul 2>&1
if errorlevel 1 (
    echo ERRORE: Git non trovato.
    echo Installa Git da https://git-scm.com/download/win
    pause & exit /b 1
)

:: ── Chiedi URL repo GitHub ────────────────────────────────────────────────────
echo Crea un repo su https://github.com/new  (nome suggerito: comp3b-reports)
echo Poi incolla qui l'URL del repo (es. https://github.com/tuonome/comp3b-reports.git)
echo.
set /p REPO_URL="> URL repo: "

if "%REPO_URL%"=="" (
    echo ERRORE: URL non inserito.
    pause & exit /b 1
)

:: ── Config git identity (se non già impostata) ────────────────────────────────
git config user.email >nul 2>&1
if errorlevel 1 (
    set /p GIT_EMAIL="> Email GitHub: "
    set /p GIT_NAME="> Nome (es. Greg): "
    git config --global user.email "%GIT_EMAIL%"
    git config --global user.name "%GIT_NAME%"
)

:: ── Init repo locale ──────────────────────────────────────────────────────────
if not exist ".git" (
    git init -b main
    echo.
    echo Repo locale inizializzato.
) else (
    echo Repo locale già esistente.
)

:: ── .gitignore + index.html ───────────────────────────────────────────────────
git add .gitignore index.html auto_report.py AGGIORNA_REPORT.bat 2>nul

:: ── Aggiungi tutti gli HTML ───────────────────────────────────────────────────
git add IR-0*_redesign.html 2>nul

:: ── Primo commit ──────────────────────────────────────────────────────────────
git commit -m "Setup iniziale EEW COMP3B reports" 2>nul

:: ── Collega remote e push ─────────────────────────────────────────────────────
git remote remove origin 2>nul
git remote add origin %REPO_URL%

echo.
echo Push verso GitHub...
git push -u origin main

if errorlevel 1 (
    echo.
    echo ERRORE nel push. Possibili cause:
    echo  - URL repo errato
    echo  - Non autenticato su GitHub (esegui: git credential-manager)
    echo  - Repo non ancora creato su GitHub
    pause & exit /b 1
)

:: ── Ricava username e nome repo dall'URL ──────────────────────────────────────
:: Es. https://github.com/gregr/comp3b-reports.git -> gregr/comp3b-reports
set REPO_PATH=%REPO_URL:https://github.com/=%
set REPO_PATH=%REPO_PATH:.git=%

echo.
echo ==========================================================
echo   SETUP COMPLETATO!
echo ==========================================================
echo.
echo   GitHub Pages URL (attiva in ~1 min dopo averla abilitata):
echo   https://%REPO_PATH:.git=%.github.io/%REPO_PATH:*/=%
echo.
echo   Per abilitare GitHub Pages:
echo   1. Vai su %REPO_URL:.git=%/settings/pages
echo   2. Source: "Deploy from a branch"
echo   3. Branch: main  ^|  Folder: / (root)
echo   4. Clicca Save
echo.
pause
