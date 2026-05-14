@echo off
REM ============================================================
REM   METFRAA Costing Engine — Push Helper
REM ============================================================
REM   Place this file in C:\metfraa-pkg (your repo root).
REM   Double-click to run, or run from CMD.
REM
REM   What it does:
REM     1. Shows what's changed
REM     2. Asks for a commit message
REM     3. Stages everything, commits, pushes
REM     4. Tells you to wait for Render auto-deploy
REM ============================================================

setlocal enabledelayedexpansion
cd /d "%~dp0"

echo.
echo ============================================================
echo   METFRAA Costing Engine — Push to GitHub / Render
echo ============================================================
echo.

REM --- Verify this is a git repo ---
if not exist ".git" (
    echo [ERROR] This folder is not a git repo.
    echo         Make sure Push.bat is in C:\metfraa-pkg ^(or wherever your repo lives^).
    echo.
    pause
    exit /b 1
)

REM --- Show current branch ---
for /f "delims=" %%b in ('git rev-parse --abbrev-ref HEAD 2^>nul') do set BRANCH=%%b
if "!BRANCH!"=="" (
    echo [ERROR] Could not determine current branch. Is git installed and the repo healthy?
    echo.
    pause
    exit /b 1
)
echo Current branch: !BRANCH!
echo.

REM --- Show what's changed ---
echo ------------------------------------------------------------
echo Files changed since last commit:
echo ------------------------------------------------------------
git status --short
echo.

REM --- Check if there's anything to commit ---
for /f %%i in ('git status --porcelain ^| find /c /v ""') do set COUNT=%%i
if "!COUNT!"=="0" (
    echo Nothing to commit. Working tree is clean.
    echo.
    pause
    exit /b 0
)

REM --- Ask for commit message ---
echo ------------------------------------------------------------
set /p MSG="Commit message (or press Enter for default): "
if "!MSG!"=="" set MSG=Update from Push.bat

echo.
echo ------------------------------------------------------------
echo Will commit with message:  !MSG!
echo Will push to:              origin/!BRANCH!
echo ------------------------------------------------------------
set /p CONFIRM="Proceed? (Y/N): "
if /i not "!CONFIRM!"=="Y" (
    echo Cancelled.
    echo.
    pause
    exit /b 0
)

REM --- Stage everything ---
echo.
echo [1/3] Staging changes...
git add .
if errorlevel 1 (
    echo [ERROR] git add failed.
    pause
    exit /b 1
)

REM --- Commit ---
echo [2/3] Committing...
git commit -m "!MSG!"
if errorlevel 1 (
    echo [ERROR] git commit failed.
    pause
    exit /b 1
)

REM --- Push ---
echo [3/3] Pushing to origin/!BRANCH!...
git push origin !BRANCH!
if errorlevel 1 (
    echo.
    echo [ERROR] git push failed.
    echo         If this is your first push, run:  git push -u origin !BRANCH!
    echo         If GitHub asked for credentials, sign in via the browser popup.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   SUCCESS
echo ============================================================
echo.
echo Next steps:
echo   1. Render auto-deploys in ~1-2 minutes
echo   2. Watch progress at https://dashboard.render.com
echo   3. Hard-refresh your app once "Live" shows green
echo.
pause
exit /b 0
