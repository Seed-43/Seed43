@echo off
title Seed43 - Git Push
color 0A

echo.
echo  ========================================
echo   Seed43 - Auto Push to GitHub
echo  ========================================
echo.

:: Move to the repo folder (edit this path if needed)
cd /d "C:\Users\Fred\Proton Drive\juvenciodas\My files\pySeeds\Seed43"

:: Check if there's anything to commit
git status --porcelain > temp_status.txt
set /p STATUS=<temp_status.txt
del temp_status.txt

if "%STATUS%"=="" (
    echo  No changes to push - everything is up to date.
    echo.
    pause
    exit /b
)

:: Show what's changed
echo  Changes detected:
git status --short
echo.

:: Ask for a commit message
set /p MSG= Enter commit message (or press Enter for default): 

if "%MSG%"=="" set MSG=Update %DATE% %TIME%

:: Stage all, commit, push
echo.
echo  Staging files...
git add .

echo  Committing: %MSG%
git commit -m "%MSG%"

echo  Pushing to GitHub...
git push origin main

echo.
if %ERRORLEVEL%==0 (
    echo  ========================================
    echo   Done! Changes pushed to GitHub.
    echo  ========================================
) else (
    echo  ========================================
    echo   Something went wrong - check above.
    echo  ========================================
)

echo.
pause
