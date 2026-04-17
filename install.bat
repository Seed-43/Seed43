@echo off
:: Seed43 installer — downloads Seed43.tab from GitHub into Seed43.extension

set GITHUB_ORG=Seed-43
set MAIN_REPO=Seed43
set BRANCH=main
set ZIP_URL=https://github.com/%GITHUB_ORG%/%MAIN_REPO%/archive/refs/heads/%BRANCH%.zip

set INSTALL_DIR=%APPDATA%\pyRevit\Extensions\Seed43.extension
set TEMP_ZIP=%TEMP%\seed43_install.zip
set TEMP_EXTRACT=%TEMP%\seed43_extracted

echo Downloading Seed43 from GitHub...
curl -L -o "%TEMP_ZIP%" "%ZIP_URL%"
if errorlevel 1 (
    echo Error: Download failed.
    pause
    exit /b 1
)

echo Extracting...
if exist "%TEMP_EXTRACT%" rmdir /s /q "%TEMP_EXTRACT%"
mkdir "%TEMP_EXTRACT%"
tar -xf "%TEMP_ZIP%" -C "%TEMP_EXTRACT%"
if errorlevel 1 (
    echo Error: Extraction failed.
    pause
    exit /b 1
)

:: GitHub extracts to reponame-branch folder
for /d %%i in ("%TEMP_EXTRACT%\*") do set EXTRACTED_ROOT=%%i

if not exist "%EXTRACTED_ROOT%\Seed43.tab" (
    echo Error: Seed43.tab not found in repo ZIP.
    pause
    exit /b 1
)

echo Installing to: %INSTALL_DIR%
if exist "%INSTALL_DIR%" rmdir /s /q "%INSTALL_DIR%"
mkdir "%INSTALL_DIR%"
xcopy /e /i /q "%EXTRACTED_ROOT%\Seed43.tab" "%INSTALL_DIR%\Seed43.tab"

echo Cleaning up...
del /q "%TEMP_ZIP%"
rmdir /s /q "%TEMP_EXTRACT%"

echo.
echo Done! Reload PyRevit in Revit to activate the Seed43 tab.
pause
