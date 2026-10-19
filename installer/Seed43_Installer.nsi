; Seed43 Installer
; NSIS Script

!include "MUI2.nsh"
!include "LogicLib.nsh"

; ── Installer Info ────────────────────────────────────────────────────────────
Name              "Seed43"
OutFile           "Seed43_Install.exe"
InstallDir        "$APPDATA\pyRevit\Extensions"
RequestExecutionLevel user
BrandingText      "Seed43 for pyRevit | seed43.org"

; ── Variables ─────────────────────────────────────────────────────────────────
Var pyRevitFound
Var ExtensionsDir

; ── MUI Settings ──────────────────────────────────────────────────────────────
!define MUI_ABORTWARNING
!define MUI_ICON                    "Seed43.ico"
!define MUI_UNICON                  "Seed43.ico"

!define MUI_WELCOMEPAGE_TITLE       "Welcome to Seed43"
!define MUI_WELCOMEPAGE_TEXT        "This will install the Seed43 toolbar extension for pyRevit.$\r$\n$\r$\npyRevit must already be installed before continuing.$\r$\n$\r$\nClick Next to continue."

!define MUI_FINISHPAGE_TITLE        "Seed43 Installed"
!define MUI_FINISHPAGE_TEXT         "Seed43 has been installed successfully.$\r$\n$\r$\nReload pyRevit in Revit to activate the toolbar."
!define MUI_FINISHPAGE_LINK         "Visit seed43.org"
!define MUI_FINISHPAGE_LINK_LOCATION "https://seed43.org"

; ── Pages ─────────────────────────────────────────────────────────────────────
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

!insertmacro MUI_LANGUAGE "English"

; ── Functions ─────────────────────────────────────────────────────────────────

Function FindPyRevit
    StrCpy $pyRevitFound "0"

    IfFileExists "$APPDATA\pyRevit\Extensions\*.*" 0 +3
        StrCpy $ExtensionsDir "$APPDATA\pyRevit\Extensions"
        StrCpy $pyRevitFound "1"

    ${If} $pyRevitFound == "0"
        IfFileExists "$LOCALAPPDATA\pyRevit\Extensions\*.*" 0 +3
            StrCpy $ExtensionsDir "$LOCALAPPDATA\pyRevit\Extensions"
            StrCpy $pyRevitFound "1"
    ${EndIf}

    ${If} $pyRevitFound == "0"
        MessageBox MB_YESNO|MB_ICONEXCLAMATION \
            "pyRevit Extensions folder not found.$\r$\n$\r$\npyRevit must be installed before installing Seed43.$\r$\n$\r$\nDownload pyRevit from https://github.com/pyrevitlabs/pyRevit/releases$\r$\n$\r$\nDo you want to continue anyway?" \
            IDYES continue IDNO abort
        abort:
            Abort
        continue:
            StrCpy $ExtensionsDir "$APPDATA\pyRevit\Extensions"
            CreateDirectory "$ExtensionsDir"
    ${EndIf}
FunctionEnd

; ── Installer Section ─────────────────────────────────────────────────────────

Section "Seed43" SecMain

    Call FindPyRevit

    SetOutPath "$ExtensionsDir"
    DetailPrint "Installing to: $ExtensionsDir"

    ${If} ${FileExists} "$ExtensionsDir\Seed43.extension\*.*"
        DetailPrint "Removing previous Seed43 installation..."
        RMDir /r "$ExtensionsDir\Seed43.extension"
    ${EndIf}

    DetailPrint "Extracting Seed43..."
    SetOverwrite on
    File /r "Seed43.extension"

    WriteUninstaller "$ExtensionsDir\Seed43.extension\Uninstall_Seed43.exe"

    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\Seed43" \
                     "DisplayName" "Seed43 for pyRevit"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\Seed43" \
                     "UninstallString" "$ExtensionsDir\Seed43.extension\Uninstall_Seed43.exe"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\Seed43" \
                     "DisplayVersion" "${VERSION}"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\Seed43" \
                     "Publisher" "Seed43"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\Seed43" \
                     "URLInfoAbout" "https://seed43.org"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\Seed43" \
                     "NoModify" "1"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\Seed43" \
                     "NoRepair" "1"

    DetailPrint "Seed43 installed successfully."
    DetailPrint "Reload pyRevit in Revit to activate the toolbar."

SectionEnd

; ── Uninstaller Section ───────────────────────────────────────────────────────

Section "Uninstall"

    RMDir /r "$APPDATA\pyRevit\Extensions\Seed43.extension"
    RMDir /r "$LOCALAPPDATA\pyRevit\Extensions\Seed43.extension"

    DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\Seed43"

    MessageBox MB_OK "Seed43 has been uninstalled.$\r$\nReload pyRevit in Revit to complete removal."

SectionEnd
