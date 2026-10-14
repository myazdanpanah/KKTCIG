; KKTCOIG Invoice Software Installer
; Full-featured installer with offline prerequisite installation
; Include Modern UI 2 and LogicLib
!include "MUI2.nsh"
!include "LogicLib.nsh"
!include "FileFunc.nsh"

; General settings
Name "KKTCOIG Invoice Software"
OutFile "KKTCOIG_Installer.exe"
InstallDir "$PROGRAMFILES\KKTCOIG"
InstallDirRegKey HKLM "Software\KKTCOIG" "Install_Dir"
RequestExecutionLevel admin

; Variables
Var InstallType          ; "SERVER" or "CLIENT"
Var Dialog
Var ProgressBar
Var ProgressText
Var PrereqBatResult

; Pages - No license page
!insertmacro MUI_PAGE_WELCOME
Page custom PageTypeSelect PageTypeSelectLeave
!insertmacro MUI_PAGE_DIRECTORY
Page custom PageInstalling
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_LANGUAGE "English"

; -------------------- Installation type selection page --------------------
Function PageTypeSelect
    !insertmacro MUI_HEADER_TEXT "Installation Type" "Please select the installation type"
    nsDialogs::Create 1018
    Pop $Dialog
    ${If} $Dialog == error
        Abort
    ${EndIf}

    ; Server option
    ${NSD_CreateRadioButton} 30 30 250 20 "Server installation (includes SQL Server Express)"
    Pop $0
    ${NSD_AddStyle} $0 ${WS_GROUP}
    ${NSD_SetState} $0 ${BST_CHECKED}
    StrCpy $InstallType "SERVER"

    ; Client option
    ${NSD_CreateRadioButton} 30 60 250 20 "Client installation (ODBC driver only)"
    Pop $1

    ; Click event handlers
    ${NSD_OnClick} $0 TypeSelected
    ${NSD_OnClick} $1 TypeSelected

    nsDialogs::Show
FunctionEnd

Function TypeSelected
    Pop $R0
    ${NSD_GetState} $R0 $0
    ${If} $R0 == $0  ; Server radio selected
        StrCpy $InstallType "SERVER"
    ${Else}
        StrCpy $InstallType "CLIENT"
    ${EndIf}
FunctionEnd

Function PageTypeSelectLeave
    ; No validation needed
FunctionEnd

; -------------------- Installing page with progress bar --------------------
Function PageInstalling
    !insertmacro MUI_HEADER_TEXT "Installing" "Please wait while the software is being installed..."

    nsDialogs::Create 1018
    Pop $Dialog

    ; Progress bar
    ${NSD_CreateProgressBar} 0 50 100% 15 ""
    Pop $ProgressBar
    SendMessage $ProgressBar ${PBM_SETRANGE32} 0 100

    ; Status text
    ${NSD_CreateLabel} 0 80 100% 30 "Initializing..."
    Pop $ProgressText

    ; Disable Cancel button
    GetDlgItem $0 $HWNDPARENT 2
    EnableWindow $0 0

    nsDialogs::Show

    ; Start installation in a separate thread (using NSIS' built-in method)
    Call DoInstall
FunctionEnd

; -------------------- Main installation routine --------------------
Function DoInstall
    ; Step 1: Install prerequisites (Python, pip, packages) using offline batch
    SendMessage $ProgressText ${WM_SETTEXT} 0 "STR:Checking and installing prerequisites..."
    SendMessage $ProgressBar ${PBM_SETPOS} 10 0

    ; Execute the combined prerequisites batch file
    nsExec::ExecToStack '"$EXEDIR\install_prereqs.bat"'
    Pop $PrereqBatResult
    ${If} $PrereqBatResult != 0
        MessageBox MB_OK|MB_ICONSTOP "Failed to install prerequisites. Please check that all required files are present."
        Abort
    ${EndIf}

    SendMessage $ProgressBar ${PBM_SETPOS} 30 0
    SendMessage $ProgressText ${WM_SETTEXT} 0 "STR:Copying application files..."

    ; Step 2: Copy software files
    SetOutPath $INSTDIR
    File /r "software\*.*"

    SendMessage $ProgressBar ${PBM_SETPOS} 50 0
    SendMessage $ProgressText ${WM_SETTEXT} 0 "STR:Installing database components..."

    ; Step 3: Install SQL Server or ODBC based on selection
    ${If} $InstallType == "SERVER"
        Call InstallSQLServer
    ${Else}
        Call InstallODBC
    ${EndIf}

    SendMessage $ProgressBar ${PBM_SETPOS} 70 0
    SendMessage $ProgressText ${WM_SETTEXT} 0 "STR:Creating configuration file..."

    ; Step 4: Create db_config.json
    Call CreateConfigFile

    SendMessage $ProgressBar ${PBM_SETPOS} 90 0
    SendMessage $ProgressText ${WM_SETTEXT} 0 "STR:Creating desktop shortcut..."

    ; Step 5: Create shortcut
    Call CreateShortcut

    SendMessage $ProgressBar ${PBM_SETPOS} 100 0
    Sleep 500

    ; Enable Next button
    GetDlgItem $0 $HWNDPARENT 1
    EnableWindow $0 1
FunctionEnd

; -------------------- Helper functions --------------------
Function InstallSQLServer
    ; Check if SQL Server setup folder exists
    IfFileExists "$EXEDIR\SQLEXPR_x64_ENU\SETUP.EXE" 0 sql_not_found
    DetailPrint "Installing SQL Server Express..."
    ExecWait '"$EXEDIR\SQLEXPR_x64_ENU\SETUP.EXE" /QS /ACTION=Install /FEATURES=SQLEngine /INSTANCENAME=KKTCOIG /SECURITYMODE=SQL /SAPWD="KktC02o26!@#" /SQLSYSADMINACCOUNTS="BUILTIN\ADMINISTRATORS" /IACCEPTSQLSERVERLICENSETERMS'
    Return
sql_not_found:
    MessageBox MB_OK|MB_ICONSTOP "SQL Server setup folder (SQLEXPR_x64_ENU) not found. Please place it next to the installer."
    Abort
FunctionEnd

Function InstallODBC
    ; Check if ODBC installer exists
    IfFileExists "$EXEDIR\msodbcsql.msi" 0 odbc_not_found
    DetailPrint "Installing ODBC driver..."
    ExecWait 'msiexec /i "$EXEDIR\msodbcsql.msi" /qn IACCEPTMSODBCSQLLICENSETERMS=YES'
    Return
odbc_not_found:
    MessageBox MB_OK|MB_ICONSTOP "ODBC installer (msodbcsql.msi) not found. Please place it next to the installer."
    Abort
FunctionEnd

Function CreateConfigFile
    ; Create db_config.json with fixed settings
    FileOpen $0 "$INSTDIR\db_config.json" w
    FileWrite $0 '{$\r$\n'
    FileWrite $0 '  "server": ".\KKTCOIG",$\r$\n'
    FileWrite $0 '  "database": "InvoiceDB",$\r$\n'
    FileWrite $0 '  "username": "KKT",$\r$\n'
    FileWrite $0 '  "password": "KktC02o26!@#"$\r$\n'
    FileWrite $0 '}'
    FileClose $0
FunctionEnd

Function CreateShortcut
    CreateShortCut "$DESKTOP\KKTCOIG Invoice.lnk" "$INSTDIR\Invoice_app.exe"
FunctionEnd

; -------------------- Main installation section --------------------
Section "Main Application" SecMain
    ; Files are copied in DoInstall, but we need this section for the installer to work
    ; We'll just set the output path again and maybe add uninstall info
    SetOutPath $INSTDIR

    ; Write installation path to registry
    WriteRegStr HKLM "Software\KKTCOIG" "Install_Dir" "$INSTDIR"

    ; Write uninstaller
    WriteUninstaller "$INSTDIR\Uninstall.exe"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\KKTCOIG" "DisplayName" "KKTCOIG Invoice Software"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\KKTCOIG" "UninstallString" '"$INSTDIR\Uninstall.exe"'
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\KKTCOIG" "InstallLocation" "$INSTDIR"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\KKTCOIG" "Publisher" "KKTCOIG"
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\KKTCOIG" "NoModify" 1
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\KKTCOIG" "NoRepair" 1
SectionEnd

; -------------------- Uninstaller --------------------
Section "Uninstall"
    ; Remove files
    Delete "$INSTDIR\*.*"
    RMDir /r "$INSTDIR"

    ; Remove shortcut
    Delete "$DESKTOP\KKTCOIG Invoice.lnk"

    ; Remove registry keys
    DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\KKTCOIG"
    DeleteRegKey HKLM "Software\KKTCOIG"
SectionEnd