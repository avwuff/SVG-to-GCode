;NSIS Modern User Interface version 1.63
;Start Menu Folder Selection Example Script
;Written by Joost Verburg

!define MUI_PRODUCT "SVGtoGCODE"
!define MUI_VERSION "1.2.6"
SetCompressor /solid lzma

!include "MUI.nsh"



; Macro - Upgrade DLL File
 ; Written by Joost Verburg
 ; ------------------------
 ;
 ; Example of usage:
 ; !insertmacro UpgradeDLL "dllname.dll" "$SYSDIR\dllname.dll"
 ;
 ; !define UPGRADEDLL_NOREGISTER if you want to upgrade a DLL which cannot be registered
 ;
 ; Note that this macro sets overwrite to ON (the default) when it has been inserted.
 ; If you are using another setting, set it again after inserting the macro.


 !macro UpgradeDLL LOCALFILE DESTFILE

   Push $R0
   Push $R1
   Push $R2
   Push $R3

   ;------------------------
   ;Check file and version

   IfFileExists "${DESTFILE}" "" "copy_${LOCALFILE}"

   ClearErrors
     GetDLLVersionLocal "${LOCALFILE}" $R0 $R1
     GetDLLVersion "${DESTFILE}" $R2 $R3
   IfErrors "upgrade_${LOCALFILE}"

   IntCmpU $R0 $R2 "" "done_${LOCALFILE}" "upgrade_${LOCALFILE}"
   IntCmpU $R1 $R3 "done_${LOCALFILE}" "done_${LOCALFILE}" "upgrade_${LOCALFILE}"

   ;------------------------
   ;Let's upgrade the DLL!

   SetOverwrite try

   "upgrade_${LOCALFILE}:"
     !ifndef UPGRADEDLL_NOREGISTER
       ;Unregister the DLL
       UnRegDLL "${DESTFILE}"
     !endif

   ;------------------------
   ;Try to copy the DLL directly

   ClearErrors
     StrCpy $R0 "${DESTFILE}"
     Call ":file_${LOCALFILE}"
   IfErrors "" "noreboot_${LOCALFILE}"

   ;------------------------
   ;DLL is in use. Copy it to a temp file and Rename it on reboot.

   GetTempFileName $R0
     Call ":file_${LOCALFILE}"
   Rename /REBOOTOK $R0 "${DESTFILE}"

   ;------------------------
   ;Register the DLL on reboot

   !ifndef UPGRADEDLL_NOREGISTER
;     WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\RunOnce" \
;     "Register ${DESTFILE}" '"$SYSDIR\rundll32.exe" "${DESTFILE},DllRegisterServer"'
     WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\RunOnce" \
     "Register ${DESTFILE}" '"$SYSDIR\regsvr32.exe" /s "${DESTFILE}"'
   !endif

   Goto "done_${LOCALFILE}"

   ;------------------------
   ;DLL does not exist - just extract

   "copy_${LOCALFILE}:"
     StrCpy $R0 "${DESTFILE}"
     Call ":file_${LOCALFILE}"

   ;------------------------
   ;Register the DLL

   "noreboot_${LOCALFILE}:"
     !ifndef UPGRADEDLL_NOREGISTER
       RegDLL "${DESTFILE}"
     !endif

   ;------------------------
   ;Done

   "done_${LOCALFILE}:"

   Pop $R3
   Pop $R2
   Pop $R1
   Pop $R0

   ;------------------------
   ;End

   Goto "end_${LOCALFILE}"

   ;------------------------
   ;Called to extract the DLL

   "file_${LOCALFILE}:"
     File /oname=$R0 "${LOCALFILE}"
     Return

   "end_${LOCALFILE}:"

  ;------------------------
  ;Set overwrite to default
  ;(was set to TRY above)

  SetOverwrite on

 !macroend














;--------------------------------
;Configuration

  ;General
  OutFile "..\releases\Windows\${MUI_PRODUCT}SetupV${MUI_VERSION}.exe"
  
  Name "${MUI_PRODUCT}"
  
  ;Folder selection page
  InstallDir "$PROGRAMFILES\${MUI_PRODUCT}"
  
  ;Remember install folder
  InstallDirRegKey HKCU "Software\${MUI_PRODUCT}" ""
  
  ;$9 is being used to store the Start Menu Folder.
  ;Do not use this variable in your script (or Push/Pop it)!

  ;To change this variable, use MUI_STARTMENUPAGE_VARIABLE.
  ;Have a look at the Readme for info about other options (default folder,
  ;registry).

  ;Remember the Start Menu Folder
  !define MUI_STARTMENUPAGE_REGISTRY_ROOT "HKCU" 
  !define MUI_STARTMENUPAGE_REGISTRY_KEY "Software\${MUI_PRODUCT}" 
  !define MUI_STARTMENUPAGE_REGISTRY_VALUENAME "Start Menu Folder Location"
  
  !define MUI_STARTMENUPAGE_DEFAULTFOLDER "${MUI_PRODUCT}"

  !define TEMP $R0

  !define VBFILESDIR "_Library"
  !define SHELLFOLDERS "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"

  Var MUI_TEMP
  Var STARTMENU_FOLDER 
  
  
;--------------------------------
;Modern UI Configuration

  ;!define MUI_ICON "C:\Program Files\NSIS\Contrib\Graphics\Icons\modern-install.ico" 
  ;!define MUI_UNICON "C:\Program Files\NSIS\Contrib\Graphics\Icons\modern-uninstall-colorful.ico"
  
  
  !define MUI_LICENSEPAGE_TEXT_TOP "Recent changes in ${MUI_PRODUCT}"
  !define MUI_LICENSEPAGE_BUTTON "Continue"
  !define MUI_LICENSEPAGE_TEXT_BOTTOM "All updates to ${MUI_PRODUCT} are listed above, seperated by date and version."
  
  !insertmacro MUI_PAGE_LICENSE "..\changelog.txt"
  !insertmacro MUI_PAGE_COMPONENTS
  !insertmacro MUI_PAGE_DIRECTORY

  
  !insertmacro MUI_PAGE_STARTMENU Application $STARTMENU_FOLDER

  !insertmacro MUI_PAGE_INSTFILES
  
  !define MUI_FINISHPAGE_RUN "$INSTDIR\${MUI_PRODUCT}.exe"
  ; !define MUI_FINISHPAGE_RUN_PARAMETERS "-newinstall"
  !insertmacro MUI_PAGE_FINISH


  !insertmacro MUI_UNPAGE_CONFIRM
  !insertmacro MUI_UNPAGE_INSTFILES

  
  
;--------------------------------
;Languages
 
  !insertmacro MUI_LANGUAGE "English"
  
;--------------------------------
;Language Strings

  ;Description

;--------------------------------
;Data
  

;--------------------------------
;Installer Sections




Section "${MUI_PRODUCT}" SecCopyUI

	SetShellVarContext "all"

  ; PUT UNINSTALL STUFF
  
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${MUI_PRODUCT}" "DisplayName" "${MUI_PRODUCT}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${MUI_PRODUCT}" "DisplayVersion" "${MUI_VERSION}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${MUI_PRODUCT}" "Publisher" "AvBrand"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${MUI_PRODUCT}" "UninstallString" "$INSTDIR\Uninstall.exe"


  ;ADD YOUR OWN STUFF HERE!

  SetOutPath "$INSTDIR"
  File "${MUI_PRODUCT}.exe"
  File "..\changelog.txt"
  
  ;Store install folder
  WriteRegStr HKCU "Software\${MUI_PRODUCT}" "" $INSTDIR
  
  !insertmacro MUI_STARTMENU_WRITE_BEGIN Application
    
    CreateDirectory "$SMPROGRAMS\$STARTMENU_FOLDER"

    CreateShortCut "$SMPROGRAMS\$STARTMENU_FOLDER\${MUI_PRODUCT}.lnk" "$INSTDIR\${MUI_PRODUCT}.exe"


    CreateShortCut "$SMPROGRAMS\$STARTMENU_FOLDER\Uninstall.lnk" "$INSTDIR\Uninstall.exe"

  !insertmacro MUI_STARTMENU_WRITE_END
  
  ;Create uninstaller
  WriteUninstaller "$INSTDIR\Uninstall.exe"

SectionEnd

Section "Runtime Files" SecRuntimes


  !insertmacro UpgradeDLL "${VBFILESDIR}\ssubtmr6.dll" 		$SYSDIR\ssubtmr6.dll
  !insertmacro UpgradeDLL "${VBFILESDIR}\vbalTbar6.ocx"   $SYSDIR\vbalTbar6.ocx
  !insertmacro UpgradeDLL "${VBFILESDIR}\ChilkatXml.dll"  $SYSDIR\ChilkatXml.dll
  !insertmacro UpgradeDLL "${VBFILESDIR}\vbalIml6.ocx"	 $SYSDIR\vbalIml6.ocx
  !insertmacro UpgradeDLL "${VBFILESDIR}\comdlg32.ocx"	 $SYSDIR\comdlg32.ocx

SectionEnd

Section "Desktop Shortcut" SecDesktop

	CreateShortCut "$DESKTOP\${MUI_PRODUCT}.lnk" "$INSTDIR\${MUI_PRODUCT}.exe"

SectionEnd



;Display the Finish header
;Insert this macro after the sections if you are not using a finish page
;!insertmacro MUI_SECTIONS_FINISHHEADER

;--------------------------------
;Descriptions

!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
  !insertmacro MUI_DESCRIPTION_TEXT ${SecCopyUI} 		"Install the ${MUI_PRODUCT} Component"
  !insertmacro MUI_DESCRIPTION_TEXT ${SecRuntimes} 	"Install the required runtime files for ${MUI_PRODUCT}"
  !insertmacro MUI_DESCRIPTION_TEXT ${SecDesktop} 		"Create a Desktop shortcut for ${MUI_PRODUCT}"
  
!insertmacro MUI_FUNCTION_DESCRIPTION_END
 
;--------------------------------
;Uninstaller Section

Section "Uninstall"

  ;ADD YOUR OWN STUFF HERE!

  Delete "$INSTDIR\${MUI_PRODUCT}.exe"
  Delete "$INSTDIR\Uninstall.exe"
  Delete "$INSTDIR\changelog.txt"
  
  
  ;Remove shortcut
  ;ReadRegStr ${TEMP} "${MUI_STARTMENUPAGE_REGISTRY_ROOT}" "${MUI_STARTMENUPAGE_REGISTRY_KEY}" "${MUI_STARTMENUPAGE_REGISTRY_VALUENAME}"
  
  
  !insertmacro MUI_STARTMENU_GETFOLDER Application $MUI_TEMP
  
  
  Delete "$SMPROGRAMS\$MUI_TEMP\Uninstall.lnk"
  Delete "$SMPROGRAMS\$MUI_TEMP\${MUI_PRODUCT}.lnk"
  RMDir "$SMPROGRAMS\$MUI_TEMP"
  
  
  ;Delete empty start menu parent diretories
  StrCpy $MUI_TEMP "$SMPROGRAMS\$MUI_TEMP"
 
  startMenuDeleteLoop:
    RMDir $MUI_TEMP
    GetFullPathName $MUI_TEMP "$MUI_TEMP\.."
    
    IfErrors startMenuDeleteLoopDone
  
    StrCmp $MUI_TEMP $SMPROGRAMS startMenuDeleteLoopDone startMenuDeleteLoop
  startMenuDeleteLoopDone:
  
  ; delete desktop shortcut
  Delete "$DESKTOP\${MUI_PRODUCT}.lnk"
  
  RMDir "$INSTDIR\myskin"
  RMDir "$INSTDIR"

  DeleteRegKey HKCR "Folder\shell\liveexpcopy"
  DeleteRegKey HKCR "Folder\shell\liveexpsync"
  
  DeleteRegKey HKCR ".LEI"
  DeleteRegKey HKCR "LiveExInstruct"
  DeleteRegKey HKCR "MIME\Database\Content Type\application/x-liveexplorerextension"


  DeleteRegKey /ifempty HKCU "Software\${MUI_PRODUCT}"

  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${MUI_PRODUCT}"

	

  ;Display the Finish header
  ;!insertmacro MUI_UNFINISHHEADER

SectionEnd