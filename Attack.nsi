# Attack.nsi installation script
# Automatically generated by VB6Converter by David Nuttall
# VB6Converter was compiled on 15/02/08
#

Name "Attack-Setup"
OutFile "Attack-Setup.exe"
RequestExecutionLevel admin
SetCompress auto
SetCompressor /SOLID /FINAL Lzma
WindowIcon on
AutoCloseWindow false
SetDatablockOptimize on
CRCCheck on
InstProgressFlags smooth
SetDateSave on
AllowSkipFiles on
SetOverwrite IfNewer
InstallDir "$ProgramFiles\Attack"

!include Library.nsh

#Page Directory
#Page Components
Page InstFiles

UninstPage UninstConfirm
UninstPage InstFiles

Var ALREADY_INSTALLED
Var Version
Var WinSockLicence

Section "Attack"
	SectionIn 1 RO
	;LogSet on	# Needs special build of NSIS to actually work.
	# Change the registry structure to reflect the directory structure you are using.
	ReadRegStr $Version HKLM "Software\Attack" Version
	IfErrors new_installation
	StrCpy $ALREADY_INSTALLED 1
	
	new_installation:
	# VB DLLs
	# http://activex.microsoft.com/controls/vb6/mswinsck.cab
	!insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED \
		"Support\mswinsck.ocx" "$SysDir\mswinsck.ocx" "$SysDir"
	!insertmacro InstallLib DLL NOTSHARED REBOOT_PROTECTED \
		"Support\RICHED32.DLL" "$SysDir\RICHED32.DLL" "$SysDir"
	# http://activex.microsoft.com/controls/vb6/RichTx32.cab
	!insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED \
		"Support\RICHTX32.OCX" "$SysDir\RICHTX32.OCX" "$SysDir"
	!insertmacro InstallLib DLL $ALREADY_INSTALLED REBOOT_PROTECTED \
		"Support\VB5DB.DLL" "$SysDir\VB5DB.DLL" "$SysDir"
	!insertmacro InstallLib DLL NOTSHARED REBOOT_PROTECTED \
		"Support\MSREPL35.DLL" "$SysDir\MSREPL35.DLL" "$SysDir"
	!insertmacro InstallLib REGDLL NOTSHARED REBOOT_PROTECTED \
		"Support\MSRD2X35.DLL" "$SysDir\MSRD2X35.DLL" "$SysDir"
	!insertmacro InstallLib DLL NOTSHARED REBOOT_PROTECTED \
		"Support\expsrv.dll" "$SysDir\expsrv.dll" "$SysDir"
	!insertmacro InstallLib DLL NOTSHARED REBOOT_PROTECTED \
		"Support\vbajet32.dll" "$SysDir\vbajet32.dll" "$SysDir"
	!insertmacro InstallLib DLL NOTSHARED REBOOT_PROTECTED \
		"Support\MSJINT35.DLL" "$SysDir\MSJINT35.DLL" "$SysDir"
	!insertmacro InstallLib DLL NOTSHARED REBOOT_PROTECTED \
		"Support\MSJTER35.DLL" "$SysDir\MSJTER35.DLL" "$SysDir"
	!insertmacro InstallLib REGDLL NOTSHARED REBOOT_PROTECTED \
		"Support\MSJET35.DLL" "$SysDir\MSJET35.DLL" "$SysDir"
	!insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED \
		"Support\DAO350.DLL" "$CommonFiles\Microsoft Shared\DAO\DAO350.DLL" "$CommonFiles\Microsoft Shared\DAO"
	
	# Register Winsock
	StrCmp $ALREADY_INSTALLED 1 AfterRegWinsock # skip if Attack already installed
	WriteRegStr HKLM "Software\Attack" ShouldDeleteWinsock "0"
	ClearErrors
	ReadRegStr $WinSockLicence HKCR "Licenses\2c49f800-c2dd-11cf-9ad6-0080c7e7b78d" ""
	IfErrors 0 AfterRegWinsock
	WriteRegStr HKCR "Licenses\2c49f800-c2dd-11cf-9ad6-0080c7e7b78d" "" "mlrljgrlhltlngjlthrligklpkrhllglqlrk"
	WriteRegStr HKLM "Software\Attack" ShouldDeleteWinsock "1"
	
	AfterRegWinsock:
	
	# Images
	SetOutPath "$InstDir\images"
	
	File /a "images\greenMonster.bmp"
	File /a "images\blackMonster.bmp"
	File /a "images\bat.bmp"
	File /a "images\tree.bmp"
	File /a "images\cloud.bmp"
	File /a "images\rabbit.bmp"
	File /a "images\ladyBug.bmp"
	File /a "images\knight.bmp"
	File /a "images\knightFlail.bmp"
	File /a "images\knightHorse.bmp"
	File /a "images\dragon.bmp"
	
	File /a "images\health.bmp"
	File /a "images\monHealth.bmp"
	File /a "images\flail.bmp"
	File /a "images\font.bmp"
	File /a "images\castle.bmp"
	File /a "images\background.bmp"
	
	# EXEs and empty database
	SetOutPath "$InstDir"
	File /a "prjATTACK.exe"
	File /a "prjATTACKSERVER.exe"
	File /a "Support\saveFiles.mdb" #empty database in Support folder
	
	#Registry
	WriteRegStr HKLM "Software\Attack" Version "0.0.0.1a"
	
	#Uninstaller
	WriteUninstaller "$InstDir\uninstaller.exe"
	
	# Put shortcuts in the start menu
	CreateDirectory "$SMPrograms\Attack"
	CreateShortCut "$SMPrograms\Attack\Attack Server.lnk" "$InstDir\prjATTACKSERVER.exe"
	CreateShortCut "$SMPrograms\Attack\Attack.lnk" "$InstDir\prjATTACK.exe"
	CreateShortCut "$SMPrograms\Attack\Uninstall Attack.lnk" "$InstDir\uninstaller.exe"
	
	# Add uninstaller information to registry.
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Attack" "DisplayName" "Attack"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Attack" "UninstallString" '"$InstDir\uninstaller.exe"'
	WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Attack" "NoModify" 1
	WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Attack" "NoRepair" 1
SectionEnd

Section Uninstall
	# Delete shortcuts
	Delete /RebootOk "$SMPrograms\Attack\Attack.lnk"
	Delete /RebootOk "$SMPrograms\Attack\Attack Server.lnk"
	Delete /RebootOk "$SMPrograms\Attack\Uninstall Attack.lnk"
	RMDir "$SMPrograms\Attack"	
	
	# Unregister Winsock
	ClearErrors
	ReadRegStr $WinSockLicence HKLM "Software\Attack" ShouldDeleteWinsock
	StrCmp $WinSockLicence 1 0 UnAfterRegWinsock
	DeleteRegKey HKCR "Licenses\2c49f800-c2dd-11cf-9ad6-0080c7e7b78d"
	UnAfterRegWinsock:
	
	# Uninstall VB6 DLLs
	!insertmacro UnInstallLib REGDLL SHARED REBOOT_PROTECTED "$CommonFiles\Microsoft Shared\DAO\DAO350.DLL"
	!insertmacro UnInstallLib REGDLL SHARED REBOOT_PROTECTED "$SysDir\MSJET35.DLL"
	!insertmacro UnInstallLib REGDLL SHARED REBOOT_PROTECTED "$SysDir\MSJTER35.DLL"
	!insertmacro UnInstallLib REGDLL SHARED REBOOT_PROTECTED "$SysDir\MSJINT35.DLL"
	!insertmacro UnInstallLib REGDLL SHARED REBOOT_PROTECTED "$SysDir\vbajet32.dll"
	!insertmacro UnInstallLib REGDLL SHARED REBOOT_PROTECTED "$SysDir\expsrv.dll"
	!insertmacro UnInstallLib REGDLL SHARED REBOOT_PROTECTED "$SysDir\MSRD2X35.DLL"
	!insertmacro UnInstallLib REGDLL SHARED REBOOT_PROTECTED "$SysDir\MSREPL35.DLL"
	!insertmacro UnInstallLib REGDLL SHARED REBOOT_PROTECTED "$SysDir\VB5DB.DLL"
	!insertmacro UnInstallLib REGDLL SHARED REBOOT_PROTECTED "$SysDir\RICHTX32.OCX"
	!insertmacro UnInstallLib REGDLL SHARED REBOOT_PROTECTED "$SysDir\RICHED32.DLL"
	!insertmacro UnInstallLib REGDLL SHARED REBOOT_PROTECTED "$SysDir\mswinsck.ocx"
	
	# Delete files
	Delete /RebootOk "$InstDir\prjATTACK.exe"
	Delete /RebootOk "$InstDir\prjATTACKSERVER.exe"
	Delete /RebootOk "$InstDir\saveFiles.mdb"
	Delete /RebootOk "$InstDir\uninstaller.exe"
	
	#Images
	Delete /RebootOk "$InstDir\images\background.bmp"
	Delete /RebootOk "$InstDir\images\castle.bmp"
	Delete /RebootOk "$InstDir\images\dragon.bmp"
	Delete /RebootOk "$InstDir\images\font.bmp"
	Delete /RebootOk "$InstDir\images\flail.bmp"
	Delete /RebootOk "$InstDir\images\health.bmp"
	Delete /RebootOk "$InstDir\images\knight.bmp"
	Delete /RebootOk "$InstDir\images\knightFlail.bmp"
	Delete /RebootOk "$InstDir\images\knightHorse.bmp"
	Delete /RebootOk "$InstDir\images\monster0.bmp"
	Delete /RebootOk "$InstDir\images\monster1.bmp"
	Delete /RebootOk "$InstDir\images\monster2.bmp"
	Delete /RebootOk "$InstDir\images\monster3.bmp"
	Delete /RebootOk "$InstDir\images\monster4.bmp"
	Delete /RebootOk "$InstDir\images\monster5.bmp"
	Delete /RebootOk "$InstDir\images\monster6.bmp"
	
	# Delete Registry key
	DeleteRegKey HKLM "Software\Attack"
	
	# Delete folders
	SetOutPath "\"
	RmDir /RebootOk "$INSTDIR\images"
	RmDir /RebootOk "$INSTDIR"
SectionEnd

Section "-Install VB6 runtimes"
	;Add code here that sets $ALREADY_INSTALLED to a non-zero value if"
	;the application is already installed. For example:
	
	ReadRegStr $Version HKLM "Software\Attack" Version
	IfErrors new_installation2
	StrCpy $ALREADY_INSTALLED 1
	
	new_installation2:
	
	!insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_NOTPROTECTED \
		"Support\msvbvm60.dll" "$SYSDIR\msvbvm60.dll" "$SYSDIR"
	!insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED \
		"Support\oleaut32.dll" "$SYSDIR\oleaut32.dll" "$SYSDIR"
	!insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED \
		"Support\olepro32.dll" "$SYSDIR\olepro32.dll" "$SYSDIR"
	!insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED \
		"Support\comcat.dll"   "$SYSDIR\comcat.dll"   "$SYSDIR"
	!insertmacro InstallLib DLL    $ALREADY_INSTALLED REBOOT_PROTECTED \
		"Support\asycfilt.dll" "$SYSDIR\asycfilt.dll" "$SYSDIR"
	!insertmacro InstallLib TLB    $ALREADY_INSTALLED REBOOT_PROTECTED \
		"Support\stdole2.tlb"  "$SYSDIR\stdole2.tlb"  "$SYSDIR"
SectionEnd

Section "-un.Uninstall VB6 runtimes"
	!insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\msvbvm60.dll"
	!insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\oleaut32.dll"
	!insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\olepro32.dll"
	!insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\comcat.dll"
	!insertmacro UnInstallLib DLL    SHARED NOREMOVE "$SYSDIR\asycfilt.dll"
	!insertmacro UnInstallLib TLB    SHARED NOREMOVE "$SYSDIR\stdole2.tlb"
SectionEnd
