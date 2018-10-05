@ECHO OFF

@REM ----------------------------------------------------------------
@ECHO ____ BAT 24- ASSOCIATION EXTENSIONS _ IRFANVIEW.BAT ____ 
@REM ----------------------------------------------------------------

TITLE %~f0
@CD\

@REM ----------------------------------------------------------------
@REM __ MATTHEW LANCASTER 
@REM __ MATT.LAN@BTINTERNET.COM
@REM ----------------------------------------------------------------
@REM __ ALL BATCH FILE COMMAND SETS 
@REM __ GOT AN IMPROVED ADMINISTRATOR LAUNCHER
@REM __ THAT DOESN'T GET STUCK IN A LOOP IF NOT SUCCESSFUL
@REM __ AT THE SACRIFICE OF COMMAND LINE ARGUMENTS BUT EASY TO 
@REM __ ADD CODE IF REQUIRE
@REM ----------------------------------------------------------------
@REM __ FROM ---- Thu 03-May-2018 01:31:51
@REM __ TO   ---- Thu 03-May-2018 03:35:00 __ 2 HOUR
@REM ----------------------------------------------------------------

@REM __ SET ADMIN A SAFE DOUBLE CHECKER WAY NOT A LOOPY ONE IF CAN'T GET IT TO GO ADMIN

IF '%1%'=='/ADMIN' (
	@REM :: Check for ADMIN Privileges Again
	>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
	@IF '%errorlevel%' GTR '2' SET NOT_GOT_ADMIN=TRUE
)

@IF "%NOT_GOT_ADMIN%"=="TRUE" (
	ECHO AN ATTEMPT WAS MADE TO SET ADMIN BUT NOT SUCCESSFUL
	ECHO YOUR USER MUST BE A STANDARD USER OR SOMETHING LIKE HER
	ECHO SAVES GET STUCK IN A LOOP ASKING OR ALSO DELAY REQUIRING
	ECHO END
	PAUSE
	GOTO ENDE
	EXIT /B 1
)
	
)
@REM :: Check for ADMIN Privileges
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
@IF '%errorlevel%' GTR '2' (
	@REM Get ADMIN Privileges
	@ECHO Set UAC = CreateObject^("Shell.Application"^) >> "%temp%\getadmin.vbs"
	@ECHO UAC.ShellExecute "%~s0", "/ADMIN", "", "runas", 1 >> "%temp%\getadmin.vbs"

	"%temp%\getadmin.vbs"
	@DEL "%temp%\getadmin.vbs"
	GOTO ENDE
	@EXIT /B
)
@ECHO Got ADMIN Privileges


@SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

@CD \
REM ECHO ON



:BEGIN_IRFANVIEW_FILE_ASSOCIATIONS
REM NOT *.TXT AND NOT .ICO


Reg.exe add "HKLM\SOFTWARE\Classes\.ani" /ve /t REG_SZ /d "IrfanView.ANI" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.ANI" /ve /t REG_SZ /d "IrfanView ANI File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.ANI\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.ANI\DefaultIcon" /ve /t REG_SZ /d "%%1" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".ANI" /t REG_SZ /d "IrfanView.ANI" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".B3D" /t REG_SZ /d "IrfanView.B3D" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".BMP" /t REG_SZ /d "IrfanView.BMP" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".CAM" /t REG_SZ /d "IrfanView.CAM" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".CLP" /t REG_SZ /d "IrfanView.CLP" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".CRW" /t REG_SZ /d "IrfanView.CRW" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".CR2" /t REG_SZ /d "IrfanView.CR2" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".CUR" /t REG_SZ /d "IrfanView.CUR" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".DCM" /t REG_SZ /d "IrfanView.DCM" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".DCX" /t REG_SZ /d "IrfanView.DCX" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".DDS" /t REG_SZ /d "IrfanView.DDS" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".DJVU" /t REG_SZ /d "IrfanView.DJVU" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".DXF" /t REG_SZ /d "IrfanView.DXF" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".ECW" /t REG_SZ /d "IrfanView.ECW" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".EMF" /t REG_SZ /d "IrfanView.EMF" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".EPS" /t REG_SZ /d "IrfanView.EPS" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".PS" /t REG_SZ /d "IrfanView.PS" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".EXR" /t REG_SZ /d "IrfanView.EXR" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".FPX" /t REG_SZ /d "IrfanView.FPX" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".G3" /t REG_SZ /d "IrfanView.G3" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".GIF" /t REG_SZ /d "IrfanView.GIF" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".HDP" /t REG_SZ /d "IrfanView.HDP" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".JXR" /t REG_SZ /d "IrfanView.JXR" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".WDP" /t REG_SZ /d "IrfanView.WDP" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".ICL" /t REG_SZ /d "IrfanView.ICL" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".IFF" /t REG_SZ /d "IrfanView.IFF" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".IMG" /t REG_SZ /d "IrfanView.IMG" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".JLS" /t REG_SZ /d "IrfanView.JLS" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".JP2" /t REG_SZ /d "IrfanView.JP2" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".JPG" /t REG_SZ /d "IrfanView.JPG" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".JPM" /t REG_SZ /d "IrfanView.JPM" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".KDC" /t REG_SZ /d "IrfanView.KDC" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".MNG" /t REG_SZ /d "IrfanView.MNG" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".JNG" /t REG_SZ /d "IrfanView.JNG" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".PBM" /t REG_SZ /d "IrfanView.PBM" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".PCD" /t REG_SZ /d "IrfanView.PCD" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".PCX" /t REG_SZ /d "IrfanView.PCX" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".PGM" /t REG_SZ /d "IrfanView.PGM" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".PNG" /t REG_SZ /d "IrfanView.PNG" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".PPM" /t REG_SZ /d "IrfanView.PPM" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".PSD" /t REG_SZ /d "IrfanView.PSD" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".PSP" /t REG_SZ /d "IrfanView.PSP" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".RAS" /t REG_SZ /d "IrfanView.RAS" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".RAW" /t REG_SZ /d "IrfanView.RAW" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".RLE" /t REG_SZ /d "IrfanView.RLE" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".SFF" /t REG_SZ /d "IrfanView.SFF" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".SFW" /t REG_SZ /d "IrfanView.SFW" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".SGI" /t REG_SZ /d "IrfanView.SGI" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".SID" /t REG_SZ /d "IrfanView.SID" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".TGA" /t REG_SZ /d "IrfanView.TGA" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".TIF" /t REG_SZ /d "IrfanView.TIF" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".WBMP" /t REG_SZ /d "IrfanView.WBMP" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".WEBP" /t REG_SZ /d "IrfanView.WEBP" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".WMF" /t REG_SZ /d "IrfanView.WMF" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".XBM" /t REG_SZ /d "IrfanView.XBM" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\IrfanView\Capabilities\FileAssociations" /v ".XPM" /t REG_SZ /d "IrfanView.XPM" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.ani\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.ANI" /f
Reg.exe add "HKCU\Software\Classes\AppXvhc4p7vz4b485xfp46hhk3fq3grkdgjg" /v "NoOpenWith" /t REG_SZ /d "" /f
Reg.exe add "HKCU\Software\Classes\AppXk0g4vb8gvt7b93tg50ybcy892pge6jmt" /v "NoOpenWith" /t REG_SZ /d "" /f
Reg.exe add "HKCU\Software\Classes\AppX43hnxtbyyps62jhe9sqpdzxn1790zetc" /v "NoOpenWith" /t REG_SZ /d "" /f
Reg.exe add "HKCU\Software\Classes\AppX9rkaq77s0jzh1tyccadx9ghba15r6t3h" /v "NoOpenWith" /t REG_SZ /d "" /f
Reg.exe add "HKCU\Software\Classes\AppX6eg8h5sxqq90pv53845wmnbewywdqq5h" /v "NoOpenWith" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.b3d" /ve /t REG_SZ /d "IrfanView.B3D" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.b3d" /ve /t REG_SZ /d "IrfanView B3D File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.b3d\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.b3d\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.B3D\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.B3D" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.bmp" /ve /t REG_SZ /d "IrfanView.BMP" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.BMP" /ve /t REG_SZ /d "IrfanView BMP File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.BMP\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.BMP\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.bmp\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.BMP" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.dib" /ve /t REG_SZ /d "IrfanView.BMP" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.dib\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.BMP" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.cam" /ve /t REG_SZ /d "IrfanView.CAM" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.cam" /ve /t REG_SZ /d "IrfanView CAM File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.cam\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.cam\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.CAM\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.CAM" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.clp" /ve /t REG_SZ /d "IrfanView.CLP" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.clp" /ve /t REG_SZ /d "IrfanView CLP File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.clp\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.clp\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.CLP\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.CLP" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.crw" /ve /t REG_SZ /d "IrfanView.CRW" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.CRW" /ve /t REG_SZ /d "IrfanView CRW File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.CRW\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.CRW\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.crw\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.CRW" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.cr2" /ve /t REG_SZ /d "IrfanView.CR2" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.CR2" /ve /t REG_SZ /d "IrfanView CR2 File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.CR2\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.CR2\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.cr2\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.CR2" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.cur" /ve /t REG_SZ /d "IrfanView.CUR" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.cur" /ve /t REG_SZ /d "IrfanView CUR File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.cur\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.cur\DefaultIcon" /ve /t REG_SZ /d "%%1" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.CUR\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.CUR" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.dcm" /ve /t REG_SZ /d "IrfanView.DCM" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.dcm" /ve /t REG_SZ /d "IrfanView DCM File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.dcm\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.dcm\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.DCM\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.DCM" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.acr" /ve /t REG_SZ /d "IrfanView.DCM" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.ACR\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.DCM" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.ima" /ve /t REG_SZ /d "IrfanView.DCM" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.IMA\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.DCM" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.dcx" /ve /t REG_SZ /d "IrfanView.DCX" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.DCX" /ve /t REG_SZ /d "IrfanView DCX File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.DCX\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.DCX\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.dcx\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.DCX" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.dds" /ve /t REG_SZ /d "IrfanView.DDS" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.dds" /ve /t REG_SZ /d "IrfanView DDS File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.dds\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.dds\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.dds\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.DDS" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.djvu" /ve /t REG_SZ /d "IrfanView.DJVU" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.djvu" /ve /t REG_SZ /d "IrfanView DJVU File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.djvu\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.djvu\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.DJVU\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.DJVU" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.iw44" /ve /t REG_SZ /d "IrfanView.DJVU" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.IW44\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.DJVU" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.dxf" /ve /t REG_SZ /d "IrfanView.DXF" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.dxf" /ve /t REG_SZ /d "IrfanView DXF File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.dxf\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.dxf\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.dxf\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.DXF" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.ecw" /ve /t REG_SZ /d "IrfanView.ECW" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.ecw" /ve /t REG_SZ /d "IrfanView ECW File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.ecw\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.ecw\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.ECW\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.ECW" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.emf" /ve /t REG_SZ /d "IrfanView.EMF" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.EMF" /ve /t REG_SZ /d "IrfanView EMF File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.EMF\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.EMF\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.emf\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.EMF" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.eps" /ve /t REG_SZ /d "IrfanView.EPS" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.EPS" /ve /t REG_SZ /d "IrfanView EPS File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.EPS\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.EPS\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.eps\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.EPS" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.ps" /ve /t REG_SZ /d "IrfanView.PS" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.PS" /ve /t REG_SZ /d "IrfanView PS File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.PS\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.PS\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.ps\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.PS" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.exr" /ve /t REG_SZ /d "IrfanView.EXR" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.exr" /ve /t REG_SZ /d "IrfanView EXR File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.exr\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.exr\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.EXR\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.EXR" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.fpx" /ve /t REG_SZ /d "IrfanView.FPX" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.fpx" /ve /t REG_SZ /d "IrfanView FPX File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.fpx\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.fpx\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.fpx\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.FPX" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.g3" /ve /t REG_SZ /d "IrfanView.G3" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.g3" /ve /t REG_SZ /d "IrfanView G3 File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.g3\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.g3\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.G3\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.G3" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.gif" /ve /t REG_SZ /d "IrfanView.GIF" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.GIF" /ve /t REG_SZ /d "IrfanView GIF File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.GIF\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.GIF\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.gif\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.GIF" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.hdp" /ve /t REG_SZ /d "IrfanView.HDP" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.hdp" /ve /t REG_SZ /d "IrfanView HDP File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.hdp\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.hdp\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.HDP\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.HDP" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.jxr" /ve /t REG_SZ /d "IrfanView.JXR" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.JXR" /ve /t REG_SZ /d "IrfanView JXR File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.JXR\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.JXR\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.jxr\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.JXR" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.wdp" /ve /t REG_SZ /d "IrfanView.WDP" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.WDP" /ve /t REG_SZ /d "IrfanView WDP File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.WDP\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.WDP\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.wdp\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.WDP" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.icl" /ve /t REG_SZ /d "IrfanView.ICL" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.icl" /ve /t REG_SZ /d "IrfanView ICL File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.icl\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.icl\DefaultIcon" /ve /t REG_SZ /d "%%1" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.ICL\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.ICL" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.ICO" /ve /t REG_SZ /d "IrfanView ICO File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.ICO\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.ICO\DefaultIcon" /ve /t REG_SZ /d "%%1" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.iff" /ve /t REG_SZ /d "IrfanView.IFF" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.iff" /ve /t REG_SZ /d "IrfanView IFF File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.iff\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.iff\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.IFF\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.IFF" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.lbm" /ve /t REG_SZ /d "IrfanView.IFF" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.lbm\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.IFF" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.img" /ve /t REG_SZ /d "IrfanView.IMG" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.img" /ve /t REG_SZ /d "IrfanView IMG File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.img\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.img\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.img\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.IMG" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.jls" /ve /t REG_SZ /d "IrfanView.JLS" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.jls" /ve /t REG_SZ /d "IrfanView JLS File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.jls\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.jls\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.JLS\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.JLS" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.jp2" /ve /t REG_SZ /d "IrfanView.JP2" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.jp2" /ve /t REG_SZ /d "IrfanView JP2 File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.jp2\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.jp2\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.jp2\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.JP2" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.jpc" /ve /t REG_SZ /d "IrfanView.JP2" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.jpc\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.JP2" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.j2k" /ve /t REG_SZ /d "IrfanView.JP2" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.j2k\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.JP2" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.JPF" /ve /t REG_SZ /d "IrfanView.JP2" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.JPF\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.JP2" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.jpg" /ve /t REG_SZ /d "IrfanView.JPG" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.JPG" /ve /t REG_SZ /d "IrfanView JPG File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.JPG\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.JPG\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.jpg\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.JPG" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.jpeg" /ve /t REG_SZ /d "IrfanView.JPG" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.jpeg\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.JPG" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.jpe" /ve /t REG_SZ /d "IrfanView.JPG" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.jpe\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.JPG" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.jpm" /ve /t REG_SZ /d "IrfanView.JPM" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.jpm" /ve /t REG_SZ /d "IrfanView JPM File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.jpm\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.jpm\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.JPM\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.JPM" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.kdc" /ve /t REG_SZ /d "IrfanView.KDC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.KDC" /ve /t REG_SZ /d "IrfanView KDC File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.KDC\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.KDC\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.kdc\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.KDC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.mng" /ve /t REG_SZ /d "IrfanView.MNG" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.mng" /ve /t REG_SZ /d "IrfanView MNG File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.mng\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.mng\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.MNG\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.MNG" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.jng" /ve /t REG_SZ /d "IrfanView.JNG" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.jng" /ve /t REG_SZ /d "IrfanView JNG File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.jng\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.jng\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.JNG\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.JNG" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.pbm" /ve /t REG_SZ /d "IrfanView.PBM" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.pbm" /ve /t REG_SZ /d "IrfanView PBM File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.pbm\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.pbm\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.PBM\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.PBM" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.pcd" /ve /t REG_SZ /d "IrfanView.PCD" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.pcd" /ve /t REG_SZ /d "IrfanView PCD File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.pcd\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.pcd\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.PCD\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.PCD" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.pcx" /ve /t REG_SZ /d "IrfanView.PCX" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.PCX" /ve /t REG_SZ /d "IrfanView PCX File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.PCX\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.PCX\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.pcx\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.PCX" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.pgm" /ve /t REG_SZ /d "IrfanView.PGM" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.PGM" /ve /t REG_SZ /d "IrfanView PGM File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.PGM\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.PGM\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.pgm\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.PGM" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.png" /ve /t REG_SZ /d "IrfanView.PNG" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.PNG" /ve /t REG_SZ /d "IrfanView PNG File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.PNG\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.PNG\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.png\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.PNG" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.ppm" /ve /t REG_SZ /d "IrfanView.PPM" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.PPM" /ve /t REG_SZ /d "IrfanView PPM File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.PPM\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.PPM\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.ppm\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.PPM" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.psd" /ve /t REG_SZ /d "IrfanView.PSD" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.PSD" /ve /t REG_SZ /d "IrfanView PSD File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.PSD\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.PSD\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.psd\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.PSD" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.psp" /ve /t REG_SZ /d "IrfanView.PSP" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.psp" /ve /t REG_SZ /d "IrfanView PSP File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.psp\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.psp\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.PSP\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.PSP" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.ras" /ve /t REG_SZ /d "IrfanView.RAS" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.ras" /ve /t REG_SZ /d "IrfanView RAS File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.ras\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.ras\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.RAS\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.RAS" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.sun" /ve /t REG_SZ /d "IrfanView.RAS" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.SUN\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.RAS" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.raw" /ve /t REG_SZ /d "IrfanView.RAW" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.RAW" /ve /t REG_SZ /d "IrfanView RAW File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.RAW\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.RAW\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.raw\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.RAW" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.rle" /ve /t REG_SZ /d "IrfanView.RLE" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.RLE" /ve /t REG_SZ /d "IrfanView RLE File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.RLE\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.RLE\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.rle\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.RLE" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.sff" /ve /t REG_SZ /d "IrfanView.SFF" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.sff" /ve /t REG_SZ /d "IrfanView SFF File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.sff\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.sff\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.SFF\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.SFF" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.sfw" /ve /t REG_SZ /d "IrfanView.SFW" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.sfw" /ve /t REG_SZ /d "IrfanView SFW File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.sfw\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.sfw\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.SFW\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.SFW" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.sgi" /ve /t REG_SZ /d "IrfanView.SGI" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.sgi" /ve /t REG_SZ /d "IrfanView SGI File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.sgi\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.sgi\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.SGI\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.SGI" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.rgb" /ve /t REG_SZ /d "IrfanView.SGI" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.rgb\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.SGI" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.sid" /ve /t REG_SZ /d "IrfanView.SID" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.sid" /ve /t REG_SZ /d "IrfanView SID File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.sid\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.sid\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.SID\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.SID" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.tga" /ve /t REG_SZ /d "IrfanView.TGA" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.tga" /ve /t REG_SZ /d "IrfanView TGA File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.tga\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.tga\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.TGA\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.TGA" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.tif" /ve /t REG_SZ /d "IrfanView.TIF" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.TIF" /ve /t REG_SZ /d "IrfanView TIF File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.TIF\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.TIF\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.tif\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.TIF" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.tiff" /ve /t REG_SZ /d "IrfanView.TIF" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.tiff\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.TIF" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.wbmp" /ve /t REG_SZ /d "IrfanView.WBMP" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.wbmp" /ve /t REG_SZ /d "IrfanView WBMP File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.wbmp\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.wbmp\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.WBMP\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.WBMP" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.webp" /ve /t REG_SZ /d "IrfanView.WEBP" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.webp" /ve /t REG_SZ /d "IrfanView WEBP File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.webp\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.webp\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.WEBP\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.WEBP" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.wmf" /ve /t REG_SZ /d "IrfanView.WMF" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.WMF" /ve /t REG_SZ /d "IrfanView WMF File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.WMF\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.WMF\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.wmf\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.WMF" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.xbm" /ve /t REG_SZ /d "IrfanView.XBM" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.XBM" /ve /t REG_SZ /d "IrfanView XBM File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.XBM\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.XBM\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.xbm\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.XBM" /f
Reg.exe add "HKLM\SOFTWARE\Classes\.xpm" /ve /t REG_SZ /d "IrfanView.XPM" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.XPM" /ve /t REG_SZ /d "IrfanView XPM File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.XPM\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.XPM\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.xpm\UserChoice" /v "Progid" /t REG_SZ /d "IrfanView.XPM" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.swf" /ve /t REG_SZ /d "IrfanView SWF File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.swf\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.swf\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.flv" /ve /t REG_SZ /d "IrfanView FLV File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.flv\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.flv\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.ttf" /ve /t REG_SZ /d "IrfanView TTF File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.ttf\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.ttf\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.asf" /ve /t REG_SZ /d "IrfanView ASF File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.asf\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.asf\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,9" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.avi" /ve /t REG_SZ /d "IrfanView AVI File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.avi\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.avi\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,9" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.mpg" /ve /t REG_SZ /d "IrfanView MPG File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.mpg\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.mpg\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,9" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.mov" /ve /t REG_SZ /d "IrfanView MOV File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.mov\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.mov\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,9" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.wmv" /ve /t REG_SZ /d "IrfanView WMV File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.wmv\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.wmv\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,9" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.aif" /ve /t REG_SZ /d "IrfanView AIF File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.aif\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.aif\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,8" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.mid" /ve /t REG_SZ /d "IrfanView MID File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.mid\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.mid\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,8" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.rmi" /ve /t REG_SZ /d "IrfanView RMI File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.rmi\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.rmi\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,8" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.med" /ve /t REG_SZ /d "IrfanView MED File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.med\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.med\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,8" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.mp3" /ve /t REG_SZ /d "IrfanView MP3 File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.mp3\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.mp3\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,8" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.ogg" /ve /t REG_SZ /d "IrfanView OGG File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.ogg\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.ogg\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,8" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.ra" /ve /t REG_SZ /d "IrfanView RA File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.ra\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.ra\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,8" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.au" /ve /t REG_SZ /d "IrfanView AU File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.au\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.au\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,8" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.wav" /ve /t REG_SZ /d "IrfanView WAV File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.wav\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.wav\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,8" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.wma" /ve /t REG_SZ /d "IrfanView WMA File" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.wma\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\IrfanView.wma\DefaultIcon" /ve /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe,8" /f
Reg.exe add "HKLM\SOFTWARE\Classes\Directory\shell\Browse with &IrfanView" /ve /t REG_SZ /d "Browse with &IrfanView" /f
Reg.exe add "HKLM\SOFTWARE\Classes\Directory\shell\Browse with &IrfanView\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\" /thumbs" /f
Reg.exe add "HKLM\SOFTWARE\Classes\Drive\shell\Browse with &IrfanView" /ve /t REG_SZ /d "Browse with &IrfanView" /f
Reg.exe add "HKLM\SOFTWARE\Classes\Drive\shell\Browse with &IrfanView\command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\IrfanView\i_view32.exe\" \"%%1\" /thumbs" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer" /v "GlobalAssocChangedCounter" /t REG_DWORD /d "39" /f

:END_IRFANVIEW_FILE_ASSOCIATION


:VERY_ENDER_BLOCK

:ENDE

PAUSE

