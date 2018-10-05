@ECHO OFF

@REM ----------------------------------------------------------------
@ECHO ____ BAT 25- ASSOCIATION EXTENSIONS _ MPC.BAT ____ 
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

:BEGIN_MEDIA_PLAYER_CLASSIC__MPC

ECHO MEDIA_PLAYER_CLASSIC __ MPC


Reg.exe add "HKLM\SOFTWARE\RegisteredApplications" /v "Media Player Classic" /t REG_SZ /d "Software\Clients\Media\Media Player Classic\Capabilities" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities" /v "ApplicationDescription" /t REG_EXPAND_SZ /d "MPC-HC is an extremely light-weight, open source media player for Windows. It supports all common video and audio file formats available for playback. We are 100%% spyware free, there are no advertisements or toolbars." /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities" /v "ApplicationIcon" /t REG_EXPAND_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities" /v "ApplicationName" /t REG_EXPAND_SZ /d "Media Player Classic Home Cinema" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.avi" /ve /t REG_SZ /d "AVI" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.avi" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.avi\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.avi\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.avi\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.avi\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.avi\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.avi\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".avi" /t REG_SZ /d "mplayerc64.avi" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".mpg" /t REG_SZ /d "mplayerc64.mpg" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".mpeg" /t REG_SZ /d "mplayerc64.mpeg" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".mpe" /t REG_SZ /d "mplayerc64.mpe" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".m1v" /t REG_SZ /d "mplayerc64.m1v" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".m2v" /t REG_SZ /d "mplayerc64.m2v" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".mpv2" /t REG_SZ /d "mplayerc64.mpv2" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".mp2v" /t REG_SZ /d "mplayerc64.mp2v" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".pva" /t REG_SZ /d "mplayerc64.pva" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".evo" /t REG_SZ /d "mplayerc64.evo" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".m2p" /t REG_SZ /d "mplayerc64.m2p" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".ts" /t REG_SZ /d "mplayerc64.ts" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".tp" /t REG_SZ /d "mplayerc64.tp" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".trp" /t REG_SZ /d "mplayerc64.trp" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".m2t" /t REG_SZ /d "mplayerc64.m2t" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".m2ts" /t REG_SZ /d "mplayerc64.m2ts" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".mts" /t REG_SZ /d "mplayerc64.mts" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".rec" /t REG_SZ /d "mplayerc64.rec" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".ssif" /t REG_SZ /d "mplayerc64.ssif" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".vob" /t REG_SZ /d "mplayerc64.vob" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".ifo" /t REG_SZ /d "mplayerc64.ifo" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".mkv" /t REG_SZ /d "mplayerc64.mkv" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".mk3d" /t REG_SZ /d "mplayerc64.mk3d" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".webm" /t REG_SZ /d "mplayerc64.webm" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".mp4" /t REG_SZ /d "mplayerc64.mp4" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".m4v" /t REG_SZ /d "mplayerc64.m4v" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".mp4v" /t REG_SZ /d "mplayerc64.mp4v" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".mpv4" /t REG_SZ /d "mplayerc64.mpv4" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".hdmov" /t REG_SZ /d "mplayerc64.hdmov" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".mov" /t REG_SZ /d "mplayerc64.mov" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".3gp" /t REG_SZ /d "mplayerc64.3gp" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".3gpp" /t REG_SZ /d "mplayerc64.3gpp" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".3ga" /t REG_SZ /d "mplayerc64.3ga" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".3g2" /t REG_SZ /d "mplayerc64.3g2" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".3gp2" /t REG_SZ /d "mplayerc64.3gp2" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".flv" /t REG_SZ /d "mplayerc64.flv" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".f4v" /t REG_SZ /d "mplayerc64.f4v" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".ogm" /t REG_SZ /d "mplayerc64.ogm" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".ogv" /t REG_SZ /d "mplayerc64.ogv" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".rm" /t REG_SZ /d "mplayerc64.rm" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".rmvb" /t REG_SZ /d "mplayerc64.rmvb" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".rt" /t REG_SZ /d "mplayerc64.rt" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".ram" /t REG_SZ /d "mplayerc64.ram" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".rpm" /t REG_SZ /d "mplayerc64.rpm" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".rmm" /t REG_SZ /d "mplayerc64.rmm" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".rp" /t REG_SZ /d "mplayerc64.rp" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".smi" /t REG_SZ /d "mplayerc64.smi" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".smil" /t REG_SZ /d "mplayerc64.smil" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".wmv" /t REG_SZ /d "mplayerc64.wmv" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".wmp" /t REG_SZ /d "mplayerc64.wmp" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".wm" /t REG_SZ /d "mplayerc64.wm" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".asf" /t REG_SZ /d "mplayerc64.asf" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".smk" /t REG_SZ /d "mplayerc64.smk" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".bik" /t REG_SZ /d "mplayerc64.bik" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".fli" /t REG_SZ /d "mplayerc64.fli" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".flc" /t REG_SZ /d "mplayerc64.flc" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".flic" /t REG_SZ /d "mplayerc64.flic" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".dsm" /t REG_SZ /d "mplayerc64.dsm" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".dsv" /t REG_SZ /d "mplayerc64.dsv" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".dsa" /t REG_SZ /d "mplayerc64.dsa" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".dss" /t REG_SZ /d "mplayerc64.dss" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".ivf" /t REG_SZ /d "mplayerc64.ivf" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".swf" /t REG_SZ /d "mplayerc64.swf" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".divx" /t REG_SZ /d "mplayerc64.divx" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".amv" /t REG_SZ /d "mplayerc64.amv" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".ac3" /t REG_SZ /d "mplayerc64.ac3" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".dts" /t REG_SZ /d "mplayerc64.dts" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".dtshd" /t REG_SZ /d "mplayerc64.dtshd" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".dtsma" /t REG_SZ /d "mplayerc64.dtsma" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".aif" /t REG_SZ /d "mplayerc64.aif" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".aifc" /t REG_SZ /d "mplayerc64.aifc" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".aiff" /t REG_SZ /d "mplayerc64.aiff" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".alac" /t REG_SZ /d "mplayerc64.alac" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".amr" /t REG_SZ /d "mplayerc64.amr" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".ape" /t REG_SZ /d "mplayerc64.ape" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".apl" /t REG_SZ /d "mplayerc64.apl" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".au" /t REG_SZ /d "mplayerc64.au" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".snd" /t REG_SZ /d "mplayerc64.snd" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".cda" /t REG_SZ /d "mplayerc64.cda" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".flac" /t REG_SZ /d "mplayerc64.flac" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".m4a" /t REG_SZ /d "mplayerc64.m4a" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".m4b" /t REG_SZ /d "mplayerc64.m4b" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".m4r" /t REG_SZ /d "mplayerc64.m4r" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".aac" /t REG_SZ /d "mplayerc64.aac" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".mid" /t REG_SZ /d "mplayerc64.mid" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".midi" /t REG_SZ /d "mplayerc64.midi" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".rmi" /t REG_SZ /d "mplayerc64.rmi" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".mka" /t REG_SZ /d "mplayerc64.mka" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".mp3" /t REG_SZ /d "mplayerc64.mp3" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".mpa" /t REG_SZ /d "mplayerc64.mpa" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".mp2" /t REG_SZ /d "mplayerc64.mp2" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".m1a" /t REG_SZ /d "mplayerc64.m1a" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".m2a" /t REG_SZ /d "mplayerc64.m2a" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".mpc" /t REG_SZ /d "mplayerc64.mpc" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".ofr" /t REG_SZ /d "mplayerc64.ofr" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".ofs" /t REG_SZ /d "mplayerc64.ofs" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".ogg" /t REG_SZ /d "mplayerc64.ogg" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".oga" /t REG_SZ /d "mplayerc64.oga" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".opus" /t REG_SZ /d "mplayerc64.opus" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".ra" /t REG_SZ /d "mplayerc64.ra" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".tak" /t REG_SZ /d "mplayerc64.tak" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".tta" /t REG_SZ /d "mplayerc64.tta" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".wav" /t REG_SZ /d "mplayerc64.wav" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".wma" /t REG_SZ /d "mplayerc64.wma" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".wv" /t REG_SZ /d "mplayerc64.wv" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".aob" /t REG_SZ /d "mplayerc64.aob" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".mlp" /t REG_SZ /d "mplayerc64.mlp" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".thd" /t REG_SZ /d "mplayerc64.thd" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".asx" /t REG_SZ /d "mplayerc64.asx" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".m3u" /t REG_SZ /d "mplayerc64.m3u" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".m3u8" /t REG_SZ /d "mplayerc64.m3u8" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".pls" /t REG_SZ /d "mplayerc64.pls" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".wvx" /t REG_SZ /d "mplayerc64.wvx" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".wax" /t REG_SZ /d "mplayerc64.wax" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".wmx" /t REG_SZ /d "mplayerc64.wmx" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".mpcpl" /t REG_SZ /d "mplayerc64.mpcpl" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".mpls" /t REG_SZ /d "mplayerc64.mpls" /f
Reg.exe add "HKLM\SOFTWARE\Clients\Media\Media Player Classic\Capabilities\FileAssociations" /v ".bdmv" /t REG_SZ /d "mplayerc64.bdmv" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.avi\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",8" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpg" /ve /t REG_SZ /d "MPEG" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpg" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpg\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpg\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpg\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpg\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpg\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpg\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpg\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",26" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpeg" /ve /t REG_SZ /d "MPEG" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpeg" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpeg\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpeg\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpeg\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpeg\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpeg\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpeg\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpeg\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",26" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpe" /ve /t REG_SZ /d "MPEG" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpe" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpe\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpe\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpe\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpe\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpe\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpe\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpe\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",26" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m1v" /ve /t REG_SZ /d "MPEG" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m1v" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m1v\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m1v\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m1v\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m1v\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m1v\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m1v\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m1v\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",26" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2v" /ve /t REG_SZ /d "MPEG" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2v" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2v\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2v\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2v\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2v\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2v\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2v\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2v\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",26" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpv2" /ve /t REG_SZ /d "MPEG" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpv2" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpv2\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpv2\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpv2\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpv2\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpv2\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpv2\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpv2\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",26" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp2v" /ve /t REG_SZ /d "MPEG" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp2v" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp2v\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp2v\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp2v\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp2v\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp2v\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp2v\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp2v\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",26" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.pva" /ve /t REG_SZ /d "MPEG" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.pva" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.pva\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.pva\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.pva\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.pva\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.pva\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.pva\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.pva\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",26" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.evo" /ve /t REG_SZ /d "MPEG" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.evo" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.evo\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.evo\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.evo\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.evo\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.evo\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.evo\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.evo\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",26" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2p" /ve /t REG_SZ /d "MPEG" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2p" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2p\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2p\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2p\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2p\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2p\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2p\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2p\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",26" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ts" /ve /t REG_SZ /d "MPEG-TS" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ts" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ts\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ts\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ts\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ts\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ts\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ts\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ts\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",36" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.tp" /ve /t REG_SZ /d "MPEG-TS" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.tp" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.tp\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.tp\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.tp\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.tp\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.tp\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.tp\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.tp\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",36" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.trp" /ve /t REG_SZ /d "MPEG-TS" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.trp" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.trp\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.trp\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.trp\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.trp\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.trp\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.trp\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.trp\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",36" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2t" /ve /t REG_SZ /d "MPEG-TS" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2t" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2t\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2t\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2t\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2t\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2t\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2t\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2t\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",36" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2ts" /ve /t REG_SZ /d "MPEG-TS" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2ts" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2ts\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2ts\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2ts\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2ts\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2ts\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2ts\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2ts\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",36" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mts" /ve /t REG_SZ /d "MPEG-TS" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mts" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mts\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mts\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mts\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mts\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mts\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mts\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mts\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",36" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rec" /ve /t REG_SZ /d "MPEG-TS" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rec" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rec\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rec\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rec\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rec\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rec\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rec\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rec\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",36" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ssif" /ve /t REG_SZ /d "MPEG-TS" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ssif" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ssif\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ssif\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ssif\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ssif\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ssif\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ssif\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ssif\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.vob" /ve /t REG_SZ /d "DVD-Video" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.vob" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.vob\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.vob\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.vob\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.vob\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.vob\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.vob\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.vob\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",38" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ifo" /ve /t REG_SZ /d "DVD-Video" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ifo" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ifo\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ifo\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ifo\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ifo\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ifo\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ifo\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ifo\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",16" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mkv" /ve /t REG_SZ /d "Matroska" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mkv" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mkv\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mkv\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mkv\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mkv\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mkv\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mkv\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mkv\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",20" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mk3d" /ve /t REG_SZ /d "Matroska" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mk3d" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mk3d\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mk3d\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mk3d\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mk3d\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mk3d\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mk3d\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mk3d\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.webm" /ve /t REG_SZ /d "WebM" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.webm" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.webm\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.webm\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.webm\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.webm\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.webm\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.webm\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.webm\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",40" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp4" /ve /t REG_SZ /d "MP4" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp4" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp4\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp4\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp4\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp4\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp4\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp4\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp4\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",23" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4v" /ve /t REG_SZ /d "MP4" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4v" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4v\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4v\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4v\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4v\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4v\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4v\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4v\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",23" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp4v" /ve /t REG_SZ /d "MP4" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp4v" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp4v\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp4v\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp4v\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp4v\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp4v\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp4v\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp4v\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",23" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpv4" /ve /t REG_SZ /d "MP4" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpv4" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpv4\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpv4\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpv4\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpv4\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpv4\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpv4\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpv4\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",23" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.hdmov" /ve /t REG_SZ /d "MP4" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.hdmov" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.hdmov\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.hdmov\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.hdmov\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.hdmov\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.hdmov\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.hdmov\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.hdmov\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",23" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mov" /ve /t REG_SZ /d "QuickTime Movie" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mov" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mov\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mov\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mov\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mov\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mov\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mov\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mov\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",21" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3gp" /ve /t REG_SZ /d "3GP" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3gp" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3gp\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3gp\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3gp\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3gp\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3gp\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3gp\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3gp\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",23" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3gpp" /ve /t REG_SZ /d "3GP" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3gpp" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3gpp\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3gpp\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3gpp\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3gpp\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3gpp\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3gpp\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3gpp\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",23" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3ga" /ve /t REG_SZ /d "3GP" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3ga" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3ga\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3ga\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3ga\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3ga\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3ga\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3ga\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3ga\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",23" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3g2" /ve /t REG_SZ /d "3G2" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3g2" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3g2\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3g2\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3g2\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3g2\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3g2\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3g2\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3g2\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",21" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3gp2" /ve /t REG_SZ /d "3G2" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3gp2" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3gp2\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3gp2\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3gp2\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3gp2\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3gp2\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3gp2\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.3gp2\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",21" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flv" /ve /t REG_SZ /d "Flash Video" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flv" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flv\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flv\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flv\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flv\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flv\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flv\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flv\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",15" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.f4v" /ve /t REG_SZ /d "Flash Video" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.f4v" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.f4v\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.f4v\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.f4v\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.f4v\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.f4v\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.f4v\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.f4v\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",15" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ogm" /ve /t REG_SZ /d "Ogg Media" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ogm" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ogm\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ogm\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ogm\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ogm\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ogm\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ogm\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ogm\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",29" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ogv" /ve /t REG_SZ /d "Ogg Media" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ogv" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ogv\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ogv\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ogv\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ogv\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ogv\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ogv\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ogv\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",29" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rm" /ve /t REG_SZ /d "Real Media" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rm" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rm\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rm\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rm\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rm\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rm\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rm\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rm\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",32" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rmvb" /ve /t REG_SZ /d "Real Media" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rmvb" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rmvb\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rmvb\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rmvb\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rmvb\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rmvb\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rmvb\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rmvb\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rt" /ve /t REG_SZ /d "Real Script" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rt" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rt\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rt\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rt\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rt\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rt\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rt\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rt\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",33" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ram" /ve /t REG_SZ /d "Real Script" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ram" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ram\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ram\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ram\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ram\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ram\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ram\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ram\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",32" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rpm" /ve /t REG_SZ /d "Real Script" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rpm" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rpm\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rpm\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rpm\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rpm\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rpm\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rpm\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rpm\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rmm" /ve /t REG_SZ /d "Real Script" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rmm" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rmm\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rmm\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rmm\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rmm\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rmm\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rmm\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rmm\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",32" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rp" /ve /t REG_SZ /d "Real Script" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rp" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rp\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rp\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rp\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rp\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rp\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rp\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rp\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",33" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.smi" /ve /t REG_SZ /d "Real Script" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.smi" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.smi\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.smi\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.smi\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.smi\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.smi\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.smi\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.smi\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.smil" /ve /t REG_SZ /d "Real Script" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.smil" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.smil\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.smil\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.smil\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.smil\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.smil\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.smil\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.smil\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",33" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wmv" /ve /t REG_SZ /d "Windows Media Video" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wmv" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wmv\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wmv\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wmv\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wmv\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wmv\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wmv\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wmv\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",42" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wmp" /ve /t REG_SZ /d "Windows Media Video" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wmp" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wmp\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wmp\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wmp\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wmp\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wmp\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wmp\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wmp\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",42" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wm" /ve /t REG_SZ /d "Windows Media Video" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wm" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wm\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wm\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wm\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wm\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wm\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wm\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wm\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",42" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.asf" /ve /t REG_SZ /d "Windows Media Video" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.asf" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.asf\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.asf\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.asf\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.asf\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.asf\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.asf\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.asf\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",42" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.smk" /ve /t REG_SZ /d "Smacker/Bink Video" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.smk" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.smk\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.smk\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.smk\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.smk\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.smk\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.smk\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.smk\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",34" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.bik" /ve /t REG_SZ /d "Smacker/Bink Video" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.bik" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.bik\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.bik\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.bik\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.bik\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.bik\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.bik\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.bik\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",9" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.fli" /ve /t REG_SZ /d "FLIC Animation" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.fli" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.fli\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.fli\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.fli\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.fli\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.fli\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.fli\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.fli\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",14" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flc" /ve /t REG_SZ /d "FLIC Animation" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flc" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flc\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flc\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flc\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flc\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flc\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flc\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flc\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",14" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flic" /ve /t REG_SZ /d "FLIC Animation" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flic" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flic\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flic\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flic\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flic\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flic\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flic\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flic\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",14" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dsm" /ve /t REG_SZ /d "DirectShow Media" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dsm" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dsm\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dsm\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dsm\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dsm\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dsm\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dsm\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dsm\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",11" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dsv" /ve /t REG_SZ /d "DirectShow Media" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dsv" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dsv\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dsv\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dsv\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dsv\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dsv\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dsv\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dsv\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",11" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dsa" /ve /t REG_SZ /d "DirectShow Media" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dsa" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dsa\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dsa\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dsa\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dsa\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dsa\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dsa\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dsa\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",11" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dss" /ve /t REG_SZ /d "DirectShow Media" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dss" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dss\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dss\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dss\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dss\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dss\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dss\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dss\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",11" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ivf" /ve /t REG_SZ /d "Indeo Video Format" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ivf" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ivf\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ivf\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ivf\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ivf\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ivf\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ivf\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ivf\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",17" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.swf" /ve /t REG_SZ /d "Shockwave Flash" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.swf" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.swf\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.swf\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.swf\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.swf\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.swf\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.swf\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.swf\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",35" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.divx" /ve /t REG_SZ /d "Other" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.divx" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.divx\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.divx\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.divx\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.divx\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.divx\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.divx\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.divx\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.amv" /ve /t REG_SZ /d "Other" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.amv" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.amv\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.amv\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.amv\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.amv\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.amv\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.amv\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.amv\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ac3" /ve /t REG_SZ /d "AC-3" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ac3" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ac3\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ac3\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ac3\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ac3\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ac3\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ac3\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ac3\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",2" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dts" /ve /t REG_SZ /d "DTS/DTS-HD" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dts" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dts\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dts\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dts\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dts\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dts\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dts\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dts\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",12" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dtshd" /ve /t REG_SZ /d "DTS/DTS-HD" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dtshd" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dtshd\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dtshd\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dtshd\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dtshd\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dtshd\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dtshd\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dtshd\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",12" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dtsma" /ve /t REG_SZ /d "DTS/DTS-HD" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dtsma" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dtsma\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dtsma\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dtsma\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dtsma\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dtsma\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dtsma\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.dtsma\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",12" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aif" /ve /t REG_SZ /d "AIFF" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aif" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aif\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aif\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aif\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aif\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aif\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aif\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aif\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",3" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aifc" /ve /t REG_SZ /d "AIFF" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aifc" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aifc\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aifc\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aifc\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aifc\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aifc\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aifc\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aifc\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",3" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aiff" /ve /t REG_SZ /d "AIFF" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aiff" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aiff\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aiff\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aiff\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aiff\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aiff\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aiff\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aiff\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",3" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.alac" /ve /t REG_SZ /d "Apple Lossless" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.alac" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.alac\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.alac\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.alac\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.alac\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.alac\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.alac\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.alac\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",4" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.amr" /ve /t REG_SZ /d "AMR" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.amr" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.amr\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.amr\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.amr\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.amr\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.amr\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.amr\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.amr\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",5" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ape" /ve /t REG_SZ /d "Monkey's Audio" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ape" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ape\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ape\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ape\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ape\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ape\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ape\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ape\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",6" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.apl" /ve /t REG_SZ /d "Monkey's Audio" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.apl" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.apl\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.apl\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.apl\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.apl\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.apl\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.apl\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.apl\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",6" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.au" /ve /t REG_SZ /d "AU/SND" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.au" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.au\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.au\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.au\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.au\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.au\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.au\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.au\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",7" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.snd" /ve /t REG_SZ /d "AU/SND" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.snd" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.snd\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.snd\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.snd\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.snd\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.snd\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.snd\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.snd\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",7" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.cda" /ve /t REG_SZ /d "Audio CD track" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.cda" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.cda\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.cda\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.cda\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.cda\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.cda\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.cda\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.cda\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",10" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flac" /ve /t REG_SZ /d "FLAC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flac" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flac\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flac\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flac\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flac\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flac\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flac\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.flac\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",13" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4a" /ve /t REG_SZ /d "MPEG-4 Audio" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4a" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4a\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4a\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4a\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4a\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4a\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4a\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4a\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",1" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4b" /ve /t REG_SZ /d "MPEG-4 Audio" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4b" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4b\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4b\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4b\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4b\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4b\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4b\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4b\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",1" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4r" /ve /t REG_SZ /d "MPEG-4 Audio" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4r" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4r\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4r\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4r\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4r\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4r\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4r\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m4r\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",1" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aac" /ve /t REG_SZ /d "MPEG-4 Audio" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aac" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aac\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aac\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aac\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aac\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aac\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aac\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aac\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",1" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mid" /ve /t REG_SZ /d "MIDI" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mid" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mid\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mid\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mid\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mid\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mid\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mid\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mid\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",18" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.midi" /ve /t REG_SZ /d "MIDI" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.midi" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.midi\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.midi\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.midi\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.midi\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.midi\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.midi\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.midi\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",18" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rmi" /ve /t REG_SZ /d "MIDI" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rmi" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rmi\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rmi\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rmi\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rmi\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rmi\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rmi\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.rmi\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",18" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mka" /ve /t REG_SZ /d "Matroska audio" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mka" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mka\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mka\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mka\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mka\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mka\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mka\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mka\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",19" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp3" /ve /t REG_SZ /d "MP3" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp3" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp3\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp3\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp3\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp3\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp3\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp3\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp3\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",22" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpa" /ve /t REG_SZ /d "MPEG audio" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpa" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpa\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpa\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpa\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpa\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpa\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpa\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpa\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",24" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp2" /ve /t REG_SZ /d "MPEG audio" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp2" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp2\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp2\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp2\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp2\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp2\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp2\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mp2\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",24" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m1a" /ve /t REG_SZ /d "MPEG audio" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m1a" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m1a\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m1a\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m1a\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m1a\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m1a\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m1a\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m1a\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",24" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2a" /ve /t REG_SZ /d "MPEG audio" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2a" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2a\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2a\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2a\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2a\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2a\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2a\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m2a\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",24" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpc" /ve /t REG_SZ /d "Musepack" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpc" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpc\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpc\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpc\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpc\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpc\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpc\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpc\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",25" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ofr" /ve /t REG_SZ /d "OptimFROG" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ofr" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ofr\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ofr\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ofr\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ofr\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ofr\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ofr\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ofr\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",27" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ofs" /ve /t REG_SZ /d "OptimFROG" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ofs" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ofs\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ofs\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ofs\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ofs\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ofs\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ofs\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ofs\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",27" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ogg" /ve /t REG_SZ /d "Ogg Vorbis" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ogg" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ogg\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ogg\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ogg\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ogg\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ogg\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ogg\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ogg\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",28" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.oga" /ve /t REG_SZ /d "Ogg Vorbis" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.oga" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.oga\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.oga\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.oga\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.oga\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.oga\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.oga\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.oga\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",28" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.opus" /ve /t REG_SZ /d "Opus Audio Codec" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.opus" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.opus\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.opus\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.opus\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.opus\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.opus\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.opus\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.opus\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ra" /ve /t REG_SZ /d "Real Audio" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ra" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ra\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ra\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ra\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ra\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ra\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ra\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.ra\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",31" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.tak" /ve /t REG_SZ /d "TAK" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.tak" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.tak\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.tak\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.tak\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.tak\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.tak\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.tak\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.tak\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.tta" /ve /t REG_SZ /d "True Audio" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.tta" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.tta\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.tta\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.tta\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.tta\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.tta\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.tta\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.tta\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",37" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wav" /ve /t REG_SZ /d "WAV" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wav" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wav\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wav\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wav\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wav\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wav\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wav\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wav\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",39" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wma" /ve /t REG_SZ /d "Windows Media Audio" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wma" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wma\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wma\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wma\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wma\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wma\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wma\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wma\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",41" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wv" /ve /t REG_SZ /d "WavPack" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wv" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wv\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wv\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wv\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wv\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wv\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wv\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wv\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",43" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aob" /ve /t REG_SZ /d "Other Audio" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aob" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aob\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aob\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aob\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aob\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aob\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aob\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.aob\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mlp" /ve /t REG_SZ /d "Other Audio" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mlp" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mlp\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mlp\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mlp\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mlp\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mlp\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mlp\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mlp\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.thd" /ve /t REG_SZ /d "Other Audio" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.thd" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.thd\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.thd\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.thd\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.thd\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.thd\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.thd\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.thd\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.asx" /ve /t REG_SZ /d "Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.asx" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.asx\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.asx\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.asx\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.asx\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.asx\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.asx\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.asx\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",30" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m3u" /ve /t REG_SZ /d "Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m3u" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m3u\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m3u\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m3u\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m3u\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m3u\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m3u\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m3u\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",30" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m3u8" /ve /t REG_SZ /d "Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m3u8" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m3u8\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m3u8\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m3u8\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m3u8\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m3u8\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m3u8\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.m3u8\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",30" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.pls" /ve /t REG_SZ /d "Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.pls" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.pls\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.pls\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.pls\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.pls\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.pls\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.pls\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.pls\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",30" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wvx" /ve /t REG_SZ /d "Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wvx" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wvx\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wvx\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wvx\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wvx\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wvx\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wvx\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wvx\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",30" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wax" /ve /t REG_SZ /d "Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wax" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wax\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wax\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wax\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wax\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wax\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wax\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wax\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",30" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wmx" /ve /t REG_SZ /d "Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wmx" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wmx\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wmx\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wmx\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wmx\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wmx\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wmx\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.wmx\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",30" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpcpl" /ve /t REG_SZ /d "Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpcpl" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpcpl\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpcpl\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpcpl\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpcpl\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpcpl\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpcpl\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpcpl\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",30" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpls" /ve /t REG_SZ /d "Blu-ray playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpls" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpls\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpls\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpls\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpls\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpls\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpls\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.mpls\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",30" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.bdmv" /ve /t REG_SZ /d "Blu-ray playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.bdmv" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.bdmv\shell\enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.bdmv\shell\enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.bdmv\shell\enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.bdmv\shell\open" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.bdmv\shell\open" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.bdmv\shell\open\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\mplayerc64.bdmv\DefaultIcon" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpciconlib.dll\",30" /f
Reg.exe add "HKLM\SOFTWARE\Classes\Directory\shell\mplayerc64.enqueue" /ve /t REG_SZ /d "Add to MPC-HC Playlist" /f
Reg.exe add "HKLM\SOFTWARE\Classes\Directory\shell\mplayerc64.enqueue" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\Directory\shell\mplayerc64.enqueue\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" /add \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\Directory\shell\mplayerc64.play" /ve /t REG_SZ /d "Play with MPC-HC" /f
Reg.exe add "HKLM\SOFTWARE\Classes\Directory\shell\mplayerc64.play" /v "Icon" /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\",0" /f
Reg.exe add "HKLM\SOFTWARE\Classes\Directory\shell\mplayerc64.play\command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" \"%%1\"" /f
Reg.exe add "HKLM\SOFTWARE\Classes\MediaPlayerClassic.Autorun\Shell\PlayVideoFiles\Command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" %%1" /f
Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers\Handlers\MPCPlayVideoFilesOnArrival" /v "Action" /t REG_SZ /d "Play Video" /f
Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers\Handlers\MPCPlayVideoFilesOnArrival" /v "Provider" /t REG_SZ /d "Media Player Classic" /f
Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers\Handlers\MPCPlayVideoFilesOnArrival" /v "InvokeProgID" /t REG_SZ /d "MediaPlayerClassic.Autorun" /f
Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers\Handlers\MPCPlayVideoFilesOnArrival" /v "InvokeVerb" /t REG_SZ /d "PlayVideoFiles" /f
Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers\Handlers\MPCPlayVideoFilesOnArrival" /v "DefaultIcon" /t REG_SZ /d "C:\Program Files\MPC-HC\mpc-hc64.exe,0" /f
Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers\EventHandlers\PlayVideoFilesOnArrival" /v "MPCPlayVideoFilesOnArrival" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\MediaPlayerClassic.Autorun\Shell\PlayDVDMovie\Command" /ve /t REG_SZ /d "\"C:\Program Files\MPC-HC\mpc-hc64.exe\" %%1 /dvd" /f
Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers\Handlers\MPCPlayDVDMovieOnArrival" /v "Action" /t REG_SZ /d "Play DVD Movie" /f
Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers\Handlers\MPCPlayDVDMovieOnArrival" /v "Provider" /t REG_SZ /d "Media Player Classic" /f
Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers\Handlers\MPCPlayDVDMovieOnArrival" /v "InvokeProgID" /t REG_SZ /d "MediaPlayerClassic.Autorun" /f
Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers\Handlers\MPCPlayDVDMovieOnArrival" /v "InvokeVerb" /t REG_SZ /d "PlayDVDMovie" /f
Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers\Handlers\MPCPlayDVDMovieOnArrival" /v "DefaultIcon" /t REG_SZ /d "C:\Program Files\MPC-HC\mpc-hc64.exe,0" /f
Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers\EventHandlers\PlayDVDMovieOnArrival" /v "MPCPlayDVDMovieOnArrival" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer" /v "GlobalAssocChangedCounter" /t REG_DWORD /d "250" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "HideCaptionMenu" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "HideNavigation" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "ControlState" /t REG_DWORD /d "19" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "DefaultVideoFrame" /t REG_DWORD /d "4" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "KeepAspectRatio" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "CompMonDeskARDiff" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "Volume" /t REG_DWORD /d "100" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "Balance" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "VolumeStep" /t REG_DWORD /d "5" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "SpeedStep" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "Mute" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "LoopNum" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "Loop" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "LoopMode" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "Zoom" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "AllowMultipleInstances" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "TitleBarTextStyle" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "TitleBarTextTitle" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "OnTop" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "TrayIcon" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "AutoZoom" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "AutoFitFactor" /t REG_DWORD /d "100" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "AfterPlayback" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "HideFullscreenControls" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "HideFullscreenControlsPolicy" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "HideFullscreenControlsDelay" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "HideFullscreenDockedPanels" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "HideWindowedControls" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "HideWindowedMousePointer" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "ExitFullscreenAtTheEnd" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "RememberWindowPos" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "RememberWindowSize" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "LastWindowRect" /t REG_BINARY /d "2804000035000000F00500006A030000" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "LastWindowType" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "SnapToDesktopEdges" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "AspectRatioX" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "AspectRatioY" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "KeepHistory" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "RecentFilesNumber" /t REG_DWORD /d "20" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "DSVidRen" /t REG_DWORD /d "11" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "RMVidRen" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "QTVidRen" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "ShufflePlaylistItems" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "RememberPlaylistItems" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "HidePlaylistFullScreen" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "APSurfaceUsage" /t REG_DWORD /d "2" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "DX9Resizer" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "VMR9MixerMode" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "VMRMixerYUV" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "VMRAlternateVSync" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "VMRVSyncOffset" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "VMRVSyncAccurate2" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "VMRFullscreenGUISupport" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "VMRVSync" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "VMRDisableDesktopComposition" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "VMRFullFloatingPointProcessing" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "VMRHalfFloatingPointProcessing" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "VMRColorManagementEnable" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "VMRColorManagementInput" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "VMRColorManagementAmbientLight" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "VMRColorManagementIntent" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "EVROutputRange" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "EVRHighColorRes" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "EVRForceInputHighColorRes" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "EVREnableFrameTimeCorrection" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "VMRFlushGPUBeforeVSync" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "VMRFlushGPUAfterPresent" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "VMRFlushGPUWait" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "SynchronizeClock" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "SynchronizeDisplay" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "SynchronizeNearest" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "LineDelta" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "ColumnDelta" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "CycleDelta" /t REG_BINARY /d "613255302AA9533F" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "TargetSyncOffset" /t REG_BINARY /d "0000000000002840" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "ControlLimit" /t REG_BINARY /d "0000000000000040" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "CacheShaders" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "ResetDevice" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "SPCSize" /t REG_DWORD /d "10" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "SPCMaxRes" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "DisableSubtitleAnimation" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "RenderAtWhenSubtitleAnimationIsDisabled" /t REG_DWORD /d "50" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "SubtitleAnimationRate" /t REG_DWORD /d "100" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "AllowDroppingSubpic" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "EVRBuffers" /t REG_DWORD /d "5" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "D3D9RenderDevice" /t REG_SZ /d "" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "AudioRendererType" /t REG_SZ /d "" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "AutoloadAudio" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "SubtitlesLanguageOrder" /t REG_SZ /d "" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "AudiosLanguageOrder" /t REG_SZ /d "" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "BlockVSFilter" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "EnableWorkerThreadForOpening" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "ReportFailedPins" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "DVDPath" /t REG_SZ /d "" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "UseDVDPath" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "MenuLang" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "AudioLang" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "SubtitlesLang" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "ClosedCaptions" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "SPDefaultStyle" /t REG_SZ /d "20;20;20;20;2;0;2.000000;2.000000;3.000000;3.000000;0xffffff;0x00ffff;0x000000;0x000000;0x00;0x00;0x00;0x80;1;Arial;18.000000;100.000000;100.000000;0.000000;700;0;0;0;0;0.000000;0.000000;0.000000;0.000000;1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "SPOverridePlacement" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "SPHorPos" /t REG_DWORD /d "50" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "SPVerPos" /t REG_DWORD /d "90" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "SubtitleARCompensation" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "SubDelayInterval" /t REG_DWORD /d "500" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "EnableSubtitles" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "PreferForcedDefaultSubtitles" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "PrioritizeExternalSubtitles" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "DisableInternalSubtitles" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "AllowOverridingExternalSplitterSubtitleChoice" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "AutoDownloadSubtitles" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "AutoDownloadScoreMovies" /t REG_DWORD /d "22" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "AutoDownloadScoreSeries" /t REG_DWORD /d "24" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "AutoDownloadSubtitlesExclude" /t REG_SZ /d "" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "AutoUploadSubtitles" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "PreferHearingImpairedSubtitles" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "SubtitlesProviders" /t REG_SZ /d "<|OpenSubtitles|||1|1|><|podnapisi|||0|0|><|SubDB|||0|0|><|Napisy24|||0|0|>" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "SubtitlePaths" /t REG_SZ /d ".;.\subtitles;.\subs" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "UseDefaultsubtitlesStyle" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "EnableAudioSwitcher" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "EnableAudioTimeShift" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "AudioTimeShift" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "DownSampleTo441" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "CustomChannelMapping" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "SpeakerToChannelMapping" /t REG_BINARY /d "010000000100000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000010000000200000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000010000000200000004000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000010000000200000000000000000000000400000008000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000010000000200000004000000000000000800000010000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000010000000200000004000000080000001000000020000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000010000000200000004000000080000001000000020000000400000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000010000000200000004000000080000001000000020000000400000008000000000000000000000000000000000000000000000000000000000000000000000000000000000000000010000000200000004000000080000001000000020000000400000008000000000010000000000000000000000000000000000000000000000000000000000000000000000000000010000000200000004000000080000001000000020000000400000008000000000010000000200000000000000000000000000000000000000000000000000000000000000000000010000000200000004000000080000001000000020000000400000008000000000010000000200000004000000000000000000000000000000000000000000000000000000000000010000000200000004000000080000001000000020000000400000008000000000010000000200000004000000080000000000000000000000000000000000000000000000000000010000000200000004000000080000001000000020000000400000008000000000010000000200000004000000080000001000000000000000000000000000000000000000000000010000000200000004000000080000001000000020000000400000008000000000010000000200000004000000080000001000000020000000000000000000000000000000000000010000000200000004000000080000001000000020000000400000008000000000010000000200000004000000080000001000000020000000400000000000000000000000000000010000000200000004000000080000001000000020000000400000008000000000010000000200000004000000080000001000000020000000400000008000000000000000000000010000000200000004000000080000001000000020000000400000008000000000010000000200000004000000080000001000000020000000400000008000000000010000000000010000000200000004000000080000001000000020000000400000008000000000010000000200000004000000080000001000000020000000400000008000000000010000000200" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "AudioNormalize" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "AudioMaxNormFactor" /t REG_DWORD /d "400" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "AudioNormalizeRecover" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "AudioBoost" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "SpeakerChannels" /t REG_DWORD /d "2" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "FullScreenMonitor" /t REG_SZ /d "" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "PreventMinimize" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "UseWin7TaskBar" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "UseSearchInFolder" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "UseTimeTooltip" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "TimeTooltipPosition" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "OSDSize" /t REG_DWORD /d "18" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "OSDFont" /t REG_SZ /d "Segoe UI" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "AssociatedWithIcon" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "LastOpenDir" /t REG_SZ /d "C:\\" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "D3DFullScreen" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "VideoBrightness" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "VideoContrast" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "VideoHue" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "VideoSaturation" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "ShowOSD" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "EnableEDLEditor" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "InterfaceLanguage" /t REG_DWORD /d "2057" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "FastSeek" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "FastSeekMethod" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "ShowChapters" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "LcdSupport" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "DefaultCapture" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "RememberDVDPos" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "RememberFilePos" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "RememberPosForLongerThan" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "RememberPosForAudioFiles" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "LastFullScreen" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "IntRealMedia" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "UseWinLirc" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "WinLircAddr" /t REG_SZ /d "127.0.0.1:8765" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "UseUICE" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "UICEAddr" /t REG_SZ /d "127.0.0.1:1234" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "UseGlobalMedia" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "JumpDistS" /t REG_DWORD /d "1000" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "JumpDistM" /t REG_DWORD /d "5000" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "JumpDistL" /t REG_DWORD /d "20000" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "LimitWindowProportions" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "LastUsedPage" /t REG_DWORD /d "10027" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "LogoFile" /t REG_SZ /d "" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "LogoID2" /t REG_DWORD /d "209" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "LogoExt" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "HideCDROMsSubMenu" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "Priority" /t REG_DWORD /d "32" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "LaunchFullScreen" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "EnableWebServer" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "WebServerPort" /t REG_DWORD /d "13579" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "WebServerUseCompression" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "WebServerLocalhostOnly" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "WebUIEnablePreview" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "WebServerPrintDebugIfo" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "WebRoot" /t REG_SZ /d "*./webroot" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "WebDefIndex" /t REG_SZ /d "index.html;index.php" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "WebServerCGI" /t REG_SZ /d "" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "SnapshotPath" /t REG_SZ /d "D:\CLOUD ONEDRIVE\OneDrive\Other\Pictures" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "SnapshotExt" /t REG_SZ /d ".jpg" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "ThumbRows" /t REG_DWORD /d "4" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "ThumbCols" /t REG_DWORD /d "4" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "ThumbWidth" /t REG_DWORD /d "1024" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "SubSaveExternalStyleFile" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "RemainingTime" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "HighPrecisionTimer" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "UpdaterAutoCheck" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "UpdaterDelay" /t REG_DWORD /d "7" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "NotifySkype" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "JpegQuality" /t REG_DWORD /d "90" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "EnableCoverArt" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "CoverArtSizeLimit" /t REG_DWORD /d "600" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "EnableLogging" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "UseLegacyToolbar" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "SubtitleRenderer" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings" /v "DefaultToolbarSize" /t REG_DWORD /d "24" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode" /v "Enable" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode" /v "ApplyDefaultModeAtFSExit" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode" /v "RestoreResAfterExit" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode" /v "Delay" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode0" /v "Checked" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode0" /v "FrameRateStart" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode0" /v "FrameRateStop" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode0" /v "AudioDelay" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode0" /v "BPP" /t REG_DWORD /d "32" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode0" /v "Freq" /t REG_DWORD /d "60" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode0" /v "SizeX" /t REG_DWORD /d "1920" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode0" /v "SizeY" /t REG_DWORD /d "1080" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode0" /v "Flags" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode1" /v "Checked" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode1" /v "FrameRateStart" /t REG_DWORD /d "23500000" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode1" /v "FrameRateStop" /t REG_DWORD /d "23981000" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode1" /v "AudioDelay" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode1" /v "BPP" /t REG_DWORD /d "32" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode1" /v "Freq" /t REG_DWORD /d "60" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode1" /v "SizeX" /t REG_DWORD /d "1920" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode1" /v "SizeY" /t REG_DWORD /d "1080" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode1" /v "Flags" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode2" /v "Checked" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode2" /v "FrameRateStart" /t REG_DWORD /d "23982000" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode2" /v "FrameRateStop" /t REG_DWORD /d "24499000" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode2" /v "AudioDelay" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode2" /v "BPP" /t REG_DWORD /d "32" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode2" /v "Freq" /t REG_DWORD /d "60" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode2" /v "SizeX" /t REG_DWORD /d "1920" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode2" /v "SizeY" /t REG_DWORD /d "1080" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode2" /v "Flags" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode3" /v "Checked" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode3" /v "FrameRateStart" /t REG_DWORD /d "24500000" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode3" /v "FrameRateStop" /t REG_DWORD /d "25499000" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode3" /v "AudioDelay" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode3" /v "BPP" /t REG_DWORD /d "32" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode3" /v "Freq" /t REG_DWORD /d "60" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode3" /v "SizeX" /t REG_DWORD /d "1920" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode3" /v "SizeY" /t REG_DWORD /d "1080" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode3" /v "Flags" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode4" /v "Checked" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode4" /v "FrameRateStart" /t REG_DWORD /d "29500000" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode4" /v "FrameRateStop" /t REG_DWORD /d "29981000" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode4" /v "AudioDelay" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode4" /v "BPP" /t REG_DWORD /d "32" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode4" /v "Freq" /t REG_DWORD /d "60" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode4" /v "SizeX" /t REG_DWORD /d "1920" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode4" /v "SizeY" /t REG_DWORD /d "1080" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode4" /v "Flags" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode5" /v "Checked" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode5" /v "FrameRateStart" /t REG_DWORD /d "29982000" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode5" /v "FrameRateStop" /t REG_DWORD /d "30499000" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode5" /v "AudioDelay" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode5" /v "BPP" /t REG_DWORD /d "32" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode5" /v "Freq" /t REG_DWORD /d "60" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode5" /v "SizeX" /t REG_DWORD /d "1920" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode5" /v "SizeY" /t REG_DWORD /d "1080" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode5" /v "Flags" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode6" /v "Checked" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode6" /v "FrameRateStart" /t REG_DWORD /d "49500000" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode6" /v "FrameRateStop" /t REG_DWORD /d "50499000" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode6" /v "AudioDelay" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode6" /v "BPP" /t REG_DWORD /d "32" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode6" /v "Freq" /t REG_DWORD /d "60" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode6" /v "SizeX" /t REG_DWORD /d "1920" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode6" /v "SizeY" /t REG_DWORD /d "1080" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode6" /v "Flags" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode7" /v "Checked" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode7" /v "FrameRateStart" /t REG_DWORD /d "59500000" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode7" /v "FrameRateStop" /t REG_DWORD /d "59945000" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode7" /v "AudioDelay" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode7" /v "BPP" /t REG_DWORD /d "32" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode7" /v "Freq" /t REG_DWORD /d "60" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode7" /v "SizeX" /t REG_DWORD /d "1920" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode7" /v "SizeY" /t REG_DWORD /d "1080" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode7" /v "Flags" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode8" /v "Checked" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode8" /v "FrameRateStart" /t REG_DWORD /d "59946000" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode8" /v "FrameRateStop" /t REG_DWORD /d "60499000" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode8" /v "AudioDelay" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode8" /v "BPP" /t REG_DWORD /d "32" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode8" /v "Freq" /t REG_DWORD /d "60" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode8" /v "SizeX" /t REG_DWORD /d "1920" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode8" /v "SizeY" /t REG_DWORD /d "1080" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\FullscreenAutoChangeMode\Mode8" /v "Flags" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Favorites" /v "RememberPosition" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Favorites" /v "RelativeDrive" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Capture" /v "VidDispName" /t REG_SZ /d "@device:pnp:\\?\usb#vid_04f2&pid_b424&mi_00#6&c1271f2&0&0000#{65e8773d-8f56-11d0-a3b9-00a0c9223196}\global" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Capture" /v "AudDispName" /t REG_SZ /d "" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Capture" /v "Country" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\DVBConfiguration" /v "BDANetworkProvider" /t REG_SZ /d "@device:sw:{71985F4B-1CA1-11D3-9CC8-00C04F7971E0}\Microsoft Network Provider" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\DVBConfiguration" /v "BDATuner" /t REG_SZ /d "" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\DVBConfiguration" /v "BDAReceiver" /t REG_SZ /d "" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\DVBConfiguration" /v "BDAScanFreqStart" /t REG_DWORD /d "474000" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\DVBConfiguration" /v "BDAScanFreqEnd" /t REG_DWORD /d "858000" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\DVBConfiguration" /v "BDABandWidth" /t REG_DWORD /d "8" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\DVBConfiguration" /v "BDAUseOffset" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\DVBConfiguration" /v "BDAOffset" /t REG_DWORD /d "166" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\DVBConfiguration" /v "BDAIgnoreEncryptedChannels" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\DVBConfiguration" /v "LastChannel" /t REG_DWORD /d "4294967295" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\DVBConfiguration" /v "RebuildFilterGraph" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\DVBConfiguration" /v "StopFilterGraph" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\PnSPresets" /v "Preset0" /t REG_SZ /d "Scale to 16:9 TV,0.500,0.500,1.000,1.333" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\PnSPresets" /v "Preset1" /t REG_SZ /d "Zoom To Widescreen,0.500,0.500,1.333,1.333" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Settings\PnSPresets" /v "Preset2" /t REG_SZ /d "Zoom To Ultra-Widescreen,0.500,0.500,1.763,1.763" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "RtspHandler" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "RtspFileExtFirst" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "avi" /t REG_SZ /d "avi \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "mpeg" /t REG_SZ /d "mpg mpeg mpe m1v m2v mpv2 mp2v pva evo m2p \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "mpegts" /t REG_SZ /d "ts tp trp m2t m2ts mts rec ssif \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "dvdvideo" /t REG_SZ /d "vob ifo \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "mkv" /t REG_SZ /d "mkv mk3d \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "webm" /t REG_SZ /d "webm \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "mp4" /t REG_SZ /d "mp4 m4v mp4v mpv4 hdmov \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "mov" /t REG_SZ /d "mov \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "3gp" /t REG_SZ /d "3gp 3gpp 3ga \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "3g2" /t REG_SZ /d "3g2 3gp2 \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "flv" /t REG_SZ /d "flv f4v \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "ogm" /t REG_SZ /d "ogm ogv \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "rm" /t REG_SZ /d "rm rmvb \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "rt" /t REG_SZ /d "rt ram rpm rmm rp smi smil \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "wmv" /t REG_SZ /d "wmv wmp wm asf \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "bink" /t REG_SZ /d "smk bik \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "flic" /t REG_SZ /d "fli flc flic \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "dsm" /t REG_SZ /d "dsm dsv dsa dss \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "ivf" /t REG_SZ /d "ivf \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "swf" /t REG_SZ /d "swf \3" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "other" /t REG_SZ /d "divx amv \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "ac3" /t REG_SZ /d "ac3 \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "dts" /t REG_SZ /d "dts dtshd dtsma \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "aiff" /t REG_SZ /d "aif aifc aiff \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "alac" /t REG_SZ /d "alac \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "amr" /t REG_SZ /d "amr \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "ape" /t REG_SZ /d "ape apl \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "au" /t REG_SZ /d "au snd \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "audiocd" /t REG_SZ /d "cda \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "flac" /t REG_SZ /d "flac \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "m4a" /t REG_SZ /d "m4a m4b m4r aac \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "midi" /t REG_SZ /d "mid midi rmi \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "mka" /t REG_SZ /d "mka \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "mp3" /t REG_SZ /d "mp3 \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "mpa" /t REG_SZ /d "mpa mp2 m1a m2a \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "mpc" /t REG_SZ /d "mpc \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "ofr" /t REG_SZ /d "ofr ofs \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "ogg" /t REG_SZ /d "ogg oga \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "opus" /t REG_SZ /d "opus \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "ra" /t REG_SZ /d "ra \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "tak" /t REG_SZ /d "tak \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "tta" /t REG_SZ /d "tta \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "wav" /t REG_SZ /d "wav \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "wma" /t REG_SZ /d "wma \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "wavpack" /t REG_SZ /d "wv \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "other_audio" /t REG_SZ /d "aob mlp thd \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "pls" /t REG_SZ /d "asx m3u m3u8 pls wvx wax wmx mpcpl \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "bdpls" /t REG_SZ /d "mpls bdmv \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\FileFormats" /v "rar" /t REG_SZ /d "rar \0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "SRC_CDDA" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "SRC_CDXA" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "SRC_VTS" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "SRC_FLIC" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "SRC_AC3" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "SRC_DTS" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "SRC_MATROSKA" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "SRC_HTTP" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "SRC_RTSP" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "SRC_UDP" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "SRC_RTP" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "SRC_MMS" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "SRC_RTMP" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "SRC_REALMEDIA" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "SRC_AVI" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "SRC_AVS" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "SRC_OGG" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "SRC_MPEG" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "SRC_MPA" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "SRC_DSM" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "SRC_SUBS" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "SRC_MP4" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "SRC_FLV" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "SRC_GIF" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "SRC_ASF" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "SRC_WTV" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "SRC_FLAC" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "SRC_RFS" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_MPEG1" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_MPEG2" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_RV" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_RA" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_MPA" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_DTS" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_LPCM" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_AC3" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_AAC" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_ALAC" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_ALS" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_PS2AUD" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_VORBIS" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_FLAC" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_NELLY" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_AMR" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_OPUS" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_PCM" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_H264" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_HEVC" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_VC1" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_FLV4" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_VP356" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_VP8" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_VP9" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_XVID" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_DIVX" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_MSMPEG4" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_WMV" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_SVQ3" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_H263" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_THEORA" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_AMVV" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_MJPEG" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_INDEO" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_SCREEN" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_FLIC" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_MSVIDEO" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters" /v "TRA_V210_V410" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Shaders" /v "Extra" /t REG_SZ /d "" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Shaders" /v "PreResize" /t REG_SZ /d "" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Shaders" /v "PostResize" /t REG_SZ /d "" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Shaders" /v "LastPreset" /t REG_SZ /d "" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters\Audio Renderer" /v "DeviceId" /t REG_SZ /d "" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters\Audio Renderer" /v "DeviceExclusive" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters\Audio Renderer" /v "DeviceBufferDuration" /t REG_DWORD /d "200" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters\Audio Renderer" /v "AllowBitstreaming" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters\Audio Renderer" /v "CrossfeedEnabled" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters\Audio Renderer" /v "CrossfeedCutoffFrequency" /t REG_DWORD /d "700" /f
Reg.exe add "HKCU\SOFTWARE\MPC-HC\MPC-HC\Internal Filters\Audio Renderer" /v "CrossfeedLevel" /t REG_DWORD /d "60" /f

:END_MEDIA_PLAYER_CLASSIC



:VERY_ENDER_BLOCK
:ENDE


PAUSE

