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

IF '%1%'=='/ADMIN' (
	@REM :: Check for ADMIN Privileges Again
	>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
	@IF '%errorlevel%' GTR '2' (
	ECHO AN ATTEMPT WAS MADE TO SET ADMIN BUT NOT SUCCESSFUL
	ECHO YOUR USER MUST BE A STANDARD USER OR SOMETHING LIKE HER
	ECHO SAVES GET STUCK IN A LOOP ASKING OR ALSO DELAY REQUIRING
	ECHO END
	PAUSE
	EXIT /B 1
	)
@ECHO Got ADMIN Privileges __ Requested and Obtained
)
@REM :: Check for ADMIN Privileges
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
@IF '%errorlevel%' GTR '2' (
	@REM Get ADMIN Privileges
	@ECHO Set UAC = CreateObject^("Shell.Application"^) >> "%temp%\getadmin.vbs"
	@ECHO UAC.ShellExecute "%~s0", "/ADMIN", "", "runas", 1 >> "%temp%\getadmin.vbs"

	"%temp%\getadmin.vbs"
	@DEL "%temp%\getadmin.vbs"
	@GOTO VERY_ENDER_BLOCK
	@EXIT /B
)
@ECHO Got ADMIN Privileges

@SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

@CD \
REM ECHO ON

REM AS SET BY __ Stop resetting my apps
REM ----
REM Windows 10 default programs keep changing - Super User
REM https://superuser.com/questions/1022063/windows-10-default-programs-keep-changing
REM --------
REM Stop resetting my apps
REM http://www.carifred.com/stop_resetting_my_apps/
REM ----

REM THIS WILL STOP 
REM MICROSOFT EDGE __ FROM TAKING OVER
REM MICROSOFT MAIL
REM PHOTOS -- NOT ALLOWED 
REM MOVIES AND TV
REM GROOVE MUSIC
REM 3D BULDER -- NOT ALLOWED
REM NOTEPAD

Reg.exe delete "HKCU\Software\Classes\AppX4hxtad77fbk3jkkeerkrm0ze94wjf3s9" /v "NoOpenWith" /f
Reg.exe delete "HKCU\Software\Classes\AppX4hxtad77fbk3jkkeerkrm0ze94wjf3s9" /v "NoStaticDefaultVerb" /f
Reg.exe add "HKCU\Software\Classes\AppX4hxtad77fbk3jkkeerkrm0ze94wjf3s9" /f
Reg.exe delete "HKCU\Software\Classes\AppX7rm9drdg8sk7vqndwj3sdjw11x96jc0y" /v "NoOpenWith" /f
Reg.exe delete "HKCU\Software\Classes\AppX7rm9drdg8sk7vqndwj3sdjw11x96jc0y" /v "NoStaticDefaultVerb" /f
Reg.exe add "HKCU\Software\Classes\AppX7rm9drdg8sk7vqndwj3sdjw11x96jc0y" /f
Reg.exe delete "HKCU\Software\Classes\AppX90nv6nhay5n6a98fnetv7tpk64pp35es" /v "NoOpenWith" /f
Reg.exe delete "HKCU\Software\Classes\AppX90nv6nhay5n6a98fnetv7tpk64pp35es" /v "NoStaticDefaultVerb" /f
Reg.exe add "HKCU\Software\Classes\AppX90nv6nhay5n6a98fnetv7tpk64pp35es" /f
Reg.exe delete "HKCU\Software\Classes\AppXcc58vyzkbjbs4ky0mxrmxf8278rk9b3t" /v "NoOpenWith" /f
Reg.exe delete "HKCU\Software\Classes\AppXcc58vyzkbjbs4ky0mxrmxf8278rk9b3t" /v "NoStaticDefaultVerb" /f
Reg.exe add "HKCU\Software\Classes\AppXcc58vyzkbjbs4ky0mxrmxf8278rk9b3t" /f
Reg.exe delete "HKCU\Software\Classes\AppXd4nrz8ff68srnhf9t5a8sbjyar1cr723" /v "NoOpenWith" /f
Reg.exe delete "HKCU\Software\Classes\AppXd4nrz8ff68srnhf9t5a8sbjyar1cr723" /v "NoStaticDefaultVerb" /f
Reg.exe add "HKCU\Software\Classes\AppXd4nrz8ff68srnhf9t5a8sbjyar1cr723" /f
Reg.exe delete "HKCU\Software\Classes\AppXde74bfzw9j31bzhcvsrxsyjnhhbq66cs" /v "NoOpenWith" /f
Reg.exe delete "HKCU\Software\Classes\AppXde74bfzw9j31bzhcvsrxsyjnhhbq66cs" /v "NoStaticDefaultVerb" /f
Reg.exe add "HKCU\Software\Classes\AppXde74bfzw9j31bzhcvsrxsyjnhhbq66cs" /f
Reg.exe delete "HKCU\Software\Classes\AppXe862j7twqs4aww05211jaakwxyfjx4da" /v "NoOpenWith" /f
Reg.exe delete "HKCU\Software\Classes\AppXe862j7twqs4aww05211jaakwxyfjx4da" /v "NoStaticDefaultVerb" /f
Reg.exe add "HKCU\Software\Classes\AppXe862j7twqs4aww05211jaakwxyfjx4da" /f
Reg.exe delete "HKCU\Software\Classes\AppXq0fevzme2pys62n3e0fbqa7peapykr8v" /v "NoOpenWith" /f
Reg.exe delete "HKCU\Software\Classes\AppXq0fevzme2pys62n3e0fbqa7peapykr8v" /v "NoStaticDefaultVerb" /f
Reg.exe add "HKCU\Software\Classes\AppXq0fevzme2pys62n3e0fbqa7peapykr8v" /f
Reg.exe delete "HKCU\Software\Classes\AppX77ghgzrbzwe6djbdyty6hp5e3z1qryqx" /v "NoOpenWith" /f
Reg.exe delete "HKCU\Software\Classes\AppX77ghgzrbzwe6djbdyty6hp5e3z1qryqx" /v "NoStaticDefaultVerb" /f
Reg.exe add "HKCU\Software\Classes\AppX77ghgzrbzwe6djbdyty6hp5e3z1qryqx" /f
Reg.exe delete "HKCU\Software\Classes\AppXcr7ve1e6y7nx8n3y5qv7m7s9kzz18jx5" /v "NoOpenWith" /f
Reg.exe delete "HKCU\Software\Classes\AppXcr7ve1e6y7nx8n3y5qv7m7s9kzz18jx5" /v "NoStaticDefaultVerb" /f
Reg.exe add "HKCU\Software\Classes\AppXcr7ve1e6y7nx8n3y5qv7m7s9kzz18jx5" /f
Reg.exe delete "HKCU\Software\Classes\AppXdyw2j75xqd0y008p4bvtxhnxp3py2t2g" /v "NoOpenWith" /f
Reg.exe delete "HKCU\Software\Classes\AppXdyw2j75xqd0y008p4bvtxhnxp3py2t2g" /v "NoStaticDefaultVerb" /f
Reg.exe add "HKCU\Software\Classes\AppXdyw2j75xqd0y008p4bvtxhnxp3py2t2g" /f
Reg.exe delete "HKCU\Software\Classes\AppXjmgntwab0zvv6ga5dkkwc34b7s4r55e9" /v "NoOpenWith" /f
Reg.exe delete "HKCU\Software\Classes\AppXjmgntwab0zvv6ga5dkkwc34b7s4r55e9" /v "NoStaticDefaultVerb" /f
Reg.exe add "HKCU\Software\Classes\AppXjmgntwab0zvv6ga5dkkwc34b7s4r55e9" /f
Reg.exe delete "HKCU\Software\Classes\AppXkn1qcbm5sq3zpckpgc0sr39rbdbentq3" /v "NoOpenWith" /f
Reg.exe delete "HKCU\Software\Classes\AppXkn1qcbm5sq3zpckpgc0sr39rbdbentq3" /v "NoStaticDefaultVerb" /f
Reg.exe add "HKCU\Software\Classes\AppXkn1qcbm5sq3zpckpgc0sr39rbdbentq3" /f
Reg.exe delete "HKCU\Software\Classes\AppXszr29fkzg3f7qt34d8bsjx2k03d4rn2v" /v "NoOpenWith" /f
Reg.exe delete "HKCU\Software\Classes\AppXszr29fkzg3f7qt34d8bsjx2k03d4rn2v" /v "NoStaticDefaultVerb" /f
Reg.exe add "HKCU\Software\Classes\AppXszr29fkzg3f7qt34d8bsjx2k03d4rn2v" /f
Reg.exe delete "HKCU\Software\Classes\AppXvqhb9dhb3dg0mpfydfbz4qp875kpspza" /v "NoOpenWith" /f
Reg.exe delete "HKCU\Software\Classes\AppXvqhb9dhb3dg0mpfydfbz4qp875kpspza" /v "NoStaticDefaultVerb" /f
Reg.exe add "HKCU\Software\Classes\AppXvqhb9dhb3dg0mpfydfbz4qp875kpspza" /f
Reg.exe delete "HKCU\Software\Classes\AppXydk58wgm44se4b399557yyyj1w7mbmvd" /v "NoOpenWith" /f
Reg.exe delete "HKCU\Software\Classes\AppXydk58wgm44se4b399557yyyj1w7mbmvd" /v "NoStaticDefaultVerb" /f
Reg.exe add "HKCU\Software\Classes\AppXydk58wgm44se4b399557yyyj1w7mbmvd" /f
Reg.exe delete "HKCU\Software\Classes\AppX6eg8h5sxqq90pv53845wmnbewywdqq5h" /v "NoOpenWith" /f
Reg.exe delete "HKCU\Software\Classes\AppX6eg8h5sxqq90pv53845wmnbewywdqq5h" /v "NoStaticDefaultVerb" /f
Reg.exe add "HKCU\Software\Classes\AppX6eg8h5sxqq90pv53845wmnbewywdqq5h" /f
Reg.exe delete "HKCU\Software\Classes\AppX6w6n4f8xch1s3vzwf3af6bfe88qhxbza" /v "NoOpenWith" /f
Reg.exe delete "HKCU\Software\Classes\AppX6w6n4f8xch1s3vzwf3af6bfe88qhxbza" /v "NoStaticDefaultVerb" /f
Reg.exe add "HKCU\Software\Classes\AppX6w6n4f8xch1s3vzwf3af6bfe88qhxbza" /f
Reg.exe delete "HKCU\Software\Classes\AppXqj98qxeaynz6dv4459ayz6bnqxbyaqcs" /v "NoOpenWith" /f
Reg.exe delete "HKCU\Software\Classes\AppXqj98qxeaynz6dv4459ayz6bnqxbyaqcs" /v "NoStaticDefaultVerb" /f
Reg.exe add "HKCU\Software\Classes\AppXqj98qxeaynz6dv4459ayz6bnqxbyaqcs" /f
Reg.exe delete "HKCU\Software\Classes\AppXtggqqtcfspt6ks3fjzyfppwc05yxwtwy" /v "NoOpenWith" /f
Reg.exe delete "HKCU\Software\Classes\AppXtggqqtcfspt6ks3fjzyfppwc05yxwtwy" /v "NoStaticDefaultVerb" /f
Reg.exe add "HKCU\Software\Classes\AppXtggqqtcfspt6ks3fjzyfppwc05yxwtwy" /f
Reg.exe delete "HKCU\Software\Classes\Applications\notepad.exe" /v "NoOpenWith" /f
Reg.exe delete "HKCU\Software\Classes\Applications\notepad.exe" /v "NoStaticDefaultVerb" /f
Reg.exe add "HKCU\Software\Classes\Applications\notepad.exe" /f

IF "%1"=="" PAUSE

