@REM ----------------------------------------------------------------
@REM __ MATTHEW LANCASTER 
@REM __ MATT.LAN@BTINTERNET.COM
@REM ----------------------------------------------------------------
@REM MY FIRST SET UP OF GITHUB 
@REM HAS TAKEN ALL-NIGHT NICE TO WAKE WITH A BRIGHT SUNNY DAY
@REM
@REM HERE THIS CODE PURELY AN EXAMPLE TO SHOW HOW YOU CAN PUT YOUR USERNAME AND PASSWORD IN TO 
@REM AUTO-UPDATER
@REM I READ IN THE ON-LINE HELP PAGES 
@REM THAT THIS METHOD DOES NOT STORE YOUR DETAILS OF PASSWORD IN ANY OF THE CONFIGURE FILE SET
@REM MANY PEOPLE WERE GOING ABOUT USING CLONE AS THE METHOD
@REM BUT SOMEBODY SAID WAS BETTER TO HAVE IT WITH A PULL COMMAND
@REM FETCH 
@REM I INCLUDE TIME NOT IN LOUSY FORMAT SO EACH UPDATED EVENT HAS A NICE TIME PRINT NEXT TO HER
@REM ----------------------------------------------------------------
@REM __ FROM ---- Fri 05-Oct-2018 01:37:29
@REM __ TO   ---- Fri 05-Oct-2018 13:04:00 __ 12 HOUR+-
@REM ----------------------------------------------------------------

ECHO DON'T RUN THIS ONE NOT FOR MY USE JUST AN EXAMPLE ONE TO GIVE AWAY
PAUSE
GOTO ENDER
EXIT 

@REM __ SET ADMIN A SAFE DOUBLE CHECKER WAY

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
	@EXIT /B
)
@ECHO Got ADMIN Privileges



@SET PATH_GIT="C:\Program Files\Git\bin"

CD /D "C:\SCRIPTOR GITHUB"

%PATH_GIT%\git init
%PATH_GIT%\git add .

@SET v=%DATE:~3,2%
@SET map=01-Jan;02-Feb;03-Mar;04-Apr;05-May;06-Jun;07-Jul;08-Aug;09-Sep;10-Oct;11-Nov;12-Dec
@CALL SET v=%%map:*%v%-=%%
@SET v=%v:;=&rem.%
@ECHO %DATE:~6,4%-%v%-%DATE:~0,2% -- %TIME:~0,8%

%PATH_GIT%\git commit -m "%DATE:~6,4%-%v%-%DATE:~0,2% -- %TIME:~0,8%"
%PATH_GIT%\git config user.email "matt.lan@btinternet.coml"
%PATH_GIT%\git config user.name "Matthew-Lancaster"
%PATH_GIT%\git pull https://<USERNAME:PASSWORD@github.com//matthew-lancaster/matthew-lancaster master
%PATH_GIT%\git push -u origin master

TIMEOUT /T 20

:ENDER
EXIT

@REM ----------------------------------------------------------------
@REM 
@REM THE FIRST PROBLEM I HAD TO GITHUB
@REM WAS MAKING A REPOSITORY ON THE WEB
@REM AND NOT CLONING IT TO THE DESKTOP
@REM WOULD OF MADE EASIER
@REM TRAWL ALONG THE DOCUMENTATION AGAIN WHEN ALL ELSE FAILED

@REM AND THEN FINE TUNING OF THIS FILE
@REM AND THAT WAS IT

@REM AND SYNC IS DONE FROM MY SCRIPT FOLDER VIA GOODSYNC

@REM A WELL WORTH WHILE ALLNIGHTER IT WAS TO I BEEN DREAMER OF THIS ONE FOR A LONG TIME
@REM 
@REM ----------------------------------------------------------------

