ECHO OFF
REM RECENT EDITOR AND FIRST CREATE
REM [ Thursday 15:59:20 Pm_11 May 2017 ] The UK
REM [ Thursday 14:59:20 Pm_11 May 2017 ] BST BDT DST DLS 1 Hr Ahead UTC+1 CET

REM ----------------------
REM ----------------------
REM DUAL DOCUMENT LOCATION
REM ----------------------
REM ----------------------
REM ______________________
REM __ 01 OF 02
REM D:\VB6\VB-NT\#__ THE SETUP OF VISUAL BASIC BEGIN FIRST TIME\START DO FIRST TIME INSTALL OF VISUAL BASIC.BAT
REM ______________________
REM __ 02 OF 02
REM C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 08-THE SETUP OF VISUAL BASIC BEGIN FIRST TIME __ 2 DLL REGSVR32__.BAT
REM ______________________
REM ----------------------
REM ----------------------

REM ----------------------------------------------
REM ----------------------------------------------
REM 01 OF 02 __ RunAS.dll
REM ----------------------------------------------
REM ----------------------------------------------
SET FILE_NAME_1="D:\VB6\VB-NT\00_Best_VB_01\Run_AS\RUN_AS.BAK\RunAS.dll"
SET FILE_NAME_2="C:\WINDOWS\SYSTEM32\RunAS.dll"

COPY %FILE_NAME_1% %FILE_NAME_2%

REGSVR32 /U /S %FILE_NAME_2%
REGSVR32 %FILE_NAME_2%

REM 'EXAMPLE -------------------------------------
REM 'HERE IS THE RUNAS CLASS
REM 'INCLUDE THE CLASS FOLLOW THE REM READ ME INFO AT BOTTOM LOCK
REM '-----------------------------
REM 'IN THE __.VBP FILE PROJECT
REM 'Reference=*\G{565783C6-CB41-11D1-8B02-00600806D9B6}#1.2#0#C:\Windows\SysWOW64\wbem\wbemdisp.TLB#Microsoft WMI Scripting V1.2 Library
REM Dim oRunas
REM Set oRunas = CreateObject("runas.clsrunas", GetComputerName) 'SERVER NAME
REM With oRunas
		REM .sDomain = "WORKGROUP"
		REM .sUserName = GetUserName
		REM .sPassword = " "
		REM .sCommand = "C:\Program Files\Siber Systems\GoodSync\GoodSync-v10.exe"
		REM .RunAs 'Call the Run As method
REM End With
REM Set oRunas = Nothing
REM '--------------------------------------------------------------------------------------------
REM 'Title: Windows 2000/XP Run As without having to Type a password class
REM 'Description: This class enhances the brillant Run as command which ships
REM 'with Windows 2000/XP Run As allows you to Run a program as any user
REM 'without the need to log on and off.
REM 'My RunAs class allows Network Admins to use Run as without having to type in a password this
REM 'means you can run a batch file using run as with no user input.
REM 'This class uses the API code from my forthcoming RunApp product the only different
REM 'is my final product will allow admins to create encrypted files which can be passed into the
REM 'excuteable so the username and password won't be exposed in a batch or VBScript file.
REM 'To use this class compile it on the server or client you wish to use it
REM 'on. Register it using regsvr32 "Path to compiled runas.dll" at the command prompt.
REM 'This file came from Planet-Source-Code.com...the home millions of lines of source code
REM 'You can view comments on this code/and or vote on it at: http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=33158&lngWId=1
REM '
REM 'The author may have retained certain copyrights to this code...please observe their request and the law by reviewing all copyright conditions at the above URL.
REM '--------------------------------------------------------------------------------------------
REM ----------------------------------------------
REM ----------------------------------------------
REM 01 OF 02 __ RunAS.dll __ END_ER
REM ----------------------------------------------
REM ----------------------------------------------

REM ----------------------------------------------
REM ----------------------------------------------
REM 02 OF 02 __ VB6IDEMouseWheelAddin.dll
REM -------------------------------------
REM -------------------------------------
SET FILE_NAME_1="D:\VB6\VB-NT\#__ MOUSE_WHEEL_FIX\VB6IDEMouseWheelAddin.dll"
SET FILE_NAME_2="C:\WINDOWS\SYSTEM32\VB6IDEMouseWheelAddin.dll"

COPY %FILE_NAME_1% %FILE_NAME_2%

REGSVR32 /U /S %FILE_NAME_2%
REGSVR32 %FILE_NAME_2%

REM ------------
REM ------------
REM READ ME TEXT
REM ------------
REM ------------
REM OPEN THE __.VPB 
REM LOOK AT TOP MENU FOR ADD IN
REM SEE IT THERE
REM SET TO LOAD ON START UP

REM FIND OUT SCROLLBAR ALSO WILL WORK WHEN HOOVER ON THEM MOUSE WHEEL
REM THE REGISTER THE DLL IS ALSO REQUIRED
REM D:\VB6\VB-NT\#__ MOUSE_WHEEL_FIX\VB6IDEMouseWheelAddin __ RUN ME.BAT
REM IF BATCH RUN OKAY IT WILL WAIT AT END RESULT
REM IF NOT AND QUICK POP GONE THEN ERROR ACCESS THE DLL LOCKED OR FOLDER

REM RELOCATE THE FOLDER
REM AND MAYBE ERROR MOUSE ADD-IN NOT FOUND WANT REMOVE
REM THAT IS OKAY
REM MOUSE WHEEL ADD-IN STILL THE
REM AND HAS TO BE SET AGAIN
REM ----------------------------------------------
REM ----------------------------------------------
REM 02 OF 02 __ VB6IDEMouseWheelAddin.dll__ END_ER
REM ----------------------------------------------
REM ----------------------------------------------
