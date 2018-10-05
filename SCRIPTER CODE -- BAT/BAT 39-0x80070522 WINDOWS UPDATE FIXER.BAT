REM ----
REM Windows Update 1709 stops with error 0x80070522 at 71%, MCT as - Microsoft Community
REM https://answers.microsoft.com/en-us/windows/forum/windows_10-update/windows-update-1709-stops-with-error-0x80070522-at/e23b98e5-1bed-4875-814c-f0fbb739473c
REM ----

REM REMOVE COMPUTER FROM DOMAIN 
REM ALREADY DONE FOR MINE

@REM 3. Stopp the Windows Update Service in command shell executed as an administrator with
net stop wuauserv

REM DEL EVERYTHING THIS FOLDER
DEL C:\Windows\SoftwareDistribution\*.* /S

REM -----------------------------------------------------------------
REM STILL ERROR GETS TO 71% AND STOPS
REM -----------------------------------------------------------------

sfc /scannow

DISM.exe /Online /Cleanup-image /Restorehealth 




REM These culimative update were tried 
REM None worked
REM Except current version one
REM Stuck at
REM A look at windows update and how far it got and try next update in line
REM Failed that was kb4345418

REM kb4457127 Failed that was the 1st one to try the culmative last one for version 1607
REM Look like install okay but result history show nothing entered for that one I was 
REM out the door so not see if fail

REM kb4457141 Next one to try the 1703 next culmative but had error not correct computer as if to ahead of time and 1607 not done proper yet

REM kb4132216 Was little one that talk was required before load the kb4457127 a SSU and talker not required already present

REM kb4345418 Was the next in line after he last update that was received from the 1607 that failed to update
REM 2018-07 Cumulative Update for Windows 10 Version 1607 for x64-based Systems (KB4345418)

REM Time to look further into trouble shooter with this far not corrected yet

REM Next

REM KB4132649 Is the SSU Required to Run kb4345418 Small File
REM Lets Try
REM This is the 
REM 2018-05 Update for Windows 10 Version 1703 for x64-based Systems (KB4132649)
REM Maybe Ready for Next Step as Still in 1607
REM These Error Messenger
REM ---------------------
REM [Window Title]
REM Windows Update Standalone Installer

REM [Main Instruction]
REM Windows Update Standalone Installer

REM [Content]
REM The update is not applicable to your computer.

REM [OK]
REM ---------------------

REM Time to look further into trouble shooter with this far not corrector yet



REM "D:\windows10.0-kb4132216-x64_9cbeb1024166bdeceff90cd564714e1dcd01296e.msu"
REM "D:\windows10.0-kb4132649-x64_7b1001d846a7a14b752a7ea4c7dd444d4c05b39b.msu"
REM "D:\windows10.0-kb4457127-x64_53e4351d1e0233b3e55ed7fd70b6dca54b0ab878.msu"
REM "D:\windows10.0-kb4457141-x64_f0223c9e41536ad7ae26bf4219ae899ab7d33112.msu"
REM "D:\windows10.0-kb4345418-x64_a636a1048a2601d9218463107bc5fde558a1055e.msu"
