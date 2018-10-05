
d:
cd \

echo here

for /f "tokens=2 delims=\" %%x in ('schtasks /query /fo list ^| findstr /C:"OneDrive Standalone Update Task"') do schtasks /delete /tn "%%x" /f

schtasks /Create /TN "\OneDrive Standalone Update Task" /TR "'%localappdata%\Microsoft\OneDrive\OneDriveStandaloneUpdater.exe'" /SC DAILY /MO 2 /ST 00:00:00 /F

for /f "tokens=2 delims=\" %%x in ('schtasks /query /fo list ^| findstr /C:"User_Feed_Synchronization"') do schtasks /delete /tn "%%x" /f

schtasks /Create /TN "\User_Feed_Synchronization" /TR "'C:\Windows\system32\msfeedssync.exe' sync" /SC DAILY /MO 1 /ST 00:00:00 /F

@SET "VALUE_RETURN="
@FOR /F "tokens=1,2 delims= " %%A IN ('REG.EXE QUERY HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run /S /F VideoDownloaderUltimate /T REG_SZ') DO @IF "%%B"=="REG_SZ" SET VALUE_RETURN=%%A
@if "%VALUE_RETURN%"=="" GOTO NOT_HERE
reg.exe delete "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" /v "%VALUE_RETURN%" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" /f
:NOT_HERE


@SET "VALUE_RETURN="
@FOR /F "tokens=1,2 delims= " %%A IN ('REG.EXE QUERY HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run /S /F ProductUpdater /T REG_SZ') DO @IF "%%B"=="REG_SZ" SET VALUE_RETURN=%%A
@if "%VALUE_RETURN%"=="" GOTO NOT_HERE
reg.exe delete "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run" /v "%VALUE_RETURN%" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run" /f
:NOT_HERE



pause