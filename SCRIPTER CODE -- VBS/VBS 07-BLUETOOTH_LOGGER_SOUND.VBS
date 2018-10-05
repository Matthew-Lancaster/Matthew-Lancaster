REM --------
REM NOT USED
REM --------

REM -----------
REM IS RUN FROM 
REM -----------
REM RUN FROM
REM C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 19-SCRIPT_TIMER_UTIL.ahk
REM EVERY HOUR
REM AND 
REM D:\VB6\VB-NT\00_Best_VB_01\BLUETOOTH SOUND\BLUETOOTH_LOGGER_SOUND.exe
REM -----------



RUN_NAME_1="D:\VB6\VB-NT\00_Best_VB_01\BLUETOOTH SOUND\BLUETOOTH_LOGGER_SOUND.exe"

RUN_NAME_2="C:\PStart\Progs\0_Nirsoft_Package\NirSoft\bluetoothview_Desc.VBS"

Set UAC = CreateObject("Shell.Application")

UAC.ShellExecute RUN_NAME_1

UAC.ShellExecute RUN_NAME_2


'UAC.ShellExecute RUN_NAME_1, "", "", "runas", 1

'UAC.ShellExecute RUN_NAME_2, "", "", "runas", 1
