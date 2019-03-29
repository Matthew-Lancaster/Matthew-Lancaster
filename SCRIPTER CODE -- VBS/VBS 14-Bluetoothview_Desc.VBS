'--------------------------------------------------------------------
' MATT.LAN@BTINTERNET.COM
'--------------------------------------------------------------------
' Fri 22 Dec 2017 03:44:39----
' Tue 24 Apr 2018 09:08:00---- EDIT ON ERROR TRAP
'--------------------------------------------------------------------

'--------------------------------------------------------------------
' RUN FROM
'--------------------------------------------------------------------
' C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\
' Autokey -- 19-SCRIPT_TIMER_UTIL.ahk
' EVERY HOUR
' AND 
' D:\VB6\VB-NT\00_Best_VB_01\BLUETOOTH SOUND\BLUETOOTH_LOGGER_SOUND.exe
'--------------------------------------------------------------------



Set wshShell = CreateObject( "WScript.Shell" )
strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )

Dim FSO
Dim Fil_1
Dim Fil_3

Set FSO = CreateObject("scripting.filesystemobject")

REM ---------------------------------------
Fil_1_PATH="C:\PStart\Progs\0_Nirsoft_Package\NirSoft\BluetoothView_Desc.txt"
Set Fil_1 = FSO.GetFile (Fil_1_PATH)
Fil_1_SIZE=Fil_1.Size

Fil_2_PATH="C:\PStart\Progs\0_Nirsoft_Package\NirSoft\BluetoothView_Desc -- " & strComputerName & ".txt"

Fil_2_PATH_TRIM=MID(Fil_2_PATH,1,INSTRREV(Fil_2_PATH,"\"))

IF FSO.FileExists(Fil_2_PATH) THEN
	Set Fil_2 = FSO.GetFile (Fil_2_PATH)
	Fil_2_SIZE=Fil_2.Size
END IF

If Fil_1_SIZE > Fil_2_SIZE Then
	ERR.CLEAR 
	ON ERROR RESUME NEXT
	FSO.CopyFile Fil_1_PATH, Fil_2_PATH
	IF ERR.NUMBER>0 THEN 
		' MAYBE AN ACCESS ERROR CLASH
		' WScript.Quit(0)
	END IF
End If

REM ---------------------------------------
Fil_1_PATH="C:\PStart\Progs\0_Nirsoft_Package\NirSoft\BluetoothView_Desc -- " & strComputerName & ".txt"
Set Fil_1 = FSO.GetFile (Fil_1_PATH)
Fil_1_SIZE=Fil_1.Size

Fil_2_PATH="C:\0 00 LOGGERS TEXT\BlueToothView Logger\BluetoothView_Desc -- " & strComputerName & ".txt"

Fil_2_PATH_TRIM=MID(Fil_2_PATH,1,INSTRREV(Fil_2_PATH,"\"))

IF FSO.FileExists(Fil_2_PATH) THEN
	Set Fil_2 = FSO.GetFile (Fil_2_PATH)
	Fil_2_SIZE=Fil_2.Size
END IF

If Fil_1_SIZE > Fil_2_SIZE Then
	ERR.CLEAR 
	ON ERROR RESUME NEXT
	FSO.CopyFile Fil_1_PATH, Fil_2_PATH
	IF ERR.NUMBER>0 THEN 
		' MAYBE AN ACCESS ERROR CLASH
		' WScript.Quit(0)
	END IF
End If

REM ---------------------------------------

Fil_1_PATH="C:\PStart\Progs\0_Nirsoft_Package\NirSoft\BluetoothLogView.cfg"
Set Fil_1 = FSO.GetFile (Fil_1_PATH)
Fil_1_SIZE=Fil_1.Size

Fil_2_PATH="C:\PStart\Progs\0_Nirsoft_Package\NirSoft\BluetoothLogView -- " & strComputerName & ".txt"

Fil_2_PATH_TRIM=MID(Fil_2_PATH,1,INSTRREV(Fil_2_PATH,"\"))

IF FSO.FileExists(Fil_2_PATH) THEN
	Set Fil_2 = FSO.GetFile (Fil_2_PATH)
	Fil_2_SIZE=Fil_2.Size
END IF

If Fil_1_SIZE > Fil_2_SIZE Then
	ERR.CLEAR 
	ON ERROR RESUME NEXT
	FSO.CopyFile Fil_1_PATH, Fil_2_PATH
	IF ERR.NUMBER>0 THEN 
		' MAYBE AN ACCESS ERROR CLASH
		' WScript.Quit(0)
	END IF
End If
