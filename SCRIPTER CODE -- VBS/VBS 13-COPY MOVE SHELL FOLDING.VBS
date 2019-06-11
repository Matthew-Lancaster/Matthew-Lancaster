' -----------------------------------------------------------------
' __ CODE BY MATTHEW LANCASTER
' __ DATE BEGINNER 5 AM -- 5 HOURS
' __ DATE BEGINNER AND ENDER --------------------------------------
' [ Saturday 10:18:00 Am_30 December 2017 ] The UK
' [ Saturday 10:18:00 Am_30 December 2017 ] GMT UTC+0 WET'-----------------------------------------------------------------
' ADDITIONAL 3 URL'S OPEN OF HELPING
' -----------------------------------------------------------------
' THIS WILL COPY SOURCE TO TARGET BY FOLDER AND OPTION DELETE AFTER 

' FOR
' DESKTOP
' SENDTO

DIM DELETE_AFTER
DIM objFSO
Dim objShell
DIM ComputerName
DIM obj_SOURCE_Folder
DIM obj_TARGET_Folder
DIM Version_OS

Set objShell = Wscript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
WindowsDir = objShell.ExpandEnvironmentStrings("%windir%")
AppDataLocation=objShell.ExpandEnvironmentStrings("%APPDATA%")
ComputerName=objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")

CALL GET_WINDOWS_VERSION

IF Version_OS=>6 THEN 
	'Wscript.Echo "begin."
	CALL USERS_DESKTOP
	CALL USERS_SEND_TO
	'CALL USERS_START_MENU
	'CALL USERS_PROGRAM_DATA_START_MENU
	'Wscript.Echo "end."
END IF

WScript.Quit 0

'-------------------------------------
Sub USERS_DESKTOP
'-------------------------------------
FOR R_X=0 TO 5
	DELETE_AFTER=TRUE
	XX=CSTR(R_X)
	IF LEN(XX)=1 THEN XX="0"&XX
	IF R_X =0 THEN obj_SOURCE_Folder = "C:\Users\Public\Desktop"
	IF R_X >0 THEN obj_SOURCE_Folder = "C:\Users\MATT "&XX&"\Desktop"
	obj_TARGET_Folder = "E:\01 Desktop\#_"&ComputerName&"\"

	SET_GO=TRUE
	If objFSO.FOLDERExists(obj_SOURCE_Folder) = FALSE Then SET_GO=FALSE
	If objFSO.FOLDERExists(obj_TARGET_Folder) = FALSE Then SET_GO=FALSE

	IF SET_GO=TRUE THEN 
		Set objSuperFolder = objFSO.GetFolder(obj_SOURCE_Folder)
		Call ShowSubfolders (objSuperFolder)
	END IF
NEXT
END SUB

'-------------------------------------
Sub USERS_SEND_TO
'-------------------------------------
FOR R_X=1 TO 5
	DELETE_AFTER=TRUE
	XX=CSTR(R_X)
	IF LEN(XX)=1 THEN XX="0"&XX
	IF R_X >0 THEN obj_SOURCE_Folder = "C:\Users\MATT "&XX&"\AppData\Roaming\Microsoft\Windows\SendTo"
	obj_TARGET_Folder = "E:\01 SendTo\#_"&ComputerName&"\"

	SET_GO=TRUE
	If objFSO.FOLDERExists(obj_SOURCE_Folder) = FALSE Then SET_GO=FALSE
	If objFSO.FOLDERExists(obj_TARGET_Folder) = FALSE Then SET_GO=FALSE

	IF SET_GO=TRUE THEN 
		Set objSuperFolder = objFSO.GetFolder(obj_SOURCE_Folder)
		Call ShowSubfolders (objSuperFolder)
	END IF
NEXT
END SUB

'-------------------------------------
Sub USERS_START_MENU
'-------------------------------------
FOR R_X=1 TO 5
	DELETE_AFTER=FALSE
	XX=CSTR(R_X)
	IF LEN(XX)=1 THEN XX="0"&XX
	IF R_X >0 THEN obj_SOURCE_Folder = "C:\Users\MATT "&XX&"\AppData\Roaming\Microsoft\Windows\Start Menu"
	obj_TARGET_Folder = "E:\01 Start Menu\#_"&ComputerName&"\"

	SET_GO=TRUE
	If objFSO.FOLDERExists(obj_SOURCE_Folder) = FALSE Then SET_GO=FALSE
	If objFSO.FOLDERExists(obj_TARGET_Folder) = FALSE Then SET_GO=FALSE

	IF SET_GO=TRUE THEN 
		Set objSuperFolder = objFSO.GetFolder(obj_SOURCE_Folder)
		Call ShowSubfolders (objSuperFolder)
	END IF
NEXT
END SUB

'-------------------------------------
Sub USERS_PROGRAM_DATA_START_MENU
'-------------------------------------
	DELETE_AFTER=TRUE
	obj_SOURCE_Folder = "C:\ProgramData\Microsoft\Windows\Start Menu"
	obj_TARGET_Folder = "E:\01 Start Menu\#_"&ComputerName&"\"

	SET_GO=TRUE
	If objFSO.FOLDERExists(obj_SOURCE_Folder) = FALSE Then SET_GO=FALSE
	If objFSO.FOLDERExists(obj_TARGET_Folder) = FALSE Then SET_GO=FALSE

	IF SET_GO=TRUE THEN 
		Set objSuperFolder = objFSO.GetFolder(obj_SOURCE_Folder)
		Call ShowSubfolders (objSuperFolder)
	END IF
END SUB

'-------------------------------------
Sub ShowSubFolders(fFolder)
'-------------------------------------
    Set objFolder = objFSO.GetFolder(fFolder.Path)
    Set colFiles = objFolder.Files
    For Each objFile in colFiles
		SET_GO=TRUE
		If INSTR(objFolder.name, "_gsdata_")>0 Then SET_GO=FALSE
		If INSTR(objFile.name, "desktop.ini")>0 Then SET_GO=FALSE
		IF SET_GO=TRUE THEN 
			ERR.CLEAR
			ON ERROR RESUME NEXT
			objFSO.CopyFile objFile.PATH, obj_TARGET_Folder
			IF DELETE_AFTER=TRUE AND ERR.NUMBER=0 THEN 
				objFSO.DELETEFILE objFile.PATH
			END IF
		End If
    Next
	
    For Each Subfolder in fFolder.SubFolders
        ShowSubFolders(Subfolder)
    Next
End Sub

'--------------------------------------------------------
SUB GET_WINDOWS_VERSION
'--------------------------------------------------------
	Version_OS=0
	ON ERROR RESUME NEXT
	Set SystemSet = GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem") 
	for each System in SystemSet 
	' WScript.Echo System.Caption 
	' WScript.Echo System.Manufacturer 
	' WScript.Echo System.BuildType 
	' WScript.Echo " Version: " + System.Version 
	' WScript.Echo " Locale: " + System.Locale 
	' WScript.Echo " Windows Directory: " + System.WindowsDirectory 
	' WScript.Echo " Total memory: " + System.TotalVisibleMemorySize + " bytes" 
	' WScript.Echo " Serial Number: " + System.SerialNumber 
	' Wscript.Echo "" 
	Version_OS = CINT(MID(System.Version,1,INSTR(System.Version,".")-1))
	'WINDOWS 7 = 6.1.7601
	next 
	IF ERR.NUMBER>0 THEN Version_OS=0
	' CATCH THE ERROR
	' ---------------------------
	' Windows Script Host
	' ---------------------------
	' Script:	C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 13-COPY MOVE SHELL FOLDING.VBS
	' Line:	150
	' Char:	2
	' Error:	ActiveX component can't create object
	' Code:	800A01AD
	' Source: 	Microsoft VBScript runtime error
	' ---------------------------
END SUB


'-------------------------------------
' JUNK 
'-------------------------------------
'Set objSuperFolder = objFSO.GetFolder(WScript.Arguments(0))
