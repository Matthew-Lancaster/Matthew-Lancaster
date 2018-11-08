' ----------------------------------------------------------------
' __ MATTHEW LANCASTER 
' __ MATT.LAN@BTINTERNET.COM
' ----------------------------------------------------------------
' __ FROM ---- Sun 27-May-2018 07:20:00
' __ TO   ---- Sun 27-May-2018 08:30:00 __ 1 HOUR 10 MINUTE
' ----------------------------------------------------------------

'-----------------------------------------------------------------
' HIKVISION CLOUD GOOGLE DRIVE IMAGES SNAPSHOT 
' CONVERT WITH IrfanView TO MAKE GOOGLE ABLE TO DISPLAY
'-----------------------------------------------------------------

const DontWaitUntilFinished = false, ShowWindow = 1, DontShowWindow = 0, WaitUntilFinished = true

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
Set OBJFSO = CreateObject("Scripting.FileSystemObject")

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
Dim oFile
Dim oFSO

DIM ARRAY(2)

ARRAY(1)="A3"
ARRAY(2)="A4"

COUNT_1=0
COUNT_2=0

i_view32="""C:\Program Files (x86)\IrfanView\i_view32.exe"""


' -------------------------------------------------------------------
' FOR R_T=0 TO 30
' R_T --  0 = NOW TODAY
' R_T -- 30 = DAYS BEFORE
' R_T -- 10 = DAYS BEFORE
' R_T -- 20 = DAYS BEFORE
' -------------------------------------------------------------------

FOR R_T=0 TO 20

	i_TIME=""
	i_TIME=i_TIME+LPad(Year(Now-R_T), "0", 4)
	i_TIME=i_TIME+LPad(Month(Now-R_T), "0", 2)
	i_TIME=i_TIME+LPad(Day(Now-R_T), "0", 2)

	FOR R_I=1 TO 2

		FOLDER_1="D:\0 00 VIDEO SNAPSHOT CCSE HIKVISION\"+i_TIME+"\"+ARRAY(R_I)
		If fso.FolderExists(FOLDER_1) Then

			Set objFolder = objFSO.GetFolder(FOLDER_1)
			Set colFiles = objFolder.Files
			For Each objFile in colFiles

				Set oFSO = CreateObject("Scripting.FileSystemObject")
				Set oFile = oFSO.GetFile(objFile.PATH)

				FILE_1=objFile.PATH
				FILE_NAME=objFile.NAME
				IF MID(FILE_NAME,1,4)="DVR_" THEN 
				
					With oFile.OpenAsTextStream()
						readBinary = .Read(119+4)
						.Close
					End With
							
					IF MID(readBinary,119,4)=STRING(4,CHR(&H1E)) THEN 
						
						WSHShell.Run i_view32 + " """+FILE_1+""" /jpgq=88 /convert=$D$F" , DontShowWindow, WaitUntilFinished
						
						COUNT_1=COUNT_1+1
						
					END IF
					
					COUNT_2=COUNT_2+1
					
				END IF
			Next
		END IF
	Next
Next

Set WSHShell = Nothing

' MSGBOX "DONE __ " & COUNT_1 & " _OF_ " & COUNT_2

WScript.Quit(0)

Function LPad (str, pad, length)
    LPad = String(length - Len(str), pad) & str
End Function

'-----------------------------------------------
FUNCTION FOLDER_SCAN_COUNTER(fFolder,FOLDER_SCAN_COUNTER_VAR)
'-----------------------------------------------
	FILE_SCAN_COUNTER_VAR=0
	Set objFolder = objFSO.GetFolder(fFolder.Path)
    Set colFiles = objFolder.Files
	Dim oFile
	For Each objFile in colFiles

		FILE_NAME=objFile.NAME
		IF MID(FILE_NAME,1,4)="DVR_" THEN 
		
			'Set oFile = oFSO.GetFile(objFile.PATH)
			FOLDER_SCAN_COUNTER_VAR=FOLDER_SCAN_COUNTER_VAR+objFile.SIZE

		END IF
		
	Next
End FUNCTION


'--------------------------------------------------------------------
' SPARE CODE
'--------------------------------------------------------------------
' FullPath=FOLDER_2
' If Not fso.FolderExists(FullPath) Then
	' 'fso.CreateFolder FullPath
' End If
'--------------------------------------------------------------------

'--------------------------------------------------------------------
' Set objSuperFolder = FSO.GetFolder(FOLDER_1)
' CALL FOLDER_SCAN_COUNTER(objSuperFolder,FOLDER_SCAN_COUNTER_1)
' Set objSuperFolder = FSO.GetFolder(FOLDER_2)
' CALL FOLDER_SCAN_COUNTER(objSuperFolder,FOLDER_SCAN_COUNTER_2)

'MSGBOX FOLDER_SCAN_COUNTER_1 & VBCRLF & FOLDER_SCAN_COUNTER_2
'--------------------------------------------------------------------



'--------------------------------------------------------------------
' REF
'--------------------------------------------------------------------

' ----
' scripting - Help To create Folder1/Folder2 in Windows using VBScript ( Both the folders not exists before, i mean to create multilevel folders @ a strech.) - Stack Overflow
' https://stackoverflow.com/questions/4407386/help-to-create-folder1-folder2-in-windows-using-vbscript-both-the-folders-not
' --------
' vba - VBScript - How to make program wait until process has finished? - Stack Overflow
' https://stackoverflow.com/questions/10279404/vbscript-how-to-make-program-wait-until-process-has-finished
' --------
' Read and write binary file in VBscript - Stack Overflow
' https://stackoverflow.com/questions/6060529/read-and-write-binary-file-in-vbscript
' ----