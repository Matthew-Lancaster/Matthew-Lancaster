Option Explicit  
Dim fso, file, recentDate, recentFile, objStartFolder, rootFolder, fol, fil
' create a global copy of the filesystem object
Set fso = CreateObject("Scripting.FileSystemObject")
Set recentFile = Nothing

objStartFolder = "D:\0 00 Art Loggers\# APP AND SCREEN\7-ASUS-GL522VW\HOT_KEY_SCREEN_SHOT"

'Set fso = CreateObject("Scripting.FileSystemObject")
'Set rootFolder = fso.GetFolder(objStartFolder)

' get current folder
Set fol = fso.GetFolder(objStartFolder)

' go thru each files in the folder
REM For Each fil In fol.Files 
	REM For Each file in fil.Files
	  REM If (recentFile is Nothing) Then
		REM Set recentFile = file
	  REM ElseIf (file.DateLastModified > recentFile.DateLastModified) Then
		REM Set recentFile = file
	  REM End If
	REM Next
REM Next





' Call the RecurseFolders routine with name of function to be performed
' Takes one argument - in this case, the Path of the folder to be searched
RecurseFolders objStartFolder, "RenameIt"


' echo the job is completed
WScript.Echo "Completed!"

Sub RecurseFolders(sPath, funcName)
Dim folder

  With fso.GetFolder(sPath)
    if .SubFolders.Count > 0 Then
      For each folder in .SubFolders
		For Each file in folder.Files

         ' Perform function's operation
         'Execute funcName & " " & chr(34) & folder.Path & chr(34)

		  If (recentFile is Nothing) Then
		  Set recentFile = file
		  ElseIf (file.DateLastModified > recentFile.DateLastModified) Then
			Set recentFile = file
		  End If
		 

       Next
         ' Recurse to check for further subfolders
         RecurseFolders folder.Path, funcName
      Next
    End if
  End With

  
If recentFile is Nothing Then
  WScript.Echo "no recent files"
Else
  WScript.Echo "Recent file is " & recentFile.Name & " " & recentFile.DateLastModified
  Shell "Explorer.exe /SELECT, """ + LAST_CLIPBOARD_FILE + """", vbMaximizedFocus
	
  REM Set objShell = CreateObject("Wscript.Shell")
  REM strCommand = "Explorer /SELECT, """ + recentFile.Name + """"
  REM objShell.Run strCommand, 0, True
  
End If

  
End Sub