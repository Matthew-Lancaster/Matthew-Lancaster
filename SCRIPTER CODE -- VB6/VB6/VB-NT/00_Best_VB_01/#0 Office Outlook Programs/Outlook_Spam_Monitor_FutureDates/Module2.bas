Attribute VB_Name = "Module1"
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long


Sub Main()
tty$ = "VB_Future_Googles - Microsoft Visual Basic "
xx1 = FindWindow(vbNullString, tty$ + "[design]")
xx2 = FindWindow(vbNullString, tty$ + "[break]")
xx3 = FindWindow(vbNullString, tty$ + "[run]")

If xx1 > 0 Or xx2 > 0 Or xx3 > 0 Then End

Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" /Runexit ""D:\VB6\VB-NT\#OffIce Outlook Programs\Outlook_Spam_Monitor_FutureDates\OutLook Spam Copy an Del-It-ER Future Dates.vbp"""

End
End Sub
