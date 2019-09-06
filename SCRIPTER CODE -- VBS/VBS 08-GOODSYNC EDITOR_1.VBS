Dim fso

Set fso = CreateObject("Scripting.FileSystemObject")

'FR1 = FreeFile
'Open "C:\SCRIPTER\jobs-groups-options - Copy.tic" For Input As FR1
'TT_EDITOR = Input(LOF(FR1), FR1)
'Close #FR1

Const ForReading = 1, ForWriting = 2, ForAppending = 8

Set File_1 = fso.OpenTextFile("C:\SCRIPTER\jobs-groups-options - Copy.tic", 1)
TT_EDITOR = File_1.ReadAll

X = 1
Do
    T1 = InStr(X, TT_EDITOR, "|M:409")
    'F_21 = InStr(T1, TT_EDITOR, "|U=")
   ' F_22 = InStr(T1, TT_EDITOR, "|B:")
   ' F_23 = InStr(T1, TT_EDITOR, "|J:")
    IF T1>0 THEN 
		T2 = InStr(T1, TT_EDITOR, "|i=")
		N3 = Mid(TT_EDITOR, T1, (T2 - T1) + 3)
		MSGBOX N3,vbMsgBoxSetForeground 
		N4 = Replace(N3, "|", "")
		' COUNT THE PIPE TOTAL INBETWEEN
		If Len(N3) - 5 = Len(N4) Then
			REASSEMBLE_VAR = "|M:0:|U=0|B:0:|J:0|i="
			TT_EDITOR = Replace(TT_EDITOR, N3, REASSEMBLE_VAR)
			Counter = Counter + 1
		End If
    End If
    X = T1 + 1
Loop Until T1=0


'FR1 = FreeFile
'Open "C:\SCRIPTER\jobs-groups-options.tic" For Output As FR1
'Print #FR1, TT_EDITOR;
'Close #FR1

Set File_2 = fso.OpenTextFile("C:\SCRIPTER\jobs-groups-options.tic", 2)
File_2.Write TT_EDITOR

MsgBox "Done Total " & Counter & " Proecesses",vbMsgBoxSetForeground 