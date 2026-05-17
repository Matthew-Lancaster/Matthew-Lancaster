Attribute VB_Name = "Module1"

''make command$(7) for a dry run

''''
lastbackupnumber = 4
If Command$ <> "" Then lastbackupnumber = Val(Command$)
lastbackn = lastbackupnumber
lastbackup$ = Chr$(64 + lastbackn) ''''=H
lastb$ = lastbackup$


tti = 2000
ff$ = Dir$("c:\callerid\2double.txt")
If ff$ = "" Then Print "2double.txt DONT EXIST SORT IT OUT:": End
ff$ = Dir$("c:\callerid\2double.txt")


If ff$ <> "" Then
tt = 0
ttx = 0
Open "c:\callerid\2double.txt" For Binary As #1
Open "c:\callerid\2doublea.txt" For Binary As #2
zap = LOF(1)
For r = 1 To zap Step tti
If zap <> LOF(2) Then Exit For
On Error GoTo ip
ert = 0
tt$ = Input$(tti, 1)
On Error GoTo 0
If 1 = 2 Then
ip:
ert = 1
tt$ = Input$(zap - Seek(1), 1)
ttx$ = Input$(LOF(2) - Seek(2), 2)
Resume Next
End If
If ert = 0 Then ttx$ = Input$(tti, 2)
If tt$ <> ttx$ Then Exit For
rew = Int((r / zap) * 100) + 1
If rew = 101 Then rew = 100
If rew <> op Then
Color 15
LOCATE CSRLIN, 1: Print LTrim$(RTrim$(Str$(rew))); "%";
Color 14
Print " Scanned ";
op = rew
End If
Next

Color 15
LOCATE CSRLIN, 1: Print "100%";
Color 14
Print " Scanned ";

If tt$ <> ttx$ Then tt = 1: ttx = 2
If LOF(1) <> LOF(2) Then tt = 1: ttx = 2
Close #1, #2
Color 15
LOCATE CSRLIN, 1: Print "100%";
Color 14
Print " Scanned ";

End If
If Command$ <> "" Then tt = 1: ttx = 2
If tt <> ttx Then
tack = 1
Color 15
Print "File Contents changed Executing Backup"

ad$ = Dir$("c:\callerid\2double" + lastb$ + ".txt")
If ad$ <> "" Then
Color 14
Print "Delete Last backup ";
Color 15
Print "c:\callerid\2double" + lastb$ + ".txt";
Color 11
Kill "c:\callerid\2double" + lastb$ + ".txt"
End If
Color 14
Print " Renaming Backup History...."

For r = lastbackn To 2 Step -1
fad$ = Dir$("c:\callerid\2double" + Chr$(63 + r) + ".txt")
Top = 0
If fad$ <> "" Then
Color 11
P$ = "  "
RSet P$ = Format$(r, "#0")
Print P$;
Color 15
Print " ";
Print "c:\callerid\";
Color 12
Print "2double";
Color 11: Print Chr$(63 + r);
Color 7: Print ".txt";
Color 10
Print " to ";
Color 12
Print "2double";
Color 11: Print Chr$(64 + r);
Color 7: Print ".txt";
Top = 1
Name "c:\callerid\2double" + Chr$(63 + r) + ".txt" As "c:\callerid\2double" + Chr$(64 + r) + ".txt"
End If



If Top = 1 Then
Color 10
freefile1 = FreeFile
Open "c:\callerid\2double" + Chr$(64 + r) + ".txt" For Input As #freefile1
h = LOF(freefile1)
Close #freefile1

Print " Size";
Color 15
Print "= "; Format$(h, "###,###,###,###")
End If

Next
r = 1
Color 11
P$ = "  "
RSet P$ = Format$(r, "#0")
Print P$;
Color 10
Print " ";
Print "c:\callerid\";
Color 15
Print "2double .txt";
Color 10
Print " to ";
Color 7
Print "2double";
Color 11: Print Chr$(64 + r);
Color 7: Print ".txt";




freefile1 = FreeFile
Open "c:\callerid\2double.txt" For Input As #freefile1
h = LOF(freefile1)
Close #freefile1
Color 10
Print " Size";
Color 15
Print "= "; Format$(h, "###,###,###,###")

Shell "copy c:\callerid\2double.txt c:\callerid\2doublea.txt >nul"

End If

If tack = 0 Then
Color 15
Print "No Need For Sequenced Backup of 2double.txt.."
End If


