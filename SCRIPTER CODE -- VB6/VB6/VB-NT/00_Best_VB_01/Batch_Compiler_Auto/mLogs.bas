Attribute VB_Name = "mLogs"
Public DaysToScan, AXx$

Public Sub AddAlertToLog(vTextBox As RichTextBox, vText As String, Optional vTime As Double)
  
  Dim S As String
  
  S = "*** ALERT (" & Now & ")" & vbCrLf
  S = S & "*** " & vText & vbCrLf
  If vTime > 0 Then
    S = S & "*** Average Compile Time: " & Format(vTime, "0.000") & " Seconds" & vbCrLf
  End If
  
  S = S & vbCrLf

S = S & AXx$ & vbCrLf
S = S & "*** Unique - VB Projects Compiled in the Last (" + Trim(Str(DaysToScan)) + ") Days = " & Str(ScanPath.ListView1.ListItems.Count) & vbCrLf
S = S & vbCrLf
  
  vTextBox.Text = S & vTextBox.Text
  
  DoEvents
  
  vTextBox.SelStart = Len(vTextBox)
End Sub

Public Sub AddAlertToLogSimple(vTextBox As RichTextBox)
  
  Dim S As String
  
S = "*** ALERT (" & Now & ")" & vbCrLf
S = S & AXx$ & vbCrLf
S = S & "*** Unique - VB Projects Compiled in the Last (" + Trim(Str(DaysToScan)) + ") Days = " & Str(ScanPath.ListView1.ListItems.Count) & vbCrLf
S = S & vbCrLf
  
'  txtLog = S & txtLog
  'txtLog.SelStart = 0
  
  vTextBox.Text = S & vTextBox.Text
  
  DoEvents
  
vTextBox.SelStart = Len(vTextBox)
End Sub

Public Sub CompileToLog(vTextBox As RichTextBox, vProject As String, vExe As String, vItemNumber As Integer, vTotalItems As Integer, Optional FromAuto As Boolean = False)
  
  Dim S As String
   
  If FromAuto Then
    S = "[CHANGE DETECTED - PROJECT COMPILED: " & vItemNumber & " of " & vTotalItems & "  (" & Now & ")]" & vbCrLf
  Else
    S = "[PROJECT COMPILED: " & vItemNumber & " of " & vTotalItems & "  (" & Now & ")]" & vbCrLf
  End If
  
  S = S & vProject & vbCrLf
  S = S & vExe & vbCrLf & vbCrLf
  
  vTextBox.Text = S & vTextBox.Text
  
  DoEvents
    vTextBox.SelStart = Len(vTextBox)
End Sub
