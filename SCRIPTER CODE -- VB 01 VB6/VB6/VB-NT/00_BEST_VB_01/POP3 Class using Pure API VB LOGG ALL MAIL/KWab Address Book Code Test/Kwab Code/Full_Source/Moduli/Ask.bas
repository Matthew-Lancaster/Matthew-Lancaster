Attribute VB_Name = "modAsk"
Option Explicit

Public Function AskLongLong(inDefault As LONGLONG, ByRef Out As LONGLONG) As Boolean
   Dim strVal As String, boolOk As Boolean
   Dim Word1 As Long, Word2 As Long, k As Integer

   AskLongLong = False
   
   For k = 1 To 2
      strVal = CStr(IIf(k = 1, inDefault.Word1, inDefault.Word2))
      Do
         boolOk = True
         Do
            strVal = Trim(InputBox("Enter new value" & vbNewLine & "Word #" & k, "Update property value", strVal))
            If Len(strVal) = 0 Then Exit Function
         Loop Until IsNumeric(strVal)
         On Error Resume Next
         If k = 1 Then
            Out.Word1 = CLng(strVal)
         Else
            Out.Word2 = CLng(strVal)
         End If
         If Err <> 0 Then
            If MsgBox(Err.Description, vbExclamation Or vbRetryCancel) <> vbRetry Then
               Err.Clear: On Error GoTo 0
               Exit Function
            End If
            boolOk = False
            Err.Clear: On Error GoTo 0
         End If
      Loop Until boolOk
   Next
   AskLongLong = True
End Function

Public Function AskLong(ByVal inDefault As Long, ByRef Out As Long) As Boolean
   Dim strVal As String, boolOk As Boolean

   AskLong = False
   strVal = CStr(inDefault)
   Do
      boolOk = True
      Do
         strVal = Trim(InputBox("Enter new value", "Update property value", strVal))
         If Len(strVal) = 0 Then Exit Function
      Loop Until IsNumeric(strVal)
      On Error Resume Next
      Out = CLng(strVal)
      If Err <> 0 Then
         If MsgBox(Err.Description, vbExclamation Or vbRetryCancel) <> vbRetry Then
            Err.Clear: On Error GoTo 0
            Exit Function
         End If
         boolOk = False
         Err.Clear: On Error GoTo 0
      End If
   Loop Until boolOk
   AskLong = True
End Function

Public Function AskFloat(ByVal inDefault As Single, ByRef Out As Single) As Boolean
   Dim strVal As String, boolOk As Boolean
   Dim v As String
   
   v = Mid(CStr(1.2), 2, 1)

   AskFloat = False
   strVal = CStr(inDefault)
   Do
      boolOk = True
      Do
         strVal = Trim(InputBox("Enter new value", "Update property value", strVal))
         If Len(strVal) = 0 Then Exit Function
         strVal = Replace(Replace(strVal, ".", v), ",", v)
      Loop Until IsNumeric(strVal)
      On Error Resume Next
      Out = CSng(strVal)
      If Err <> 0 Then
         If MsgBox(Err.Description, vbExclamation Or vbRetryCancel) <> vbRetry Then
            Err.Clear: On Error GoTo 0
            Exit Function
         End If
         boolOk = False
         Err.Clear: On Error GoTo 0
      End If
   Loop Until boolOk
   AskFloat = True
End Function

Public Function AskDouble(ByVal inDefault As Double, ByRef Out As Double) As Boolean
   Dim strVal As String, boolOk As Boolean
   Dim v As String
   
   v = Mid(CStr(1.2), 2, 1)

   AskDouble = False
   strVal = CStr(inDefault)
   Do
      boolOk = True
      Do
         strVal = Trim(InputBox("Enter new value", "Update property value", strVal))
         If Len(strVal) = 0 Then Exit Function
         strVal = Replace(Replace(strVal, ".", v), ",", v)
      Loop Until IsNumeric(strVal)
      On Error Resume Next
      Out = CDbl(strVal)
      If Err <> 0 Then
         If MsgBox(Err.Description, vbExclamation Or vbRetryCancel) <> vbRetry Then
            Err.Clear: On Error GoTo 0
            Exit Function
         End If
         boolOk = False
         Err.Clear: On Error GoTo 0
      End If
   Loop Until boolOk
   AskDouble = True
End Function

Public Function AskBoolean(ByVal inDefault As Boolean, ByRef Out As Boolean) As Boolean
   Dim r As VbMsgBoxResult

   r = MsgBox("Select new value:" & vbNewLine & "YES: true" & vbNewLine & "NO: false" & vbNewLine & "CANCEL: abort operation", vbQuestion Or vbYesNoCancel)
   AskBoolean = (r <> vbCancel)
   Out = (r <> vbNo)
End Function

Public Function AskCurrency(ByVal inDefault As Currency, ByRef Out As Currency) As Boolean
   Dim strVal As String, boolOk As Boolean
   Dim v As String
   
   v = Mid(CStr(1.2), 2, 1)

   AskCurrency = False
   strVal = CStr(inDefault)
   Do
      boolOk = True
      Do
         strVal = Trim(InputBox("Enter new value", "Update property value", strVal))
         If Len(strVal) = 0 Then Exit Function
         strVal = Replace(Replace(strVal, ".", v), ",", v)
      Loop Until IsNumeric(strVal)
      On Error Resume Next
      Out = CCur(strVal)
      If Err <> 0 Then
         If MsgBox(Err.Description, vbExclamation Or vbRetryCancel) <> vbRetry Then
            Err.Clear: On Error GoTo 0
            Exit Function
         End If
         boolOk = False
         Err.Clear: On Error GoTo 0
      End If
   Loop Until boolOk
   AskCurrency = True
End Function

Public Function AskCLSID(inDefault As CLSID, ByRef Out As CLSID) As Boolean
   Dim strVal As String, boolOk As Boolean

   AskCLSID = False
   strVal = CLSID2String(inDefault)
   Do
      boolOk = True
      Do
         strVal = Trim(InputBox("Enter new value" & vbNewLine & "(respect string structure)", "Update property value", strVal))
         If Len(strVal) = 0 Then Exit Function
         
      Loop Until IsCLSID(strVal)
      On Error Resume Next
      Out = CCLSID(strVal)
      If Err <> 0 Then
         If MsgBox(Err.Description, vbExclamation Or vbRetryCancel) <> vbRetry Then
            Err.Clear: On Error GoTo 0
            Exit Function
         End If
         boolOk = False
         Err.Clear: On Error GoTo 0
      End If
   Loop Until boolOk
   AskCLSID = True
End Function

Public Function AskShort(ByVal inDefault As Integer, ByRef Out As Integer) As Boolean
   Dim strVal As String, boolOk As Boolean

   AskShort = False
   strVal = CStr(inDefault)
   Do
      boolOk = True
      Do
         strVal = Trim(InputBox("Enter new value", "Update property value", strVal))
         If Len(strVal) = 0 Then Exit Function
      Loop Until IsNumeric(strVal)
      On Error Resume Next
      Out = CInt(strVal)
      If Err <> 0 Then
         If MsgBox(Err.Description, vbExclamation Or vbRetryCancel) <> vbRetry Then
            Err.Clear: On Error GoTo 0
            Exit Function
         End If
         boolOk = False
         Err.Clear: On Error GoTo 0
      End If
   Loop Until boolOk
   AskShort = True
End Function

Public Function AskString(ByVal inDefault As String, ByRef Out As String) As Boolean
   Dim strVal As String, NEWEMAIL

   AskString = False
   strVal = InputBox("Enter new value", "Update property value", inDefault)
   NEWEMAIL = "HELLO3"
   If Len(strVal) = 0 Then
      If MsgBox("Set the value to a zero-length string?", vbQuestion Or vbYesNo Or vbDefaultButton2) <> vbYes Then Exit Function
   End If
   Out = strVal
   AskString = True
End Function

Public Function AskSysTime(inDefault As SysTime, ByRef Out As SysTime) As Boolean
   Dim dof As Integer, dy As Integer, mh As Integer, yr As Integer, hr As Integer, mn As Integer, ss As Integer, ms As Integer
   Dim frm As frmAskDateTime
   
   AskSysTime = False
   If Not ExpandSysTime(inDefault, yr, mh, dof, dy, hr, mn, ss, ms) Then Exit Function
   Set frm = New frmAskDateTime
   With frm
      .inUseMilliseconds = True
      .inDay = dy
      .inMonth = mh
      .inYear = yr
      .inHour = hr
      .inMinute = mn
      .inSecond = ss
      .inMillisecond = ms
      .Show vbModal
      If Not .outOk Then
         Unload frm
         Exit Function
      End If
      dy = .outDay
      mh = .outMonth
      yr = .outYear
      hr = .outHour
      mn = .outMinute
      ss = .outSecond
      ms = .outMillisecond
   End With
   Unload frm
   AskSysTime = True
End Function

Public Function AskAppTime(ByVal inDefault As Date, ByRef Out As Date) As Boolean
   Dim frm As frmAskDateTime
   
   AskAppTime = False
   Set frm = New frmAskDateTime
   With frm
      .inUseMilliseconds = False
      .inDay = Day(inDefault)
      .inMonth = Month(inDefault)
      .inYear = Year(inDefault)
      .inHour = Hour(inDefault)
      .inMinute = Minute(inDefault)
      .inSecond = Second(inDefault)
      .Show vbModal
      If Not .outOk Then
         Unload frm
         Exit Function
      End If
      Out = DateSerial(.outYear, .outMonth, .outDay) + TimeSerial(.outHour, .outMinute, .outSecond)
   End With
   Unload frm
   AskAppTime = True
End Function
Public Function AskBinary(inDefault() As Byte, ByRef Out() As Byte) As Boolean
   Dim strVal As String, lb As Long, ub As Long, k As Long, ItsOk As Boolean
   Dim baVal As Variant

   AskBinary = False
   strVal = ""
   On Error Resume Next
   lb = LBound(inDefault)
   ub = UBound(inDefault)
   If Err <> 0 Then
      Err.Clear
      On Error GoTo 0
   Else
      On Error GoTo 0
      For k = lb To ub
         If k > lb Then strVal = strVal & " "
         strVal = strVal & Right("00" & Hex(inDefault(k)), 2)
      Next
   End If
   Do
      strVal = Trim(InputBox("Enter new value" & vbNewLine & "(separate each byte with a space)", "Update property value", strVal))
      If Len(strVal) = 0 Then
         If MsgBox("Set the value to a zero-length byte array?", vbQuestion Or vbYesNo Or vbDefaultButton2) <> vbYes Then Exit Function
         Exit Function
      End If
      Do While InStr(1, strVal, "  ", vbTextCompare) > 0
         strVal = Replace(strVal, "  ", " ")
      Loop
      If Len(strVal) = 0 Then
         lb = 0
         ub = -1
         ItsOk = True
      Else
         baVal = Split(strVal, " ")
         lb = LBound(baVal)
         ub = UBound(baVal)
         ReDim Out(1 To (ub - lb + 1))
         ItsOk = True
         For k = lb To ub
            On Error Resume Next
            Out(k - lb + 1) = CByte("&h" & baVal(k))
            If Err <> 0 Then
               If MsgBox(Err.Description, vbExclamation Or vbRetryCancel) <> vbRetry Then
                  Err.Clear: On Error GoTo 0
                  Exit Function
               End If
               Err.Clear
               Err.Clear: On Error GoTo 0
               ItsOk = False
               Exit For
            End If
            On Error GoTo 0
         Next
      End If
   Loop Until ItsOk
   AskBinary = True
End Function
Public Function IsCLSID(s As String) As Boolean
   Dim DoveMeno(1 To 4) As Long, k As Long, k2 As Long, QuestoIsMeno As Boolean

   DoveMeno(1) = Len("00000000-")
   DoveMeno(2) = Len("00000000-0000-")
   DoveMeno(3) = Len("00000000-0000-0000-")
   DoveMeno(4) = Len("00000000-0000-0000-0000-")
   IsCLSID = False
   If Len(s) <> Len("00000000-0000-0000-0000-000000000000") Then Exit Function
   For k = 1 To Len(s)
      QuestoIsMeno = False
      For k2 = LBound(DoveMeno) To UBound(DoveMeno)
         If k = DoveMeno(k2) Then
            If Mid(s, k, 1) <> "-" Then Exit Function
            QuestoIsMeno = True
            Exit For
         End If
      Next
      If Not QuestoIsMeno Then
         Select Case UCase(Mid(s, k, 1))
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F"
            Case Else
               Exit Function
         End Select
      End If
   Next
   IsCLSID = True
End Function
Public Function CCLSID(s As String) As CLSID
   Dim r As CLSID
   If Not IsCLSID(s) Then
      Err.Raise 521, , "The string is not a CLSID"
      Exit Function
   End If
   On Error GoTo CCLSID_Err
   Err.Clear
   r.Data1 = CLng("&h" & Mid(s, 1, 8))
   r.Data2 = CInt("&h" & Mid(s, 10, 4))
   r.Data3 = CInt("&h" & Mid(s, 15, 4))
   r.Data4_0 = CByte("&h" & Mid(s, 20, 2))
   r.Data4_1 = CByte("&h" & Mid(s, 22, 2))
   r.Data4_2 = CByte("&h" & Mid(s, 25, 2))
   r.Data4_3 = CByte("&h" & Mid(s, 27, 2))
   r.Data4_4 = CByte("&h" & Mid(s, 29, 2))
   r.Data4_5 = CByte("&h" & Mid(s, 31, 2))
   r.Data4_6 = CByte("&h" & Mid(s, 33, 2))
   r.Data4_7 = CByte("&h" & Mid(s, 35, 2))
   CCLSID = r
CCLSID_Exit:
   Exit Function
CCLSID_Err:
   Dim en As Long, ed As String
   en = Err.Number
   ed = Err.Description
   Err.Clear
   On Error GoTo 0
   Err.Raise en, , ed
   Resume CCLSID_Exit
End Function

Public Function AskPropTag(ByRef PropTag As Long, Optional ByVal inDefault As Long = 0) As Boolean
   Dim f As New frmAskPropTag
   
   f.inDefault = inDefault
   f.Show vbModal
   
   AskPropTag = f.outOk
   AskPropTag = True
   
   If AskPropTag Then
      PropTag = f.outValue
   End If
   'Unload f
End Function
