VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Outlook Spam an File in Future Deleter"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   Icon            =   "attachEx.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   4890
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3000
      Top             =   1650
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "Start Extraction of Text..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5085
      Width           =   3210
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Delete Attachments"
      Height          =   270
      Left            =   1515
      TabIndex        =   9
      Top             =   4425
      Width           =   1785
   End
   Begin VB.ListBox lstExtractionStatus 
      Enabled         =   0   'False
      Height          =   1620
      Left            =   135
      TabIndex        =   6
      Top             =   2790
      Width           =   4560
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   135
      TabIndex        =   2
      Top             =   2115
      Width           =   4560
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   135
      TabIndex        =   1
      Top             =   360
      Width           =   4515
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Start Extraction..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4725
      Width           =   3210
   End
   Begin VB.Label lblExtract 
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2070
      TabIndex        =   8
      Top             =   2520
      Width           =   825
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status of the Extraction..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   135
      TabIndex        =   7
      Top             =   2520
      Width           =   4515
   End
   Begin VB.Label Label3 
      Caption         =   "Application developed by Arun  Nair"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   990
      TabIndex        =   5
      Top             =   5625
      Width           =   2625
   End
   Begin VB.Label Label2 
      Caption         =   "Select the drive "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   135
      TabIndex        =   4
      Top             =   1890
      Width           =   3930
   End
   Begin VB.Label Label1 
      Caption         =   "Select Folder where you want the Attachment to be extracted"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   135
      TabIndex        =   3
      Top             =   45
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MailFieldToSort = "Received"       'language dependant





Sub SendMail()

    Exit Sub
    

 '   Generate and send E-mail
   Set olOutlookApp = GetObject(, "Outlook.Application")
   If Err <> 0 Then
       '   Outlook not running - start it
       Set olOutlookApp = CreateObject("Outlook.Application")
       blnNewOutlookApp = True
   End If
   '   Create E-mail
   Set olEMail = olOutlookApp.CreateItem(olMailItem)
   
   Set olOutlookApp = GetObject(, "Outlook.Application")
   If Err <> 0 Then
       '   Outlook not running - start it
       Set olOutlookApp = CreateObject("Outlook.Application")
       blnNewOutlookApp = True
   End If
   '   Create E-mail
   Set olEMail = olOutlookApp.CreateItem(olMailItem)
   With olEMail
        .SenderName = "Matt Lan"
        .SenderEmailAddress = "matt.lan@btinternet.com"
        .To = "Matt Lan1 <matt.lan@btinternet.com>; Matt Lan 2<matt.lan@btinternet.com>"
        .Subject = "Mail From Matthew to Sophie With Attach"
        .Body = "Hello"
        .ReadReceiptRequested = True
        .Attachments.FileName , "0807"
   End With

   With olEMail
    ' Add attachments to the message.
    If Not IsMissing(AttachmentPath) Then
        Set objOutlookAttach = .Attachments.Add(AttachmentPath)
    End If
   End With


End

End Sub


Sub SendMailOld()









Dim MyItems


'On Error Resume Next
    
Me.lblExtract.Caption = ""
Me.lstExtractionStatus.Enabled = True
Me.lstExtractionStatus.Clear

Dim oApp As New Outlook.Application
Dim oNameSpace As NameSpace
Dim oMailitem As Object
Dim sMessage As String

Set oNameSpace = oApp.GetNamespace("MAPI")



Set Myitems9 = MyItems.Item(Rrr) 'oApp.CreateItem(olMailItem) 'Just Dummy

Myitems9.ReadReceiptRequested = False
'MyItems9.senderemail = a2$
'MyItems9.SenderName = a3$
Myitems9.Subject = a4$ + " GA xxxNo5xxx"

'MyItems9.Body = MyItems.Item(rrr).Body
Myitems9.To = "matt.lan@btinternet.com"  ' && First Security Dialog

'MyItems9.Recipients.ResolveAll



Myitems9.htmlBody = a9$


Myitems9.Send '        && Another Security Dialog - Need to Wait for timer.
Set Myitems9 = Nothing



End Sub




Private Sub Form_Activate()


Exit Sub
Call xy
End Sub

Private Sub Form_Load()

Call SendMail

Exit Sub

If App.PrevInstance = True Then
Open App.Path + "\Timer.txt" For Output As #1
Close #1
End
End If
    
rr1 = FindWindow(vbNullString, "VB_Future_Googles - Microsoft Visual Basic [design]")
rr2 = FindWindow(vbNullString, "VB_Future_Googles - Microsoft Visual Basic [run]")
rr3 = FindWindow(vbNullString, "VB_Future_Googles - Microsoft Visual Basic [break]")

If (rr1 > 0 Or rr2 > 0 Or rr3 > 0) And IsIDE = False Then End

Timer4 = 3
If IsIDE = False Then Timer4 = 50

If Gr2 = True Then Timer4 = 0

Form3.Show
Call xy


    
    'Me.Drive1.Drive = "C"
    'Me.Dir1.Path = "c:\Art\My Pictures\Fractals"
    'Me.Dir1.Path = "D:\# MY DOCS\# 01 My Documents"
    'Me.Dir1.Path = "D:\My Webs\MatthewLan.Com Web\LoveFolder\Alt.Sz_Borg_Roid_Judy\Alt_Sz_2003_Onwards_TxT"
    'Me.Dir1.Path = "D:\My Webs\MatthewLan.Com Web\LoveFolder"

'Timered Now
'Call Command2_Click
End Sub



Private Sub Command2_Click()
'this will last time
'Extract txt for whole list of sz.group from outlook 2000
Dim MyItems


'On Error Resume Next
    
Me.lblExtract.Caption = ""
Me.lstExtractionStatus.Enabled = True
Me.lstExtractionStatus.Clear

Dim oApp As New Outlook.Application
Dim oApp2 As New Outlook.Application
Dim oNameSpace As NameSpace
Dim oFolder As MAPIFolder
Dim oFolder2 As MAPIFolder
Dim oFolder3 As MAPIFolder
Dim SentMailoFolder As MAPIFolder
'Dim oFolderJunk As MAPIFolder
Dim oMailitem As Object
Dim sMessage As String

'Set oApp = New Outlook.Application
'Set oApp2 = New Outlook.Application

Set oNameSpace = oApp.GetNamespace("MAPI")
Set oNameSpace2 = oApp2.GetNamespace("MAPI")
Set oNameSpace3 = oApp2.GetNamespace("MAPI")
Set oNameSpace5 = oApp2.GetNamespace("MAPI")

Set oFolder2 = oNameSpace2.GetDefaultFolder(23) 'olFolderJunkFolder
Set oFolder3 = oNameSpace3.GetDefaultFolder(3) 'olFolderDeleted
Set SentMailoFolder = oNameSpace3.GetDefaultFolder(olFolderSentMail) 'olFolderDeleted
'olFolderSentMail



Set oMAPI = GetObject("", "Outlook.application").GetNamespace("MAPI")
Set oParentFolder = oMAPI.Folders("Personal Folders")
If oParentFolder.Folders.Count Then
  For i1 = 1 To oParentFolder.Folders.Count
    'Set oFolder5 = oParentFolder.Folders(i1) 'Neto
    'tt1$ = oFolder5.Name
    If Trim(oParentFolder.Folders(i1).Name) = "00-NetOMatic" Then
        Exit For
    End If
  Next i1
End If
'Set oFolder5 = oParentFolder.Folders(i1) 'Neto
'tt1$ = oFolder5.Name

Set oPart2 = oMAPI.Folders("Personal Folders")
If oPart2.Folders.Count Then
  For i2 = 1 To oPart2.Folders.Count
    If Trim(oPart2.Folders(i2).Name) = "00-Future spam" Then
        Exit For
    End If
  Next
End If

'Set oPart2 = oMAPI.Folders("Personal Folders")
If oPart2.Folders.Count Then
  For i3 = 1 To oPart2.Folders.Count
    If Trim(oPart2.Folders(i3).Name) = "00-Google_Copy" Then
        Exit For
    End If
  Next
End If

If oPart2.Folders.Count Then
  For i4 = 1 To oPart2.Folders.Count
    If Trim(oPart2.Folders(i4).Name) = "00-Grab-Test" Then
        Exit For
    End If
  Next
End If

On Local Error Resume Next

Set oFolder5 = oParentFolder.Folders(i1) 'Neto
tt1$ = oFolder5.Name
Set oFolder8 = oPart2.Folders(i2) 'Future Spam
tt2$ = oFolder8.Name
Set oFolder10 = oPart2.Folders(i3) 'Google Copy
tt3$ = oFolder10.Name

Set oFolder11 = oPart2.Folders(i4) ' Grab Test For google
tt4$ = oFolder11.Name


'Set oFolder12 = oPart2.Folders(i4) ' ForWards Folder
'tt3$ = oFolder12.Name

Set oFolder = oNameSpace.GetDefaultFolder(olFolderInbox) '6) 'olFolderInBox


Set MyItems = oFolder.Items
Set MyItems2 = oFolder2.Items
Set MyItems3 = oFolder3.Items
Set myitems4 = oApp.CreateItem(olMailItem) 'Just Dummy

'Set myItem = oApp.CreateItem(olMailItem)


Set myitems5 = oFolder5.Items 'NetOMatic
Set myitems8 = oFolder8.Items 'Future Spam
Set MyItems10 = oFolder10.Items 'Google Copy
Set Myitems11 = oFolder11.Items 'Test Google Copy
'Set MYGC = oFolder11.Items

'Set MYGC = oApp.CreateItem(olMailItem) 'Just Dummy

Set mygc = oApp.CreateItem(olMailItem) 'Just Dummy
Set Myitems9 = oApp.CreateItem(olMailItem) 'Just Dummy
Set MyItemsFrm4 = oApp.CreateItem(olMailItem) 'Just Dummy
'Set MyItems9 = oApp.CreateItem(CreateItem(0))   'Just Dummy

Set MyItemsentMail = SentMailoFolder.Items

'inbox
Call MyItems.Sort("[" & MailFieldToSort & "]", True)
Call MyItems2.Sort("[" & MailFieldToSort & "]", True)
Call MyItems3.Sort("[" & MailFieldToSort & "]", True)
Call MyItemsentMail.Sort("[" & MailFieldToSort & "]", True)















Form3.Hide
'Form4.Show

'Del All Del Folder at start
Title$ = "Del DelBox-------------------------------"
Call Frm4


Rrr2 = MyItems3.Count
If Rrr2 > 0 Then Form2.Show
Form2.Lb1 = "Del the del Box"
Form2.Show
For Rrr = 1 To MyItems3.Count
    Set MyItemsFrm4 = MyItems3.Item(Rrr)
    Call Frm4
'Form2.Lb2 = Str(Rrr) + " of " + Str(MyItems.Count)
tt1$ = oFolder.Name

Form2.Lb2 = Str(Rrr) + " of " + Str(MyItems3.Count)
   
    MyItems3.Item(Rrr).Delete
Next

'Dim myCopiedItem As Outlook.MailItem

Form4.Hide
Form2.Show
Call xy

'inbox
Form2.Lb1 = "ChK Inbox For 1St Gog Lert"
'Form2.Lb2 = Str(Rrr) + " of " + Str(MyItems.Count)
tt1$ = oFolder.Name
    For Rrr = 1 To MyItems.Count

Form2.Lb2 = Str(Rrr) + " of " + Str(MyItems.Count)

    a1$ = MyItems.Item(Rrr).Subject
'    MyItems.Item(rrr).Display
    If InStr(a1$, "Google Alert - ") = 1 Then Exit For
Next

rtt1 = MyItems.Item(Rrr).ReceivedTime
rty1 = Rrr

'MyItems.Item(rty1).Display


















'------------------------
Dim Gr2

GrabOneTest = IsIDE
'GrabOneTest = True
GrabOneTest = False

Gr2 = GrabOneTest


'If IsIDE = False And Gr2 = True Then MsgBox "Test Mode": End

'If GrabOneTest = True Then GoTo jump1

'sent box
tt1$ = SentMailoFolder.Name
Form2.Lb1 = "Get latest date Of Google Alerts Sent Sent Items Folder"

' get latest date Of Google Alerts Sent Sent Items Folder
For Rrr = 1 To MyItemsentMail.Count

Form2.Lb2 = Str(Rrr) + " of " + Str(MyItemsentMail.Count)

'qq = MyItems.Item(rrr).ReceivedTime
    d1$ = MyItemsentMail.Item(Rrr).To
    a1$ = MyItemsentMail.Item(Rrr).Subject
    b1$ = MyItemsentMail.Item(Rrr).Senderemail
'    MyItemsentMail.Item(rrr).Display
    tyx = InStr(a1$, "Google Alert - ")
    If (tyx > 0 And tyx < 8 And d1$) Or Gr2 = True Then
        If d1$ <> "'NetOMatic@btinternet.com'" Then
            rtt2 = MyItemsentMail.Item(Rrr).ReceivedTime
            rtu = Rrr
            Exit For
        End If
    End If

Next



Form2.Lb1 = "ChK Inbox an Compare 1St Gog Lert From Sent Box"
'Form2.Lb2 = Str(Rrr) + " of " + Str(MyItems.Count)

If rtt2 = 0 And MyItemsentMail.Count = 0 Then rtt2 = Now - 2
'inbox
tt1$ = oFolder.Name
For Rrr = 1 To MyItems.Count
Form2.Lb2 = Str(Rrr) + " of " + Str(MyItems.Count)
        xrtt1 = MyItems.Item(Rrr).ReceivedTime

    If xrtt1 < rtt2 Then rr7 = Rrr: Exit For
Next

'MyItems.Item(rr7).Display
'MyItems.Item(1).Display
 '       xrtt1 = MyItems.Item(rrr).ReceivedTime


Form2.Lb1 = "Do The Google File Manips"
'Form2.Lb2 = Str(Rrr) + " of " + Str(MyItems.Count)
'inbox
tt1$ = oFolder.Name


For Rrr = rr7 To 1 Step -1
If hake > 0 Then
Form2.Lb2 = Str(Rrr) + " of " + Str(rr7) + " and " + Str(hake) + " Done"
Else
Form2.Lb2 = Str(Rrr) + " of " + Str(rr7)
End If

'MyItems.Item(rty1).Display
'MyItems.Item(rtt2).Display

'd1$ = MyItemsentMail.Item(rrr).To
a1$ = MyItems.Item(Rrr).Subject

jump1:




If InStr(a1$, "Google Alert - ") = 1 Or Gr2 = True Then

a5$ = Trim(MyItems.Item(Rrr).htmlBody)
tyuo = InStr(a5$, "Filthy Fucking File Manipulator")

If tyuo = 0 Or Gr2 = True Then

If Gr2 = False Then
    Set MyItems10 = MyItems.Item(Rrr).Copy
    MyItems.Item(Rrr).SaveAs App.Path + "\zzEmail_Orgi.txt", olTXT
    MyItems.Item(Rrr).SaveAs App.Path + "\zzEmail_Orgi.html", olHTML
    MyItems.Item(Rrr).SaveAs App.Path + "\zzEmail_Orgi Text.txt", olHTML
End If

If Gr2 = True Then
    Set oFolder11 = oPart2.Folders(i4) ' Grab Test For google
    tt4$ = oFolder11.Name
    Set Myitems11 = oFolder11.Items 'Test Google Copy
    'rrt$ = Myitems11.Name
'    t9 = Myitems11.Count
 '   rr$ = Myitems11.Item(1).Subject
    'Or That if 1 item
    'rr$ = Myitems11.Subject
    
    hhht$ = App.Path + "\zzEmail_Orgi.txt"
    hhhh$ = App.Path + "\zzEmail_Orgi.html"
    hhh2$ = App.Path + "\zzEmail_Orgi Htm.txt"
    Myitems11.Item(1).SaveAs hhht$, olTXT
    Myitems11.Item(1).SaveAs hhhh$, olHTML
    Myitems11.Item(1).SaveAs hhh2$, olHTML
    Set MyItems10 = Myitems11.Item(1).Copy
End If

'Set MYGC = myitems11.Copy

a1$ = MyItems.Item(Rrr).Subject
a2$ = Myitems11.Item(1).Subject
'a3$ = mygc.Subject



'myitems11.SaveAs App.Path + "\Email Orgi.txt", olTXT
'myitems11.SaveAs App.Path + "\Email Orgi.html", olHTML


'Set mygc = Nothing

'myitems10.Display
If Gr2 = False Then MyItems10.Move oFolder10 'Google Copy
'myitems10.Display

If Gr2 = False Then
a4$ = Trim(MyItems.Item(Rrr).Subject)
a5$ = Trim(MyItems.Item(Rrr).htmlBody)
End If
If Gr2 = True Then
a4$ = Trim(Myitems11.Item(1).Subject)
a5$ = Trim(Myitems11.Item(1).htmlBody)
End If

'len(a5$)
ffds1 = InStr(a5$, "<p style=")
ffds2 = InStr(ffds1, a5$, "<a")

ffds3 = InStr(a5$, "1>Google ")


a7$ = Mid$(a5$, 1, ffds1 - 1)
a7$ = a7$ + Mid$(a5$, ffds2)

a8$ = "Filthy Fucking Warrior GangStar<br>Filthy Fucking File Manipulator<br>Military Militia Maffia Matt<br>AlCaPhone Does His Own Shopping<br>Yheeeeeer<br>"
a8$ = a8$ + "<font  color=""0000ff"">Rx--- " + Format$(MyItems.Item(Rrr).ReceivedTime, "DDD DD-MMM-YYYY HH:MM:SSa/p") + "<br></font>"
a8$ = a8$ + "<font  color=""0000ff"">Now " + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SSa/p") + "<br></font>"

'a10$ = "<br><br><font  color=""0000FF"">&lt;p </font><font  color=""FF0000"">style=</font><font color=""0000FF"">""width:600px""></font><font color=""#000000""><br>Eye There  Google<br>2+2+2 = 3 * 2 = 5 - Matey<br>Military Mean Machine Matty Roids<br>Har Har Harder Ha Hard Ha Yer<br>Matty Roiders Rimmers<br>Ratty Roids" + vbCrLf
If Dir$(App.Path + "\Mach One Bugg Free.txt") = "" Then
    a10$ = "<br><br><font color=""#000000"">Military Mean Machine Matty Roids<br>Har Har Harder Ha Hard Ha Yer<br>Matty Roiders Rimmers<br>Ratty Roids" + vbCrLf
End If

If Dir$(App.Path + "\Mach One Bugg Free.txt") <> "" Then
    a10$ = "<br><br><font color=""#000000"">Military Mean Machine Matty Roids<br>Har Har Harder Ha Hard Ha Yer<br><br>First of The Mach-a-Roids Buggers Free-ers<br>Matty Roiders Rimmers<br>Ratty Roids" + vbCrLf
    Open App.Path + "\Mach One Bugg Free.txt" For Output As #1
    Print #1, Format$(Now, "ddd dd-mm-yyyy hh:mm:ss")
    Close #1
End If

a9$ = Mid$(a7$, 1, ffds3 + 1) + a8$
a9$ = a9$ + Mid$(a7$, ffds3 + 2)

ffds4 = InStr(a9$, "</div></body></html>")
'a22$ = Mid$(a9$, 1, ffds4 - 5)
'a23$ = Mid$(a22$, Len(a22$) - 10)
a9$ = Mid$(a9$, 1, ffds4 - 5) + a10$ + vbCrLf + UCase$("</font></p></div></body></html>")

Do

'style="WIDTH: 600px"
ii = InStr(UCase$(a9$), UCase$("style=""Width"))
If ii > 0 Then
a9B$ = Mid$(a9$, 1, ii - 2)
'yy$ = Mid$(a9B$, Len(a9B$) - 10)
a9C$ = a9B$ + Mid$(a9$, ii + 19)
'yy$ = Mid$(A9C$, II - 10)
'yy2$ = Mid$(A9C$, II - 8)
a9$ = a9C$
End If

Loop Until ii = 0

'Open App.Path + "\EmailTxt.txt" For Output As #1
'Print #1, a9$
'Close #1
'Open App.Path + "\EmailTxt.html" For Output As #1
'Print #1, a9$
'Close #1

'If Gr2 = True Then Shell "explorer " + App.Path, vbNormalFocus



usacount = 0: ii = 0
Do
'-
ii = InStr(ii + 1, UCase$(a9$), UCase$("#666666>"))
If ii = 0 Then ii = InStr(ii + 1, UCase$(a9$), UCase$("style=""COLOR: blue"))

ii3 = InStrRev(UCase$(a9$), "<BR>RX-", ii)
ii2 = InStr(ii + 1, UCase$(a9$), UCase$("</font><br>"))
If ii > 0 Then
chkusa$ = Mid$(a9$, ii, (ii2 - ii))
tto = InStr(chkusa$, "<font color=#666666>")
If tto = 0 Then tto = InStr(chkusa$, "style=""COLOR: blue")

usaspam2$ = ",PA,NY,CA,NJ,WA,MA,SC,IL,NC,AZ,CT,TX,NV,TN,WA,USA"
yy$ = UCase$(Mid$(chkusa$, Len(chkusa$) - 2))
If InStr(usaspam2$, yy$) > 0 Then 'xxy = 1
usacount = usacount + 1
End If
End If
Loop Until ii = 0 Or tto = 0




xxr = 0


If InStr(a4$, "Google Alert - Mental-Health") > 0 Then xxr = 1
If InStr(a4$, "Google Alert - Intelligent knowledge") > 0 Then xxr = 1
If InStr(a4$, "Google Alert - Power Of Thought") > 0 Then xxr = 1

If InStr(a4$, "Google Alert - Highly Organized Organised") > 0 Then xxr = 1
If InStr(a4$, "Google Alert - Needle Razor") > 0 Then xxr = 1
If InStr(a4$, "Google Alert - Olanzapine") > 0 Then xxr = 1
If InStr(a4$, "Google Alert - Orodispersible") > 0 Then xxr = 1
If InStr(a4$, "Google Alert - Risperdal OR Risperidone") > 0 Then xxr = 1
If InStr(a4$, "Google Alert - Schizophrenia UK") > 0 Then xxr = 1
If InStr(a4$, "Google Alert - Uk Government") > 0 Then xxr = 1
If InStr(a4$, "Google Alert - ""Schizophrenia""") > 0 Then xxr = 1

'If InStr(a4$, "") > 0 Then xxr = 1
'If InStr(a4$, "") > 0 Then xxr = 1
'If InStr(a4$, "") > 0 Then xxr = 1
'If InStr(a4$, "") > 0 Then xxr = 1
'If InStr(a4$, "") > 0 Then xxr = 1
'If InStr(a4$, "") > 0 Then xxr = 1
'If InStr(a4$, "") > 0 Then xxr = 1
'If InStr(a4$, "") > 0 Then xxr = 1
'If InStr(a4$, "") > 0 Then xxr = 1

ii = 0
xxc = 0
itemcount2 = 0
usacount = 0: Iidf = 0
blue = 3
Do
If blue = 3 Or blue = 1 Then
ii = InStr(ii + 1, UCase$(a9$), UCase$("color=#666666"))
If ii > 0 Then blue = 1
itemcount2 = itemcount2 + 1
End If
If blue = 3 Or blue = 2 Then
ii = InStr(ii + 1, UCase$(a9$), UCase$("style=""COLOR: blue"))
blue = 2
itemcount2 = itemcount2 + 1
End If

Loop Until ii = 0

itemcount2 = itemcount2 - 1


a9B$ = Mid$(a9$, 1, ii - 2)
'yy$ = Mid$(a9B$, Len(a9B$) - 10)
a9C$ = a9B$ + a45$ + Mid$(a9$, ii2 - 1)

usacount4 = 0
usacount5 = 0
ii = 0
blue = 3
Do

If blue = 3 Or blue = 1 Then
ii = InStr(ii + 1, UCase$(a9$), UCase$("#666666>"))
If ii > 0 Then blue = 1
End If

If blue = 3 Or blue = 2 Then
ii = InStr(ii + 1, UCase$(a9$), UCase$("style=""COLOR: blue"))
blue = 2
End If

If ii = 0 Then Exit Do

ii3 = InStrRev(UCase$(a9$), "<BR>RX-", ii)
ii2 = InStr(ii + 1, UCase$(a9$), UCase$("</font><br>"))
If ii > 0 Then
chkusa$ = Mid$(a9$, ii, (ii2 - ii))
'tto = InStr(chkusa$, "<font color=#666666>")
'If tto = 0 Then tto = InStr(chkusa$, "style=""COLOR: blue")
usaspam2$ = ",PA,NY,CA,NJ,WA,MA,SC,IL,NC,AZ,CT,TX,NV,TN,WA,USA"
yy$ = UCase$(Mid$(chkusa$, Len(chkusa$) - 2))
If InStr(usaspam2$, yy$) > 0 Then 'xxy = 1
usacount4 = usacount4 + 1
End If
End If
'Exit Do
Loop Until ii = 0 ' Or tto = 0


blue = 3
ii = 0
Do
itc = itc + 1
ii = InStr(ii + 1, UCase$(a9$), UCase$("<a style="))
Loop Until ii = 0

If blue = 3 Or blue = 1 Then
    ii = InStr(ii + 1, UCase$(a9$), UCase$("#666666>"))
    If ii > 0 Then blue = 1
End If
If blue3 Or blue = 2 Then
    ii = InStr(ii + 1, UCase$(a9$), UCase$("style=""COLOR: blue"))
    blue = 2
End If

ii2 = InStr(ii + 1, UCase$(a9$), UCase$("</font><br>"))
as3$ = Mid$(a9$, 1, ii - 2)
as4$ = Mid$(as3$, InStrRev(as3$, "Google "))
a45$ = "Rx- " + Format$(MyItems.Item(Rrr).ReceivedTime, "DDD DD-MMM-YYYY HH:MM:SSa/p") + "<br>"
a45$ = a45$ + vbCrLf + Mid$(as4$, 1, InStr(as4$, "</b>") + 3) + "<br>"
ii = 0
itc = 0
itg = 0
blue = 3
iidg2 = 0

Do
itc = itc + 1
ii = InStr(ii + 1, UCase$(a9$), UCase$("<a style="))
If ii = 0 Then Exit Do
a9B$ = Mid$(a9$, 1, ii - 2)
'yy$ = Mid$(a9B$, Len(a9B$) - 10)
'iidg2 = 0
'blue = 3

If blue = 3 Or blue = 1 Then
iidg2 = InStr(ii + 1, UCase$(a9$), UCase$("#666666>"))
If iidg2 > 0 Then blue = 1
End If
If blue = 3 Or blue = 2 Then
iidg2 = InStr(ii + 1, UCase$(a9$), UCase$("style=""COLOR: blue"))
blue = 2
End If

Open App.Path + "\zz EmailTxt.txt" For Output As #1
Print #1, a9$
Close #1
Open App.Path + "\zz EmailTxt.html" For Output As #1
Print #1, a9$
Close #1
If Gr2 = True Then Shell "explorer " + App.Path, vbNormalFocus

'If blue = 1 Then
'ii3dg2 = InStrRev(UCase$(a9$), "<BR>RX-", ii)
If blue = 1 Then ii2dg2 = InStr(ii + 1, UCase$(a9$), UCase$("</font><br>"))
If blue = 2 Then ii2dg2 = InStr(ii + 1, UCase$(a9$), UCase$("<font color="))
'ii2dg2 = InStr(ii + 1, UCase$(a9$), UCase$("</font><br>"))
If iidg2 > 0 Then
chkusa$ = Mid$(a9$, iidg2, (ii2dg2 - iidg2))

If blue = 1 Then
tto = InStr(chkusa$, "<font color=#666666>")
usaspam2$ = ",PA,NY,CA,NJ,WA,MA,SC,IL,NC,AZ,CT,TX,NV,TN,WA,USA"
yy$ = UCase$(Mid$(chkusa$, Len(chkusa$) - 2))
If InStr(usaspam2$, yy$) > 0 Then 'xxy = 1
usacount5 = usacount5 + 1
End If
End If

End If
'End If




Open App.Path + "\zz Small String.txt" For Output As #1
Print #1, Mid$(a9$, ii)
Close #1
Open App.Path + "\zz Small String.html" For Output As #1
Print #1, Mid$(a9$, ii)
Close #1
'Shell "notepad " + App.Path + "\zz Small String.txt"

If blue = 1 Then
'Stop
itg = itg + 1
 ggh$ = Format$(itg, "0") + " of " + Format$(itemcount2, "0") + "<br>" + a45$
a9C$ = a9B$ + ggh$ + Mid$(a9$, ii)
End If
If blue = 2 Then
ggh$ = Format$(itc - usacount5, "0") + " of " + Format$(itemcount2 - usacount4, "0") + a45$
a9C$ = a9B$ + ggh$ + Mid$(a9$, ii)
End If
ii = ii + Len(ggh$) + 1


a9$ = a9C$
Loop Until ii = 0
'Stop
Open App.Path + "\zz String.txt" For Output As #1
Print #1, a9$
Close #1
Open App.Path + "\zz String.html" For Output As #1
Print #1, a9$
Close #1



'Itemcount = Itemcount2

'If Itemcount = 0 Then
'    Do
'    ii = InStr(ii + 1, UCase$(a9$), UCase$("style=""COLOR: blue"))
'    If ii > 0 Then Itemcount = Itemcount + 1
'    Loop Until ii = 0
'End If


'style="COLOR: blue"
blue = 3
usacount = 0
ii = 0
Do


'usacount = 0:ii=0
'-
'ii = InStr(ii + 1, UCase$(a9$), UCase$("#666666>"))
'ii3 = InStrRev(UCase$(a9$), "<BR>RX-", ii)
'ii2 = InStr(ii + 1, UCase$(a9$), UCase$("</font><br>"))
'If ii > 0 Then
'chkUSA$ = Mid$(a9$, ii, (ii2 - ii))
'tto = InStr(chkUSA$, "<font color=#666666>")
'usaspam2$ = ",PA,NY,CA,NJ,WA,MA,SC,IL,NC,AZ,CT,TX,NV,TN,WA,USA"
'yy$ = UCase$(Mid$(chkUSA$, Len(chkUSA$) - 2))
'If InStr(usaspam2$, yy$) > 0 Then 'xxy = 1
'usacount = usacount + 1




'Itemcount = Itemcount + 1
'If InStr(a4$, "Google Alert - Mental-Health") = 0 Then Exit Do
If xxr = 0 Then Exit Do


'ii = InStr(UCase$(A9$), UCase$("<a style=""color: blue"))


If blue = 3 Or blue = 1 Then
ii = InStr(ii + 1, UCase$(a9$), UCase$("#666666>"))
If ii > 0 Then blue = 1
End If
If blue = 3 Or blue = 2 Then
ii = InStr(ii + 1, UCase$(a9$), UCase$("style=""COLOR: blue"))
blue = 2
End If



ii3 = InStrRev(UCase$(a9$), "<BR>RX-", ii)

'<br>Rx-
'#666666>
ii2 = InStr(ii + 1, UCase$(a9$), UCase$("</font><br>"))
'<hr noshade size=1>
If ii > 0 Then
chkusa$ = Mid$(a9$, ii, (ii2 - ii))

'always a 666666

tto = InStr(chkusa$, "<font color=#666666>")
'Open App.Path + "\zz String.txt" For Output As #1
'Print #1, chkusa$
'Close #1

Open App.Path + "\zz String.txt" For Output As #1
Print #1, a9$
Close #1


usaspam2$ = ",PA,NY,CA,NJ,WA,MA,SC,IL,NC,AZ,CT,TX,NV,TN,WA,USA"

yy$ = UCase$(Mid$(chkusa$, Len(chkusa$) - 2))

'xxy = 0

If InStr(usaspam2$, yy$) > 0 Then 'xxy = 1

usacount = usacount + 1

'ii = InStrRev(UCase$(a9$), UCase$("<a style=""color: blue"), ii)
ii = InStrRev(UCase$(a9$), UCase$("</font></p>"), ii3)
'ii2 = InStr(ii + 1, UCase$(a9$), UCase$("<a style=""color: blue"))
ii2 = InStr(ii + 1, UCase$(a9$), UCase$("topic</font></a>"))
'topic</font></a>
'</font></p>
xi2 = 0
If ii2 = 0 Then
    ii2 = InStr(ii + 1, UCase$(a9$), UCase$("</font></p><p>"))
    ii2 = ii2 - 1
    xi2 = 1
End If
'<b>...</b><br>
a10$ = vbCrLf + Mid$(a9$, ii2 + 18)


If Gr2 = True Then
'    Shell "explorer " + App.Path, vbNormalFocus
    Open App.Path + "\zz String.txt" For Output As #1
    Print #1, a9$
    Close #1
End If

a9B$ = Mid$(a9$, 1, ii - 2)
'yy$ = Mid$(a9B$, Len(a9B$) - 10)
a9C$ = a9B$ + a10$ 'Mid$(a9$, ii2 - 1)

'Open App.Path + "\zz String.txt" For Output As #1
'Print #1, a9b$
'Close #1


yy$ = Mid$(a9$, ii2)
'yy2$ = Mid$(A9C$, II - 8)
a9$ = a9C$
End If

If Gr2 = True Then
    'Shell "explorer " + App.Path, vbNormalFocus
    Open App.Path + "\zz String.txt" For Output As #1
    Print #1, a9$
    Close #1
    Open App.Path + "\zz String.html" For Output As #1
    Print #1, a9$
    Close #1
End If



End If
'End If

'Open App.Path + "\zz String.txt" For Output As #1
'Print #1, A9$
'Close #1

Loop Until ii = 0

If xxr = 1 Then
ii = InStr(UCase$(a9$), UCase$("Yheeeeeer<br>"))
If ii > 0 Then ii = ii + 12
xx3 = 0
If itemcount2 = usacount Then xx3 = 2
If usacount = 0 Then xx3 = 3

If xx3 = 0 Then a9B$ = Mid$(a9$, 1, ii) + "</font><font color=#000000>" + Trim(Str(usacount)) + "</font><font color=#800080> US Items Deleted"
If xx3 = 2 Then a9B$ = Mid$(a9$, 1, ii) + "</font><font color=#000000>" + Str(usacount) + "</font><font color=#800080> Opps All US Items Gone"
If xx3 = 3 Then a9B$ = Mid$(a9$, 1, ii) + "<font color=#800080>Oh - No Do US Items Today"

a9B$ = a9B$ + vbCrLf + "</font>"
a9C$ = a9B$ + Mid$(a9$, ii - 3)

a9$ = a9C$
End If

ii = InStr(UCase$(a9$), UCase$("Yheeeeeer<br>"))
If ii > 0 Then ii = ii + 12

xxc = 0

ttbb = itemcount2 - usacount4

If itemcount2 > usacount4 Or itemcount2 = usacount4 Then
ttg1$ = "<font color=#000000>" + Str(itemcount2) + "</font><font color=#800080> Items - Today</font>"
End If

If itemcount2 < usacount4 And xxc = 0 Then
ttg2$ = "<br><font color=#000000>" + Str(itemcount2) + "</font><font color=#800080> Items - Were Here"
End If



a9B$ = Mid$(a9$, 1, ii) + ttg1$ + vbCrLf + ttg2$

a9C$ = a9B$ + Mid$(a9$, ii - 3)

a9$ = a9C$

ii = InStrRev(a9$, "&nbsp;")
a9B$ = Mid$(a9$, 1, ii - 1)
'Open App.Path + "\zz String.txt" For Output As #1
'Print #1, a9b$
'Close #1

'a9c$ = Mid$(a9$, ii + 6)
a9C$ = a9B$ + Mid$(a9$, ii + 6)
'Open App.Path + "\zz String.txt" For Output As #1
'Print #1, a9c$
'Close #1

a9$ = a9C$

Open App.Path + "\zz String2.txt" For Output As #1
Print #1, a9$
Close #1
Open App.Path + "\zz String2.html" For Output As #1
Print #1, a9$
Close #1


'Yheeeeeer<br>

'a = a

'<a style="color: blue



Open App.Path + "\EmailTxt.txt" For Output As #1
Print #1, a9$
Close #1
Open App.Path + "\EmailTxt.html" For Output As #1
Print #1, a9$
Close #1
If Gr2 = True Then Shell "explorer " + App.Path, vbNormalFocus




'Shell "explorer " + App.Path, vbNormalFocus


'Google Groups Alert
'<p style="width:600px">
Set Myitems9 = MyItems.Item(Rrr) 'oApp.CreateItem(olMailItem) 'Just Dummy

Myitems9.ReadReceiptRequested = False
'MyItems9.senderemail = a2$
'MyItems9.SenderName = a3$
Myitems9.Subject = a4$ + " GA xxxNo5xxx"

'MyItems9.Body = MyItems.Item(rrr).Body
Myitems9.To = "matt.lan@btinternet.com"  ' && First Security Dialog

'MyItems9.Recipients.ResolveAll



Myitems9.htmlBody = a9$
'MyItems9.Display

If Gr2 = True Then End

    hhht$ = App.Path + "\zzEmail_Final.txt"
    hhhh$ = App.Path + "\zzEmail_Final.html"
    hhh2$ = App.Path + "\zzEmail_Final Htm.txt"
    Myitems9.SaveAs hhht$, olTXT
    Myitems9.SaveAs hhhh$, olHTML
    Myitems9.SaveAs hhh2$, olHTML

hake = hake + 1
Form2.Lb2 = Str(Rrr) + " of " + Str(rr7) + " and " + Str(hake) + " Done"


Shell "explorer " + hhhh$
'Stop
'End

If xx3 <> 2 Then Myitems9.Send '        && Another Security Dialog - Need to Wait for timer.




'tt$ = oFolder10.Name

'tt3$ = oFolder10.Name
'a1$ = MyItem10.Subject
'If xx3 <> 2 Then Myitems9.Send '        && Another Security Dialog - Need to Wait for timer.
Set Myitems9 = Nothing
DoEvents
'Set MyItems = Nothing

'Set MyItems = oFolder.Items

MyItems.Item(Rrr).Display
MyItems.Item(Rrr).Delete


End If
End If
'End
Next









'------------




If hake > 0 Then
    Form2.Lb1 = "Del The Google File Manips - Ones That Were Copyied Before"
    'inbox
    tt1$ = oFolder.Name
    For Rrr = rr7 To 1 Step -1
        Form2.Lb2 = Str(Rrr) + " of " + Str(rr7)
        'MyItems.Item(rty1).Display
        'MyItems.Item(rtt2).Display

        'd1$ = MyItemsentMail.Item(rrr).To
        a1$ = MyItems.Item(Rrr).Subject

        If InStr(a1$, "Google Alert - ") = 1 Then
        a5$ = Trim(MyItems.Item(Rrr).htmlBody)
        tyuo = InStr(a5$, "Filthy Fucking File Manipulator")
            If tyuo = 0 Then MyItems.Item(Rrr).Delete
        End If
    Next
End If


Form2.Lb1 = "Chk For Future Spam..."
tty = MyItems.Count

For Rrr = 1 To MyItems.Count

    Form2.Lb2 = Str(Rrr) + " of " + Str(tty)

    qq = MyItems.Item(Rrr).ReceivedTime

    'inbox
    tt$ = oFolder.Name

    If qq > Now Then

        'FSPam
        Set myitems8 = MyItems.Item(Rrr).Copy

        'FSPam
        myitems8.Move oFolder8 'FSPam
        tt$ = oFolder.Name

        MyItems.Item(Rrr).Delete

    End If

    If qq < Now Then Exit For

Next


Form2.Lb1 = "Copy Junk Mail To NetOMatic Then Del It From Junk Folder.."
'Form2.Lb2 = Str(Rrr) + " of " + Str(tty)
tty = MyItems2.Count

For Rrr = MyItems2.Count To 1 Step -1

    Form2.Lb2 = Str(Rrr) + " of " + Str(tty)
    
    Set myitems4 = MyItems2.Item(Rrr).Copy
    myitems4.Move oFolder5 'NetOMatic

    MyItems2.Item(Rrr).Delete


Next


Set oMailitem = Nothing
Set oFolder = Nothing
Set oNameSpace = Nothing
Set oApp = Nothing
    
End
End Sub

Private Sub Drive1_Change()
    Me.Dir1.Path = Drive1.Drive

End Sub




Public Function OutlookFolderNames(MailboxName As String) _
    As String()

'********************************************************
'PARAMETER: MailboxName = Name of Parent Outlook Folder for
'the current user:  Usually in the form of
'"Mailbox - Doe, John" or
'"Public Folders

'RETURNS: Array of SubFolders in Current User's Mailbox
'Or unitialized array if error occurs

'Because it returns an array, it is for VB6 only.
'Change to return a variant or a delimited list for
'previous versions of vb

'EXAMPLE:
'Dim sArray() As String
'Dim ictr As Integer

'sArray = OutlookFolderNames("Public Folders'") Mailbox - Doe, John")
'On Error Resume Next
'For ictr = 0 To UBound(sArray)
 '   Debug.Print sArray(ictr)
'Next

'*********************************************************

Dim sArray() As String
Dim oMAPI As Outlook.NameSpace
Dim oParentFolder As Outlook.MAPIFolder
Dim i As Integer
Dim iElement As Integer

ReDim sArray(0) As String
    
On Error GoTo ErrorHandler
    
Set oMAPI = GetObject("", "Outlook.application").GetNamespace("MAPI")
Set oParentFolder = oMAPI.Folders(MailboxName)
If oParentFolder.Folders.Count Then
             
  For i = 1 To oParentFolder.Folders.Count
    If Trim(oParentFolder.Folders(i).Name) <> "" Then
        iElement = IIf(sArray(0) = "", 0, UBound(sArray) + 1)
        ReDim Preserve sArray(iElement) As String
        sArray(iElement) = oParentFolder.Folders(i).Name
    End If
  Next i
Else
  
  sArray(0) = oParentFolder.Name

End If

OutlookFolderNames = sArray
Set oMAPI = Nothing
Exit Function

ErrorHandler:
    Set oMAPI = Nothing
End Function

Private Sub myItem_Send(Cancel As Boolean)
    myItem.ExpiryTime = #2/2/2020 4:00:00 PM#
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Public Sub Timer1_Timer()
Form3.Label1.Caption = "Chk OutLook Email in " + Trim(Str(Timer4)) + " Secs" + vbCrLf + "Press an Start Now.."
If Timer4 <= 0 Then
Form3.Label1.Caption = "Chk OutLook Email Now - Zero - Now."
Form3.Label1.Refresh
Timer1.Enabled = False
Call Command2_Click
End If
Timer4 = Timer4 - 1

If Dir$(App.Path + "\Timer.txt") <> "" Then
Timer4 = 90
Kill App.Path + "\Timer.txt"
End If

End Sub



Public Sub Frm4()

If Title$ <> "" Then
Form4.List1_Click
Exit Sub
End If

    t1$ = Str(Rrr) + " - of - "
    t2$ = Str(Rrr2) + " - "
    t3$ = Format$(MyItemsFrm4.ReceivedTime, "DD-MM-YY HH:MM:SS") + " - "
    t4$ = "SendName- " + MyItemsFrm4.SenderName + " - "
    t5$ = "SendEmail- " + MyItemsFrm4.SenderEmailAddress + " - "
    t6$ = "To- " + MyItemsFrm4.To + " - "
    t7$ = "SubJ- " + MyItemsFrm4.Subject


Call Form4.List1_Click
Set MyItemsFrm4 = Nothing

End Sub


Sub Capture_Folders()


'--------------------------------------------------------
'--------------------------------------------------------
'--------------------------------------------------------
'--------------------------------------------------------
'--------------------------------------------------------
'--------------------------------------------------------
'--------------------------------------------------------
'--------------------------------------------------------
'--------------------------------------------------------
'--------------------------------------------------------

Set oFolder = oNameSpace.GetDefaultFolder(olFolderInbox)

'Set oFolder = oNameSpace.PickFolder
'MsgBox oNameSpace.PickFolder

Dim exCnt As Integer
Me.Command2.Enabled = False
zag = 0


Set MyItems = oFolder.Items
  
Call MyItems.Sort("[" & MailFieldToSort & "]", True)

t$ = MyItems.Item(1).Subject

'For Each oMailitem In oFolder.Items
'    With oMailitem

'For Each MyItems In MyItems.Items
'    With MyItems
ert$ = Dir1.Path & "\" & Trim(MyItems.Parent) & "~~"
'MkDir (ert$)
            
'Open "D:\# MY DOCS\# 01 My Documents\Michelle.txt" For Output As #3
            
For Rrr = 1 To MyItems.Count
DoEvents
zag = zag + 1
            
'          Look in Object Brower for info on oMailitem
            
           qtag = 0
           qws$ = ""

'                If InStr(oMailItem.Attachments.Item(rrr).FileName, "0807") Then Stop
                
                
                    'If rrr > 1 Then qws$ = " (" + Trim(Str(rrr)) + ")"
                
                    'If qtag > 0 Then qws$ = " (" + Trim(Str(rrr + qtag)) + ")"
                
                    'ert$ = Dir1.Path & "\" & MyItems.Parent & "~~"
                    ert$ = Dir1.Path '& "\" & MyItems.Parent & "~~"
                    
                    UHT$ = MyItems.Item(Rrr).Subject
                
                
                    For r = 1 To Len(UHT$)
                        tyh$ = Mid$(UHT$, r, 1)
                        If tyh$ = """" Then Mid$(UHT$, r, 1) = "-"
                        If tyh$ = ">" Then Mid$(UHT$, r, 1) = "-"
                        If tyh$ = "<" Then Mid$(UHT$, r, 1) = "-"
                        If tyh$ = ":" Then Mid$(UHT$, r, 1) = "-"
                        If tyh$ = "\" Then Mid$(UHT$, r, 1) = "-"
                        If tyh$ = "/" Then Mid$(UHT$, r, 1) = "-"
                        If tyh$ = "?" Then Mid$(UHT$, r, 1) = "-"
                        If tyh$ = "*" Then Mid$(UHT$, r, 1) = "-"
                        If tyh$ = "|" Then Mid$(UHT$, r, 1) = "-"
                    Next
                
                If Len(MyItems.Item(Rrr).SenderName) < 5 Then
                fr2$ = Trim(MyItems.Item(Rrr).SenderName)
                Else
                fr2$ = Trim(Mid$(MyItems.Item(Rrr).SenderName, 1, 4))
                End If
                
                qq = MyItems.Item(Rrr).ReceivedTime
                
                If qq - TimeSerial(2, 0, 0) > Now Then
                
                    MyItems.Item(Rrr).Delete
                
                End If
                
                tyh$ = Format$(qq, "YYYYMMDDHHMMSS")
                at$ = "Sz" + tyh$ + "-" + UCase$(Mid$(fr2$, 1, 1)) + Mid$(fr2$, 2, 3) + "-" + UHT$ + ".txt"
                'att2$ = att2$ + "Sz" + TYH$ + "-" + UCase$(Mid$(fr2$, 1, 1)) + Mid$(fr2$, 2, 3) + "-" + UHT$ + ".html"

                
'put back after michelle stuff or other stuff
'                MyItems.Item(rrr).SaveAs ert$ + "\" + at$, olTXT
                
a1$ = MyItems.Item(Rrr).SenderName
a2$ = MyItems.Item(Rrr).SenderEmailAddress
a3$ = Trim(MyItems.Item(Rrr).Subject)
a4$ = Trim(MyItems.Item(Rrr).Body)

att1 = InStr(a4$, "@")
att2 = InStrRev(a4$, "news:")
att3 = InStr(att2, a4$, ">")
att2 = att3

If att1 > 0 Then
   att4 = InStrRev(a4$, vbCrLf, att1)
   att5 = InStr(att1, a4$, vbCrLf) - att4
End If


rr$ = ""
If att1 > att2 And att2 > 0 Then
rr$ = Mid$(a4$, att4 + 2, att5)
End If

'GoTo jumpm

qq = MyItems.Item(Rrr).ReceivedTime
tyh$ = Format$(qq, "DDD DD MMM YYYY HH:MM:SSp")

uu$ = UCase("Get a 2 week free trial at")

If InStr(UCase(a4$), uu$) Then
    ww = InStr(UCase(a4$), uu$)
    tty$ = Space$(Len(uu$))
    Mid$(a4$, ww, Len(tty$)) = tty$
End If

uu$ = UCase("Saving Trust")
If InStr(UCase(a4$), uu$) Then
    ww = InStr(UCase(a4$), uu$)
    tty$ = Space$(Len(uu$))
    Mid$(a4$, ww, Len(tty$)) = tty$
End If

rr$ = ""
rr$ = rr$ + String$(Len(tyh$), "-") + vbCrLf
rr$ = rr$ + tyh$ + vbCrLf
rr$ = rr$ + String$(Len(tyh$), "-") + vbCrLf
rr$ = rr$ + "Subject - " + a3$ + vbCrLf
rr$ = rr$ + a4$


If rr$ <> "" Then rr2$ = rr2$ + Trim(rr$) + vbCrLf

jumpm:

If InStr(a5$, a2$) = 0 Then a5$ = a5$ + a2$ + "," + vbCrLf
                
                
'lstExtractionStatus.AddItem (a3$)
exCnt = exCnt + 1
lblExtract.Caption = exCnt & " Saved"
    
            
Next

'================
'Open "D:\# MY DOCS\# 01 My Documents\Michelle.txt" For Output As #3
'Print #3, a5$
'Print #3, ","
'Print #3, rr2$
'Close #3
'----------------
'================
'Open "D:\# MY DOCS\# 01 My Documents\Alt.Gods.txt" For Output As #3
'Print #3, "Total Msg's -- "; MyItems.Count; " ------------------------------"
'Print #3,
'Print #3, "Alt.Gods ------------- Date Order Oldest First ----"
'Print #3,
'Print #3, rr2$
'Close #3
'----------------
'================
'Open "D:\# MY DOCS\# 01 My Documents\Alt.Roids Newest.txt" For Output As #3
'Print #3, "Total Msg's -- "; MyItems.Count; " ------------------------------"
'Print #3,
'Print #3, "Alt.Roids ------------- Date Order Newest First ----"
'Print #3,
'Print #3, rr2$
'close #3
'----------------
Open "D:\# MY DOCS\# 01 My Documents\News Outlook All News.txt" For Output As #3
Print #3, "Total Msg's -- "; MyItems.Count; " ------------------------------"
Print #3,
Print #3, "News Papers Such ------------- Date Order Newest First ----"
Print #3,
Print #3, rr2$
Close #3
'----------------


'End With

'Next oMailitem

Set oMailitem = Nothing
Set oFolder = Nothing
Set oNameSpace = Nothing
Set oApp = Nothing
    
'Me.Command2.Enabled = True

End Sub


Public Sub xy()

Top = 500
Left = 100
Form1.Top = Top
Form1.Left = Left
Form2.Top = Top
Form2.Left = Left
Form3.Top = Top
Form3.Left = Left
Form4.Top = Top
Form4.Left = Left

End Sub
