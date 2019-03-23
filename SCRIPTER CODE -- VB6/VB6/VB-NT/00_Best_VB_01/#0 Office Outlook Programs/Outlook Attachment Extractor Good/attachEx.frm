VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Outlook Spam an File in Future Deleter"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   Icon            =   "attachEx.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   4890
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "Unet Posts Into Text"
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
      Left            =   735
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5805
      Width           =   3210
   End
   Begin VB.CommandButton Extract_Work_UNet_Folder 
      BackColor       =   &H000080FF&
      Caption         =   "Extract Work Folder..."
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
      Left            =   735
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5445
      Width           =   3210
   End
   Begin VB.CommandButton Email_Addys 
      BackColor       =   &H000080FF&
      Caption         =   "Extract Email Addys"
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
      Left            =   735
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5085
      Width           =   3210
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Delete Attachments"
      Height          =   270
      Left            =   1515
      TabIndex        =   8
      Top             =   4425
      Width           =   1785
   End
   Begin VB.ListBox lstExtractionStatus 
      Enabled         =   0   'False
      Height          =   1620
      Left            =   135
      TabIndex        =   5
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
      Caption         =   "Start EXtraction Orginal Code..."
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
      Left            =   735
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4725
      Width           =   3210
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   900
      TabIndex        =   12
      Top             =   7110
      Width           =   2640
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   915
      TabIndex        =   11
      Top             =   6735
      Width           =   2625
   End
   Begin VB.Label lblExtract 
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   2085
      TabIndex        =   7
      Top             =   2535
      Width           =   2610
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
      TabIndex        =   6
      Top             =   2535
      Width           =   1920
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
Dim MyItems

Const MailFieldToSort = "Received"       'language dependant

Private Sub Command1_Click()
On Error Resume Next
    
Me.lblExtract.Caption = ""
Me.lstExtractionStatus.Enabled = True
Me.lstExtractionStatus.Clear

Dim oApp As New Outlook.Application
Dim oNameSpace As NameSpace
Dim oFolder As MAPIFolder
Dim oMailitem As Object
Dim sMessage As String

Set oApp = New Outlook.Application
Set oNameSpace = oApp.GetNamespace("MAPI")
'Set oFolder = oNameSpace.GetDefaultFolder(olFolderInbox)
oNameSpace.PickFolder
'msgbox onamespace.PickFolder.

Set oFolder = oNameSpace.PickFolder
Dim exCnt As Integer
Me.Command1.Enabled = False

'Set myItems = oFolder.Items
'Call oFolder.Items.Sort("[" & CalendarFieldToSort & "]", True)

For Each oMailitem In oFolder.Items
    With oMailitem
        If oMailitem.Attachments.Count > 0 Then
            
            
            
            
'          Look in Object Brower for info on oMailitem
            
            
            
            qtag = 0
            qws$ = ""
            For rrr = oMailitem.Attachments.Count To 1 Step -1
'                If InStr(oMailItem.Attachments.Item(rrr).FileName, "0807") Then Stop
                
                
                If rrr > 1 Then qws$ = " (" + Trim(Str(rrr)) + ")"
                
                Do
                    If qtag > 0 Then qws$ = " (" + Trim(Str(rrr + qtag)) + ")"
                
                    ert$ = Dir1.Path & "\" & _
                    oMailitem.Attachments.Item(rrr).Parent & "~~"
                    
                    UHT$ = oMailitem.Attachments.Item(rrr).FileName
                    uyt = InStrRev(UHT$, ".")
                    ert$ = ert$ + Mid$(UHT$, 1, uyt - 1) + qws$ + Mid$(UHT$, uyt)
                
                
                    For R = Len(Dir1.Path) + 2 To Len(ert$)
                        tyh$ = Mid$(ert$, R, 1)
                        If tyh$ = "|" Then Mid$(ert$, R, 1) = "-"
                        If tyh$ = ">" Then Mid$(ert$, R, 1) = "-"
                        If tyh$ = "<" Then Mid$(ert$, R, 1) = "-"
                        If tyh$ = ":" Then Mid$(ert$, R, 1) = "-"
                        If tyh$ = "\" Then Mid$(ert$, R, 1) = "-"
                        If tyh$ = "/" Then Mid$(ert$, R, 1) = "-"
                        If tyh$ = "?" Then Mid$(ert$, R, 1) = "-"
                        If tyh$ = "*" Then Mid$(ert$, R, 1) = "-"
                        If tyh$ = """" Then Mid$(ert$, R, 1) = "-"
                    Next
                
                    DoEvents
                
                    If Dir$(ert$) = "" Or qtag > 10 Then Exit Do
                    qtag = qtag + 1
                Loop Until 1 = 1
                
                perd$ = oMailitem.Attachments.Item(rrr).FileName
                oMailitem.Attachments.Item(rrr).SaveAsFile ert$
                'ee$ = oMailitem.NoteItem.Item(rrr).Body
                
                
                oMailitem.Attachments.Item(rrr).SaveAsFile ert$
                'oMailItem.Attachments
                    
                If Check1.Value = vbChecked Then
                oMailitem.Attachments.Item(rrr).Delete
                oMailitem.Save
                End If
                        
                DoEvents
                lstExtractionStatus.AddItem (oMailitem.Attachments.Item(rrr).Parent)
                exCnt = exCnt + 1
                lblExtract.Caption = exCnt & " extracted"
 
           
            
            
            
            Next
        End If
    End With
Next oMailitem

Set oMailitem = Nothing
Set oFolder = Nothing
Set oNameSpace = Nothing
Set oApp = Nothing
    
    
Me.Command1.Enabled = True
End Sub


Sub KillDupes()

Open App.Path + "\New Folder\Email Addys Matt.txt" For Input As #1
Open App.Path + "\New Folder\Email Addys Matt out.txt" For Output As #2

Do
Line Input #1, h$
If InStr(tt$, h$) = 0 Then
tt$ = tt$ + h$
Print #2, h$
End If
Loop Until EOF(1)
Close #1, #2
End

End Sub
Sub KillDupes2()
'Forgien Domains

Open App.Path + "\New Folder\Email Addys Matt.txt" For Input As #1
Open App.Path + "\New Folder\Email Addys Matt out.txt" For Output As #2
Open App.Path + "\New Folder\Email Addys Matt out-Forgien.txt" For Output As #3
Dim b1(20, 20)
Do
Line Input #1, h$
If InStr(tt$, h$) = 0 Then
tt$ = tt$ + h$

C = 1
b1(C, 1) = ".com": b1(C, 2) = Len(b1(C, 1)) - 1
C = 2
b1(C, 1) = ".co.uk": b1(C, 2) = Len(b1(C, 1)) - 1
C = 3
b1(C, 1) = ".net": b1(C, 2) = Len(b1(C, 1)) - 1
C = 4
b1(C, 1) = ".gov": b1(C, 2) = Len(b1(C, 1)) - 1
C = 5
b1(C, 1) = ".ssl1.us": b1(C, 2) = Len(b1(C, 1)) - 1
C = 6
b1(C, 1) = ".org": b1(C, 2) = Len(b1(C, 1)) - 1
C = 7
b1(C, 1) = ".org.uk": b1(C, 2) = Len(b1(C, 1)) - 1
C = 8
b1(C, 1) = ".edu": b1(C, 2) = Len(b1(C, 1)) - 1
C = 9
b1(C, 1) = ".info": b1(C, 2) = Len(b1(C, 1)) - 1
C = 10
b1(C, 1) = ".net.uk": b1(C, 2) = Len(b1(C, 1)) - 1
C = 11
b1(C, 1) = ".nhs.uk": b1(C, 2) = Len(b1(C, 1)) - 1
C = 12
b1(C, 1) = ".ru": b1(C, 2) = Len(b1(C, 1)) - 1
C = 13
b1(C, 1) = ".ath.cx": b1(C, 2) = Len(b1(C, 1)) - 1
C = 14
b1(C, 1) = ".gr": b1(C, 2) = Len(b1(C, 1)) - 1
C = 15
b1(C, 1) = ".ltd.uk": b1(C, 2) = Len(b1(C, 1)) - 1
C = 16
b1(C, 1) = ".ac.uk": b1(C, 2) = Len(b1(C, 1)) - 1

t = 0
For R = 1 To C
If InStr(LCase(h$), LCase(b1(R, 1))) = Len(h$) - b1(R, 2) Then
    t = 1
End If
Next
If t = 1 Then Print #2, h$
If t = 0 Then Print #3, h$

End If

Loop Until EOF(1)
Close #1, #2
End

End Sub

Private Sub Command2_Click()
Call Extract_Work_UNet_Chunck_Text
End Sub

Private Sub Email_Addys_Click()
Exit Sub
'dont Run By Accident


'this will last time
'Extract txt for whole list of sz.group from outlook 2000

Call KillDupes2

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
'Dim oFolderJunk As MAPIFolder
Dim oMailitem As Object
Dim sMessage As String

'Set oApp = New Outlook.Application
'Set oApp2 = New Outlook.Application

Set oNameSpace = oApp.GetNamespace("MAPI")
Set oNameSpace2 = oApp2.GetNamespace("MAPI")
Set oNameSpace3 = oApp2.GetNamespace("MAPI")

Set oFolder = oNameSpace.PickFolder
'Set oFolder = oNameSpace.GetDefaultFolder(6) 'olFolderInBox
Set oFolder2 = oNameSpace2.GetDefaultFolder(23) 'olFolderJunkFolder
Set oFolder3 = oNameSpace3.GetDefaultFolder(3) 'olFolderDeleted

'r = olFolderDeleted
t$ = oFolder.Name

Set MyItems = oFolder.Items
Set MyItems2 = oFolder2.Items
Set MyItems3 = oFolder3.Items
Set myitems4 = oFolder2.Items


Call MyItems.Sort("[" & MailFieldToSort & "]", True)
Call MyItems2.Sort("[" & MailFieldToSort & "]", True)
Call MyItems3.Sort("[" & MailFieldToSort & "]", True)



Dim myCopiedItem As Outlook.MailItem


'GoTo jump1

Open App.Path + "\Email Addys -" + t$ + "-.txt" For Output As #1
Open App.Path + "\Email Addys -" + t$ + "-Detailed Version-.txt" For Output As #2

On Local Error Resume Next
Form2.Label2.Caption = Trim(Str$(MyItems.Count))
Form2.Show

For rrr = 1 To MyItems.Count
DoEvents

Form2.Label1.Caption = Trim(Str$(rrr))
t1$ = MyItems.Item(rrr).SenderEmailAddress
g1 = InStr(t1$, "@")
g$ = Mid$(t1$, g1)
If InStr(r2$, g$) = 0 Then
    r2$ = r2$ + " --" + g$

        
        
    Print #1, g$
End If

Print #2, g$;
Print #2, vbTab;
Print #2, MyItems.Item(rrr).SenderEmailAddress;
Print #2, vbTab;
Print #2, MyItems.Item(rrr).SenderName;
Print #2, vbTab;
Print #2, MyItems.Item(rrr).Subject;
Print #2, vbTab;
Print #2, MyItems.Item(rrr).ReceivedTime;
Print #2, vbTab;
qq = MyItems.Item(rrr).ReceivedTime
tyh$ = Format$(qq, "DDD DD MMM YYYY HH:MM:SSp")
Print #2, MyItems.Item(rrr).tyh$;
Print #2, vbTab;
Print #2, MyItems.Item(rrr).SenderName;
Print #2, vbTab;
Print #2, MyItems.Item(rrr).To;
Print #2, vbTab;
Print #2, MyItems.Item(rrr).Size;
Print #2, vbTab;
Print #2, t$; 'Folder Name
Print #2, vbTab;
Print #2, MyItems.Item(rrr).SentOn;
Print #2, vbTab;
qq = MyItems.Item(rrr).SentOn
tyh$ = Format$(qq, "DDD DD MMM YYYY HH:MM:SSp")
Print #2, tyh$



'Print #2, MyItems.Item(rrr).Senton
'Print #2, MyItems.Item(rrr).Size
'Print #2, MyItems.Item(rrr).ReceivedTime
'Exit For
'If rrr > 5 Then Exit For
Next

On Local Error GoTo 0

Close #1, #2

'a2$ = MyItems.Item(rrr).SenderEmailAddress


End
jump1:


For rrr = 1 To MyItems.Count

qq = MyItems.Item(rrr).ReceivedTime


If qq < Now Then Exit For

If qq - TimeSerial(2, 0, 0) > Now Then

    Set myitems4 = MyItems.Item(rrr)
    MyItems.Item(rrr).Delete



'MyItems.Item(rrr).Copy
'MyItems.Item(rrr).SaveAsFile ert$

'    Set myItems2 = oApp2.CreateItem(MyItems.Item(rrr))
    'myItems.item(rrr).Subject = "Speeches"
'    Set myCopiedItem = myItems2.Copy
    
    
    
    'Set myItem2 = myolApp.CreateItem(olMailItem)

    'Set myCopiedItem = MyItems.Item(rrr)
    'myCopiedItem.Copy ofolder2
    
    
    
   Do
    DoEvents
   Loop Until MyItems3.Item(1) = myitems4
   
       

    
    Set MyItems2 = MyItems3.Item(1).Copy



    MyItems2.Move oFolder2

'Dim myItems2 As Integer


    'myCopiedItem.Move myNewFolder

    MyItems3.Item(1).Delete




'MyItems.Item(rrr).Delete


End If

Next


End


Set oFolder = oNameSpace.GetDefaultFolder(olFolderInbox)
'oNameSpace.PickFolder

'Set oFolder = oNameSpace.PickFolder
'MsgBox oNameSpace.PickFolder

'oNameSpace.PickFolder
Dim exCnt As Integer
Extract_Work_UNet_Folder.Enabled = False
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
            
For rrr = 1 To MyItems.Count
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
                    
                    UHT$ = MyItems.Item(rrr).Subject
                
                
                    For R = 1 To Len(UHT$)
                        tyh$ = Mid$(UHT$, R, 1)
                        If tyh$ = """" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = ">" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = "<" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = ":" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = "\" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = "/" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = "?" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = "*" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = "|" Then Mid$(UHT$, R, 1) = "-"
                    Next
                
                If Len(MyItems.Item(rrr).SenderName) < 5 Then
                fr2$ = Trim(MyItems.Item(rrr).SenderName)
                Else
                fr2$ = Trim(Mid$(MyItems.Item(rrr).SenderName, 1, 4))
                End If
                
                qq = MyItems.Item(rrr).ReceivedTime
                
                If qq - TimeSerial(2, 0, 0) > Now Then
                
                    MyItems.Item(rrr).Delete
                
                End If
                
                tyh$ = Format$(qq, "YYYYMMDDHHMMSS")
                at$ = "Sz" + tyh$ + "-" + UCase$(Mid$(fr2$, 1, 1)) + Mid$(fr2$, 2, 3) + "-" + UHT$ + ".txt"
                'att2$ = att2$ + "Sz" + TYH$ + "-" + UCase$(Mid$(fr2$, 1, 1)) + Mid$(fr2$, 2, 3) + "-" + UHT$ + ".html"

                
'put back after michelle stuff or other stuff
'                MyItems.Item(rrr).SaveAs ert$ + "\" + at$, olTXT
                
a1$ = MyItems.Item(rrr).SenderName
a2$ = MyItems.Item(rrr).SenderEmailAddress
a3$ = Trim(MyItems.Item(rrr).Subject)
a4$ = Trim(MyItems.Item(rrr).Body)

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

qq = MyItems.Item(rrr).ReceivedTime
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
    
Extract_Work_UNet_Folder.Enabled = True

End Sub

Private Sub Drive1_Change()
    Me.Dir1.Path = Drive1.Drive

End Sub

Private Sub Email_Addys2_Click()
'Exit Sub
'dont Run By Accident

'This will get the news put in file


'this will last time
'Extract txt for whole list of sz.group from outlook 2000

'Call KillDupes2
'Not This Yet

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
'Dim oFolderJunk As MAPIFolder
Dim oMailitem As Object
Dim sMessage As String

'Set oApp = New Outlook.Application
'Set oApp2 = New Outlook.Application

Set oNameSpace = oApp.GetNamespace("MAPI")
Set oNameSpace2 = oApp2.GetNamespace("MAPI")
Set oNameSpace3 = oApp2.GetNamespace("MAPI")

Set oFolder = oNameSpace.PickFolder
'Set oFolder = oNameSpace.GetDefaultFolder(6) 'olFolderInBox
Set oFolder2 = oNameSpace2.GetDefaultFolder(23) 'olFolderJunkFolder
Set oFolder3 = oNameSpace3.GetDefaultFolder(3) 'olFolderDeleted

'r = olFolderDeleted
t$ = oFolder.Name

Set MyItems = oFolder.Items
Set MyItems2 = oFolder2.Items
Set MyItems3 = oFolder3.Items
Set myitems4 = oFolder2.Items


Call MyItems.Sort("[" & MailFieldToSort & "]", True)
Call MyItems2.Sort("[" & MailFieldToSort & "]", True)
Call MyItems3.Sort("[" & MailFieldToSort & "]", True)



Dim myCopiedItem As Outlook.MailItem


'Set oFolder = oNameSpace.GetDefaultFolder(olFolderInbox)
'oNameSpace.PickFolder

'Set oFolder = oNameSpace.PickFolder
'MsgBox oNameSpace.PickFolder

'oNameSpace.PickFolder
Dim exCnt As Integer
Extract_Work_UNet_Folder.Enabled = False
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
            
MyItemsCnt = MyItems.Count
For rrr = 1 To MyItems.Count
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
                    
                    UHT$ = MyItems.Item(rrr).Subject
                
                
                    For R = 1 To Len(UHT$)
                        tyh$ = Mid$(UHT$, R, 1)
                        If tyh$ = """" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = ">" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = "<" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = ":" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = "\" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = "/" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = "?" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = "*" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = "|" Then Mid$(UHT$, R, 1) = "-"
                    Next
                
                If Len(MyItems.Item(rrr).SenderName) < 5 Then
                fr2$ = Trim(MyItems.Item(rrr).SenderName)
                Else
                fr2$ = Trim(Mid$(MyItems.Item(rrr).SenderName, 1, 4))
                End If
                
                qq = MyItems.Item(rrr).ReceivedTime
                
                If qq - TimeSerial(2, 0, 0) > Now Then
                
                    'MyItems.Item(rrr).Delete
                
                End If
                
                tyh$ = Format$(qq, "YYYYMMDDHHMMSS")
                at$ = "Sz" + tyh$ + "-" + UCase$(Mid$(fr2$, 1, 1)) + Mid$(fr2$, 2, 3) + "-" + UHT$ + ".txt"
                'att2$ = att2$ + "Sz" + TYH$ + "-" + UCase$(Mid$(fr2$, 1, 1)) + Mid$(fr2$, 2, 3) + "-" + UHT$ + ".html"

                
'put back after michelle stuff or other stuff
'                MyItems.Item(rrr).SaveAs ert$ + "\" + at$, olTXT
                
a1$ = MyItems.Item(rrr).SenderName
a2$ = MyItems.Item(rrr).SenderEmailAddress
a3$ = Trim(MyItems.Item(rrr).Subject)
a4$ = Trim(MyItems.Item(rrr).Body)

att1 = InStr(a4$, "@")
att2 = InStrRev(a4$, "news:")
If att2 > 0 Then att2 = InStr(att2, a4$, ">")
'att2 = att3

If att1 > 0 Then
   att4 = InStrRev(a4$, vbCrLf, att1)
   att5 = InStr(att1, a4$, vbCrLf) - att4
End If


rr$ = ""
If att1 > att2 And att2 > 0 Then
rr$ = Mid$(a4$, att4 + 2, att5)
End If

'GoTo jumpm

qq = MyItems.Item(rrr).ReceivedTime
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
lblExtract.Caption = Str(exCnt) + " Of" + Str(MyItemsCnt) + " Saved"
Label3 = exCnt
Label3 = MyItemsCnt
            
Next

'----------------
t$ = oFolder.Name
Open "D:\# MY DOCS\# 01 My Documents\" + t$ + ".txt" For Output As #3
'Open "D:\# MY DOCS\# 01 My Documents\News Outlook All News.txt" For Output As #3

Print #3, "Total Msg's -- "; MyItems.Count; " ------------------------------"
Print #3,
Print #3, t$ + " ------------- Date Order Newest First ----"
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
    
Extract_Work_UNet_Folder.Enabled = True

End Sub

'Private Sub Drive1_Change()
    
'    Me.Dir1.Path = Drive1.Drive

'End Sub

Private Sub Extract_Work_UNet_Folder_Click()





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

Set oMAPI = GetObject("", "Outlook.application").GetNamespace("MAPI")
'Set oParentFolder = oMAPI.Folders("Personal Folders")
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

Set oPart2 = oMAPI.Folders("Personal Folders")
If oPart2.Folders.Count Then
  For i2 = 1 To oPart2.Folders.Count
    If Trim(oPart2.Folders(i2).Name) = "00-Future spam" Then
        Exit For
    End If
  Next i2
End If

On Local Error Resume Next

Set oFolder5 = oParentFolder.Folders(i1) 'Neto
tt1$ = oFolder5.Name
Set oFolder8 = oPart2.Folders(i2) 'Future Spam
tt2$ = oFolder8.Name

'Set oFolder = oNameSpace.GetDefaultFolder(olFolderInbox) '6) 'olFolderInBox
Set oFolder = oNameSpace.PickFolder

Set MyItems = oFolder.Items


Set MyItems2 = oFolder2.Items
Set MyItems3 = oFolder3.Items
Set myitems4 = myolApp.CreateItem(olMailItem) 'Just Dummy

'Set myItem = myolApp.CreateItem(olMailItem)


Set myitems5 = oFolder5.Items 'NetOMatic
Set myitems8 = oFolder8.Items 'Future Spam


Call MyItems.Sort("[" & MailFieldToSort & "]", True)
Call MyItems2.Sort("[" & MailFieldToSort & "]", True)
Call MyItems3.Sort("[" & MailFieldToSort & "]", True)



'Del All Del Folder at start
'For rrr = 1 To MyItems3.Count
'    MyItems3.Item(rrr).Delete
'Next


Dim myCopiedItem As Outlook.MailItem


'GoTo jump1

'Open App.Path + "\Email Addys -" + T$ + "-.txt" For Output As #1
'Open App.Path + "\Email Addys -" + T$ + "-Detailed Version-.txt" For Output As #2

On Local Error Resume Next
'Form1.Label3.Caption = Trim(Str$(MyItems.Count))
'Form2.Show

ttx = MyItems.Count

For rrr = 1 To MyItems.Count
DoEvents

Form1.Label3.Caption = Trim(Str(rrr)) + " of " + Trim(Str$(ttx))
Form1.Label4.Caption = MyItems.Item(rrr).ReceivedTime

't1$ = MyItems.Item(rrr).SenderEmailAddress
'g1 = InStr(t1$, "@")
'g$ = Mid$(t1$, g1)
'If InStr(r2$, g$) = 0 Then
'   r2$ = r2$ + " --" + g$

        
        
'    Print #1, g$
'End If

'Print #2, g$;
'Print #2, vbTab;

Print #2, "SenderEmailAddress " + MyItems.Item(rrr).SenderEmailAddress
'Print #2, vbTab;
Print #2, "SenderName " + MyItems.Item(rrr).SenderName
'Print #2, vbTab;
Print #2, "Subject " + MyItems.Item(rrr).Subject
'Print #2, vbTab;
Print #2, "ReceivedTime " + MyItems.Item(rrr).ReceivedTime
'Print #2, vbTab;
qq = MyItems.Item(rrr).ReceivedTime
tyh$ = Format$(qq, "DDD DD MMM YYYY HH:MM:SSp")
Print #2, " ReceivedTime" + MyItems.Item(rrr).tyh$
'Print #2, vbTab;
'Print #2, " " + MyItems.Item(rrr).SenderName;
'Print #2, vbTab;
Print #2, "Sent To " + MyItems.Item(rrr).To
'Print #2, vbTab;
Print #2, "Size " + MyItems.Item(rrr).Size
'Print #2, vbTab;
Print #2, "OutLook Folder Name " + t$ 'Folder Name
'Print #2, vbTab;
Print #2, "SentOn " + MyItems.Item(rrr).SentOn
'Print #2, vbTab;
qq = MyItems.Item(rrr).SentOn
tyh$ = Format$(qq, "DDD DD MMM YYYY HH:MM:SSp")
Print #2, "SentOn  " + tyh$



'Print #2, MyItems.Item(rrr).Senton
'Print #2, MyItems.Item(rrr).Size
'Print #2, MyItems.Item(rrr).ReceivedTime
'Exit For
'If rrr > 5 Then Exit For
Next

On Local Error GoTo 0

Close #1, #2

'a2$ = MyItems.Item(rrr).SenderEmailAddress


End
jump1:


For rrr = 1 To MyItems.Count

qq = MyItems.Item(rrr).ReceivedTime


If qq < Now Then Exit For

If qq - TimeSerial(2, 0, 0) > Now Then

    Set myitems4 = MyItems.Item(rrr)
    MyItems.Item(rrr).Delete



'MyItems.Item(rrr).Copy
'MyItems.Item(rrr).SaveAsFile ert$

'    Set myItems2 = oApp2.CreateItem(MyItems.Item(rrr))
    'myItems.item(rrr).Subject = "Speeches"
'    Set myCopiedItem = myItems2.Copy
    
    
    
    'Set myItem2 = myolApp.CreateItem(olMailItem)

    'Set myCopiedItem = MyItems.Item(rrr)
    'myCopiedItem.Copy ofolder2
    
    
    
   Do
    DoEvents
   Loop Until MyItems3.Item(1) = myitems4
   
       

    
    Set MyItems2 = MyItems3.Item(1).Copy



    MyItems2.Move oFolder2

'Dim myItems2 As Integer


    'myCopiedItem.Move myNewFolder

    MyItems3.Item(1).Delete




'MyItems.Item(rrr).Delete


End If

Next


End


Set oFolder = oNameSpace.GetDefaultFolder(olFolderInbox)
'oNameSpace.PickFolder

'Set oFolder = oNameSpace.PickFolder
'MsgBox oNameSpace.PickFolder

'oNameSpace.PickFolder
Dim exCnt As Integer
Extract_Work_UNet_Folder.Enabled = False
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
            
For rrr = 1 To MyItems.Count
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
                    
                    UHT$ = MyItems.Item(rrr).Subject
                
                
                    For R = 1 To Len(UHT$)
                        tyh$ = Mid$(UHT$, R, 1)
                        If tyh$ = """" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = ">" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = "<" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = ":" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = "\" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = "/" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = "?" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = "*" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = "|" Then Mid$(UHT$, R, 1) = "-"
                    Next
                
                If Len(MyItems.Item(rrr).SenderName) < 5 Then
                fr2$ = Trim(MyItems.Item(rrr).SenderName)
                Else
                fr2$ = Trim(Mid$(MyItems.Item(rrr).SenderName, 1, 4))
                End If
                
                qq = MyItems.Item(rrr).ReceivedTime
                
                If qq - TimeSerial(2, 0, 0) > Now Then
                'h
                    MyItems.Item(rrr).Delete
                
                End If
                
                tyh$ = Format$(qq, "YYYYMMDDHHMMSS")
                at$ = "Sz" + tyh$ + "-" + UCase$(Mid$(fr2$, 1, 1)) + Mid$(fr2$, 2, 3) + "-" + UHT$ + ".txt"
                'att2$ = att2$ + "Sz" + TYH$ + "-" + UCase$(Mid$(fr2$, 1, 1)) + Mid$(fr2$, 2, 3) + "-" + UHT$ + ".html"

                
'put back after michelle stuff or other stuff
'                MyItems.Item(rrr).SaveAs ert$ + "\" + at$, olTXT
                
a1$ = MyItems.Item(rrr).SenderName
a2$ = MyItems.Item(rrr).SenderEmailAddress
a3$ = Trim(MyItems.Item(rrr).Subject)
a4$ = Trim(MyItems.Item(rrr).Body)

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

qq = MyItems.Item(rrr).ReceivedTime
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
'Open "D:\# MY DOCS\# 01 My Documents\News Outlook All News.txt" For Output As #3
'Print #3, "Total Msg's -- "; MyItems.Count; " ------------------------------"
'Print #3,
'Print #3, "News Papers Such ------------- Date Order Newest First ----"
'Print #3,
'Print #3, rr2$
'Close #3
'----------------
'----------------
xx$ = Format$(Now, " YYYY MM DD HH MM SS")
Open "D:\# MY DOCS\# 01 My Documents\00 Email Unet Posts Alt.Philo " + xx$ + ".txt" For Output As #3
Print #3, "Total Msg's -- "; MyItems.Count; " ------------------------------"
Print #3,
'Print #3, "News Papers Such ------------- Date Order Newest First ----"
Print #3, "Alt.Philo 2003 2007 ------------- Date Order Newest First ----"
Print #3,
Print #3, rr2$
Close #3
'----------------


Shell "notepad2 " + xx$

'End With

'Next oMailitem


Set oMailitem = Nothing
Set oFolder = Nothing
Set oNameSpace = Nothing
Set oApp = Nothing
    
Extract_Work_UNet_Folder.Enabled = True

End Sub

Private Sub Form_Load()
    
    Me.Drive1.Drive = "C"
    'Me.Dir1.Path = "c:\Art\My Pictures\Fractals"
    Me.Dir1.Path = "D:\# MY DOCS\# 01 My Documents"
    'Me.Dir1.Path = "D:\My Webs\MatthewLan.Com Web\LoveFolder\Alt.Sz_Borg_Roid_Judy\Alt_Sz_2003_Onwards_TxT"
    'Me.Dir1.Path = "D:\My Webs\MatthewLan.Com Web\LoveFolder"

    Form1.Show
    'Call Extract_Work_UNet_Folder_Click
    
    'Call Email_Addys2_Click


End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub






Private Sub Extract_Work_UNet_Chunck_Text()





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

Set oMAPI = GetObject("", "Outlook.application").GetNamespace("MAPI")
'Set oParentFolder = oMAPI.Folders("Personal Folders")
Set oParentFolder = oMAPI.Folders("Personal Folders")
If oParentFolder.Folders.Count Then
  For i1 = 1 To oParentFolder.Folders.Count
    'Set oFolder5 = oParentFolder.Folders(i1) 'Neto
    'tt1$ = oFolder5.Name
    If Trim(oParentFolder.Folders(i1).Name) = "02 RoidsRim Mine" Then
        Exit For
    End If
  Next i1
End If

Set oPart2 = oMAPI.Folders("Personal Folders")
If oPart2.Folders.Count Then
  For i2 = 1 To oPart2.Folders.Count
    If Trim(oPart2.Folders(i2).Name) = "02 RoidsRim Mine Done" Then
        Exit For
    End If
  Next i2
End If

On Local Error Resume Next

Set oFolder5 = oParentFolder.Folders(i1) 'Neto
tt1$ = oFolder5.Name
Set oFolder8 = oPart2.Folders(i2) 'Future Spam
tt2$ = oFolder8.Name

'Set oFolder = oNameSpace.GetDefaultFolder(olFolderInbox) '6) 'olFolderInBox
'Set oFolder = oNameSpace.PickFolder

Set MyItems = oFolder.Items


Set MyItems2 = oFolder2.Items
Set MyItems3 = oFolder3.Items
Set myitems4 = myolApp.CreateItem(olMailItem) 'Just Dummy

'Set myItem = myolApp.CreateItem(olMailItem)


Set myitems5 = oFolder5.Items 'NetOMatic
Set myitems8 = oFolder8.Items 'Future Spam


Call MyItems.Sort("[" & MailFieldToSort & "]", True)
Call MyItems2.Sort("[" & MailFieldToSort & "]", True)
Call MyItems3.Sort("[" & MailFieldToSort & "]", True)



'Del All Del Folder at start
'For rrr = 1 To MyItems3.Count
'    MyItems3.Item(rrr).Delete
'Next


Dim myCopiedItem As Outlook.MailItem


GoTo jump1

'Open App.Path + "\Email Addys -" + T$ + "-.txt" For Output As #1
'Open App.Path + "\Email Addys -" + T$ + "-Detailed Version-.txt" For Output As #2

On Local Error Resume Next
'Form1.Label3.Caption = Trim(Str$(MyItems.Count))
'Form2.Show

ttx = MyItems.Count

For rrr = 1 To MyItems.Count
DoEvents

Form1.Label3.Caption = Trim(Str(rrr)) + " of " + Trim(Str$(ttx))
Form1.Label4.Caption = MyItems.Item(rrr).ReceivedTime

't1$ = MyItems.Item(rrr).SenderEmailAddress
'g1 = InStr(t1$, "@")
'g$ = Mid$(t1$, g1)
'If InStr(r2$, g$) = 0 Then
'   r2$ = r2$ + " --" + g$

        
        
'    Print #1, g$
'End If

'Print #2, g$;
'Print #2, vbTab;

Print #2, "SenderEmailAddress " + MyItems.Item(rrr).SenderEmailAddress
'Print #2, vbTab;
Print #2, "SenderName " + MyItems.Item(rrr).SenderName
'Print #2, vbTab;
Print #2, "Subject " + MyItems.Item(rrr).Subject
'Print #2, vbTab;
Print #2, "ReceivedTime " + MyItems.Item(rrr).ReceivedTime
'Print #2, vbTab;
qq = MyItems.Item(rrr).ReceivedTime
tyh$ = Format$(qq, "DDD DD MMM YYYY HH:MM:SSp")
Print #2, " ReceivedTime" + MyItems.Item(rrr).tyh$
'Print #2, vbTab;
'Print #2, " " + MyItems.Item(rrr).SenderName;
'Print #2, vbTab;
Print #2, "Sent To " + MyItems.Item(rrr).To
'Print #2, vbTab;
Print #2, "Size " + MyItems.Item(rrr).Size
'Print #2, vbTab;
Print #2, "OutLook Folder Name " + t$ 'Folder Name
'Print #2, vbTab;
Print #2, "SentOn " + MyItems.Item(rrr).SentOn
'Print #2, vbTab;
qq = MyItems.Item(rrr).SentOn
tyh$ = Format$(qq, "DDD DD MMM YYYY HH:MM:SSp")
Print #2, "SentOn  " + tyh$



'Print #2, MyItems.Item(rrr).Senton
'Print #2, MyItems.Item(rrr).Size
'Print #2, MyItems.Item(rrr).ReceivedTime
'Exit For
'If rrr > 5 Then Exit For
Next

On Local Error GoTo 0

Close #1, #2

'a2$ = MyItems.Item(rrr).SenderEmailAddress


End
jump1:


For rrr = 1 To MyItems.Count

qq = MyItems.Item(rrr).ReceivedTime


If qq < Now Then Exit For

If qq - TimeSerial(2, 0, 0) > Now Then

    Set myitems4 = MyItems.Item(rrr)
'    MyItems.Item(rrr).Delete



'MyItems.Item(rrr).Copy
'MyItems.Item(rrr).SaveAsFile ert$

'    Set myItems2 = oApp2.CreateItem(MyItems.Item(rrr))
    'myItems.item(rrr).Subject = "Speeches"
'    Set myCopiedItem = myItems2.Copy
    
    
    
    'Set myItem2 = myolApp.CreateItem(olMailItem)

    'Set myCopiedItem = MyItems.Item(rrr)
    'myCopiedItem.Copy ofolder2
    
    
    
   Do
    DoEvents
   Loop Until MyItems3.Item(1) = myitems4
   
       

    
    Set MyItems2 = MyItems3.Item(1).Copy



    MyItems2.Move oFolder2

'Dim myItems2 As Integer


    'myCopiedItem.Move myNewFolder

    MyItems3.Item(1).Delete




'MyItems.Item(rrr).Delete


End If

Next


End


Set oFolder = oNameSpace.GetDefaultFolder(olFolderInbox)
'oNameSpace.PickFolder

'Set oFolder = oNameSpace.PickFolder
'MsgBox oNameSpace.PickFolder

'oNameSpace.PickFolder
Dim exCnt As Integer
Extract_Work_UNet_Folder.Enabled = False
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
            
For rrr = 1 To MyItems.Count
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
                    
                    UHT$ = MyItems.Item(rrr).Subject
                
                
                    For R = 1 To Len(UHT$)
                        tyh$ = Mid$(UHT$, R, 1)
                        If tyh$ = """" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = ">" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = "<" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = ":" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = "\" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = "/" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = "?" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = "*" Then Mid$(UHT$, R, 1) = "-"
                        If tyh$ = "|" Then Mid$(UHT$, R, 1) = "-"
                    Next
                
                If Len(MyItems.Item(rrr).SenderName) < 5 Then
                fr2$ = Trim(MyItems.Item(rrr).SenderName)
                Else
                fr2$ = Trim(Mid$(MyItems.Item(rrr).SenderName, 1, 4))
                End If
                
                qq = MyItems.Item(rrr).ReceivedTime
                
                If qq - TimeSerial(2, 0, 0) > Now Then
                'h
                    MyItems.Item(rrr).Delete
                
                End If
                
                tyh$ = Format$(qq, "YYYYMMDDHHMMSS")
                at$ = "Sz" + tyh$ + "-" + UCase$(Mid$(fr2$, 1, 1)) + Mid$(fr2$, 2, 3) + "-" + UHT$ + ".txt"
                'att2$ = att2$ + "Sz" + TYH$ + "-" + UCase$(Mid$(fr2$, 1, 1)) + Mid$(fr2$, 2, 3) + "-" + UHT$ + ".html"

                
'put back after michelle stuff or other stuff
'                MyItems.Item(rrr).SaveAs ert$ + "\" + at$, olTXT
                
a1$ = MyItems.Item(rrr).SenderName
a2$ = MyItems.Item(rrr).SenderEmailAddress
a3$ = Trim(MyItems.Item(rrr).Subject)
a4$ = Trim(MyItems.Item(rrr).Body)

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

qq = MyItems.Item(rrr).ReceivedTime
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
'Open "D:\# MY DOCS\# 01 My Documents\News Outlook All News.txt" For Output As #3
'Print #3, "Total Msg's -- "; MyItems.Count; " ------------------------------"
'Print #3,
'Print #3, "News Papers Such ------------- Date Order Newest First ----"
'Print #3,
'Print #3, rr2$
'Close #3
'----------------
'----------------
xx$ = Format$(Now, " YYYY MM DD HH MM SS")
Open "D:\# MY DOCS\# 01 My Documents\00 Email Unet Posts Alt.Philo Alt.SZ Roids " + xx$ + ".txt" For Output As #3
Print #3, "Total Msg's -- "; MyItems.Count; " ------------------------------"
Print #3,
'Print #3, "News Papers Such ------------- Date Order Newest First ----"
Print #3, "Alt.UNET ------------- Date Order Newest First ----"
Print #3,
Print #3, rr2$
Close #3
'----------------


Shell "notepad2 " + xx$

'End With

'Next oMailitem


Set oMailitem = Nothing
Set oFolder = Nothing
Set oNameSpace = Nothing
Set oApp = Nothing
    
Extract_Work_UNet_Folder.Enabled = True

End Sub



