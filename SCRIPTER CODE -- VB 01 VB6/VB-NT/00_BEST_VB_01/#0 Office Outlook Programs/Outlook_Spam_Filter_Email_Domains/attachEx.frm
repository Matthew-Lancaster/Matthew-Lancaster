VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Outlook Attachment Extractor"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4815
   Icon            =   "attachEx.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "Start Extraction of Text Unet ..."
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
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5550
      Width           =   3210
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Delete Attachments"
      Enabled         =   0   'False
      Height          =   270
      Left            =   135
      TabIndex        =   8
      Top             =   4890
      Width           =   1785
   End
   Begin VB.ListBox lstExtractionStatus 
      Enabled         =   0   'False
      Height          =   1620
      Left            =   135
      TabIndex        =   5
      Top             =   3255
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
      Enabled         =   0   'False
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
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5190
      Width           =   3210
   End
   Begin VB.Label Label5 
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   180
      TabIndex        =   12
      Top             =   2970
      Width           =   4380
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   3465
      TabIndex        =   11
      Top             =   5535
      Width           =   1245
   End
   Begin VB.Label Label3 
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2175
      TabIndex        =   10
      Top             =   2715
      Width           =   2400
   End
   Begin VB.Label lblExtract 
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2175
      TabIndex        =   7
      Top             =   2505
      Width           =   2400
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
      Height          =   195
      Left            =   135
      TabIndex        =   6
      Top             =   2520
      Width           =   1965
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
'Use Command 2 for Unet
'Set Path Manual Override

Dim exCnt As Long
Dim MyItems

Const MailFieldToSort = "Received"       'language dependant


Private Sub Command2_Click()
'this will
'Extract txt for whole list of sz.group from outlook 2000

Command2.Enabled = False

Dir1.Path = "D:\I\A.Philo\2005"

    
Me.lblExtract.Caption = ""
Me.lstExtractionStatus.Enabled = True
Me.lstExtractionStatus.Clear

Dim oApp As New Outlook.Application
Dim oNameSpace As NameSpace
Dim oFolder As MAPIFolder
Dim oMailitem As Object
Dim sMessage As String
'Set oNameSpace = oApp.GetNamespace("MAPI")
Set oApp = New Outlook.Application

Set oNameSpace = oApp.GetNamespace("MAPI")

'Set oFolder = oNameSpace.GetDefaultFolder(olFolderDeletedItems)

'oFolder.Delete
'Set oFolder = oNameSpace.GetDefaultFolder(olFolderInbox)
'oNameSpace.PickFolder

'Set oFolder = oNameSpace.PickFolder
'MsgBox oNameSpace.PickFolder

'tt$ = oNameSpace.PickFolder
'Set oFolder = "InBox"

'Set oFolder = oNameSpace.GetDefaultFolder(olFolderInbox)

'oNameSpace.PickFolder


Fol$ = "00 A.Phiol all 2003 - 2007 Copy"
Set oMAPI = GetObject("", "Outlook.application").GetNamespace("MAPI")
Set oParentFolder = oMAPI.Folders("Personal Folders")
If oParentFolder.Folders.Count Then
  For i1 = 1 To oParentFolder.Folders.Count
    'Set oFolder5 = oParentFolder.Folders(i1) 'Neto
    'tt1$ = oFolder5.Name
    If Trim(oParentFolder.Folders(i1).Name) = Fol$ Then
        Exit For
    End If
  Next i1
End If

Set oFolder = oParentFolder.Folders(i1)
tt1$ = oFolder.Name
Label5.Caption = "Extracting From Folder " + tt1$


On Error Resume Next











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


'Dir1.Path = "D:\I\A.Philo\" + Trim(Str(Year(MyItems.Item(rrr).ReceivedTime)))


'ert$ = Dir1.Path & "\" & Trim(MyItems.Parent) & "~~"


'MkDir (ert$)
            
'Open "D:\# MY DOCS\# 01 My Documents\Michelle.txt" For Output As #3
xxe = MyItems.Count
            
For rrr = MyItems.Count To 1 Step -1
DoEvents
zag = zag + 1
            
'          Look in Object Brower for info on oMailitem
            
           qtag = 0
           qws$ = ""

'                If InStr(oMailItem.Attachments.Item(rrr).FileName, "0807") Then Stop
                
                
                    'If rrr > 1 Then qws$ = " (" + Trim(Str(rrr)) + ")"
                
                    'If qtag > 0 Then qws$ = " (" + Trim(Str(rrr + qtag)) + ")"
                
                    'ert$ = Dir1.Path & "\" & MyItems.Parent & "~~"
                    Dir1.Path = "D:\I\A.Philo\" + Trim(Str(Year(MyItems.Item(rrr).ReceivedTime)))
                    
                    ert$ = Dir1.Path '& "\" & MyItems.Parent & "~~"
                    
                    UHT$ = MyItems.Item(rrr).Subject
                
                
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
                
                If Len(MyItems.Item(rrr).SenderName) < 5 Then
                fr2$ = Trim(MyItems.Item(rrr).SenderName)
                Else
                fr2$ = Trim(Mid$(MyItems.Item(rrr).SenderName, 1, 4))
                End If
                
                qq = MyItems.Item(rrr).ReceivedTime
                
                
                If qq - TimeSerial(2, 0, 0) > Now Then
                
                'MyItems.Item(rrr).Delete
                'MyItems.Item(rrr).
                
                'oMailitem.Attachments.Item(rrr).Delete
                
                End If
                
                'MyItems.Item(rrr).Delete
                
                tyh$ = Format$(qq, "YYYYMMDDHHMMSS")
                at$ = "Pi" + tyh$ + "-" + UCase$(Mid$(fr2$, 1, 1)) + Mid$(fr2$, 2, 3) + "-" + UHT$ + ".txt"
                'att2$ = att2$ + "Sz" + TYH$ + "-" + UCase$(Mid$(fr2$, 1, 1)) + Mid$(fr2$, 2, 3) + "-" + UHT$ + ".html"

On Local Error Resume Next
lstExtractionStatus.AddItem at$
If Err.Number > 0 Then lstExtractionStatus.Clear
On Local Error GoTo 0

lstExtractionStatus.ListIndex = lstExtractionStatus.ListCount - 1
'put back after michelle stuff or other stuff
                
                On Local Error Resume Next
                Err.Clear
                MyItems.Item(rrr).SaveAs ert$ + "\" + at$, olTXT
                On Local Error GoTo 0
                
                If Err.Number = 0 Then MyItems.Item(rrr).Delete
                
                DoEvents
'lstExtractionStatus.AddItem (a3$)
exCnt = exCnt + 1
'lblExtract.Caption = exCnt & " Saved"
lblExtract.Caption = Trim(Str(exCnt)) & " Of" + Str(xxe) + " Saved"
Label3.Caption = Format$(qq, "DD-MMM-YYYY HH:MM:SSa/p")
                   
    
    
            
Next



Set oMailitem = Nothing
Set oFolder = Nothing
Set oNameSpace = Nothing
Set oApp = Nothing
    
Me.Command2.Enabled = True

End Sub

Private Sub Drive1_Change()
    Me.Dir1.Path = Drive1.Drive

End Sub

Private Sub Form_Load()
    
    Me.Drive1.Drive = "C"
    'Me.Dir1.Path = "c:\Art\My Pictures\Fractals"
    Me.Dir1.Path = "D:\# MY DOCS\# 01 My Documents"
    'Me.Dir1.Path = "D:\My Webs\MatthewLan.Com Web\LoveFolder\Alt.Sz_Borg_Roid_Judy\Alt_Sz_2003_Onwards_TxT"
    'Me.Dir1.Path = "D:\My Webs\MatthewLan.Com Web\LoveFolder"
End Sub

Sub Old_Code_Dump()
GoTo jump12
                
                
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
                
jump12:
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

End Sub

Private Sub Label4_Click()
Label4.BackColor = &HFF00&
DoEvents
End
End Sub

Private Sub Command1xx_Click()
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
'oNameSpace.PickFolder
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
                
                
                    For r = Len(Dir1.Path) + 2 To Len(ert$)
                        tyh$ = Mid$(ert$, r, 1)
                        If tyh$ = "|" Then Mid$(ert$, r, 1) = "-"
                        If tyh$ = ">" Then Mid$(ert$, r, 1) = "-"
                        If tyh$ = "<" Then Mid$(ert$, r, 1) = "-"
                        If tyh$ = ":" Then Mid$(ert$, r, 1) = "-"
                        If tyh$ = "\" Then Mid$(ert$, r, 1) = "-"
                        If tyh$ = "/" Then Mid$(ert$, r, 1) = "-"
                        If tyh$ = "?" Then Mid$(ert$, r, 1) = "-"
                        If tyh$ = "*" Then Mid$(ert$, r, 1) = "-"
                        If tyh$ = """" Then Mid$(ert$, r, 1) = "-"
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



