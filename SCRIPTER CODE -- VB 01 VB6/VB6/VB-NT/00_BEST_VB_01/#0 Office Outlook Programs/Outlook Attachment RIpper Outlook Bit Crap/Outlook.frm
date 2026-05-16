VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form OutlookAttachmentRipper 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Outlook Attachment Ripper"
   ClientHeight    =   3630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9465
   Icon            =   "Outlook.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3630
   ScaleWidth      =   9465
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   2535
      TabIndex        =   18
      Top             =   165
      Width           =   2220
      Begin VB.DriveListBox DriveList 
         Height          =   315
         Left            =   60
         TabIndex        =   2
         Top             =   2520
         Width           =   2115
      End
      Begin VB.DirListBox DirList 
         Height          =   2340
         Left            =   60
         TabIndex        =   1
         Top             =   150
         Width           =   2100
      End
      Begin VB.CommandButton SetLocalFolder 
         Caption         =   "Set Folder"
         Height          =   240
         Left            =   45
         TabIndex        =   3
         Top             =   2865
         Width           =   2145
      End
      Begin VB.CommandButton ExtractEmail 
         Caption         =   "Process"
         Enabled         =   0   'False
         Height          =   240
         Left            =   45
         TabIndex        =   4
         Top             =   3105
         Width           =   2145
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Outlook Folders"
      Height          =   3360
      Left            =   75
      TabIndex        =   17
      Top             =   180
      Width           =   2445
      Begin MSComctlLib.TreeView Outlookflist 
         Height          =   3105
         Left            =   45
         TabIndex        =   0
         Top             =   195
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   5477
         _Version        =   393217
         Indentation     =   11
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "FolderImages"
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3360
      Left            =   4770
      TabIndex        =   14
      Top             =   165
      Width           =   4575
      Begin VB.CheckBox DelFromDeleted 
         Caption         =   "Delete From Deleted Items"
         Enabled         =   0   'False
         Height          =   330
         Left            =   2355
         TabIndex        =   11
         Top             =   1605
         Width           =   1980
      End
      Begin VB.TextBox FolderName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1425
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   150
         Width           =   3105
      End
      Begin VB.CheckBox Delete 
         Caption         =   "Delete Attachments (Only)   From Mail "
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   2355
         TabIndex        =   9
         Top             =   1140
         Width           =   2055
      End
      Begin VB.TextBox MailFolderName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1425
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   3105
      End
      Begin VB.CheckBox SaveAttachment 
         Caption         =   "Save Attachments"
         Height          =   225
         Left            =   2355
         TabIndex        =   7
         Top             =   870
         Width           =   2055
      End
      Begin VB.CheckBox SaveMail 
         Caption         =   "Save Mail Contents"
         Height          =   225
         Left            =   75
         TabIndex        =   8
         Top             =   885
         Width           =   2040
      End
      Begin VB.CheckBox Dbstore 
         Caption         =   "Store Details in Database (Time  Consuming) "
         Enabled         =   0   'False
         Height          =   375
         Left            =   75
         TabIndex        =   12
         Top             =   1575
         Width           =   2115
      End
      Begin VB.CheckBox Delete 
         Caption         =   "Delete Mails and Attachments"
         Enabled         =   0   'False
         Height          =   405
         Index           =   1
         Left            =   75
         TabIndex        =   10
         Top             =   1140
         Width           =   1980
      End
      Begin VB.TextBox MyDetail 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C00000&
         Height          =   1365
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   13
         Text            =   "Outlook.frx":0442
         Top             =   1965
         Width           =   3075
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   " Help Guide"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   3165
         TabIndex        =   20
         Top             =   2235
         Width           =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "Attachment Folder"
         Height          =   240
         Left            =   75
         TabIndex        =   16
         Top             =   180
         Width           =   1365
      End
      Begin VB.Label Label2 
         Caption         =   "Mail Folder"
         Height          =   240
         Left            =   75
         TabIndex        =   15
         Top             =   525
         Width           =   1365
      End
   End
   Begin MSComctlLib.ImageList FolderImages 
      Left            =   870
      Top             =   4725
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outlook.frx":04F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outlook.frx":0950
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView OutlookFolderList 
      Height          =   870
      Left            =   5715
      TabIndex        =   19
      Top             =   5775
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   1535
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Folder Name"
         Object.Width           =   2822
      EndProperty
   End
   Begin VB.Image BorderBottom 
      Height          =   15
      Left            =   0
      Picture         =   "Outlook.frx":0DA8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15
   End
   Begin VB.Image BorderRight 
      Height          =   15
      Left            =   0
      Picture         =   "Outlook.frx":0DE5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15
   End
   Begin VB.Image BorderLeft 
      Height          =   15
      Left            =   0
      Picture         =   "Outlook.frx":0E22
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15
   End
   Begin VB.Image Menu 
      Height          =   150
      Left            =   35
      Picture         =   "Outlook.frx":0E5F
      Top             =   0
      Width           =   150
   End
   Begin VB.Image CloseButton 
      Height          =   150
      Left            =   0
      Picture         =   "Outlook.frx":0EBB
      Top             =   0
      Width           =   150
   End
   Begin VB.Image Title 
      Height          =   150
      Left            =   0
      Picture         =   "Outlook.frx":0F19
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15
   End
End
Attribute VB_Name = "OutlookAttachmentRipper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OutlookApp As New outlook.Application
Public SubFolderList As outlook.MAPIFolder
Dim MynameSpace As outlook.NameSpace
Dim FolderList As outlook.MAPIFolder
Dim StoreFolder
Dim ListX As ListItem
Dim NodeX As Node
Public FullLogCounter

Dim MoveFlag
Dim OldX
Dim OldY

Private Sub DirList_Change()
    StoreFolder = DirList.Path & "\"
    If Right(DirList.Path, 1) = "\" Then StoreFolder = DirList.Path
    FolderName.Text = StoreFolder
End Sub

Private Sub Form_Resize()
    Title.Width = Me.Width
    CloseButton.Left = Me.Width - 180
    BorderLeft.Height = Me.Height
    BorderRight.Left = Me.Width - 8
    BorderRight.Height = Me.Height
    BorderBottom.Top = Me.Height - 8
    BorderBottom.Width = Me.Width
End Sub

Private Sub CloseButton_Click()
    End
End Sub

Private Sub Label3_Click()
MsgBox "Only Seems to work on inbox select two top options an one down on the right" + vbCr + "Doesnt seem to save images just good at stripping them"
End Sub

Private Sub Menu_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub Title_DblClick()
'    If me.WindowState = vbMaximized Then
'        me.WindowState = vbNormal
'    Else
'        me.WindowState = vbMaximized
'    End If
End Sub

Private Sub Title_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveFlag = 1
    OldX = X
    OldY = Y
End Sub

Private Sub Title_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MoveFlag = 1 Then
        Me.Left = X + Me.Left - OldX
        Me.Top = Y + Me.Top - OldY
    End If
End Sub

Private Sub Title_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveFlag = 0
End Sub




Private Sub Delete_Click(Index As Integer)
    If Index = 1 Then
        If Delete(1).Value = 1 Then
            Delete(0).Value = 1
            Delete(0).Enabled = False
        Else
            Delete(0).Value = 0
            Delete(0).Enabled = True
        End If
    End If
        
End Sub

Private Sub DriveList_Change()
    DirList.Path = DriveList.Drive
End Sub

Private Sub ExtractEmail_Click()
    
    'Set MynameSpace = OutlookApp.GetNamespace("MAPI")

    'Set NameSpace = NameSpace.PickFolder
    
    If Not SubFolderList Is Nothing Then
        Load OutlookRipperLog
        OutlookRipperLog.Show 1
    Else
        MsgBox "Please Select an Outlook Folder ", vbOKOnly & vbInformation
    End If
End Sub

Private Sub Form_Load()
    DirList.Path = "D:\# MY DOCS\# 01 My DocumentsEmail Bits\Email Archive\Outlook 2000\Outlook Ripper Sent Items"
    FullLogCounter = 1
    StoreFolder = "D:\TEMP\"
    
    

    FolderName.Text = StoreFolder
    MailFolderName.Text = "Not Selected"
    Set MynameSpace = OutlookApp.GetNamespace("MAPI")
    Set FolderList = MynameSpace.GetDefaultFolder(olFolderInbox)
    Set NodeX = Outlookflist.Nodes.Add(, , "Root", FolderList.Name, 1)
    Set SubFolderList = MynameSpace.GetDefaultFolder(olFolderInbox)
    Call GetFolderListinTree(FolderList)
    
    Set SubFolderList = Nothing
    'Set SubFolderList = olNamespace.PickFolder

End Sub

Sub SaveFiles(flist As outlook.MAPIFolder)

Dim ListFilename As outlook.MailItem
Dim RecurFolder As outlook.MAPIFolder
Dim Counter
Dim LeftMostChar
Dim Extension
Dim IsdirExist
Dim IsExtDirExist
Dim Finalfolder
Dim FinalFileName
Dim AttachmentFilename
Dim NumberofMails

If flist.Folders.Count > 0 Then
    For Each RecurFolder In flist.Folders
        Call SaveFiles(RecurFolder)
    Next
End If

NumberofMails = flist.Items.Count
While NumberofMails <> 0
    DoEvents
    Set ListFilename = flist.Items(NumberofMails)
    
    If SaveMail.Value = 1 Then
        IsdirExist = Dir(StoreFolder & "Mails", vbDirectory)
        If Len(Trim(IsdirExist)) = 0 Then
            MkDir StoreFolder & "Mails"
            Set ListX = OutlookRipperLog.LogList.ListItems.Add(FullLogCounter, FullLogCounter & "Log", "Creating Folder : " & StoreFolder & "Mails")
            FullLogCounter = FullLogCounter + 1
        End If
        ChDir (StoreFolder & "Mails")
        Set ListX = OutlookRipperLog.LogList.ListItems.Add(FullLogCounter, FullLogCounter & "Log", "Changing Folder : " & StoreFolder & "Mails\")
        FullLogCounter = FullLogCounter + 1
        FinalFileName = ReplaceText(ListFilename.SenderName & Date & Time & NumberofMails & ".txt", "[\\\/:*?""<>|]", "")
        ListFilename.SaveAs StoreFolder & "Mails" & "\" & FinalFileName, 0
        Set ListX = OutlookRipperLog.LogList.ListItems.Add(FullLogCounter, FullLogCounter & "Log", "Saving Mail : " & StoreFolder & "Mails\" & FinalFileName)
        FullLogCounter = FullLogCounter + 1
    End If
    
    If SaveAttachment.Value = 1 Then
        If ListFilename.Attachments.Count > 0 Then
            Counter = ListFilename.Attachments.Count
            While Counter <> 0
                AttachmentFilename = ListFilename.Attachments.Item(Counter).FileName
                LeftMostChar = Left(AttachmentFilename, 1)
                LeftMostChar = LeftMostChar & "-Ripped"
                    IsdirExist = Dir(StoreFolder & LeftMostChar, vbDirectory)
                    If Len(Trim(IsdirExist)) = 0 Then
                        MkDir StoreFolder & LeftMostChar
                        Set ListX = OutlookRipperLog.LogList.ListItems.Add(FullLogCounter, FullLogCounter & "Log", "Creating Folder : " & StoreFolder & LeftMostChar)
                        FullLogCounter = FullLogCounter + 1
                    End If
                    ChDir (StoreFolder & LeftMostChar)
                    Set ListX = OutlookRipperLog.LogList.ListItems.Add(FullLogCounter, FullLogCounter & "Log", "Changing folder : " & StoreFolder & LeftMostChar)
                    FullLogCounter = FullLogCounter + 1
                    Extension = Mid(AttachmentFilename, InStrRev(AttachmentFilename, ".") + 1, Len(AttachmentFilename) - InStrRev(AttachmentFilename, ".") + 1)
                    IsExtDirExist = Dir(StoreFolder & LeftMostChar & "\" & Extension, vbDirectory)
                    If Len(Trim(IsExtDirExist)) = 0 Then
                         MkDir StoreFolder & LeftMostChar & "\" & Extension
                         Set ListX = OutlookRipperLog.LogList.ListItems.Add(FullLogCounter, FullLogCounter & "Log", "Creating Folder : " & StoreFolder & LeftMostChar & "\" & Extension)
                         FullLogCounter = FullLogCounter + 1
                    End If
                    ChDir (StoreFolder & LeftMostChar & "\" & Extension)
                    Set ListX = OutlookRipperLog.LogList.ListItems.Add(FullLogCounter, FullLogCounter & "Log", "Changing Folder : " & StoreFolder & LeftMostChar & "\" & Extension)
                    FullLogCounter = FullLogCounter + 1
                    
                    Finalfolder = StoreFolder & LeftMostChar & "\" & Extension & "\"
                    FinalFileName = Left(AttachmentFilename, InStrRev(AttachmentFilename, ".") - 1) & ReplaceText(Date, "[\\\/:*?""<>|]", "") & ReplaceText(Time, "[\\\/:*?""<>|]", "") & Counter & "." & Extension
                    ListFilename.Attachments.Item(Counter).SaveAsFile Finalfolder & FinalFileName
                    Set ListX = OutlookRipperLog.LogList.ListItems.Add(FullLogCounter, FullLogCounter & "Log", "Saving file : " & FinalFileName)
                    FullLogCounter = FullLogCounter + 1
                    
                    If Delete(0).Value = 1 Then
                        ListFilename.Attachments.Item(Counter).Delete
                        ListFilename.Save
                        Set ListX = OutlookRipperLog.LogList.ListItems.Add(FullLogCounter, FullLogCounter & "Log", "Deteleting Attachment : " & Counter)
                        FullLogCounter = FullLogCounter + 1
                    End If
                    Counter = Counter - 1
                    ChDir (StoreFolder)
            Wend
         End If
         
        If Delete(1).Value = 1 Then
            Set ListX = OutlookRipperLog.LogList.ListItems.Add(FullLogCounter, FullLogCounter & "Log", "Deteleting Mail : " & ListFilename.Subject)
            FullLogCounter = FullLogCounter + 1
            ListFilename.Delete
        End If
    End If
    NumberofMails = NumberofMails - 1
Wend
End Sub

Sub GetFolderList(fldFolder As outlook.MAPIFolder)
   Dim objItem As outlook.MAPIFolder
    If fldFolder.Folders.Count > 0 Then
      For Each objItem In fldFolder.Folders
        Set ListX = OutlookFolderList.ListItems.Add(, objItem.Name, objItem.Name)
      Next objItem
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    OutlookApp.Quit
End Sub

Private Sub Outlookflist_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim SubObjitem As outlook.MAPIFolder
    Dim Splittedvalue
    Dim Counter
    
    Set SubFolderList = MynameSpace.GetDefaultFolder(olFolderInbox)
    
    If SaveMail.Value = 1 Or SaveAttachment.Value = 1 Then
        ExtractEmail.Enabled = True
    Else
        ExtractEmail.Enabled = False
    End If
    
    If Node.Children = 0 Then
        Splittedvalue = Split(Node.FullPath, "\")
    
        Set SubFolderList = MynameSpace.GetDefaultFolder(olFolderInbox)
        For Counter = 1 To UBound(Splittedvalue)
            Set SubFolderList = SubFolderList.Folders(Splittedvalue(Counter))
        Next
    
        For Each SubObjitem In SubFolderList.Folders
            Set NodeX = Outlookflist.Nodes.Add(Node.Key, tvwChild, SubObjitem.Name, SubObjitem.Name, 1)
        Next
        Node.Expanded = True
        Node.ExpandedImage = 2
    End If

End Sub

Private Sub OutlookFolderList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set FolderList = MynameSpace.GetDefaultFolder(olFolderInbox)
    Set FolderList = FolderList.Folders(Item.Key)
End Sub

Private Sub SaveAttachment_Click()
    If SaveMail.Value = 1 And SaveAttachment.Value = 1 Then
        Delete(1).Enabled = True
    Else
        Delete(1).Enabled = False
    End If
    
    If SaveMail.Value = 1 Or SaveAttachment.Value = 1 Then
        ExtractEmail.Enabled = True
    Else
        ExtractEmail.Enabled = False
    End If
    
    If SaveAttachment.Value = 1 Then
        Delete(0).Enabled = True
    Else
        Delete(0).Enabled = False
    End If
End Sub

Private Sub SaveMail_Click()
    If SaveMail.Value = 1 Then
        MailFolderName.Text = StoreFolder & "Mails"
    Else
        MailFolderName.Text = "Not Selected"
    End If
        
    If SaveMail.Value = 1 Or SaveAttachment.Value = 1 Then
        ExtractEmail.Enabled = True
    Else
        ExtractEmail.Enabled = False
    End If
        
    If SaveMail.Value = 1 And SaveAttachment.Value = 1 Then
        Delete(1).Enabled = True
    Else
        Delete(1).Enabled = False
    End If
End Sub

Private Sub SetLocalFolder_Click()
    StoreFolder = DirList.Path & "\"
    If Right(DirList.Path, 1) = "\" Then StoreFolder = DirList.Path
    FolderName.Text = StoreFolder
End Sub


Sub GetFolderListinTree(fldFolder As outlook.MAPIFolder)
   Dim objItem            As outlook.MAPIFolder
    If fldFolder.Folders.Count > 0 Then
      For Each objItem In fldFolder.Folders
        Set NodeX = Outlookflist.Nodes.Add("Root", tvwChild, objItem.Name, objItem.Name, 1)
      Next objItem
    End If
End Sub

Function ReplaceText(OrgStr, patrn, replStr)
  Dim regEx, str1               ' Create variables.
  str1 = OrgStr
  Set regEx = New RegExp             ' Create regular expression.
  regEx.Pattern = patrn                 ' Set pattern.
  regEx.IgnoreCase = True            ' Make case insensitive.
  regEx.Global = True
  ReplaceText = regEx.Replace(str1, replStr)  ' Make replacement.
End Function
