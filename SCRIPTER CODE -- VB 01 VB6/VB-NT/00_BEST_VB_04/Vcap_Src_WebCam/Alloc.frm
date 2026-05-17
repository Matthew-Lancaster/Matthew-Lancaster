VERSION 5.00
Begin VB.Form frmAlloc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set File Size"
   ClientHeight    =   2445
   ClientLeft      =   1575
   ClientTop       =   4860
   ClientWidth     =   3465
   ControlBox      =   0   'False
   Icon            =   "Alloc.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   1800
      TabIndex        =   8
      Top             =   1950
      Width           =   960
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   360
      Left            =   720
      TabIndex        =   7
      Top             =   1950
      Width           =   960
   End
   Begin VB.TextBox txtAlloc 
      Height          =   330
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   3
      Text            =   "1"
      Top             =   1455
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "MBytes"
      Height          =   270
      Left            =   2685
      TabIndex        =   6
      Top             =   1485
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "MBytes"
      Height          =   270
      Left            =   2655
      TabIndex        =   5
      Top             =   1065
      Width           =   735
   End
   Begin VB.Label lblFreeDisk 
      Height          =   270
      Left            =   1830
      TabIndex        =   4
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Capture file size:"
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Top             =   1455
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Free disk space:"
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   1065
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the Amount of disk space to set aside for the capture file.  Existing video data in the file will be lost."
      Height          =   765
      Left            =   195
      TabIndex        =   0
      Top             =   210
      Width           =   3060
   End
End
Attribute VB_Name = "frmAlloc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private available As Long

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()

Call capFileAlloc(frmMain.capwnd, txtAlloc.Text * ONE_MEGABYTE)

Unload Me

End Sub

Private Sub Form_Load()

Dim capfilesize As Long

'Dim path As String ' MADE PUBLIC IN .BAS

Dim DNAME
Dim SubPath, TS, FS

Set FS = CreateObject("Scripting.FileSystemObject")

On Error Resume Next 'if user has deleted file this is necessary

path = capFileGetCaptureFile(frmMain.capwnd)

'path = left$(path, 3)

DNAME = capDriverGetName(frmMain.capwnd)

SubPath = "D:\0 00 Art Loggers - WEBCAM"
If FS.folderexists(SubPath) = False Then MkDir SubPath
SubPath = "D:\0 00 Art Loggers - WEBCAM\VIDEO"
If FS.folderexists(SubPath) = False Then MkDir SubPath

TS = SubPath + "\Web_Video_Capture- " + Format$(Now, "YYYY-MM-DD HH-MM-SS DDD")
TS = TS + " -- " + DNAME + ".avi"

path = TS



txtAlloc.Text = 1


Exit Sub


'OLD CODE
'lblFreeDisk.Caption = vbGetAvailableMBytesAsString(path) 'use GetFree.bas to handle large ( > 2GB ) volumes...
Dim D, S

Set D = FS.GetDrive(FS.GetDriveName(path))
lblFreeDisk.Caption = Format(D.FreeSpace, 0)
'lblFreeDisk.Caption = FormatNumber(D.FreeSpace, 0)


capfilesize = FileLen(capFileGetCaptureFile(frmMain.capwnd))




capfilesize = "0"

'capfilesize = 20 * ONE_MEGABYTE

If capfilesize > (ONE_MEGABYTE / 2) Then
    txtAlloc.Text = capfilesize / ONE_MEGABYTE
Else
    txtAlloc.Text = 1
End If

'txtAlloc = "120"

txtAlloc.SelStart = 0
txtAlloc.SelLength = Len(txtAlloc.Text)



End Sub

Private Sub txtAlloc_Change()
If Val(txtAlloc.Text) < 0 Then txtAlloc.Text = 1
If Val(lblFreeDisk.Caption) < 1 Then Exit Sub
If Val(txtAlloc.Text) > Val(lblFreeDisk.Caption) Then txtAlloc.Text = lblFreeDisk.Caption
End Sub

Private Sub txtAlloc_KeyPress(KeyAscii As Integer)
'allow backspace key
If KeyAscii = 8 Then Exit Sub
'logic to keep the user input valid
If KeyAscii < 48 Then KeyAscii = 0
If KeyAscii > 57 Then KeyAscii = 0
End Sub
