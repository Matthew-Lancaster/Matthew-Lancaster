VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000008&
   Caption         =   "Matt Run AS"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11970
   Icon            =   "Matt_RunAS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   11970
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Process Explorer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   615
      Left            =   6420
      TabIndex        =   26
      Top             =   6285
      Width           =   4260
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Norton Ghost"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   615
      Left            =   30
      TabIndex        =   25
      Top             =   7170
      Width           =   3330
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "TWEAK XP Pro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   615
      Left            =   6420
      TabIndex        =   24
      Top             =   5640
      Width           =   3735
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "SOUND VOL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   615
      Left            =   30
      TabIndex        =   23
      Top             =   5280
      Width           =   3060
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "AutoRuns"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   615
      Left            =   30
      TabIndex        =   22
      Top             =   4650
      Width           =   2400
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "KeyBoard"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   615
      Left            =   6420
      TabIndex        =   21
      Top             =   1860
      Width           =   2445
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "PartitionMagic 8.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   615
      Left            =   30
      TabIndex        =   20
      Top             =   6540
      Width           =   4410
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "SUPER COPIER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   615
      Left            =   30
      TabIndex        =   19
      Top             =   5910
      Width           =   3900
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "TWEAK UI"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   615
      Left            =   6375
      TabIndex        =   18
      Top             =   4965
      Width           =   2535
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "WinDowse 01 RunAs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   615
      Left            =   6420
      TabIndex        =   17
      Top             =   3120
      Width           =   5100
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Magnifier"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   615
      Left            =   6420
      TabIndex        =   16
      Top             =   1230
      Width           =   2280
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "FB Photo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   615
      Left            =   6420
      TabIndex        =   15
      Top             =   600
      Width           =   2280
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Google Earth"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   555
      Left            =   6420
      TabIndex        =   14
      Top             =   30
      Width           =   3015
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Elite Spy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   615
      Left            =   6420
      TabIndex        =   13
      Top             =   4380
      Width           =   2190
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "WinDowse 02 Shell"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   615
      Left            =   6420
      TabIndex        =   12
      Top             =   3750
      Width           =   4695
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Google Earth"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   480
      Left            =   165
      TabIndex        =   11
      Top             =   4110
      Width           =   2610
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Feed Demon"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   615
      Left            =   6420
      TabIndex        =   10
      Top             =   2490
      Width           =   3090
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Winamp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   480
      Left            =   3450
      TabIndex        =   9
      Top             =   3555
      Width           =   1620
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Encrypt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   510
      Left            =   1725
      TabIndex        =   8
      Top             =   3570
      Width           =   1620
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "EXE's"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   510
      Left            =   225
      TabIndex        =   7
      Top             =   3570
      Width           =   1200
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Cool 3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   510
      Left            =   3390
      TabIndex        =   6
      Top             =   3030
      Width           =   1350
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Cool 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   510
      Left            =   1770
      TabIndex        =   5
      Top             =   3030
      Width           =   1350
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Cool 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   510
      Left            =   150
      TabIndex        =   4
      Top             =   3030
      Width           =   1350
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Opus"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   1125
      Left            =   3210
      TabIndex        =   3
      Top             =   30
      Width           =   3180
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Encarta Encyclopedia 2001"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   645
      Left            =   30
      TabIndex        =   2
      Top             =   2355
      Width           =   6360
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Palm && Sync"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   1140
      Left            =   30
      TabIndex        =   1
      Top             =   1185
      Width           =   6360
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Opus"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   1125
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   3150
   End
   Begin VB.Menu Mnu_VB 
      Caption         =   "VB ME"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'NEED BLUETOOTH CONFIG


'"C:\Program Files\ViceVersa Pro 2\ViceVersa.exe"
'"C:\Program Files\WinHTTrack\WinHTTrack.exe"

'SAFE REMOVE HARDWARE
'%windir%\system32\RunDll32.exe shell32.dll,Control_RunDLL hotplug.dll

'"C:\Program Files\Norton 360\Engine\3.0.0.135\uiStub.exe" /page {5ABC34AE-1037-4f5d-BF93-B2B74C80B5F7}
'"C:\Program Files\MIKSOFT\Mobile Media Converter\mmc.exe"

'sWindowsFolder+"\system32\PowerCalc.exe

'"D:\Program Files\WinRAR\WinRAR.exe"

'"C:\Program Files\GoogleImageDownloader-v1.1-win\GoogleImgDownloader.exe"

'"C:\Program Files\FastStone Image Viewer\FSViewer.exe"

'"C:\Program Files\Outlook Express\wab.exe"
'"C:\Program Files\TrueCrypt\TrueCrypt.exe"
'"C:\Program Files\WordNet\2.1\bin\"
'"C:\Program Files\Windows NT\Accessories\wordpad.exe"
'"C:\Program Files\Media Player Classic\mplayerc.exe"
'sWindowsFolder+"\system32\NotePad\Notepad.exe

'"C:\Program Files\Easy CD-DA Extractor 5.1\ezcddaxfc.exe"

'Code to Auto Size Form based on controls used Including Menu Bar Sketchy Style
'Make sure form set to scale Twips
'put in your form load

Set FS = CreateObject("Scripting.FileSystemObject")
'Sys Object only uses 1 2 3 = 3 temp

' System Folder - Windows\System32
'sSystemFolder = oFS.GetSpecialfolder(1)

' Windows Folder Path
sWindowsFolder = FS.GetSpecialfolder(0)



x = 1
y = 1
On Error Resume Next
For Each Control In Controls
    If Control.Enabled = True And Control.Visible = True Then
        If Control.Width + Control.Left > x Then x = Control.Width + Control.Left
        If Control.Height + Control.Top > y Then y = Control.Height + Control.Top
        If InStr(Control.Name, "Mnu_") > 0 Then mnu = 1
    End If
Next
On Error GoTo 0

Me.Width = (x + 80)
If mnu = 1 Then
    pluso = 740 ': pluso = 1100 'Sometimes Different
Else
    pluso = 450
End If

Me.Height = (y + pluso)
Me.Refresh
DoEvents

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2





End Sub

Private Sub Label1_Click()

'matt-555roids
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
'Set oRunas = CreateObject("runas.clsrunas", "Your Server")
With oRunas
    .sDomain = "WORKGROUP"
'    .sUserName = "Username you want to Run the Program  "
    .sUserName = "Matt"
    .sPassword = "matto28"
    .sCommand = "C:\Program Files\Olympus\DSSPlayer2002\DictWnd.exe"
'    .sCommand = "Program you want to run i.e. c:\winnt\notepad.exe remember"
'             you must explictly use the relevant program for example if you wanted
'             to open a text document you can't use c:\text.txt or "notepad c:\text.txt"
'             you would have to use
'             "c:\winnt\notepad.exe c:\text.txt"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing


End
End Sub

Private Sub Label10_Click()
Shell "U:\VB6\VB-NT\00_Best_VB_01\Encrypt Email Send\Send Encrypt.exe", vbNormalFocus
End
End Sub

Private Sub Label11_Click()
Shell "C:\Program Files\Greatis\WinDowse\WinDowse.exe", vbNormalFocus

End

Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
'Set oRunas = CreateObject("runas.clsrunas", "Your Server")
With oRunas
    .sDomain = "WORKGROUP"
'    .sUserName = "Username you want to Run the Program  "
    .sUserName = "Matt"
    .sPassword = "matto28"
    .sCommand = "C:\Program Files\Greatis\WinDowse\WinDowse.exe"
'    .sCommand = "Program you want to run i.e. c:\winnt\notepad.exe remember"
'             you must explictly use the relevant program for example if you wanted
'             to open a text document you can't use c:\text.txt or "notepad c:\text.txt"
'             you would have to use
'             "c:\winnt\notepad.exe c:\text.txt"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing

End

End Sub



Private Sub Label12_Click()
Exit Sub
fr1 = FreeFile
Open "U:\VB6\VB-NT\00_Best_VB_01\Cid-Run-Me-Ace\0TextData\WinAmpEXEFileActive.txt" For Input As #fr1
Line Input #fr1, tt$
Close #fr1

'TY = cProcesses.Convert(Pid, hwnd3, TT$)
'If FindWindow("Winamp v1.x", vbNullString) = 0 Then

'matt-555roids
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
'Set oRunas = CreateObject("runas.clsrunas", "Your Server")
With oRunas
    .sDomain = "WORKGROUP"
'    .sUserName = "Username you want to Run the Program  "
    .sUserName = "Matt2"
    .sPassword = "matto28"
    .sCommand = tt$
'    .sCommand = "Program you want to run i.e. c:\winnt\notepad.exe remember"
'             you must explictly use the relevant program for example if you wanted
'             to open a text document you can't use c:\text.txt or "notepad c:\text.txt"
'             you would have to use
'             "c:\winnt\notepad.exe c:\text.txt"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing
End
End Sub

Private Sub Label13_Click()
    
'matt-555roids
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
'Set oRunas = CreateObject("runas.clsrunas", "Your Server")
With oRunas
    .sDomain = "WORKGROUP"
'    .sUserName = "Username you want to Run the Program  "
    .sUserName = "Matt3"
    .sPassword = " "
    .sCommand = "C:\Program Files\FeedDemon\FeedDemon.exe"
'    .sCommand = "Program you want to run i.e. c:\winnt\notepad.exe remember"
'             you must explictly use the relevant program for example if you wanted
'             to open a text document you can't use c:\text.txt or "notepad c:\text.txt"
'             you would have to use
'             "c:\winnt\notepad.exe c:\text.txt"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing

End
    

End
End Sub

Private Sub Label14_Click()



Shell "D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe", vbNormalFocus
End

End Sub



Private Sub Label15_Click()

Shell "C:\Program Files\Google\Google Earth\client\googleearth.exe", vbNormalFocus
End

'matt-555roids
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
'Set oRunas = CreateObject("runas.clsrunas", "Your Server")
With oRunas
    .sDomain = "WORKGROUP"
'    .sUserName = "Username you want to Run the Program  "
    .sUserName = "Matt"
    .sPassword = "matto28"
    .sCommand = "C:\Program Files\Google\Google Earth\client\googleearth.exe"
'    .sCommand = "Program you want to run i.e. c:\winnt\notepad.exe remember"
'             you must explictly use the relevant program for example if you wanted
'             to open a text document you can't use c:\text.txt or "notepad c:\text.txt"
'             you would have to use
'             "c:\winnt\notepad.exe c:\text.txt"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing
End
End Sub

Private Sub Label16_Click()

'C:\Program Files\Adobe Photo UpLoader FaceBook\Adobe Photo Uploader for Facebook\Adobe Photo Uploader for Facebook.exe

'matt-555roids
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
'Set oRunas = CreateObject("runas.clsrunas", "Your Server")
With oRunas
    .sDomain = "WORKGROUP"
'    .sUserName = "Username you want to Run the Program  "
    .sUserName = "Matt"
    .sPassword = "matto28"
    .sCommand = "C:\Program Files\Adobe Photo UpLoader FaceBook\Adobe Photo Uploader for Facebook\Adobe Photo Uploader for Facebook.exe"
'    .sCommand = "Program you want to run i.e. c:\winnt\notepad.exe remember"
'             you must explictly use the relevant program for example if you wanted
'             to open a text document you can't use c:\text.txt or "notepad c:\text.txt"
'             you would have to use
'             "c:\winnt\notepad.exe c:\text.txt"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing

End

End Sub

Private Sub Label17_Click()

'matt-555roids
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
'Set oRunas = CreateObject("runas.clsrunas", "Your Server")
With oRunas
    .sDomain = "WORKGROUP"
'    .sUserName = "Username you want to Run the Program  "
    .sUserName = "Matt"
    .sPassword = "matto28"
    .sCommand = sWindowsFolder + "\system32\magnify.exe"
'    .sCommand = "Program you want to run i.e. c:\winnt\notepad.exe remember"
'             you must explictly use the relevant program for example if you wanted
'             to open a text document you can't use c:\text.txt or "notepad c:\text.txt"
'             you would have to use
'             "c:\winnt\notepad.exe c:\text.txt"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing
End

End Sub

Private Sub Label18_Click()
'"C:\Program Files\Greatis\WinDowse\WinDowse.exe"


Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
'Set oRunas = CreateObject("runas.clsrunas", "Your Server")
With oRunas
    .sDomain = "WORKGROUP"
'    .sUserName = "Username you want to Run the Program  "
    .sUserName = "Matt"
    .sPassword = "matto28"
    .sCommand = "C:\Program Files\Greatis\WinDowse\WinDowse.exe"
'    .sCommand = "Program you want to run i.e. c:\winnt\notepad.exe remember"
'             you must explictly use the relevant program for example if you wanted
'             to open a text document you can't use c:\text.txt or "notepad c:\text.txt"
'             you would have to use
'             "c:\winnt\notepad.exe c:\text.txt"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing

End

End Sub

Private Sub Label19_Click()

Shell sWindowsFolder + "\system32\TweakUI.exe", vbNormalFocus
End

End Sub

Private Sub Label2_Click()
'matt-555roids
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
'Set oRunas = CreateObject("runas.clsrunas", "Your Server")
With oRunas
    .sDomain = "WORKGROUP"
'    .sUserName = "Username you want to Run the Program  "
    .sUserName = "Matt"
    .sPassword = "matto28"
    .sCommand = "C:\Program Files\Palm\Palm.exe"
'    .sCommand = "Program you want to run i.e. c:\winnt\notepad.exe remember"
'             you must explictly use the relevant program for example if you wanted
'             to open a text document you can't use c:\text.txt or "notepad c:\text.txt"
'             you would have to use
'             "c:\winnt\notepad.exe c:\text.txt"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing


'matt-555roids
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
'Set oRunas = CreateObject("runas.clsrunas", "Your Server")
With oRunas
    .sDomain = "WORKGROUP"
'    .sUserName = "Username you want to Run the Program  "
    .sUserName = "Matt"
    .sPassword = "matto28"
    .sCommand = "C:\Program Files\Palm\Hotsync.exe"
'    .sCommand = "Program you want to run i.e. c:\winnt\notepad.exe remember"
'             you must explictly use the relevant program for example if you wanted
'             to open a text document you can't use c:\text.txt or "notepad c:\text.txt"
'             you would have to use
'             "c:\winnt\notepad.exe c:\text.txt"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing


MsgBox "Press Button in Sync USB Wire to Start"

End



End
End Sub

Private Sub Label20_Click()

Shell "C:\Program Files\SuperCopier2\SuperCopier2.exe", vbNormalFocus

End

End Sub

Private Sub Label21_Click()
Shell "C:\Program Files\Symantec\Norton PartitionMagic 8.0\PMagic.exe", vbNormalFocus
End
End Sub

Private Sub Label22_Click()
Shell sWindowsFolder + "\system32\osk.exe", vbNormalFocus
End

End Sub

Private Sub Label23_Click()
Shell "C:\Program Files\Autoruns Microsoft\AutoRuns.exe", vbNormalFocus
End
End Sub

Private Sub Label24_Click()
Shell sWindowsFolder + "\system32\sndvol32.exe", vbNormalFocus
End

End Sub

Private Sub Label25_Click()
Shell "C:\Program Files\Tweak-XP Pro\Tweak-xp.exe", vbNormalFocus
End

End Sub

Private Sub Label26_Click()
Shell "C:\Program Files\Norton Ghost\Console\VProConsole_.exe", vbNormalFocus
End
End Sub

Private Sub Label27_Click()
Shell "C:\Program Files\ProcessExplorer\Procexp.exe", vbNormalFocus
End
End Sub

Private Sub Label3_Click()
'---------------------------
'ENC2001.EXE - Unable To Locate Component
'---------------------------
'This application has failed to start because MVCL53N.dll was not found. Re-installing the application may fix this problem.
'---------------------------
'OK
'---------------------------

ChDir "C:\Program Files\Microsoft Encarta\Encarta Encyclopedia 2001"
ChDir "C:\Program Files\Common Files\Microsoft Shared\Reference 2001"
'matt-555roids
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
'Set oRunas = CreateObject("runas.clsrunas", "Your Server")
With oRunas
    .sDomain = "WORKGROUP"
'    .sUserName = "Username you want to Run the Program  "
    .sUserName = "Matt"
    .sPassword = "matto28"
    .sCommand = "C:\Program Files\Microsoft Encarta\Encarta Encyclopedia 2001\ENC2001.EXE"
'    .sCommand = "Program you want to run i.e. c:\winnt\notepad.exe remember"
'             you must explictly use the relevant program for example if you wanted
'             to open a text document you can't use c:\text.txt or "notepad c:\text.txt"
'             you would have to use
'             "c:\winnt\notepad.exe c:\text.txt"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing

End
End Sub

Private Sub Label4_Click()

'matt-555roids
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
'Set oRunas = CreateObject("runas.clsrunas", "Your Server")
With oRunas
    .sDomain = "WORKGROUP"
'    .sUserName = "Username you want to Run the Program  "
    .sUserName = "Matt"
    .sPassword = "matto28"
    .sCommand = "C:\Program Files\Olympus\Digital Wave Player\DWP.exe"
'    .sCommand = "Program you want to run i.e. c:\winnt\notepad.exe remember"
'             you must explictly use the relevant program for example if you wanted
'             to open a text document you can't use c:\text.txt or "notepad c:\text.txt"
'             you would have to use
'             "c:\winnt\notepad.exe c:\text.txt"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing

End

End Sub

Private Sub Label5_Click()
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
With oRunas
    .sDomain = "WORKGROUP"
    .sUserName = "Matt"
    .sPassword = "matto28"
    .sCommand = "C:\Program Files\Cool2000\cool2000.exe"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing
End
End Sub

Private Sub Label6_Click()
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
With oRunas
    .sDomain = "WORKGROUP"
    .sUserName = "Admin"
    .sPassword = "matto28"
    .sCommand = "C:\Program Files\Cool2000\cool2000.exe"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing
End

End Sub

Private Sub Label7_Click()
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
With oRunas
    .sDomain = "WORKGROUP"
    .sUserName = "Administrator"
    .sPassword = "matto28"
    .sCommand = "C:\Program Files\Cool2000\cool2000.exe"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing
End
End Sub

'Shell "U:\VB6\VB-NT\00 MP3 Sound\DSS_Olympus\DSS_Olympus.exe"
'End

Private Sub Label8_Click()
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
With oRunas
    .sDomain = "WORKGROUP"
    .sUserName = "Matt"
    .sPassword = "matto28"
    .sCommand = "C:\Program Files\Google\Google Earth\GoogleEarth.exe"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing
End

End Sub

Private Sub Label9_Click()
Shell "D:\VB6\VB-NT\00_Best_VB_01\50LastEXE's\50LastEXE.exe", vbNormalFocus
End
End Sub

Private Sub Mnu_VB_Click()
'If IsIDE = False Then
    Shell "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
'    End
'End If
End
End Sub
