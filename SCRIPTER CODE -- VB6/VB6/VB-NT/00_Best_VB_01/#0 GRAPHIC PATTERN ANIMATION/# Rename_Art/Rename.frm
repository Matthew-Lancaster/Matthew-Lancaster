VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   -45
      TabIndex        =   0
      Top             =   270
      Width           =   6195
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dups 2 = "
      Height          =   255
      Left            =   4905
      TabIndex        =   4
      Top             =   0
      Width           =   1230
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dups 1 = "
      Height          =   255
      Left            =   3660
      TabIndex        =   3
      Top             =   0
      Width           =   1230
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Progress = "
      Height          =   255
      Left            =   1830
      TabIndex        =   2
      Top             =   0
      Width           =   1800
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Cue

Private Sub Form_Activate()

If Cue = 1 Then Exit Sub
Cue = 1

Load ScanPath
ScanPath.Show

putpath$ = "R:\Art\AutoPix\AutoPix (Main)\AutoPix-2006--10\"
putpath2$ = "R:\Art\AutoPix\AutoPix (Main)\AutoPix-2006--10"
'putpath2$ = "R:\Art\AutoPix\AutoPix (Main)\AutoPix-2006\"
'putpath$ = "R:\Art\AutoPix\AutoPix (Main)\AutoPix 2005--07-12\"
'putpath2$ = "R:\Art\AutoPix\AutoPix (Main)\AutoPix 2005--07-12\"

putname$ = "Sx " + Format$(Now, "YYYY-MM-DD") + "--"
ScanPath.txtPath = "R:\Art\AutoPix\AutoPix (Main)\AutoPix-2006--10"
'ScanPath.txtPath = putpath$
ScanPath.cboMask.Text = "*.jpg"

'Call ScanPath.cmdScan_Click


End Sub

Public Sub AfterScanPathClick()

putpath$ = ScanPath.txtPath + "\"
putpath2$ = ScanPath.txtPath

putname$ = "Sx " + Format$(Now, "YYYY-MM-DD") + "--"


Books = Val(ScanPath.lblCount.Caption)
Label1.Caption = ("Count = " + Str$(Books))

For we = 1 To ScanPath.ListView1.ListItems.Count
    a1$ = ScanPath.ListView1.ListItems.Item(we)
    b1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)

    r$ = Format$(we, "00000")
    c$ = a1$
    'a1$ = Mid$(a1$, 1, InStrRev(a1$, ".") - 9) + ".jpg"
    'b$ = Mid$(b$, 1, 16) + r$ + Mid$(b$, 17)
'    Name a1$ + c$ As a1$ + b$
    'If UCase$(a1$) = f$ Then
    '    dup = dup + 1
    '    r2$ = Format$(dup, "00")
    '    a1$ = Mid$(a1$, 1, InStrRev(a1$, ".") - 1) + "-" + r2$ + ".jpg"
    'Else
    '    dup = 0
    'End If
    
    'f$ = UCase$(a1$)
    
'    a1$ = Mid$(a1$, 1, InStrRev(a1$, ".") - 8) + ".jpg"
    'a1$ = Mid$(a1$, 1, 14) + Mid$(a1$, 20)
    a1$ = putname$ + a1$
    
    List1.AddItem a1$
   
Next

Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")

For we = 1 To ScanPath.ListView1.ListItems.Count
    Label2.Caption = "Progess =" + Str$(we)
    DoEvents
    a1$ = ScanPath.ListView1.ListItems.Item(we)
    b1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
'    On Error Resume Next
'    Name b1$ + a1$ As b1$ + List1.List(we - 1)
'    w$ = Err.Description
    dup = 0
    k2$ = Mid$(List1.List(we - 1), 1, InStrRev(List1.List(we - 1), ".") - 1) + "-"
    Do
        erg$ = Dir$(putpath2$ + List1.List(we - 1))
        erg2$ = Dir$(putpath$ + List1.List(we - 1))
        If erg <> "" Or erg2$ <> "" Then
            If erg$ <> "" Then
                dups1 = dups1 + 1
                Label3.Caption = "Dups 1=" + Str$(dups1)
            End If
            If erg2$ <> "" And erg$ = "" Then
                dups2 = dups2 + 1
                Label4.Caption = "Dups 2=" + Str$(dups2)
            End If
            dup = dup + 1
            r2$ = Format$(dup, "00")
            k$ = k2$ + r2$ + ".jpg"
            List1.List(we - 1) = k$
        End If
    Loop Until erg$ = ""
    
'putpath$ = "R:\Art\AutoPix\AutoPix (Main)\AutoPix-2006--10\"

'putname$ = "Sx " + Format$(Now, "YYYY-MM-DD") + "--"
    
    fs.Copyfile b1$ + a1$, putpath$ + List1.List(we - 1)
    
    'Name b1$ + a1$ As b1$ + List1.List(we - 1)
    
'    On Local Error GoTo 0
Next

End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub
