VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmOperatingSystem 
   Caption         =   "Operating System"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7905
   ScaleWidth      =   11175
   Visible         =   0   'False
   Begin MSFlexGridLib.MSFlexGrid msfgOperatingSystem 
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   13785
      _Version        =   393216
      Rows            =   53
      FixedCols       =   0
      AllowUserResizing=   3
   End
End
Attribute VB_Name = "frmOperatingSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public Sub ShowForm()

Private Sub Form_Load()
    'frmOperatingSystem.Height = MDIProcServ.Height - 200
    'frmOperatingSystem.Width = MDIProcServ.Width - 275
    'Call frmOperatingSystem.ShowForm

End Sub

