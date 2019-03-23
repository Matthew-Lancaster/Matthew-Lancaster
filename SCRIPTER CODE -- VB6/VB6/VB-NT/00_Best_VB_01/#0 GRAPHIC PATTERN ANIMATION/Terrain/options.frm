VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form options 
   Caption         =   "Options"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2385
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   2385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Load Map"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   720
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save/Load Map"
      Filter          =   "Load/Save Map|*.map"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save Map"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Load Bitmap"
      Filter          =   "Bitmap File|*.bmp"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Background Picture..."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CommonDialog1.ShowOpen
If CommonDialog1.filename <> "" Then
TerrainForm.Picture2.Picture = LoadPicture(CommonDialog1.filename)
FileString = CommonDialog1.filename
End If
End Sub

Private Sub Command2_Click()
Dim temp As Integer
CommonDialog2.ShowSave
If CommonDialog2.filename <> "" Then
Open CommonDialog2.filename For Binary As 1
temp = Len(FileString)
Put #1, , temp
Put #1, , FileString
Put #1, , ViewX
Put #1, , ViewY
temp = TerrainForm.ZoomBar.Value
Put #1, , temp
temp = TerrainForm.smooth.Text
Put #1, , temp
temp = TerrainForm.txtDetail.Text
Put #1, , temp
temp = TerrainForm.HScroll1.Value
Put #1, , temp
temp = TerrainForm.scrwaterdiff.Value
Put #1, , temp
temp = TerrainForm.scrbeachheight.Value
Put #1, , temp
temp = TerrainForm.scrollX.Value
Put #1, , temp
temp = TerrainForm.scrollY.Value
Put #1, , temp
temp = TerrainForm.scrollZ.Value
Put #1, , temp
temp = TerrainForm.scrLum.Value
Put #1, , temp
temp = TerrainForm.thresh.Text
Put #1, , temp
temp = TerrainForm.chkshallow.Value
Put #1, , temp
Put #1, , MapArray
Close #1
End If
End Sub

Private Sub Command3_Click()
Dim temp As Integer
CommonDialog2.ShowOpen
If CommonDialog2.filename <> "" Then
Open CommonDialog2.filename For Binary As 1
Get #1, , temp
FileString = Input(temp, 1)
TerrainForm.Picture2.Picture = LoadPicture(FileString)
Get #1, , ViewX
Get #1, , ViewY
Get #1, , temp
TerrainForm.ZoomBar.Value = temp
Get #1, , temp
TerrainForm.smooth.Text = temp
Get #1, , temp
TerrainForm.txtDetail.Text = temp
Get #1, , temp
TerrainForm.HScroll1.Value = temp
Get #1, , temp
TerrainForm.scrwaterdiff.Value = temp
Get #1, , temp
TerrainForm.scrbeachheight.Value = temp

Get #1, , temp
TerrainForm.scrollX.Value = temp
Get #1, , temp
TerrainForm.scrollY.Value = temp
Get #1, , temp
TerrainForm.scrollZ.Value = temp
Get #1, , temp
TerrainForm.scrLum.Value = temp
Get #1, , temp
TerrainForm.thresh.Text = temp
Get #1, , temp
TerrainForm.chkshallow.Value = temp
Get #1, , MapArray
Close #1
water = TerrainForm.HScroll1.Value
For Y = 0 To 127
For X = 0 To 127
HeightArray(X, Y) = MapArray(X, Y)
If HeightArray(X, Y) <= water Then HeightArray(X, Y) = water
Next X
Next Y
End If
End Sub
