VERSION 5.00
Begin VB.Form frmAskPropTag 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select a property tag"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstTypes 
      Height          =   3630
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   2
      Top             =   390
      Width           =   2205
   End
   Begin VB.ListBox lstPropTags 
      Height          =   3630
      IntegralHeight  =   0   'False
      Left            =   2520
      TabIndex        =   3
      Top             =   390
      Width           =   4605
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Accept"
      Default         =   -1  'True
      Height          =   435
      Left            =   1950
      TabIndex        =   6
      Top             =   4500
      Width           =   1035
   End
   Begin VB.CommandButton cmdKo 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   4290
      TabIndex        =   7
      Top             =   4500
      Width           =   1035
   End
   Begin VB.Label lblPropTag 
      AutoSize        =   -1  'True
      Caption         =   "Value:"
      Height          =   195
      Left            =   2520
      TabIndex        =   5
      Top             =   4050
      UseMnemonic     =   0   'False
      Width           =   450
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      Caption         =   "Value:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   4050
      UseMnemonic     =   0   'False
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Property tag"
      Height          =   195
      Left            =   2520
      TabIndex        =   1
      Top             =   150
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Property type"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   930
   End
End
Attribute VB_Name = "frmAskPropTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Public inDefault As Long
Public outOk As Boolean, outValue As Long

Private ptA() As Long, ptL As Long, ptU As Long

Private Sub Form_Load()
   Dim K As Long, si As Long
   Dim tA() As Long, tL As Long, tU As Long

   outOk = False
   modkWabGetAll.GetAllTypes tA
   tL = LBound(tA)
   tU = UBound(tA)

   modkWabGetAll.GetAllTags ptA
   ptL = LBound(ptA)
   ptU = UBound(ptA)

   si = -1
   With Me.lstTypes
      .Clear
      For K = tL To tU
         .AddItem modkWabTypes.kWabPropTypeName(tA(K))
         .ItemData(.NewIndex) = tA(K)
         If kWabGetPropType(inDefault) = tA(K) Then si = .NewIndex
      Next
      If si >= 0 Then .ListIndex = si
   End With
   checkOk
End Sub

Private Sub cmdOk_Click()
   With Me.lstPropTags
      If .ListCount <= 0 Then Exit Sub
      If .ListIndex < 0 Then Exit Sub
      Me.outValue = .ItemData(.ListIndex)
   End With
   Me.outOk = True
   Me.Hide
End Sub
Private Sub cmdKo_Click()
   Me.outOk = False
   Me.Hide
End Sub

Public Sub lstPropTags_Click()
   Me.lblPropTag = "Value: &H3003001E"
'code
   checkOk
   Exit Sub
   Me.lblPropTag = "Value:"
   With Me.lstPropTags
      If .ListCount > 0 Then
         If .ListIndex >= 0 Then
            Me.lblPropTag = "Value: &H" & Right(String(8, "0") & Hex(.ItemData(.ListIndex)), 8)
         End If
      End If
   End With
   checkOk
End Sub
Public Sub lstTypes_Click()
   Dim K As Long, PropType As Long, si As Long
   
   PropType = -1
   Me.lblType = "Value:"
   With Me.lstTypes
      If .ListCount < 1 Then Exit Sub
      If .ListIndex < 0 Then Exit Sub
      PropType = .ItemData(.ListIndex)
   End With
   Me.lblType = "Value: &H" & Right("0000" & Hex(PropType), 4)
   si = -1
   With Me.lstPropTags
      .Clear
      For K = ptL To ptU
         If kWabGetPropType(ptA(K)) = PropType Then
            .AddItem modkWabTags.kWabPropTagName(ptA(K))
            .ItemData(.NewIndex) = ptA(K)
            If kWabGetPropType(inDefault) = ptA(K) Then si = .NewIndex
         End If
      Next
   End With
   checkOk
End Sub

Private Sub checkOk()
   Dim ok As Boolean

   ok = False
   With Me.lstPropTags
      If .ListCount > 0 Then
         ok = (.ListIndex >= 0)
      End If
   End With
   Me.cmdOk.Enabled = ok
End Sub
