VERSION 5.00
Begin VB.Form frmAddFilter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Filter"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   267
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   383
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraParams 
      Caption         =   "Parameters"
      Height          =   3255
      Left            =   2880
      TabIndex        =   3
      Top             =   45
      Width           =   2775
      Begin VB.TextBox txtP 
         Height          =   285
         Index           =   4
         Left            =   960
         TabIndex        =   13
         Top             =   2475
         Width           =   1575
      End
      Begin VB.TextBox txtP 
         Height          =   285
         Index           =   3
         Left            =   960
         TabIndex        =   12
         Top             =   1995
         Width           =   1575
      End
      Begin VB.TextBox txtP 
         Height          =   285
         Index           =   2
         Left            =   960
         TabIndex        =   11
         Top             =   1515
         Width           =   1575
      End
      Begin VB.TextBox txtP 
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   10
         Top             =   1035
         Width           =   1575
      End
      Begin VB.TextBox txtP 
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   4
         Top             =   555
         Width           =   1575
      End
      Begin VB.Label lblParam 
         Caption         =   "Param 5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblParam 
         Caption         =   "Param 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label lblParam 
         Caption         =   "Param 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblParam 
         Caption         =   "Param 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblParam 
         Caption         =   "Param 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.ListBox lstFilters 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3765
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   3390
      Width           =   975
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   3360
      Width           =   975
   End
End
Attribute VB_Name = "frmAddFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Filters_Out() As Integer, Filters_In() As Integer

Private Sub btnAdd_Click()
Dim MyFilter As tFilterVars
Dim nFilter As Integer, nIndex As Integer, Cnt As Integer

If lstFilters.ListIndex < 0 Then Exit Sub
nIndex = lstFilters.ListIndex

Select Case Me.Tag
Case 0
    'Out
    nFilter = Filters_Out(nIndex)
Case 1
    'In
    nFilter = Filters_In(nIndex)
End Select

With MyFilter
    .IsEnabled = True
    .nType = nFilter
    .nVars = aData.Filters(nFilter).nVars
    For Cnt = 0 To aData.Filters(nFilter).nVars - 1
        .Var(Cnt) = Val(txtP(Cnt).Text)
    Next Cnt
End With

Select Case Me.Tag
Case 0
    AddFilterIndirect Frac(aData.nActive), MyFilter, True
Case 1
    AddFilterIndirect Frac(aData.nActive), MyFilter, False
End Select

End Sub

Private Sub btnClose_Click()

Me.Hide

End Sub

Private Sub Form_Activate()
Dim Cnt As Long

'Make Filter lists:

ReDim Filters_Out(0)
ReDim Filters_In(0)

For Cnt = 0 To aData.nFilters - 1
    Select Case aData.Filters(Cnt).nComp
    Case FILTER_COMP_2D
        'Add to both Out and In filter lists:
        Filters_Out(UBound(Filters_Out)) = Cnt
        ReDim Preserve Filters_Out(0 To UBound(Filters_Out) + 1)
        
        Filters_In(UBound(Filters_In)) = Cnt
        ReDim Preserve Filters_In(0 To UBound(Filters_In) + 1)
    Case FILTER_COMP_2DOUT
        'Add to Out filter list:
        Filters_Out(UBound(Filters_Out)) = Cnt
        ReDim Preserve Filters_Out(0 To UBound(Filters_Out) + 1)
    Case FILTER_COMP_2DIN
        'Add to In filter list:
        Filters_In(UBound(Filters_In)) = Cnt
        ReDim Preserve Filters_In(0 To UBound(Filters_In) + 1)
    End Select
Next Cnt

'Clear Filter list control:
lstFilters.Clear

Select Case Me.Tag
Case 0
    'Outside filters:
    For Cnt = 0 To UBound(Filters_Out) - 1
        lstFilters.AddItem aData.Filters(Filters_Out(Cnt)).szName
    Next Cnt
Case 1
    'Inside filters:
    For Cnt = 0 To UBound(Filters_In) - 1
        lstFilters.AddItem aData.Filters(Filters_In(Cnt)).szName
    Next Cnt
End Select

For Cnt = 0 To 4
    lblParam(Cnt).Enabled = False
    txtP(Cnt).Enabled = False
Next Cnt

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Cancel = 1
Me.Hide

End Sub

Private Sub Form_Unload(Cancel As Integer)

Cancel = 1
Me.Hide

End Sub

Private Sub lstFilters_Click()
Dim nVars As Integer, Cnt As Integer

If lstFilters.ListIndex < 0 Then Exit Sub
Select Case Me.Tag
Case 0
    nVars = aData.Filters(Filters_Out(lstFilters.ListIndex)).nVars
Case 1
    nVars = aData.Filters(Filters_In(lstFilters.ListIndex)).nVars
End Select

For Cnt = 0 To 4
    If Cnt + 1 <= nVars Then
        lblParam(Cnt).Enabled = True
        txtP(Cnt).Enabled = True
    Else
        lblParam(Cnt).Enabled = False
        txtP(Cnt).Enabled = False
    End If
Next Cnt

End Sub
