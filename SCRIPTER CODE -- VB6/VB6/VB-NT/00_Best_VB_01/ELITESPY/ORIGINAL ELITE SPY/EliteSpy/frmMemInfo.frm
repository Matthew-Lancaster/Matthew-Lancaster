VERSION 5.00
Begin VB.Form frmMemInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Memory Information's"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMemInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   315
      Left            =   3360
      TabIndex        =   0
      Top             =   4020
      Width           =   1155
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   4140
      Top             =   60
   End
   Begin VB.Frame Frame3 
      Caption         =   "Page File Usage (K)"
      Height          =   1035
      Left            =   60
      TabIndex        =   13
      Top             =   2880
      Width           =   4455
      Begin VB.TextBox txtAPageFile 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   660
         Width           =   2055
      End
      Begin VB.TextBox txtPageFile 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   300
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Avaible Page File:"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   660
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Total Page File:"
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   300
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Virtual Memory (K)"
      Height          =   1035
      Left            =   60
      TabIndex        =   8
      Top             =   1740
      Width           =   4455
      Begin VB.TextBox txtVirtualMemory 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   300
         Width           =   2055
      End
      Begin VB.TextBox txtAVirtualMemory 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   660
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Total Virtual Memory:"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   300
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Avaible Virtual Memory:"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   660
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Psysical Memory (K)"
      Height          =   1035
      Left            =   60
      TabIndex        =   3
      Top             =   600
      Width           =   4455
      Begin VB.TextBox txtPhysicalMemory 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   300
         Width           =   2055
      End
      Begin VB.TextBox txtAPhysicalMemory 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   660
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Total Physical Memory:"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   300
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Avaible Physical Memory:"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   660
         Width           =   1815
      End
   End
   Begin VB.TextBox txtMemoryUsage 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   180
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Memory Usage %:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   1695
   End
End
Attribute VB_Name = "frmMemInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmMemInfo
'    Project    : EliteSpy
'
'    Description: Contains code for retriving memory information's
'
'    Author     : Andrea Batina
'    Modified   : 31/10/2001
'--------------------------------------------------------------------------------
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ' Get global memory status
    GetMemoryStatus
End Sub

Private Sub Timer1_Timer()
    ' Get global memory status
    GetMemoryStatus
End Sub

Private Sub GetMemoryStatus()
    Dim tMS As MEMORYSTATUS
    
    ' Length of structure
    tMS.dwLength = Len(tMS)
    ' Get global memory status
    GlobalMemoryStatus tMS
        
    ' Print memory status
    txtPhysicalMemory.Text = Format$(tMS.dwTotalPhys / 1024, "###,###,###") & " Kb"
    txtAPhysicalMemory.Text = Format$(tMS.dwAvailPhys / 1024, "###,###,###") & " Kb"
    txtVirtualMemory.Text = Format$(tMS.dwTotalVirtual / 1024, "###,###,###") & " Kb"
    txtAVirtualMemory.Text = Format$(tMS.dwAvailVirtual / 1024, "###,###,###") & " Kb"
    txtPageFile.Text = Format$(tMS.dwTotalPageFile / 1024, "###,###,###") & " Kb"
    txtAPageFile.Text = Format$(tMS.dwAvailPageFile / 1024, "###,###,###") & " Kb"
    txtMemoryUsage.Text = tMS.dwMemoryLoad & " %"
End Sub
