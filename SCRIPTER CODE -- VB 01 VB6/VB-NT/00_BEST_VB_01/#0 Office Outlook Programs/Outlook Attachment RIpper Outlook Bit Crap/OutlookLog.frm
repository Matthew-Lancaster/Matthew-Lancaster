VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form OutlookRipperLog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log . . . "
   ClientHeight    =   4920
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5625
   Icon            =   "OutlookLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   4995
      Left            =   0
      TabIndex        =   0
      Top             =   -75
      Width           =   5625
      Begin MSComctlLib.ListView LogList 
         Height          =   4830
         Left            =   30
         TabIndex        =   1
         Top             =   135
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   8520
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Log"
            Object.Width           =   9816
         EndProperty
      End
   End
End
Attribute VB_Name = "OutlookRipperLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Activate()
    
    Call OutlookAttachmentRipper.SaveFiles(OutlookAttachmentRipper.SubFolderList)
    MsgBox "Completed ...... ", vbInformation
End Sub

Private Sub Form_Unload(Cancel As Integer)
    OutlookAttachmentRipper.Delete(0).Value = 0
    OutlookAttachmentRipper.Delete(1).Value = 0
    OutlookAttachmentRipper.Dbstore.Value = 0
    OutlookAttachmentRipper.FullLogCounter = 1
    OutlookAttachmentRipper.SaveMail.Value = 0
    OutlookAttachmentRipper.SaveAttachment.Value = 0
    OutlookAttachmentRipper.ExtractEmail = False
End Sub
