VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim TPath$

ScanPath.chkSubFolders = vbChecked
ScanPath.cboMask.Text = "*.*"

ScanPath.txtPath.Text = "E:\01 FAVS\#00 Palm\DownLoad Sites\2009-08 Aug"

'ScanPath.Show

For we = 1 To ScanPath.ListView1.ListItems.Count
    a1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    tt = LastModifiedToCurrent(a1$ + B1$)
    
Next

'MsgBox "Done" + vbCrLf + dd$
End
End Sub
