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

ScanPath.txtPath.Text = "J:\ZIPS"

'ScanPath.txtPath.Text = "M:\00-C-Drive\ZIPS\TrueCrypt"
'ScanPath.txtPath.Text = "M:\00-C-Drive-Week\ZIPS\TrueCrypt"
'ScanPath.txtPath.Text = "N:\00-C-Drive\ZIPS\TrueCrypt"
'ScanPath.txtPath.Text = "N:\00-C-Drive-Week\ZIPS\TrueCrypt"

ScanPath.txtPath.Text = "M:\ZIPS"
'ScanPath.txtPath.Text = "N:\ZIPS"
ScanPath.txtPath.Text = "N:\ZIPS M Drive\True Crypt"

'ScanPath.Show

For we = 1 To ScanPath.ListView1.ListItems.Count
    a1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    xc = 1
    If InStr(LCase(B1$), ".lnk") > 0 Then xc = 0
    If InStr(LCase(a1$), "not used") > 0 Then xc = 0
    If xc = 1 Then
        tt = LastModifiedToCurrent(a1$ + B1$)
        xx = ""
        If tt = False Then xx = " = xx"
        dd$ = dd$ + B1$ + xx + vbCrLf
    End If
Next

'MsgBox "Done" + vbCrLf + dd$
End
End Sub
