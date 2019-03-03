VERSION 5.00
Begin VB.Form FormStart 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FormStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub FormStartLoader()

FontSizez = 13

ScanPath.txtPath.Text = "E:\01 VB Shell Folders\00 " + App.EXEName
ScanPath.cboMask.Text = "*.lnk"
Call ScanPath.cmdScan_Click

ScanPath.chkSubFolders.Value = vbUnchecked
ScanPath.txtPath.Text = "D:\# MY DOCS\# 01 My Documents"
ScanPath.cboMask.Text = "*.xls"
Call ScanPath.cmdScan_Click

ScanPath.chkSubFolders.Value = vbChecked
ScanPath.txtPath.Text = "D:\# MY DOCS\# 01 My Documents\03 Excel Worksheets"
ScanPath.cboMask.Text = "*.xls"
Call ScanPath.cmdScan_Click

End Sub

Sub LabelClick(index)


A1$ = ScanPath.ListView1.ListItems.Item(index).SubItems(1)
B1$ = ScanPath.ListView1.ListItems.Item(index)

If Sqidge = True Then Shell "Explorer.exe /e," + A1$, vbNormalFocus: End

vLaunch A1$ + B1$, vbNullString
End

End Sub

