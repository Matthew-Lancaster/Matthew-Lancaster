VERSION 5.00
Begin VB.Form Form3_Extra 
   Caption         =   "Form2"
   ClientHeight    =   2436
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   3744
   LinkTopic       =   "Form2"
   ScaleHeight     =   2436
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form3_Extra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'CODE TO MOVE TO SECOND FORM

'-----------------------------------------------
'-----------------------------------------------
'-----------------------------------------------
'BOTTOM OF CODE


'MNU ITEM SET WITHDRAW

'      Begin VB.Menu MNU_EL_122
'         Caption         =   "FOLDER EXPLORER -- C:\images"
'      End
'      Begin VB.Menu MNU_EL_120
'         Caption         =   "FOLDER EXPLORER -- C:\ISO"
'      End
'      Begin VB.Menu MNU_EL_119
'         Caption         =   "FOLDER EXPLORER -- C:\Make_PE3"
'      End
'      Begin VB.Menu MNU_EL_118
'         Caption         =   "FOLDER EXPLORER -- C:\MINIPE"
'      End
'      Begin VB.Menu MNU_EL_117
'         Caption         =   "FOLDER EXPLORER -- C:\Program Files"
'      End
'      Begin VB.Menu MNU_EL_115
'         Caption         =   "FOLDER EXPLORER -- C:\Program Files (x86)"
'      End
 ''     Begin VB.Menu MNU_EL_114
'         Caption         =   "FOLDER EXPLORER -- C:\PStart"
'      End
'      Begin VB.Menu MNU_EL_15
'         Caption         =   "VB FOLDER EXPLORER -- D:\VB6\VB-NT\00_Best_VB_05_COMMON"
'      End
'      Begin VB.Menu MNU_EL_14
'         Caption         =   "VB FOLDER EXPLORER -- D:\VB6\VB-NT\00_Best_VB_04"
'      End
'      Begin VB.Menu MNU_EL_13
'         Caption         =   "VB FOLDER EXPLORER -- D:\VB6\VB-NT\00_Best_VB_05_COMMON"
'      End
'      Begin VB.Menu MNU_EL_2
'         Caption         =   "VB FOLDER EXPLORER -- D:\VB6\VB-NT\00_Send_To"
'      End




Sub FILE_OPERATIONS_HANDLER(VAR1, VAR2)

Exit Sub
'--------------------------------------
'NOT USE
'--------------------------------------

'XGO = False
'If GetComputerName = "1-ASUS-EEE" Then XGO = True
'If GetComputerName = "1-ASUS-X5DIJ" Then XGO = True


'GIVE UP ON THIS A BIT

VAR2 = "D:\0 00 ART LOGGERS\# APP AND SCREEN\" + GetComputerName + "\AUTO_Form_Shot\"

VAR2 = "D:\0 00 MUSIC ---"
VAR2 = "D:\0 00 MUSIC -Z0x"
VAR2 = "\\4-asus-gl522vw\4-asus-gl522vw_02_d-drive\0 00 MUSIC ---"
VAR2 = "\\4-asus-gl522vw\4-asus-gl522vw_02_d-drive\0 00 MUSIC -Z0x"
VAR2 = ""

MNU_SCANPATH_COUNTER.Visible = True
ScanPath.cboMask = "0#0*.MP3"
ScanPath.chkSubFolders = vbChecked
ScanPath.TxtPath = VAR2

Call ScanPath.cmdScan_Click
'Case 4
'sort of PATH then FILE

For WE = 1 To ScanPath.ListView1.ListItems.Count
    
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    
    If Mid(B1, 1, 4) = "0#01" Then
        For WE2 = 1 To ScanPath.ListView1.ListItems.Count
            A12 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
            B12 = ScanPath.ListView1.ListItems.Item(WE)
            If A1 = A12 Then
                If Len(B12) > Len(B1) Then
                
                End If
            End If
        Next
    End If
    
    'YY$ = A1 + B1
    'Print #1, YY$

Next



End Sub

