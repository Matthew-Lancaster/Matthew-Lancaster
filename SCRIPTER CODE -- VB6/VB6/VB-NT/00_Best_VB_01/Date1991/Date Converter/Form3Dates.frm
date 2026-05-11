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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Open "D:\VB6\VB-NT\00_Best_VB_01\Date1991\00 Data\1991DAT4.TXT" For Input As #1
Open "D:\VB6\VB-NT\00_Best_VB_01\Date1991\00 Data\1991DAT5.TXT" For Output As #2
Do
Line Input #1, hh$
hh$ = Trim(hh$)
tt$ = ""
If InStr("0123456789", Mid$(hh$, 1, 1)) = 0 Then
    gg$ = hh$
Else
tj = InStr(hh$, ":")
tt$ = Mid$(hh$, 1, tj) + gg$ + " --" + Mid$(hh$, tj + 1)

End If
If tt$ <> "" And hh$ <> "" Then
    ff$ = Mid$(tt$, 1, 12)
    kk = DateValue(ff$)
    vv$ = Format$(kk, "DD-MM-YYYY ")
    tt$ = vv$ + Mid$(tt$, 14)
    
End If
If tt$ <> "" And hh$ <> "" Then
    
    Print #2, tt$
End If

Loop Until EOF(1)
Close #1, #2

End

End Sub
