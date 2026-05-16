VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   105
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   195
      Width           =   8175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Open "R:\email\usenet~4.dbx" For Binary As #1

T = 0
Do
    dd$ = Input(2000, #1)

    If ee$ <> "" Then
        If Len(ee$) > 3999 Then ee$ = Mid$(ee$, 2000) + dd$: tip = tip - 2000
        If Len(ee$) < 2001 Then ee$ = ee$ + dd$
        If tip < 1 Then tip = 1
    Else
        ee$ = dd$
        tip = 1
    End If
    
    keep = 0
    Do
        tip2 = InStr(tip, ee$, "From: ")
        tip3 = InStr(tip2 + 1, ee$, vbCr)
        If tip2 > 0 And tip3 > 0 Then
            'List1.AddItem Mid$(ee$, tip2, tip3 - tip2)
            plij$ = Mid$(ee$, tip2, tip3 - tip2)
            tip = tip2 + 1
            keep = 1
        End If
        
        If keep = 1 Then
            tip2 = InStr(tip, ee$, "Date: ")
            tip3 = InStr(tip2 + 1, ee$, vbCr)
            If tip2 > 0 And tip3 > 0 Then
                'List1.AddItem Mid$(ee$, tip2, tip3 - tip2)
                plij$ = plij$ + " -*- " + Mid$(ee$, tip2, tip3 - tip2)
                tip = tip2 + 1
                keep = 2
            End If
        End If
        
        If keep = 2 Then
            tip2 = InStr(tip, ee$, "X-Newsreader: ")
            tip3 = InStr(tip2 + 1, ee$, vbCr)
            If tip2 > 0 And tip3 > 0 Then
                'List1.AddItem Mid$(ee$, tip2, tip3 - tip2)
                plij$ = plij$ + " -*- " + Mid$(ee$, tip2, tip3 - tip2)
                tip = tip2 + 1
                keep = 3
            End If
        End If
        
        If keep = 4 Then
            tip2 = InStr(tip, ee$, "X-Complaints-To: ")
            tip3 = InStr(tip2 + 1, ee$, vbCr)
            If tip2 > 0 And tip3 > 0 Then
                'List1.AddItem Mid$(ee$, tip2, tip3 - tip2)
                plij$ = plij$ + " -*- " + Mid$(ee$, tip2, tip3 - tip2)
                List1.Refresh
                tip = tip2 + 1
            End If
        End If
    
        If keep = 1 Then plij$ = plij$ + " -*- "
        If keep > 0 Then List1.AddItem plij$
    
    Loop Until tip2 = 0 Or tip3 = 0
T = T + 1
Loop Until EOF(1) ' Or T > 6000
Close #1


For r = List1.ListCount - 2 To 1 Step -1
    et1$ = Mid$(List1.List(r), 1, InStr(List1.List(r), " -*- "))
    et2$ = Mid$(List1.List(r + 1), 1, InStr(List1.List(r + 1), " -*- "))
    If et1$ = et2$ Then
        List1.RemoveItem r
    End If
Next

Open App.Path + "\usenet.txt" For Output As #1
For r = 1 To List1.ListCount - 2
    Print #1, List1.List(r)
Next
Close #1

End Sub

