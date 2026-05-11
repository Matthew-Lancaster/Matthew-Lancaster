VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Bits Code


'None USed
'Not Used
Private Sub Command1_Click()
On Error Resume Next
    
Me.lblExtract.Caption = ""
Me.lstExtractionStatus.Enabled = True
Me.lstExtractionStatus.Clear

Dim oApp As New Outlook.Application
Dim oNameSpace As NameSpace
Dim oFolder As MAPIFolder
Dim oMailitem As Object
Dim sMessage As String

Set oApp = New Outlook.Application
Set oNameSpace = oApp.GetNamespace("MAPI")
'Set oFolder = oNameSpace.GetDefaultFolder(olFolderInbox)
'oNameSpace.PickFolder
'msgbox onamespace.PickFolder.

Set oFolder = oNameSpace.PickFolder
Dim exCnt As Integer
Me.Command1.Enabled = False

'Set myItems = oFolder.Items
'Call oFolder.Items.Sort("[" & CalendarFieldToSort & "]", True)

For Each oMailitem In oFolder.Items
    With oMailitem
        If oMailitem.Attachments.Count > 0 Then
            
            
            
            
'          Look in Object Brower for info on oMailitem
            
            
            
            qtag = 0
            qws$ = ""
            For rrr = oMailitem.Attachments.Count To 1 Step -1
'                If InStr(oMailItem.Attachments.Item(rrr).FileName, "0807") Then Stop
                
                
                If rrr > 1 Then qws$ = " (" + Trim(Str(rrr)) + ")"
                
                Do
                    If qtag > 0 Then qws$ = " (" + Trim(Str(rrr + qtag)) + ")"
                
                    ert$ = Dir1.Path & "\" & _
                    oMailitem.Attachments.Item(rrr).Parent & "~~"
                    
                    UHT$ = oMailitem.Attachments.Item(rrr).FileName
                    uyt = InStrRev(UHT$, ".")
                    ert$ = ert$ + Mid$(UHT$, 1, uyt - 1) + qws$ + Mid$(UHT$, uyt)
                
                
                    For r = Len(Dir1.Path) + 2 To Len(ert$)
                        tyh$ = Mid$(ert$, r, 1)
                        If tyh$ = "|" Then Mid$(ert$, r, 1) = "-"
                        If tyh$ = ">" Then Mid$(ert$, r, 1) = "-"
                        If tyh$ = "<" Then Mid$(ert$, r, 1) = "-"
                        If tyh$ = ":" Then Mid$(ert$, r, 1) = "-"
                        If tyh$ = "\" Then Mid$(ert$, r, 1) = "-"
                        If tyh$ = "/" Then Mid$(ert$, r, 1) = "-"
                        If tyh$ = "?" Then Mid$(ert$, r, 1) = "-"
                        If tyh$ = "*" Then Mid$(ert$, r, 1) = "-"
                        If tyh$ = """" Then Mid$(ert$, r, 1) = "-"
                    Next
                
                    DoEvents
                
                    If Dir$(ert$) = "" Or qtag > 10 Then Exit Do
                    qtag = qtag + 1
                Loop Until 1 = 1
                
                perd$ = oMailitem.Attachments.Item(rrr).FileName
                oMailitem.Attachments.Item(rrr).SaveAsFile ert$
                'ee$ = oMailitem.NoteItem.Item(rrr).Body
                
                
                oMailitem.Attachments.Item(rrr).SaveAsFile ert$
                'oMailItem.Attachments
                    
                If Check1.Value = vbChecked Then
                oMailitem.Attachments.Item(rrr).Delete
                oMailitem.Save
                End If
                        
                DoEvents
                lstExtractionStatus.AddItem (oMailitem.Attachments.Item(rrr).Parent)
                exCnt = exCnt + 1
                lblExtract.Caption = exCnt & " extracted"
 
           
            
            
            
            Next
        End If
    End With
Next oMailitem

Set oMailitem = Nothing
Set oFolder = Nothing
Set oNameSpace = Nothing
Set oApp = Nothing
    
    
Me.Command1.Enabled = True
End Sub




Sub ttt()

GoTo jump1

Open App.Path + "\Email Addys InBox.txt" For Output As #1

On Local Error Resume Next
For rrr = 1 To MyItems.Count
t1$ = MyItems.Item(rrr).SenderEmailAddress
g1 = InStr(t1$, "@")
g$ = Mid$(t1$, g1)
If InStr(r2$, g$) = 0 Then
    r2$ = r2$ + " --" + g$

        
        
    Print #1, g$
End If

'Print #1, MyItems.Item(rrr).Sunject
'Print #1, MyItems.Item(rrr).SenderName
'Print #1, MyItems.Item(rrr).
'Print #1, MyItems.Item(rrr).
'Print #1, MyItems.Item(rrr).ReceivedTime
'Exit For
'If rrr > 2 Then Exit For
Next

On Local Error GoTo 0

Close #1

'a2$ = MyItems.Item(rrr).SenderEmailAddress


End
jump1:

End Sub



