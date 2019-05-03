Attribute VB_Name = "Module1"
'====================================================================
'# __ C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 51-OUTLOOK EMAIL SAVE ATTACHMENT & DELETE.VBS
'# __
'# __ VBS 51-OUTLOOK EMAIL SAVE ATTACHMENT & DELETE.VBS
'# __
'# BY Matthew __ Matt.Lan@Btinternet.com
'# __
'# START     TIME [ Fri 03-May-2019 18:10:00 ]
'# END FINAL TIME [ Fri 03-May-2019 18:10:00 ]
'# __
'====================================================================

'# ------------------------------------------------------------------
' DESCRIPTION
' -------------------------------------------------------------------
' IDEA GMAIL WITH ONLY 15 GIG KEEP FILL UP FULL WITH ATTACHMENT
' REQUIRE SAVE ATTACHMENT AND DELETE EMAIL AFTER A DATE
' BEEN SPEND SOME TIME CHECKER SOFTWARE TO DO THIS FROM
' Fri 03-May-2019 12:01:57
' ALL VERY EXPENSIVE SOFTWARE
' EASIER WAY IS USE OUTLOOK AND VB-SCRIPTOR
' IS SOME LEARN DONE
' IMAP FOLDER IS BEST WAY TO GO
' WHEN READ THE GMAIL
' THE FOLDER STRUCTURE APPEAR BUT ALL IS IN INBOX
' GET OVER LIMIT OF POP3 MAIL ONLY READ INBOX NOT WHAT IN FOLDER
' OOPS WELL TRY FIND VSB SCRIPT BUT SEEM HARD MUST USE VBA AS HERE
'#-------------------------------------------------------------------
' SESSION 001 __ CREATED
' -------------------------------------------------------------------
'# START     TIME [ Fri 03-May-2019 18:10:00 ]
'# END FINAL TIME [ Fri 03-May-2019 18:10:00 ]
'#-------------------------------------------------------------------

' -------------------------------------------------------------------
' SOURCE HERE
' -------------------------------------------------------------------
' Save Attachments to the Hard Drive
' https://www.slipstick.com/developer/save-attachments-to-the-hard-drive/
' -------------------------------------------------------------------

Sub MAIN()

Call SaveAttachments

End Sub

Public Sub SaveAttachments()
Dim objOL As Outlook.Application
Dim objMsg As Outlook.MailItem 'Object
Dim objAttachments As Outlook.Attachments
Dim objSelection As Outlook.Selection
Dim i As Long
Dim lngCount As Long
Dim strFile As String
Dim strFolderpath As String
Dim strDeletedFiles As String

' Get the path
strFolderpath = "D:\GM OUTLOOK\"
On Error Resume Next

' Instantiate an Outlook Application object.
Set objOL = Application

' Get the collection of selected objects.
Set objSelection = objOL.ActiveExplorer.Selection

' Check each selected item for attachments.
For Each objMsg In objSelection

        Set objAttachments = objMsg.Attachments
        lngCount = objAttachments.Count

        'Use the MsgBox command to troubleshoot. Remove it from the final code.
        MsgBox objAttachments.Count

        GoTo ExitSub
        
        If lngCount > 0 Then

                ' Use a count down loop for removing items
                ' from a collection. Otherwise, the loop counter gets
                ' confused and only every other item is removed.

                For i = lngCount To 1 Step -1

                        ' Get the file name.
                        strFile = objAttachments.Item(i).FileName

                        ' Combine with the path to the Temp folder.
                        strFile = strFolderpath & strFile

                        ' Save the attachment as a file.
                        objAttachments.Item(i).SaveAsFile strFile

                Next i
        End If

Next
    
ExitSub:

Set objAttachments = Nothing
Set objMsg = Nothing
Set objSelection = Nothing
Set objOL = Nothing
End Sub
