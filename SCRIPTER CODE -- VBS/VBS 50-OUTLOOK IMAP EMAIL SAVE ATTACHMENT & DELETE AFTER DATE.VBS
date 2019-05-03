'====================================================================
'# __ C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 50-OUTLOOK IMAP EMAIL SAVE ATTACHMENT & DELETE AFTER DATE.VBS
'# __ 
'# __ VBS 50-OUTLOOK IMAP EMAIL SAVE ATTACHMENT & DELETE AFTER DATE.VBS
'# __ 
'# BY Matthew __ Matt.Lan@Btinternet.com
'# __ 
'# START     TIME [ Fri 22-Mar-2019 10:31:08 ]
'# END FINAL TIME [ Fri 22-Mar-2019 11:50:00 ]
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
' EASIER WAY IS USE OUTLOOK AND VBSCRIPTOR
' IS FOR VBA -- WE DON'T WANT THAT - WHAT WANTED IS VBS
'#-------------------------------------------------------------------
' SESSION 001 __ CREATED
' -------------------------------------------------------------------
'# START     TIME [ Fri 03-May-2019 17:40:00 ]
'# END FINAL TIME [ Fri 03-May-2019 17:40:00 ]
'#-------------------------------------------------------------------

' -------------------------------------------------------------------
' SOURCE HERE
' -------------------------------------------------------------------
' vba - Remove email attachments after saving them and adding save location to email in Outlook - Stack Overflow
' https://stackoverflow.com/questions/26745239/remove-email-attachments-after-saving-them-and-adding-save-location-to-email-in
' -------------------------------------------------------------------


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
Set objOL = CreateObject("Outlook.Application")

' Get the collection of selected objects.
Set objSelection = objOL.ActiveExplorer.Selection

'Use the MsgBox command to troubleshoot. Remove it from the final code.
MsgBox strFolderpath

' Check each selected item for attachments. 
For Each objMsg In objSelection

    ' This code only strips attachments from mail items.
    If objMsg.class=olMail Then
    ' Get the Attachments collection of the item.
        Set objAttachments = objMsg.Attachments
        lngCount = objAttachments.Count

        'Use the MsgBox command to troubleshoot. Remove it from the final code.
        MsgBox objAttachments.Count

		Set objAttachments = Nothing
		Set objMsg = Nothing
		Set objSelection = Nothing
		Set objOL = Nothing
		
		WScript.Quit(0)
		
		
        If lngCount > 0 Then

            ' We need to use a count down loop for removing items
            ' from a collection. Otherwise, the loop counter gets
            ' confused and only every other item is removed.

            For i = lngCount To 1 Step -1

                ' Save attachment before deleting from item.
                ' Get the file name.
                strFile = objAttachments.Item(i).FileName

                ' Combine with the path to the folder.
                strFile = strFolderpath & strFile

                ' Save the attachment as a file.
                objAttachments.Item(i).SaveAsFile strFile

                ' Delete the attachment.
                objAttachments.Item(i).Delete

                'write the save as path to a string to add to the message
                'check for html and use html tags in link
                If objMsg.BodyFormat <> olFormatHTML Then
                    strDeletedFiles = strDeletedFiles & vbCrLf & "<file://" & strFile & ">"
                Else
                    strDeletedFiles = strDeletedFiles & "<br>" & "<a href='file://" & _
                    strFile & "'>" & strFile & "</a>"
                 End If

                 'Use the MsgBox command to troubleshoot. Remove it from the final code.
                  MsgBox strDeletedFiles

            Next i
        End If

        ' Adds the filename string to the message body and save it
        ' Check for HTML body
        If objMsg.BodyFormat <> olFormatHTML Then
            objMsg.Body = objMsg.Body & vbCrLf & _
              "The file(s) were saved to " & strDeletedFiles
        Else
            objMsg.HTMLBody = objMsg.HTMLBody & "<p>" & _
              "The file(s) were saved to " & strDeletedFiles & "</p>"
        End If

        objMsg.Save
        'sets the attachment path to nothing before it moves on to the next message.
        strDeletedFiles = ""

    End If
Next

ExitSub:

Set objAttachments = Nothing
Set objMsg = Nothing
Set objSelection = Nothing
Set objOL = Nothing

End Sub