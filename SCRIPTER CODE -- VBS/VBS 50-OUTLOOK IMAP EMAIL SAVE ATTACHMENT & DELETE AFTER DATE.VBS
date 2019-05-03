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

' Get the path to your My Documents folder
strFolderpath = CreateObject("WScript.Shell").SpecialFolders(16)
On Error Resume Next

' Instantiate an Outlook Application object.
Set objOL = CreateObject("Outlook.Application")

' Get the collection of selected objects.
Set objSelection = objOL.ActiveExplorer.Selection

' Set the Attachment folder.
strFolderpath = strFolderpath & "OLAttachments"

'Use the MsgBox command to troubleshoot. Remove it from the final code.
MsgBox strFolderpath

' Check each selected item for attachments. If attachments exist,
' save them to the Temp folder and strip them from the item.
For Each objMsg In objSelection

    ' This code only strips attachments from mail items.
    If objMsg.class=olMail Then
    ' Get the Attachments collection of the item.
        Set objAttachments = objMsg.Attachments
        lngCount = objAttachments.Count

        'Use the MsgBox command to troubleshoot. Remove it from the final code.
        MsgBox objAttachments.Count

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