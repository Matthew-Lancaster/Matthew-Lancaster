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

Call SaveAttachments_Good

End

End Sub

Public Sub SaveAttachments_Good()

' Instantiate an Outlook Application object.
Dim objOL As New Outlook.Application

' METHOD THROW ERROR IS NOT SAME
' Dim objOL As Outlook.Application



Dim objAttachments As Outlook.Attachments
' Dim objSelection As Outlook.Selection
Dim objSelection As MAPIFolder
Dim objMsg As Outlook.MailItem 'Object
Dim oNameSpace As NameSpace

Dim i As Long
Dim lngCount As Long
Dim strFile As String
Dim strFolderpath As String
Dim strDeletedFiles As String

Dim RR

' Get the path
strFolderpath = "D:\GM OUTLOOK\"


' Set oFolder = oNameSpace.PickFolder
' Set oFolder = oNameSpace2.GetDefaultFolder(23)  'olFolderJunkFolder
' Set oFolder = oNameSpace3.GetDefaultFolder(3)   'olFolderDeleted
' Set oFolder = oNameSpace.GetDefaultFolder(6)    'olFolderInBox


Set oNameSpace = objOL.GetNamespace("MAPI")

Set oMAPI = GetObject("", "Outlook.application").GetNamespace("MAPI")
' Set oParentFolder = oMAPI.Folders("nvr1000.roidsrim@gmail.com")
' Set oParentFolder = oMAPI.Folders(oNameSpace.GetDefaultFolder(0).Name)
' Set oParentFolder = oNameSpace.GetDefaultFolder(1)
Set oParentFolder = oMAPI.Folders(1)
If oParentFolder.Folders.Count Then
  For RR = 1 To oParentFolder.Folders.Count
    If Trim(oParentFolder.Folders(RR).Name) = "Inbox" Then
        Exit For
    End If
    ' MsgBox Trim(oParentFolder.Folders(RR).Name)
  Next RR
End If

' MsgBox oParentFolder.Folders(RR)

Set objSelection = oParentFolder.Folders(RR)

'r = olFolderDeleted
TT_VAR = objSelection.Name

'MsgBox objSelection.Name

'MsgBox Str(objSelection.Items.Count)

' Set MyItems = oFolder.Items

' Set objMsg = objSelection.Items

' Check each selected item for attachments.
' For Each objMsg In objSelection


For RR = objSelection.Items.Count To 1 Step -1
'For rr = 1 To objSelection.Items.Count

        Set objMsg = objSelection.Items.Item(RR)

        Set objAttachments = objMsg.Attachments
        lngCount = objAttachments.Count
        A1 = objMsg.Subject
        D1 = objMsg.ReceivedTime
        D2 = " " + Format(D1, "YYYYMMDDHHMMSS") + " "
        D3 = Format(D1, "YYYYMMDD")
        
        
        
        
'        D3 = objMsg.EntryID
'        D3 = objMsg.ReceivedByEntryID
'        D3 = objMsg.ReceivedOnBehalfOfEntryID
'        D3 = objMsg.Class
'        D3 = objMsg.MessageClass
'        D3 = objMsg.ConversationIndex
'
'        MsgBox D3
        

        
        SET_GO = False
        If InStr(A1, "Alarm Message From Device EDVR") > 0 Then
            SET_GO = True
            H3 = "Motion Alarm On Device EDVR_"
            D2 = "" ' REMOVE THE TIME FOR THIS ONE AS IT IN .JPG ATTACHMENT
        End If
        ' Embedded Net DVR: Motion Detected On Channel A3
        If InStr(A1, "Embedded Net DVR: Motion Detected On Channel") > 0 Then
            SET_GO = True
            H1 = Mid(A1, Len(A1) - 5)
            H2 = Format(Val(Mid(A1, Len(A1) - 4)), "00")
            ' H3 = "DVR_A03_Motion_Detected"
            H3 = "DVR_" + H1 + H2 + "_Motion_Detected"
        End If
        If lngCount = 0 Then
            SET_GO = False
        End If
        If D1 >= Now - 10 Then
            SET_GO = False
        End If
        
        'a1 = MyItems.Item(rrr).SenderName
        'a2 = MyItems.Item(rrr).SenderEmailAddress
        'a3 = Trim(MyItems.Item(rrr).Subject)
        'a4 = Trim(MyItems.Item(rrr).Body)


        ' Use the MsgBox command to troubleshoot. Remove it from the final code.
        ' MsgBox A1
        ' MsgBox Str(lngCount)
                    
        If SET_GO = True Then
            If lngCount > 0 Then
    
                    ' Use a count down loop for removing items
                    ' from a collection. Otherwise, the loop counter gets
                    ' confused and only every other item is removed.
                    For i = lngCount To 1 Step -1
                            UNIQUE = 0
    
                            ' Get the file name.
                            strFile = objAttachments.Item(i).FileName
                            Do
                                UNIQUE_VAR = Format(UNIQUE, "0")
                                
                                If InStr(A1, "Alarm Message From Device EDVR") > 0 Then
                                    A2 = "A" + Mid(strFile, InStr(strFile, "_") - 1, 1)
                                    strFolderpath_2 = strFolderpath + D3 + "\" + A2 + "\"
                                    If Dir(strFolderpath_2, vbDirectory) = "" Then
                                        On Error Resume Next
                                            MkDir strFolderpath + D3
                                        On Error GoTo 0
                                        MkDir strFolderpath + D3 + "\" + A2
                                    End If
                                    If Dir(strFolderpath_2, vbDirectory) = "" Then
                                        MsgBox "PATH FAILED TO MAKE" + vbCrLf + strFolderpath_2
                                    End If
                                    
                                    Do
                                    
                                        strFile_F0 = Mid(strFile, 1, 5) + "20" + Mid(strFile, 6)
                                        
                                        strFile_F1 = Mid(strFile_F0, 1, InStr(strFile_F0, ".") - 1)
                                        If InStr(strFile_F0, ".") > 0 Then
                                            strFile_F1 = Mid(strFile_F0, 1, InStr(strFile_F0, ".") - 1)
                                        End If
                                        If InStr(10, strFile_F0, "_") > 0 Then
                                            strFile_F1 = Mid(strFile_F0, 1, InStr(10, strFile_F0, "_", 10) - 1)
                                        End If
                                        If UNIQUE = 0 Then
                                            If InStr(10, strFile_F0, "_") > 0 Then
                                                UNIQUE = Mid(strFile_F0, InStr(10, strFile_F0, "_") + 1)
                                            End If
                                        End If
                                        If InStrRev(UNIQUE, ".") > 0 Then
                                            UNIQUE = Mid(UNIQUE, 1, InStrRev(UNIQUE, ".") - 1)
                                        End If
                                        UNIQUE_VAR = Format(Val(UNIQUE), "0")
                                        strFile_F1 = strFile_F1 + "_" + UNIQUE_VAR + ".jpg"
                                        strFile_NAME = strFolderpath_2 + H3 + D2 + " " + strFile_F1
                                        If Dir(strFile_NAME) <> "" Then
                                            UNIQUE = UNIQUE + 1
                                        End If
                                    Loop Until Dir(strFile_NAME) = ""
                                End If
                                
                                If UNIQUE = 0 Then UNIQUE_VAR = ""
                                
                                If InStr(A1, "Embedded Net DVR: Motion Detected On Channel") > 0 Then
                                    Stop
                                    strFile_NAME = strFolderpath_2 + H3 + D2 + strFile
                                End If
                                
                                If Dir(strFile_NAME) <> "" Then
                                    UNIQUE = UNIQUE + 1
                                End If
                            Loop Until Dir(strFile_NAME) = ""
    
                            ' Save the attachment as a file.
                            objAttachments.Item(i).SaveAsFile strFile_NAME
                            

                            
    
                    Next i
            
                    
                    If Dir(strFile_NAME) <> "" Then
                        If D1 < Now - 10 Then
                            objMsg.Delete
                            
                            RR_X = RR_X + 1
                            If RR_X >= 10000 Then GoTo ExitSub
                        End If
                    End If
            
            
            End If
        End If
Next
    
ExitSub:

Set objAttachments = Nothing
Set objMsg = Nothing
Set objSelection = Nothing
Set objOL = Nothing

End Sub

