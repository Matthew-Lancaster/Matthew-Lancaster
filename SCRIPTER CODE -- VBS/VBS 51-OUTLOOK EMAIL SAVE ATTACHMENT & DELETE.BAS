'====================================================================
'# __ C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 51-OUTLOOK EMAIL SAVE ATTACHMENT & DELETE.VBS
'# __
'# __ VBS 51-OUTLOOK EMAIL SAVE ATTACHMENT & DELETE.VBS
'# __
'# BY Matthew __ Matt.Lan@Btinternet.com
'# __
'# 1st Work
'# START     TIME [ Fri 03-May-2019 17:55:00 ]
'# END FINAL TIME [ Fri 03-May-2019 18:10:00 ]
'# 1st Work Total Day Resume After
'# START     TIME [ Fri 03-May-2019 17:55:00 ]
'# END FINAL TIME [ Sat 04-May-2019 00:00:49 ]
'# __
'# Last Work -- SESSION 04 OF 04
'# START     TIME [ Sun 05-May-2019 09:15:00 ]
'# END FINAL TIME [ Sun 05-May-2019 11:32:00 ]
'# __
'# Time Span Begin to Ender
'# START     TIME [ Fri 03-May-2019 18:10:00 ]
'# END FINAL TIME [ Sun 05-May-2019 11:32:00 ]
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
'
' WELL I INCLUDE IT IDEALLY AS A VBS BUT WITH EXTENSION .BAS 
' IT BETTER SUIT TO RUN AS BAS RATHER THAN VBS

'#-------------------------------------------------------------------
' SESSION 001 __ CREATED
' -------------------------------------------------------------------
'# START     TIME [ Fri 03-May-2019 18:10:00 ]
'# END FINAL TIME [ Fri 03-May-2019 18:10:00 ]
'#-------------------------------------------------------------------


'#-------------------------------------------------------------------
' SESSION 002 __ CREATE DAY -- LATE NIGHT CODE 
' & OTHER CODE WAS 12:20 BED AND UP AGAIN MORE CODE AT 3:20 AM
' FRIDAY DAY WAKE UP WAS 7 AM
' -------------------------------------------------------------------
'# START     TIME [ Fri 03-May-2019 20:50:05 ]
'# END FINAL TIME [ Sat 04-May-2019 00:00:49 ]
'#-------------------------------------------------------------------


'#-------------------------------------------------------------------
' SESSION 003 __
' -------------------------------------------------------------------
' FEW MORE WORK THE NEXT DAY
' SEEM FAULT AS NOT ABLE TO READ A IMAP FOLDER
' SO USE POP3 FOR NOW
' ADDED ON EMPTY THE DELETED FOLDER AFTER USER AND BEFORE FUNNY
' YET TO SEE OF DELETE ON THIS CLIENT SIDE HAPPEN ON-LINE WEB SERVER
' TO USE THE CODE
' IT IS MORE LIKE VBA VB6 THAN VBSCRIPT
' BUT STILL SAME ALMOST
' THE IDE IN OUTLOOK IS QUITE A GOOD ONE
' FOR PROPER-ING SYNTAX AND THING
' AND USUAL HOT KEY TO RUN IT ALL
' LOOK UP MACRO SET SECURITY LOW SO ABLE RUN THE SCRIPT HERE
' WHEN GO TO VISUAL BASIC EDITOR
' PASTE THE CODE INTO THIS_OUTLOOK_SESSION ALL ONE WORD
' YOU CAN HAVE ANOTHER MODULE IF CODE BIGGER
' AND THEN LOOK AT MACRO
' EVERY ROUTINE YOU PUT IN WILL CAN AS MACRO TO RUN
' SEEM TO DO IN ORDER FIRST COME
' SO TOP MAIN ROUTINE COULD SUB CALL INTO A FEW OTHER
' HAD DIFFICULTY AS WANT TO SCHEDULE IN SOMEHOW AND ON
' FOUND ANSWER HERE
' ----
' Outlook's Rules and Alerts: Run a Script
' https://www.slipstick.com/outlook/rules/outlooks-rules-and-alerts-run-a-script/
' ----
' YOU GOT TO INSERT THIS SNIPPET INTO YOUR CODE AS MAIN EXECUTION ROUTINE CALL
' TO USER RUN BY RULE
' -------------------------------------------------------------------
' Public Sub ShowMessage(Item As Outlook.MailItem)
' Call MAIN
' End Sub
'
' EMAIL ARRIVE DOESN'T SHOW RUN SCRIPT WITH MACRO NOT SHOW THERE TO SELECT IN
' HAD TO FIND MY OWN CODE WHICH USED TO RUN ENTIRELY WITHIN VBA AS A STAND ALONE APP
' AS WEBSITE PROVIDE CODE MAY OF WORK ONE TIME BUT BIT BUGGY FOR CLEVER
' ANOTHER NICE THING TO LEARN
' JUST LIKE THE VB IDE ENVIRONMENT
' YOU CAN CHANGE CONTROL BUTTON CUSTOMIZE QUITE MANICALLY TO WHATEVER YOU WANTER
' AND HAVE THE SCRIPT TO RUN AS A BUTTON ANYTIME WANT
' REMEMBER TO EXIT PROPERLY LET IT DO IT CLOSE DOWN WHICH TAKE TIME EVERY NOW AND THEN
' OR ELSE ALL CUSTOMIZATION GONE & AND YOUR CODE UP-TO-DATE
' JUST DRAG AND DROP THEM CONTROL CUSTOMIZE
' AND SET ICON HAVE LABEL ALSO
' YOU CAN EVEN DRAG A MAIN MENU ITEM FROM PULL DOWN
' AND DRAG DROP AS A BUTTON THERE EASIER THAN VB6 WOULD BE
' -------------------------------------------------------------------
'# START     TIME [ Sat 04-May-2019 05:43:27 ]
'# END FINAL TIME [ Sat 04-May-2019 13:28:00 ]
'#-------------------------------------------------------------------

'#-------------------------------------------------------------------
' SESSION 004 __
' -------------------------------------------------------------------
' NEW 3RD DAY
' LOT MORE CODE TODAY
' ERROR CHECKER IS MORE IMPORTANT
' ERROR TRAPPING DOESN'T QUITE REPORT A LOT
' IF ERROR IT BREAK THE CODE AND JUMP BACK TO SOME PART BEFORE LIKE IN NEST LOOP
' MUST CHECK COUNT OF ITEM BEFORE ASK COLLECT THEM
' SOUND EFFECT INCLUDED RATHER THAN SET RULE
' WAS A DESIRE NIGHT BEFORE TO ALLOW CODE RUN UNTIL FINISH
' AND THAT DOES HAPPEN WITH ERROR TRAP ALL CLEAR
' ANOTHER THING IS LESS SPAGHETTI CODE ON UNIQUE NAME
' AND A STOP THING
' WITH COMPARE SO THAT ANY IMAGE WITH SAME FILENAME BUT 
' DIFFERENT CONTENT ARE CHECKED OVER
' THAT IS IMPORTANT BECAUSE ALL ATTACHMENT ARE COLLECTING
' BUT NOT WANT DELETE ANY EMAIL UNLESS WEEK OLDER
' MAYBE DEVISE A SYSTEM FOR THAT ONE AS DELAY EACH CHECK MAIL LATER
' WHEN GOT OVER THE INITIAL BACKLOG TO DO WHICH OKAY AT MOMENT
' NOT GOING TO BE SPEEDY FOR EACH MAIL ARRIVE CHECK ON TEN DAY WORTH
' OKAY FOR NOW WORKING -- GET ALL DONE
' -------------------------------------------------------------------
'# START     TIME [ Sun 05-May-2019 09:15:00 ]
'# END FINAL TIME [ Sun 05-May-2019 11:32:00 ]
'#-------------------------------------------------------------------

' -------------------------------------------------------------------
' SOURCE HERE
' -------------------------------------------------------------------
' Save Attachments to the Hard Drive
' https://www.slipstick.com/developer/save-attachments-to-the-hard-drive/
' -------------------------------------------------------------------

'# ------------------------------------------------------------------
' LOCATION ON-LINE 
'# ------------------------------------------------------------------
' Matthew-Lancaster/VBS 51-OUTLOOK EMAIL SAVE ATTACHMENT & DELETE.BAS
' https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20VBS/VBS%2051-OUTLOOK%20EMAIL%20SAVE%20ATTACHMENT%20%26%20DELETE.BAS
' -------------------------------------------------------------------


Private Declare Function sndPlaySound32 _
    Lib "winmm.dll" _
    Alias "sndPlaySoundA" ( _
        ByVal lpszSoundName As String, _
        ByVal uFlags As Long) As Long

'sndPlaySound32 "C:\Windows\Media\Chimes.wav", 0&
'If the file cannot be found, Windows will play the default sound, unless this is disabled by using the value SND_NODEFAULT in the uFlags parameter (see below).
'
'The uFlags parameter specifies how the sound is to be played. You can combine multiple options by adding their respective values together and passing the sum in the uFlags parameter.
'
'Const SND_SYNC = &H0        ' (Default) Play the sound synchronously. Code execution
'                            ' pauses until sound is complete.
'
'Const SND_ASYNC = &H1       ' Play the sound asynchronously. Code execution
'                            ' does not wait for sound to complete.
'
'Const SND_NODEFAULT = &H2   ' If the specified sound is not found, do not play
'                            ' the default sound (no sound is played).
'
'Const SND_MEMORY = &H4      ' lpszSoundName is a memory file of the sound.
'                            ' Not used in VBA/VB6.
'
'Const SND_LOOP = &H8        ' Continue playing sound in a loop until the next
'                            ' call to sndPlaySound.
'
'Const SND_NOSTOP = &H10     ' Do not stop playing the current sound before playing
'                            ' the specified sound.

Sub MAIN()

sndPlaySound32 "C:\Windows\Media\Windows Ringout.wav", 10&

Call RemoveAllItemsAndFoldersInDeletedItems

Call SaveAttachments_Good

'Call RemoveAllItemsAndFoldersInDeletedItems

End

End Sub

Public Sub ShowMessage(Item As Outlook.MailItem)

Call MAIN

End Sub

Sub RemoveAllItemsAndFoldersInDeletedItems()
 
Dim oDeletedItems As Outlook.Folder
Dim oFolders As Outlook.Folders
Dim oItems As Outlook.Items
Dim i As Long
'Obtain a reference to deleted items folder
Set oDeletedItems = Application.Session.GetDefaultFolder(olFolderDeletedItems)
If oDeletedItems.Items.Count > 0 Then
    Set oItems = oDeletedItems.Items
    For i = oItems.Count To 1 Step -1
        oItems.Item(i).Delete
        X = X + 1
        If X Mod 10 = 0 Then DoEvents
    Next
    If oDeletedItems.Folders.Count > 0 Then
        Set oFolders = oDeletedItems.Folders
        For i = oFolders.Count To 1 Step -1
            oFolders.Item(i).Delete
            X = X + 1
            If X Mod 10 = 0 Then DoEvents
        Next
    End If
End If
'----
'Delete All Items and Subfolders in the Deleted Items Folder | Microsoft Docs
'https://docs.microsoft.com/en-us/office/vba/outlook/how-to/items-folders-and-stores/delete-all-items-and-subfolders-in-the-deleted-items-folder
'----

End Sub

Public Sub SaveAttachments_Good()


' Init an Outlook Application object.
Dim objOL As New Outlook.Application

' ---------------------------------
' METHOD THROW ERROR IS NOT SAME
' GLAD HAD MY OLDER ARCHIVE OF CODE
' THESE THING TAKE FOREVER TO DEBUG
' ---------------------------------
' Dim objOL As Outlook.Application
' Dim objSelection As Outlook.Selection

Dim objAttachments As Outlook.Attachments
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
' Set oParentFolder = oMAPI.Folders("nvr1000.roidsrim@gmail.com")
' Set oParentFolder = oMAPI.Folders(oNameSpace.GetDefaultFolder(0).Name)
' Set oParentFolder = oNameSpace.GetDefaultFolder(1)

Set oNameSpace = objOL.GetNamespace("MAPI")

Set oMAPI = GetObject("", "Outlook.application").GetNamespace("MAPI")

Set oParentFolder = oMAPI.Folders(1)
If oParentFolder.Folders.Count Then
  For RR = 1 To oParentFolder.Folders.Count
    If Trim(oParentFolder.Folders(RR).Name) = "Inbox" Then
        ' MsgBox Trim(oParentFolder.Folders(RR).Name)
        Exit For
    End If
  Next RR
End If

Set objSelection = oParentFolder.Folders(RR)

Set DeletedFolder = oNameSpace.GetDefaultFolder(olFolderDeletedItems)

TT_VAR = objSelection.Name

'MsgBox objSelection.Name
'MsgBox Str(objSelection.Items.Count)

' Check each selected item for attachments.
' For Each objMsg In objSelection

X_COUNT = objSelection.Items.Count
If X_COUNT > 0 Then
For RR = X_COUNT To 1 Step -1

        Set objMsg = objSelection.Items.Item(RR)
    
        
        lngCount = objMsg.Attachments.Count
        If lngCount > 0 Then
            Set objAttachments = objMsg.Attachments
        End If
        A1 = objMsg.Subject
        D1 = objMsg.ReceivedTime
        D2 = Format(D1, "YYYYMMDDHHMMSS")
        D3 = Format(D1, "YYYYMMDD")
        
        'X = X + 1
        'If X Mod 10 = 0 Then DoEvents

'        TRY TO GET EMAIL MSG ID LIKE SOME OTHER APP ABLE TO - NOT HERE
'        --------------------------------------------------------------
'        D3 = objMsg.EntryID
'        D3 = objMsg.ReceivedByEntryID
'        D3 = objMsg.ReceivedOnBehalfOfEntryID
'        D3 = objMsg.Class
'        D3 = objMsg.MessageClass
'        D3 = objMsg.ConversationIndex
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
    
                    ' ---------------------------------------------------
                    ' Use a count down loop for removing items
                    ' from a collection. Otherwise, the loop counter gets
                    ' confused and only every other item is removed.
                    ' ---------------------------------------------------
                    For i = lngCount To 1 Step -1
                            X = X + 1
                            If X Mod 10 = 0 Then DoEvents
'                            UNIQUE = 0
    
                            ' Get the file name.
                            strFile = objAttachments.Item(i).FileName
'                            Do
                                'X = X + 1
                                'If X Mod 10 = 0 Then DoEvents
'                                UNIQUE_VAR = Format(UNIQUE, "0")
                                
                                ' -----------------------------------------------------
                                ' HOW BIG WAS THAT STRING MANIPULATE
                                ' WITHOUT ERROR
                                ' -----------------------------------------------------
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
                                    
'                                    Do
                                        'X = X + 1
                                        'If X Mod 10 = 0 Then DoEvents

                                    
                                        strFile_F0 = Mid(strFile, 1, 5) + "20" + Mid(strFile, 6)
                                        
'                                        strFile_F1 = Mid(strFile_F0, 1, InStr(strFile_F0, ".") - 1)
'                                        If InStr(strFile_F0, ".") > 0 Then
'                                            strFile_F1 = Mid(strFile_F0, 1, InStr(strFile_F0, ".") - 1)
'                                        End If
'                                        If InStr(10, strFile_F0, "_") > 0 Then
'                                            strFile_F1 = Mid(strFile_F0, 1, InStr(10, strFile_F0, "_", 10) - 1)
'                                        End If
'                                        If UNIQUE = 0 Then
'                                            If InStr(10, strFile_F0, "_") > 0 Then
'                                                UNIQUE = Mid(strFile_F0, InStr(10, strFile_F0, "_") + 1)
'                                            End If
'                                        End If
'                                        If InStrRev(UNIQUE, ".") > 0 Then
'                                            UNIQUE = Mid(UNIQUE, 1, InStrRev(UNIQUE, ".") - 1)
'                                        End If
'                                        UNIQUE_VAR = Format(Val(UNIQUE), "0")
'                                        strFile_F1 = strFile_F1 + "_" + UNIQUE_VAR + ".jpg"
'                                        strFile_NAME = strFolderpath_2 + H3 + D2 + " " + strFile_F1
                                        
                                        strFile_NAME = strFolderpath_2 + H3 + D2 + strFile_F0
                                        strFile_TEMP = "D:\GM OUTLOOK\TEMP\" + strFile_F0
                                                

'                                        If Dir(strFile_NAME) <> "" Then
'                                            UNIQUE = UNIQUE + 1
'                                        End If
'                                    Loop Until Dir(strFile_NAME) = ""
                                End If
                                
'                                If UNIQUE = 0 Then UNIQUE_VAR = ""
                                
                                If InStr(A1, "Embedded Net DVR: Motion Detected On Channel") > 0 Then
                                    MsgBox "STOP REQUIRE -- GOT HERE - Embedded Net DVR: Motion Detected On Channel"
                                    Stop
                                    strFile_NAME = strFolderpath_2 + H3 + D2 + strFile
                                    strFile_TEMP = "D:\GM OUTLOOK\TEMP\" + strFile
                                End If
                                
                                If Dir(strFile_NAME) <> "" Then
'                                    UNIQUE = UNIQUE + 1
                                    If Dir("D:\GM OUTLOOK\TEMP", vbDirectory) = "" Then
                                        On Error Resume Next
                                            MkDir "D:\GM OUTLOOK"
                                            MkDir "D:\GM OUTLOOK\TEMP"
                                        On Error GoTo 0
                                    End If
                                    If Dir("D:\GM OUTLOOK\TEMP", vbDirectory) = "" Then
                                        MsgBox "PATH FAILED TO MAKE" + vbCrLf + "D:\GM OUTLOOK\TEMP"
                                    End If
                                    
                                    objAttachments.Item(i).SaveAsFile strFile_TEMP
                                    
                                    FR1 = FreeFile
                                    Open strFile_TEMP For Binary As #FR1
                                        PIC1_COMPARE = Input$(LOF(FR1), FR1)
                                    Close #FR1
                                    FR1 = FreeFile
                                    Open strFile_NAME For Binary As #FR1
                                        PIC2_COMPARE = Input$(LOF(FR1), FR1)
                                    Close #FR1
                                    If PIC2_COMPARE <> PIC1_COMPARE Then
                                        MsgBox "PIC1_COMPARE<>PIC2_COMPARE" + vbCrLf + "IMAGE NOT THE SAME AT COMPARE" + vbCrLf + "OUTLOOK SCRIPTOR -- NOW STOP"
                                        Stop
                                    End If
                                End If
    
                            ' Save the attachment as a file.
                            objAttachments.Item(i).SaveAsFile strFile_NAME
                            If Dir(strFile_NAME) = "" Then Stop
                    Next i
                    
                    If D1 < Now - 10 Then
                        On Error Resume Next
                        'objMsg.Delete
                        'objMsg.Item(RR).Delete
                        objMsg.Move DeletedFolder
                        If Err.Number > 0 Then
                            Stop
                        End If
                        On Error GoTo 0
                        
                        X = X + 1
                        If X Mod 10 = 0 Then
                            DoEvents
                        End If
                        RR_X = RR_X + 1
                        If RR_X >= 100 Then
                            GoTo ExitSub
                        End If
                    End If
            
            End If
        End If
Next
End If
ExitSub:

Set objAttachments = Nothing
Set objMsg = Nothing
Set objSelection = Nothing
Set objOL = Nothing

'----
'Save Attachments to the Hard Drive
'https://www.slipstick.com/developer/save-attachments-to-the-hard-drive/
'----


End Sub


