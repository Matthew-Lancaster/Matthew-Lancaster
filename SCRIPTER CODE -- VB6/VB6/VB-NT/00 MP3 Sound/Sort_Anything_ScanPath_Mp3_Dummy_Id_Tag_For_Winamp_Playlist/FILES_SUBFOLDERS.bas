Attribute VB_Name = "FILES_SUBFOLDERS"
Dim RR


Public Function Get_All_Directory_Files_With_Wildcard_1ST_MP3_BIGGER_ZERO_SIZE(ByVal tfolder As String, _
 ByVal getsubdirs As Boolean, _
 ByVal wildcard As String) _
 As Long

'made by Alexander Triantafyllou
'alextriantf@yahoo.gr
'Athens-Greece


'FILE AND FOLDER TYPE REQUIRES THIS
'Reference=*\G{420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#H:\WINNT\system32\scrrun.dll#Microsoft Scripting Runtime

    
Set fso = CreateObject("Scripting.FileSystemObject")
Dim objfile As File
Dim objfolder As Folder
Dim kokovar As Variant
Dim k As Long
Dim wildext As String
Dim wildexts As String
Dim wildfirst As String
Dim wildexte As String
Dim wildfirsts As String
Dim wildfirste As String
Dim examfirst As String
Dim examext As String
Dim afl_filetext As String

kokovar = Split(wildcard, ",")

    If tfolder <> "" Then


        For Each objfile In fso.GetFolder(tfolder).Files
            'do the stuff we want with the files
            For k = 0 To UBound(kokovar)
         wildext = LCase(cutgetExtension(kokovar(k)))
         wildfirst = LCase(Mid(kokovar(k), 1, Len(kokovar(k)) - Len(wildext) - 1))
            
       If InStr(1, wildext, "*") = 0 Then
       wildexts = "888NONE888"
       wildexte = "888NONE888"
       Else
       wildexts = Mid(wildext, 1, InStr(1, wildext, "*") - 1)
       wildexte = Mid(wildext, InStr(1, wildext, "*") + 1, Len(wildext) - InStr(1, wildext, "*"))
       End If
       
        If InStr(1, wildfirst, "*") = 0 Then
       wildfirsts = "888NONE888"
       wildfirste = "888NONE888"
       Else
       wildfirsts = Mid(wildfirst, 1, InStr(1, wildfirst, "*") - 1)
       wildfirste = Mid(wildfirst, InStr(1, wildfirst, "*") + 1, Len(wildfirst) - InStr(1, wildfirst, "*"))
       End If
            
        examfirst = LCase(cutgetName(cutfilename(CStr(objfile))))
        examext = LCase(cutgetExtension(CStr(objfile)))
            
        If wildexts = "888NONE888" Then
'we do not have a wildcard in the extension
        If wildfirsts = "888NONE888" Then
        'we do not have a wildcard neither on the beggining or the

'extension
        If examfirst = wildfirst And examext = wildext Then
        afl_filetext = afl_filetext + objfile + vbNewLine
        End If
                
        Else
        
        
        '###############
        'we do have a wildcard in the beggining but not in
        'the extension
        If Mid(examfirst, 1, Len(wildfirsts)) = wildfirsts And _
        Mid(examfirst, Len(wildfirst) - Len(wildfirste) + 1, Len(wildfirste)) = wildfirste And wildext = examext Then
        
        'afl_filetext = afl_filetext + objfile + vbNewLine
        
        If objfile.Size > 0 Then
            
            Get_All_Directory_Files_With_Wildcard_1ST_MP3_BIGGER_ZERO_SIZE = True
            Exit Function
        
        End If
        End If
        End If
        
        Else
        'we do not have a wildcard in the extension
        If wildfirsts = "888NONE888" Then
        'we do have a wildcard in the beggining but not in the
        'extension
        If Mid(examext, 1, Len(wildexts)) = wildexts And _
        Mid(examext, Len(wildext) - Len(wildexte) + 1, Len(wildexte)) = wildexte Then
        afl_filetext = afl_filetext + objfile + vbNewLine
        
        End If
            
        Else
        'we have a wildcard in both beggining and extension
        
            If Mid(examext, 1, Len(wildexts)) = wildexts And _
        Mid(examext, Len(wildext) - Len(wildexte) + 1, Len(wildexte)) = wildexte _
         And Mid(examfirst, 1, Len(wildfirsts)) = wildfirsts And _
        Mid(examfirst, Len(wildfirst) - Len(wildfirste) + 1, Len(wildfirste)) = wildfirste Then
        afl_filetext = afl_filetext + objfile + vbNewLine
        End If
            
            End If
            
            End If
            
            'telos if
                    
            Next k
            
        Next

If getsubdirs Then

        For Each objfolder In fso.GetFolder(tfolder).SubFolders
           RR = Get_All_Directory_Files_With_Wildcard_1ST_MP3_BIGGER_ZERO_SIZE(CStr(objfolder), getsubdirs, wildcard)
            DoEvents
            If RR = True Then
                
                Exit Function
            End If
        Next
       
    End If
End If

Set fso = Nothing

'Get_All_Directory_Files_With_Wildcard = afl_filetext



End Function


Public Function Get_All_Directory_Files_With_Wildcard_ORIGINAL(ByVal tfolder As String, _
 ByVal getsubdirs As Boolean, _
 ByVal wildcard As String) _
 As String

'made by Alexander Triantafyllou
'alextriantf@yahoo.gr
'Athens-Greece


'FILE AND FOLDER TYPE REQUIRES THIS
'Reference=*\G{420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#H:\WINNT\system32\scrrun.dll#Microsoft Scripting Runtime
    
Set fso = CreateObject("Scripting.FileSystemObject")
Dim objfile As File
Dim objfolder As Folder
Dim kokovar As Variant
Dim k As Long
Dim wildext As String
Dim wildexts As String
Dim wildfirst As String
Dim wildexte As String
Dim wildfirsts As String
Dim wildfirste As String
Dim examfirst As String
Dim examext As String
Dim afl_filetext As String

kokovar = Split(wildcard, ",")

    If tfolder <> "" Then


        For Each objfile In fso.GetFolder(tfolder).Files
            'do the stuff we want with the files
            For k = 0 To UBound(kokovar)
         wildext = LCase(cutgetExtension(kokovar(k)))
         wildfirst = LCase(Mid(kokovar(k), 1, Len(kokovar(k)) - Len(wildext) - 1))
            
       If InStr(1, wildext, "*") = 0 Then
       wildexts = "888NONE888"
       wildexte = "888NONE888"
       Else
       wildexts = Mid(wildext, 1, InStr(1, wildext, "*") - 1)
       wildexte = Mid(wildext, InStr(1, wildext, "*") + 1, Len(wildext) - InStr(1, wildext, "*"))
       End If
       
        If InStr(1, wildfirst, "*") = 0 Then
       wildfirsts = "888NONE888"
       wildfirste = "888NONE888"
       Else
       wildfirsts = Mid(wildfirst, 1, InStr(1, wildfirst, "*") - 1)
       wildfirste = Mid(wildfirst, InStr(1, wildfirst, "*") + 1, Len(wildfirst) - InStr(1, wildfirst, "*"))
       End If
            
        examfirst = LCase(cutgetName(cutfilename(CStr(objfile))))
        examext = LCase(cutgetExtension(CStr(objfile)))
            
        If wildexts = "888NONE888" Then
'we do not have a wildcard in the extension
        If wildfirsts = "888NONE888" Then
        'we do not have a wildcard neither on the beggining or the

'extension
        If examfirst = wildfirst And examext = wildext Then
        afl_filetext = afl_filetext + objfile + vbNewLine
        End If
                
        Else
        
        'we do have a wildcard in the beggining but not in
'the extension
        If Mid(examfirst, 1, Len(wildfirsts)) = wildfirsts And _
        Mid(examfirst, Len(wildfirst) - Len(wildfirste) + 1, Len(wildfirste)) = wildfirste And wildext = examext Then
        afl_filetext = afl_filetext + objfile + vbNewLine
        End If
        
        End If
        
        Else
        'we do not have a wildcard in the extension
        If wildfirsts = "888NONE888" Then
        'we do have a wildcard in the beggining but not in the
'extension
        If Mid(examext, 1, Len(wildexts)) = wildexts And _
        Mid(examext, Len(wildext) - Len(wildexte) + 1, Len(wildexte)) = wildexte Then
        afl_filetext = afl_filetext + objfile + vbNewLine
        End If
            
        Else
        'we have a wildcard in both beggining and extension
        
            If Mid(examext, 1, Len(wildexts)) = wildexts And _
        Mid(examext, Len(wildext) - Len(wildexte) + 1, Len(wildexte)) = wildexte _
         And Mid(examfirst, 1, Len(wildfirsts)) = wildfirsts And _
        Mid(examfirst, Len(wildfirst) - Len(wildfirste) + 1, Len(wildfirste)) = wildfirste Then
        afl_filetext = afl_filetext + objfile + vbNewLine
        End If
            
            End If
            
            End If
            
            'telos if
                    
            Next k
            
        Next

If getsubdirs Then

        For Each objfolder In fso.GetFolder(tfolder).SubFolders
           afl_filetext = afl_filetext & Get_All_Directory_Files_With_Wildcard_ORIGINAL(CStr(objfolder), getsubdirs, wildcard)
        Next
       
    End If
End If

Set fso = Nothing
Get_All_Directory_Files_With_Wildcard_ORIGINAL = afl_filetext
End Function




Public Function cutfilename(ByVal fname As String) As String
Dim spos As Integer
Dim ffn As String
spos = InStrRev(fname, "\")
ffn = Mid(fname, spos + 1, Len(fname) - spos)
cutfilename = ffn

End Function

Public Function cutgetExtension(ByVal fname As String)
Dim spos As Integer
Dim koko As String

spos = InStrRev(fname, ".")
If spos <> 0 Then
koko = Mid(fname, spos + 1, Len(fname) - spos)
End If

cutgetExtension = koko

End Function


Public Function cutgetName(ByVal fname As String)
Dim spos As Integer
Dim koko As String

spos = InStrRev(fname, ".")
If spos <> 0 Then
koko = Mid(fname, 1, spos - 1)
End If
cutgetName = koko

End Function






