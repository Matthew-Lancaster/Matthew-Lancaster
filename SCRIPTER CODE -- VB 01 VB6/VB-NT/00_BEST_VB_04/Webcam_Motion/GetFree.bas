Attribute VB_Name = "mGetFree"
Option Explicit
Option Private Module
'****************************************************************
'*  VB file:   GetFree.bas... Portable Disk Space Functions for
'*                              Win9x, WinNT (Fat16, Fat32)
'*
'* created 8/5/98 by Ray Mercer <raymer@macnica.co.jp>
'*
'* modified 8/25/98 by Ray Mercer
'*
'* modified 1/07/99 By Ray Mercer
'*  -changed vbGetAvailableBytesEx
'*  -changed vbGetTotalBytesEx
'*  (now this sample actually works- duh!)
'*
'* modified 1/08/99 by Ray Mercer
'*  -changed some Public functions to Private
'*  -made all path arguments in Public functions Optional
'*  -added Option Private Module for use in ActiveX Dlls
'*  -changed vbNullString arguments to "" to avoid VB component bugs
'*  -added vbGetPercentAvailable() method
'*
'* modified 2/24/99 by Ray Mercer
'*  -fixed routines to correctly handle drives with 0 bytes free
'*
'* Copyright (c) 1998-1999 by Ray Mercer.  All rights reserved.
'****************************************************************
'
'//PRIVATE DECLARES SECTION (Not callable outside of this module)
'////////////////////////////////////////////////////////////////
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" _
                            (ByVal lpLibFileName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32" _
                            (ByVal hLibModule As Long) As Long

Private Declare Function GetProcAddress Lib "kernel32" _
                            (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" _
                            (ByVal lpRootPathName As String, _
                            lpSectorsPerCluster As Long, _
                            lpBytesPerSector As Long, _
                            lpNumberOfFreeClusters As Long, _
                            lpTtoalNumberOfClusters As Long) As Long 'C Bool

Private Declare Function GetDiskFreeSpaceExAsCurrency Lib "kernel32" Alias "GetDiskFreeSpaceExA" _
                            (ByVal lpDirectoryName As String, _
                             lpFreeBytesAvailableToCaller As Currency, _
                             lpTotalNumberOfBytes As Currency, _
                             lpTotalNumberOfFreeBytes As Currency) As Long 'C Bool

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
                                                    (lpvDest As Any, _
                                                    lpvSource As Any, _
                                                    ByVal cbCopy As Long)


'****************************************************************
'* FUNCTION vbGetAvailableBytesAsString()
'* ===============
'* This routine will return a VB string containing the number of
'* free bytes on the drive pointed to by sPath.  This function will
'* correctly call either GetDiskFreeSpace() or GetDiskFreeSpaceEx()
'* from the Win32 API as appropriate
'*
'* INPUTS:
'*  (Optional) sPath
'* -notes from MSDN documentation-
'* Pointer to a null-terminated string that specifies the root directory
'* of the disk to return information about. If lpRootPathName is omitted,
'* the function uses the root of the current directory. If this parameter
'* is a UNC name, you must follow it with an additional backslash. For
'* example, you would specify \\MyServer\MyShare as \\MyServer\MyShare\.
'* Windows 95: The initial release of Windows 95 does not support UNC paths
'* for the lpszRootPathName parameter. To query the free disk space using a
'* UNC path, temporarily map the UNC path to a drive letter, query the free
'* disk space on the drive, then remove the temporary mapping. Windows 95
'* OSR2 and later: UNC paths are supported.
'*
'* RETURNS:
'*  on error returns vbNullChar ("")
'*  otherwise returns number of free (available) bytes as string
'****************************************************************
Public Function vbGetAvailableBytesAsString(Optional ByVal sPath As String = "") As String
    Dim lo As Long, hi As Long
    Dim sOut As String
    
    If ExistGetDiskFreeSpaceEx() Then
        sOut = vbGetAvailableBytesEx(sPath)

    Else
        sOut = CStr(vbGetAvailableBytes(sPath))
    End If
    vbGetAvailableBytesAsString = sOut
    
End Function

'****************************************************************
'* FUNCTION vbGetAvailableKBytesAsString()
'* ===============
'* This routine will return a VB string containing the number of
'* free kilobytes on the drive pointed to by sPath.  This function will
'* correctly call either GetDiskFreeSpace() or GetDiskFreeSpaceEx()
'* from the Win32 API as appropriate
'*
'* INPUTS:
'*  (Optional) sPath (see notes in vbGetAvailableBytesAsString function)
'*
'* RETURNS:
'*  on error returns vbNullChar ("")
'*  otherwise returns number of free (available) kilobytes as string
'*  (rounded up to nearest kbyte)
'****************************************************************
Public Function vbGetAvailableKBytesAsString(Optional ByVal sPath As String = "") As String
    Dim bytes As Currency, kBytes As Currency
    Dim sTmp As String
    
    sTmp = vbGetAvailableBytesAsString(sPath)
    bytes = CCur(sTmp)
    If bytes Then 'avoid divide by 0 errors
        kBytes = bytes / 1024
        kBytes = Fix(kBytes)
    Else
        kBytes = 0
    End If
    vbGetAvailableKBytesAsString = CStr(kBytes)
        
End Function

'****************************************************************
'* FUNCTION vbGetAvailableMBytesAsString()
'* ===============
'* This routine will return a VB string containing the number of
'* free megabytes on the drive pointed to by sPath.  This function will
'* correctly call either GetDiskFreeSpace() or GetDiskFreeSpaceEx()
'* from the Win32 API as appropriate
'*
'* INPUTS:
'*  (Optional) sPath (see notes in vbGetAvailableBytesAsString function)
'*
'* RETURNS:
'*  on error returns vbNullChar ("")
'*  otherwise returns number of free (available) megabytes as string
'*  (rounded up to nearest MB)
'****************************************************************
Public Function vbGetAvailableMBytesAsString(Optional ByVal sPath As String = "") As String
    Dim kBytes As Currency, mBytes As Currency
    Dim sTmp As String
    
    sTmp = vbGetAvailableKBytesAsString(sPath)
    kBytes = CCur(sTmp)
    If kBytes Then 'avoid divide by 0 errors
        mBytes = kBytes / 1024
        mBytes = Fix(mBytes)
    Else
        mBytes = 0
    End If
    vbGetAvailableMBytesAsString = CStr(mBytes)
        
End Function

'****************************************************************
'* FUNCTION vbGetTotalBytesAsString()
'* ===============
'* This routine will return a VB string containing the total number
'* of bytes on the drive pointed to by sPath.  This function will
'* correctly call either GetDiskFreeSpace() or GetDiskFreeSpaceEx()
'* from the Win32 API as appropriate
'*
'* INPUTS:
'*  (Optional) sPath
'* -notes from MSDN documentation-
'* Pointer to a null-terminated string that specifies the root directory
'* of the disk to return information about. If lpRootPathName is omitted,
'* the function uses the root of the current directory. If this parameter
'* is a UNC name, you must follow it with an additional backslash. For
'* example, you would specify \\MyServer\MyShare as \\MyServer\MyShare\.
'* Windows 95: The initial release of Windows 95 does not support UNC paths
'* for the lpszRootPathName parameter. To query the free disk space using a
'* UNC path, temporarily map the UNC path to a drive letter, query the free
'* disk space on the drive, then remove the temporary mapping. Windows 95
'* OSR2 and later: UNC paths are supported.
'*
'* RETURNS:
'*  on error returns vbNullChar ("")
'*  otherwise returns total number of bytes as a string
'****************************************************************
Public Function vbGetTotalBytesAsString(Optional ByVal sPath As String = "") As String
    Dim lo As Long, hi As Long
    Dim sOut As String
    
    If ExistGetDiskFreeSpaceEx() Then
        sOut = vbGetTotalBytesEx(sPath)
    Else
        sOut = CStr(vbGetTotalBytes(sPath))
    End If
    vbGetTotalBytesAsString = sOut
    
End Function

'****************************************************************
'* FUNCTION vbGetTotalKBytesAsString()
'* ===============
'* This routine will return a VB string containing the total number of
'* kilobytes on the drive pointed to by sPath.  This function will
'* correctly call either GetDiskFreeSpace() or GetDiskFreeSpaceEx()
'* from the Win32 API as appropriate
'*
'* INPUTS:
'*  (Optional) sPath (see notes in vbGetAvailableBytesAsString function)
'*
'* RETURNS:
'*  on error returns vbNullChar ("")
'*  otherwise returns total number of kilobytes as string (rounded up to
'*  nearest kbyte)
'****************************************************************
Public Function vbGetTotalKBytesAsString(Optional ByVal sPath As String = "") As String
    Dim numbytes As Currency, kBytes As Currency
    Dim sTmp As String
    
    sTmp = vbGetTotalBytesAsString(sPath)
    numbytes = CCur(sTmp)
    If numbytes Then 'avoid divide by 0 errors
        kBytes = numbytes / 1024
        kBytes = Fix(kBytes)
    Else
        kBytes = 0
    End If
    vbGetTotalKBytesAsString = CStr(kBytes)
        
End Function

'****************************************************************
'* FUNCTION vbGetTotalMBytesAsString()
'* ===============
'* This routine will return a VB string containing the total number of
'* megabytes on the drive pointed to by sPath.  This function will
'* correctly call either GetDiskFreeSpace() or GetDiskFreeSpaceEx()
'* from the Win32 API as appropriate
'*
'* INPUTS:
'*  (Optional) sPath (see notes in vbGetAvailableBytesAsString function)
'*
'* RETURNS:
'*
'* String-
'*  on error returns vbNullChar ("")
'*  otherwise returns total number of megabytes as string (rounded up
'*  to nearest MB)
'****************************************************************
Public Function vbGetTotalMBytesAsString(Optional ByVal sPath As String = "") As String
    Dim kBytes As Currency, mBytes As Currency
    Dim sTmp As String
    
    sTmp = vbGetTotalKBytesAsString(sPath)
    kBytes = CCur(sTmp)
    If kBytes Then 'avoid divide by 0 errors
        mBytes = kBytes / 1024
        mBytes = Fix(mBytes)
    Else
        mBytes = 0
    End If
    vbGetTotalMBytesAsString = CStr(mBytes)
        
End Function

'****************************************************************
'* FUNCTION ExistGetDiskFreeSpaceEx()
'* ===============
'* This routine used the Microsoft-recommended way to determine if
'* the Win32 API function GetDiskFreeSpaceEx() exists on the current
'* OS platform. (should be available on all Win32 systems after OSr.2)
'*
'* INPUTS: none
'*
'* RETURNS:
'*  TRUE - if the GetDiskFreeSpaceEx() function is available
'*  FALSE - if the GetDiskFreeSpaceEx() function is available
'*          in this case you should call the older GetDiskFreeSpace()
'****************************************************************
Public Function ExistGetDiskFreeSpaceEx() As Boolean
    Dim hInst As Long
    Dim procAddress As Long
    
    hInst = LoadLibrary("kernel32.dll")
    If hInst Then
        procAddress = GetProcAddress(hInst, "GetDiskFreeSpaceExA")
        Call FreeLibrary(hInst)
    End If
    ExistGetDiskFreeSpaceEx = CBool(procAddress)
    
End Function


'****************************************************************
'* FUNCTION vbGetAvailableBytesEx()
'* ===============
'* This routine will return a String containing the the available
'* bytes as reported by the GetDiskFreeSpaceEX() API
'*
'* This function will correctly return values for large disk partitions (i.e., Fat32)
'*
'* INPUTS:
'* sPath - (see notes in vbGetAvailableBytesAsString function)
'*
'* RETURNS:
'*
'* String - Available bytes on disk pointed to by sPath
'****************************************************************
Private Function vbGetAvailableBytesEx(ByVal sPath As String) As String
    Dim BytesAvailable As Currency
    Dim TotalBytes As Currency
    Dim TotalFreeBytes As Currency
    Dim tmp As Currency

    On Error GoTo APIfailed
    If "" = sPath Then
        Call GetDiskFreeSpaceExAsCurrency(vbNullString, BytesAvailable, TotalBytes, TotalFreeBytes)
    Else
        Call GetDiskFreeSpaceExAsCurrency(sPath, BytesAvailable, TotalBytes, TotalFreeBytes)
    End If

    'If BytesAvailable Then
        BytesAvailable = BytesAvailable * 10000
        vbGetAvailableBytesEx = CStr(BytesAvailable)
    'End If
    Exit Function
APIfailed:
    'returns false
    Debug.Print "GetDiskFreeSpaceEx() API Failed!"
End Function

'****************************************************************
'* FUNCTION vbGetTotalBytesEx()
'* ===============
'* This routine will return a String containing the the available
'* bytes as reported by the GetDiskFreeSpaceEX() API
'*
'* Before calling this function you should call ExistGetDiskFreeSpaceEx()
'* to see whether it will work on the current OS platform.  This function
'* will correctly return values for large disk partitions (i.e., Fat32)
'*
'* INPUTS:
'* sPath - (see notes in vbGetAvailableBytesAsString function)
'*
'* RETURNS:
'*
'* String - Available bytes on disk pointed to by sPath
'****************************************************************
Private Function vbGetTotalBytesEx(ByVal sPath As String) As String
    Dim BytesAvailable As Currency
    Dim TotalBytes As Currency
    Dim TotalFreeBytes As Currency
    
    On Error GoTo APIfailed
    If "" = sPath Then
        Call GetDiskFreeSpaceExAsCurrency(vbNullString, BytesAvailable, TotalBytes, TotalFreeBytes)
    Else
        Call GetDiskFreeSpaceExAsCurrency(sPath, BytesAvailable, TotalBytes, TotalFreeBytes)
    End If
    
    If TotalBytes Then
        TotalBytes = TotalBytes * 10000
    Else
        TotalBytes = 0
    End If
    vbGetTotalBytesEx = CStr(TotalBytes)
    Exit Function
APIfailed:
    'returns false
    Debug.Print "GetDiskFreeSpaceEx() API Failed!"
End Function
'****************************************************************
'* FUNCTION vbGetAvailableBytes()
'* ===============
'* This routine will return the number of free bytes on the
'* specified drive (does not handle drive partitions over
'* 2GB)
'*
'* INPUTS:
'* sPath - (see notes in vbGetAvailableBytesAsString function)
'*
'* RETURNS:
'* Long - free disk space in bytes
'****************************************************************
Private Function vbGetAvailableBytes(ByVal sPath As String) As Long
    Dim lSpc As Long 'sectors per cluster
    Dim lBps As Long 'bytes per sector
    Dim lNfc As Long 'number of free clusters
    Dim lTnc As Long 'total number of clusters
    
    
    Call GetDiskFreeSpace(sPath, lSpc, lBps, lNfc, lTnc)
    vbGetAvailableBytes = lSpc * lBps * lNfc
    
End Function

'****************************************************************
'* FUNCTION vbGetTotalBytes()
'* ===============
'* This routine will return the total number of bytes on the
'* specified drive (does not handle drive partitions over
'* 2GB)
'*
'* INPUTS:
'* sPath - (see notes in vbGetAvailableBytesAsString function)
'*
'* RETURNS:
'* Long - total disk space in bytes
'****************************************************************
Private Function vbGetTotalBytes(ByVal sPath As String) As Long
    Dim lSpc As Long 'sectors per cluster
    Dim lBps As Long 'bytes per sector
    Dim lNfc As Long 'number of free clusters
    Dim lTnc As Long 'total number of clusters
    
    
    Call GetDiskFreeSpace(sPath, lSpc, lBps, lNfc, lTnc)
    vbGetTotalBytes = lSpc * lBps * lTnc
    
End Function

'****************************************************************
'* FUNCTION vbGetPercentAvailable()
'* ===============
'* This routine will return the percentage of disk space available
'* it will work transparently across Fat16 & Fat32 volumes of any size
'* (not yet tested on NTFS volumes)
'*
'* INPUTS:
'* sPath - (see notes in vbGetAvailableBytesAsString function)
'*
'* RETURNS:
'* Long - percent free on drive
'****************************************************************

Public Function vbGetPercentAvailable(Optional ByVal sPath As String = "") As Long
Dim freeEX As Currency
Dim totalEX As Currency
Dim availEX As Currency
Dim percent As Long

On Error Resume Next 'if API fails there will be divide by zero errors

If ExistGetDiskFreeSpaceEx() Then
    If "" = sPath Then
        Call GetDiskFreeSpaceExAsCurrency(vbNullString, availEX, totalEX, freeEX)
    Else
        Call GetDiskFreeSpaceExAsCurrency(sPath, availEX, totalEX, freeEX)
    End If
Else
   totalEX = vbGetTotalBytes(sPath)
   availEX = vbGetAvailableBytes(sPath)
    
End If

totalEX = totalEX * 10000
availEX = availEX * 10000
percent = (availEX * 100) / totalEX

vbGetPercentAvailable = percent

End Function
