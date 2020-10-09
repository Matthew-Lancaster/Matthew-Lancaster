






' ActiveXperts » Administration » Scripts » WMI » VBScript
' WIN32_LOGICALDISK - WMI SAMPLE IN VBSCRIPT
' The foundations for Manageability in Windows 7/2008/Vista/XP/2000 and Millennium Edition/'98 are Windows Management Instrumentation (WMI; formerly WBEM) and WMI extensions for Windows Driver Model (WDM).

' ActiveXperts Network Monitor provides the ability to build monitor check routines based on WMI. ActiveXperts has collected more than a hundred WMI samples. You can use these samples as a base for new check routines you can write yourself.

' On this site, you can find many WMI samples.

' The Win32_LogicalDisk WMI class can be used in ActiveXperts Network Monitor to monitor your servers.

' Win32_LogicalDisk
' Description
' The Win32_LogicalDisk WMI class represents a data source that resolves to an actual local storage device on a Windows system.
' Example(s)
On Error Resume Next
STRCOMPUTER = "."
DRIVE_NAME="E:"
Set objWMIService = GetObject("winmgmts:\\" & STRCOMPUTER & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_LogicalDisk where DeviceID='"+DRIVE_NAME+"'",,48)
For Each objItem in colItems
	I=""
    I=I+  "Access:  ---- " & objItem.Access & VBCR
    I=I+  "Availability:  ---- " & objItem.Availability & VBCR
    I=I+  "BlockSize:  ---- " & objItem.BlockSize & VBCR
    I=I+  "Caption:  ---- " & objItem.Caption & VBCR
    I=I+  "Compressed:  ---- " & objItem.Compressed & VBCR
    I=I+  "ConfigManagerErrorCode:  ---- " & objItem.ConfigManagerErrorCode & VBCR
    I=I+  "ConfigManagerUserConfig:  ---- " & objItem.ConfigManagerUserConfig & VBCR
    I=I+  "CreationClassName:  ---- " & objItem.CreationClassName & VBCR
    I=I+  "Description:  ---- " & objItem.Description & VBCR
    I=I+  "DeviceID:  ---- " & objItem.DeviceID & VBCR
    I=I+  "DriveType:  ---- " & objItem.DriveType & VBCR
    I=I+  "ErrorCleared:  ---- " & objItem.ErrorCleared & VBCR
    I=I+  "ErrorDescription:  ---- " & objItem.ErrorDescription & VBCR
    I=I+  "ErrorMethodology:  ---- " & objItem.ErrorMethodology & VBCR
    I=I+  "FileSystem:  ---- " & objItem.FileSystem & VBCR
    I=I+  "FreeSpace:  ---- " & objItem.FreeSpace & VBCR
    I=I+  "InstallDate:  ---- " & objItem.InstallDate & VBCR
    I=I+  "LastErrorCode:  ---- " & objItem.LastErrorCode & VBCR
    I=I+  "MaximumComponentLength:  ---- " & objItem.MaximumComponentLength & VBCR
    I=I+  "MediaType:  ---- " & objItem.MediaType & VBCR
    I=I+  "Name:  ---- " & objItem.Name & VBCR
    I=I+  "NumberOfBlocks:  ---- " & objItem.NumberOfBlocks & VBCR
    I=I+  "PNPDeviceID:  ---- " & objItem.PNPDeviceID & VBCR
    I=I+  "PowerManagementCapabilities:  ---- " & objItem.PowerManagementCapabilities & VBCR
    I=I+  "PowerManagementSupported:  ---- " & objItem.PowerManagementSupported & VBCR
    I=I+  "ProviderName:  ---- " & objItem.ProviderName & VBCR
    I=I+  "Purpose:  ---- " & objItem.Purpose & VBCR
    I=I+  "Size:  ---- " & objItem.Size & VBCR
    I=I+  "Status:  ---- " & objItem.Status & VBCR
    I=I+  "StatusInfo:  ---- " & objItem.StatusInfo & VBCR
    I=I+  "SupportsFileBasedCompression:  ---- " & objItem.SupportsFileBasedCompression & VBCR
    I=I+  "SystemCreationClassName:  ---- " & objItem.SystemCreationClassName & VBCR
    I=I+  "SystemName:  ---- " & objItem.SystemName & VBCR
    I=I+  "VolumeName:  ---- " & objItem.VolumeName & VBCR
    I=I+  "VolumeSerialNumber:  ---- " & objItem.VolumeSerialNumber

NEXT

MSGBOX I

' WSCRIPT.QUIT

Set wmi = GetObject("winmgmts://./root/cimv2")
' qry = "SELECT * FROM Win32_Process WHERE CommandLine LIKE '% ---- " & DRIVENAME & "%'"
' QRY = "SELECT DriveLetter FROM Win32_Volume WHERE Label=' ---- " & volumeName & "'"
STRCOMPUTER = "."
DRIVE_NAME="D:"
Set objWMIService = GetObject("winmgmts:\\" & STRCOMPUTER & "\root\cimv2")
' %" & filename & "%'"
' QRY = "Select * from Win32_LogicalDisk where DeviceID='"+DRIVE_NAME+"'"
' QRY = "Select * from Win32_LogicalDisk where DeviceID='%" & DRIVE_NAME & "%'"
' Set colItems = objWMIService.ExecQuery(QRY,,48)
Set colItems = objWMIService.ExecQuery("Select * from Win32_LogicalDisk where DeviceID='"+DRIVE_NAME+"'",,48)

' For Each p In wmi.ExecQuery(qry)
For Each objItem in colItems
	I=I+  "Media file in use."
	WScript.Quit 0
Next

MSGBOX I


