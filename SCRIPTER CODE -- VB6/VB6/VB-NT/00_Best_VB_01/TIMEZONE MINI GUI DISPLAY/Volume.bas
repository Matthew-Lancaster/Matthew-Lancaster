Attribute VB_Name = "Volume"
Const MMSYSERR_NOERROR = 0
Const MAXPNAMELEN = 32
Const MIXER_LONG_NAME_CHARS = 64
Const MIXER_SHORT_NAME_CHARS = 16
Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&
Const MIXER_SETCONTROLDETAILSF_VALUE = &H0&
Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = &H4
Const MIXERCONTROL_CONTROLTYPE_VOLUME = &H50030001

Private Declare Function mixerOpen Lib "winmm.dll" (phmx As Long, _
    ByVal uMxId As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, _
    ByVal fdwOpen As Long) As Long
Private Declare Function mixerGetLineInfo Lib "winmm.dll" Alias _
    "mixerGetLineInfoA" (ByVal hmxobj As Long, pmxl As MIXERLINE, _
    ByVal fdwInfo As Long) As Long
Private Declare Function mixerGetLineControls Lib "winmm.dll" Alias _
    "mixerGetLineControlsA" (ByVal hmxobj As Long, pmxlc As MIXERLINECONTROLS, _
    ByVal fdwControls As Long) As Long
Private Declare Function mixerSetControlDetails Lib "winmm.dll" (ByVal hmxobj _
    As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Private Declare Function mixerClose Lib "winmm.dll" (ByVal hmx As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
    ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hmem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hmem As Long) As Long

Private Type MIXERCONTROL
    cbStruct As Long
    dwControlID As Long
    dwControlType As Long
    fdwControl As Long
    cMultipleItems As Long
    szShortName As String * MIXER_SHORT_NAME_CHARS
    szName As String * MIXER_LONG_NAME_CHARS
    lMinimum As Long
    lMaximum As Long
    reserved(10) As Long
End Type

Private Type MIXERCONTROLDETAILS
    cbStruct As Long
    dwControlID As Long
    cChannels As Long
    item As Long
    cbDetails As Long
    paDetails As Long
End Type

Private Type MIXERCONTROLDETAILS_UNSIGNED
    dwValue As Long
End Type

Private Type MIXERLINE
    cbStruct As Long
    dwDestination As Long
    dwSource As Long
    dwLineID As Long
    fdwLine As Long
    dwUser As Long
    dwComponentType As Long
    cChannels As Long
    cConnections As Long
    cControls As Long
    szShortName As String * MIXER_SHORT_NAME_CHARS
    szName As String * MIXER_LONG_NAME_CHARS
    dwType As Long
    dwDeviceID As Long
    wMid  As Integer
    wPid As Integer
    vDriverVersion As Long
    szPname As String * MAXPNAMELEN
End Type

Private Type MIXERLINECONTROLS
    cbStruct As Long
    dwLineID As Long
    dwControl As Long
    cControls As Long
    cbmxctrl As Long
    pamxctrl As Long
End Type

' Set the master volume level.
'
' VolumeLevel is the level value in percentage (0 = min, 100 = max)
' Returns True if successful

Public Function SetVolume(VolumeLevel As Long) As Boolean
    Dim hmx As Long
    Dim uMixerLine As MIXERLINE
    Dim uMixerControl As MIXERCONTROL
    Dim uMixerLineControls As MIXERLINECONTROLS
    Dim uDetails As MIXERCONTROLDETAILS
    Dim uUnsigned As MIXERCONTROLDETAILS_UNSIGNED
    Dim RetValue As Long
    Dim hmem As Long
    'uUnsigned.dwValue
    'uDetails.cbDetails
    'uMixerControl.szName
    'A = uMixerControl.cbStruct
    'a = uMixerLine.
    'A = hmem
    
    
    
    
    ' VolumeLevel value must be between 0 and 100
    If VolumeLevel < 0 Or VolumeLevel > 100 Then GoTo error
   
   
   
    ' Open the mixer
    RetValue = mixerOpen(hmx, 0, 0, 0, 0)
    If RetValue <> MMSYSERR_NOERROR Then GoTo error
    
    ' Initialize MIXERLINE structure and call mixerGetLineInfo
    uMixerLine.cbStruct = Len(uMixerLine)
    uMixerLine.dwComponentType = MIXERLINE_COMPONENTTYPE_DST_SPEAKERS
    RetValue = mixerGetLineInfo(hmx, uMixerLine, _
        MIXER_GETLINEINFOF_COMPONENTTYPE)
    If RetValue <> MMSYSERR_NOERROR Then GoTo error
    
    ' Initialize MIXERLINECONTROLS strucure and
    ' call mixerGetLineControls
    uMixerLineControls.cbStruct = Len(uMixerLineControls)
    uMixerLineControls.dwLineID = uMixerLine.dwLineID
    uMixerLineControls.dwControl = MIXERCONTROL_CONTROLTYPE_VOLUME
    uMixerLineControls.cControls = 1
    uMixerLineControls.cbmxctrl = Len(uMixerControl)
    
    ' Allocate a buffer to receive the properties of the master volume control
    ' and put his address into uMixerLineControls.pamxctrl
    hmem = GlobalAlloc(&H40, Len(uMixerControl))
    uMixerLineControls.pamxctrl = GlobalLock(hmem)
    uMixerControl.cbStruct = Len(uMixerControl)
    RetValue = mixerGetLineControls(hmx, uMixerLineControls, _
        MIXER_GETLINECONTROLSF_ONEBYTYPE)
    If RetValue <> MMSYSERR_NOERROR Then GoTo error
           
    ' Copy data buffer into the uMixerControl structure
    CopyMemory uMixerControl, ByVal uMixerLineControls.pamxctrl, _
        Len(uMixerControl)
    GlobalFree hmem
    hmem = 0

    uDetails.item = 0
    uDetails.dwControlID = uMixerControl.dwControlID
    'uDetails.paDetails
    'uMixerLineControls.cbmxctrl
    uDetails.cbStruct = Len(uDetails)
    uDetails.cbDetails = Len(uUnsigned)
    
    ' Allocate a buffer in which properties for the volume control are set
    ' and put his address into uDetails.paDetails
    hmem = GlobalAlloc(&H40, Len(uUnsigned))
    uDetails.paDetails = GlobalLock(hmem)
    uDetails.cChannels = 1
    uUnsigned.dwValue = CLng((VolumeLevel * uMixerControl.lMaximum) / 100)
    CopyMemory ByVal uDetails.paDetails, uUnsigned, Len(uUnsigned)
   
    ' Set new volume level
    RetValue = mixerSetControlDetails(hmx, uDetails, _
        MIXER_SETCONTROLDETAILSF_VALUE)
    GlobalFree hmem
    hmem = 0
    If RetValue <> MMSYSERR_NOERROR Then GoTo error
    
    mixerClose hmx
    ' signal success
    SetVolume = True
    Exit Function
    
error:
    ' An error occurred
    
    ' Release resources
    If hmx <> 0 Then mixerClose hmx
    If hmem Then GlobalFree hmem
    ' signal failure
    SetVolume = False

End Function


