'#==============================================================================
'#==============================================================================
'#  SCRIPT.........:  setTime.vbs
'#  AUTHOR.........:  UNKNOWN
'#  VERSION........:  1.5
'#  DATE...........:  02-21-06
'#  REVISED BY.....:  Joe Glessner
'#  REVISION DATE..:  03-13-07
'#  REQUIREMENTS...:  Windows OS with at least IE5 installed.
'#
'#  DESCRIPTION....:  Synchronizes the system time with an internet time server.
'#
'#  NOTES..........:  
'# 
'#  CUSTOMIZE......:  
'#
'#==============================================================================
'#==============================================================================
'**Start Encode**

' ----
' Sync System Time with Internet Time Server - Script Center - Spiceworks
' https://community.spiceworks.com/scripts/show/467-sync-system-time-with-internet-time-server
' ----

' -------------------------------------------------------------------
' [ Tuesday 15:28:20 Pm_02 April 2019 ]
' -------------------------------------------------------------------
' EASIER FIND AND USE THIS ONE
' ----
' NetTime - Network Time Synchronization Tool
' http://www.timesynctool.com/
' ----
' -------------------------------------------------------------------

'#==============================================================================
'#  START CODE
'#==============================================================================
    '#--------------------------------------------------------------------------
    '#  Declare Variables
    '#--------------------------------------------------------------------------
    Option Explicit
    Dim ws, strTitle
    Set ws = CreateObject("WScript.Shell")
    strTitle = "SetTime.vbs © Joe Glessner"

    '#--------------------------------------------------------------------------
    '#  Check system compatibility.
    '#--------------------------------------------------------------------------
    Dim http
    Call ChkCompat

    '#--------------------------------------------------------------------------
    '#  Read time zone offset hex value from Registry.
    '#--------------------------------------------------------------------------
    Dim TimeOffset, HexVal
    TimeOffset = ws.RegRead("HKLM\SYSTEM\CurrentControlSet\" & _
                 "Control\TimeZoneInformation\ActiveTimeBias")
                 
    '#--------------------------------------------------------------------------
    '#  Reg value format varies between Win9x and NT
    '#--------------------------------------------------------------------------
    If IsArray(TimeOffset) Then
      'Win9x uses a reversed 4 element array of Hex values.
      HexVal = Hex(TimeOffset(3)) & Hex(TimeOffset(2)) & _
               Hex(TimeOffset(1)) & Hex(TimeOffset(0))
    Else 'Must be a NT system.
      HexVal = Hex(TimeOffset)
    End If

    '#--------------------------------------------------------------------------
    '#  Convert to minutes of time zone offset.
    '#--------------------------------------------------------------------------
    TimeOffset = - CLng("&H" & HexVal)

    '#--------------------------------------------------------------------------
    '#  Get time from server.  Recheck up to 5 times if lagged.
    '#--------------------------------------------------------------------------
    Dim n, timechk, localdate, lag, gmttime
    Dim timeserv

    '#--------------------------------------------------------------------------
    '#  Check primary server, US Naval Observatory
    '#--------------------------------------------------------------------------
    timeserv = "http://tycho.usno.navy.mil/cgi-bin/timer.pl"& now()
    http.open "GET",timeserv,false
    On Error Resume Next
    http.send
    If Err Then
      'Use backup server, National Institute of Standards and Technology (NIST)
      timeserv = "http://www.nist.gov/"
      err.Clear
    End If
    On Error GoTo 0

    For n = 0 to 4
      http.open "GET",timeserv,false
      'Check response time to avoid invalid errors.
      timechk = Now
      On Error Resume Next
      http.send
      If Err Then
        If Err =  -2146697211 Then
          MsgBox "Both Time Servers unavailable!"
        Else
          MsgBox "Unknown Error occurred, " & Err
        End If
        Wscript.Quit
      End If
      On Error GoTo 0
      localdate = Now
      lag = DateDiff("s", timechk, localdate)

      'Key concept for script is reading header date.
      gmttime = http.getResponseHeader("Date")

      'Trim results to valid date format.
      gmttime = right(gmttime, len(gmttime) - 5)
      gmttime = left(gmttime, len(gmttime) - 3)

      'If less than 2 seconds lag we can use the results.
      If lag < 2 Then Exit For
    Next

    '#--------------------------------------------------------------------------
    '#  If still too much lag after 5 attemps, quit.
    '#--------------------------------------------------------------------------
    If n = 4 then
      ws.Popup "Unable to establish a reliable connection " & _
               "with time server.  This could be due to the " & _
               "time server being too busy, your connection " & _
               "already in use, or a poor connection." & vbcrlf & _
               vbcrlf & "Please try again later.", 5, strTitle
      Cleanup
    End If

    '#--------------------------------------------------------------------------
    '#  Time and date error calculations.
    '#--------------------------------------------------------------------------
    Dim remotedate, diff, newnow, newdate, newtime, ddiff, sdiff

    'Add local time zone offset to GMT returned from USNO server.
    remotedate = DateAdd("n", timeoffset, gmttime)

    'Calculate seconds difference betweed remote and local.
    diff = DateDiff("s", localdate, remotedate)

    'Adjust for difference and lag to get actual time.
    newnow = DateAdd("s", diff + lag, now)

    'Split out date and calculate any difference.
    newdate = FormatDateTime(DateValue(newnow))
    ddiff = DateDiff("d", Date, newdate)

    'Split out time.
    newtime = TimeValue(newnow)

    'Convert time to 24 hr format required for OS compatibility.
    newtime = Right(0 & Hour(newtime), 2) & ":" & _
              Right(0 & Minute(newtime), 2) & ":" & _
              Right(0 & Second(newtime), 2)

    'Calculate time difference.
    sdiff = DateDiff("s", time, newtime)

    '#--------------------------------------------------------------------------
    '#  If off by 1 or more seconds, adjust local time
    '#--------------------------------------------------------------------------
    Dim tmsg
    If sdiff < 2 and sdiff > -2 Then
      tmsg = "System is accurate to within " & _
        "1 second.  System time not changed."
    Else
      'Run DOS Time command in hidden window.
      ws.Run "%comspec% /c time " & newtime, 0
      tmsg = "System time off by " & sdiff & _
             " seconds.  System time changed to " & _
             CDate(newtime)
    End If

    'If date off, change it.
    Dim dmsg
    If ddiff <> 0 Then
      ws.Run "%comspec% /c date " & newdate, 0
      dmsg = "Date off by " & ddiff & " days.  System date changed " & _
             "to " & FormatDateTime(newdate,1) & vbcrlf & vbcrlf
    End If

    '#--------------------------------------------------------------------------
    '#  Show the changes
    '#--------------------------------------------------------------------------
    ws.Popup "Time syncronizion using " & timeserv & vbcrlf & _
              vbcrlf & dmsg & tmsg, 5, strTitle, 4096
    Call Cleanup
    
'#==============================================================================
'#  Subs
'#==============================================================================

    '#--------------------------------------------------------------------------
    '#  Check compatability
    '#--------------------------------------------------------------------------
    Sub ChkCompat
      On Error Resume Next
      Set http = CreateObject("microsoft.xmlhttp")
      If Err.Number <> 0 Then
        ws.Popup "Process Aborted!" & vbcrlf & vbcrlf & _
                 "Minimum system requirements to run this " & _
                 "script are Windows 95 or Windows NT 4.0 " & _
                 "with Internet Explorer 5.", , strTitle
        Cleanup
      End If
    End Sub
    
    '#--------------------------------------------------------------------------
    '#  Cleanup
    '#--------------------------------------------------------------------------
    Sub Cleanup
      Set ws = Nothing
      Set http = Nothing
      WScript.Quit
    End Sub

'#==============================================================================
'#  End CODE
'#================