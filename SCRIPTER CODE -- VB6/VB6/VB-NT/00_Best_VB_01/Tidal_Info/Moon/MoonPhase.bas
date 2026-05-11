Attribute VB_Name = "MoonPhase"
Option Explicit
Dim ShiftPressed_True_Debug_Helper_SORT_MOON_CALIBRATE_PROPER
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Public MoonPhaseDate

Private Const TIME_ZONE_ID_INVALID = -1
Private Const TIME_ZONE_ID_UNKNOWN = 0
Private Const TIME_ZONE_ID_STANDARD = 1
Private Const TIME_ZONE_ID_DAYLIGHT = 2

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
    End Type

Public Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName As String * 64
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName As String * 64
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

Private Const A_SECOND = 0.00001158 ' one second as a fraction of a day
Private Const LPERIOD = 29.530589 ' average days between lunations
Private Const EPOCH = 8388.51399305556 ' days from 01/01/1900 til 12/18/1922 12:20:09 UT, lunation 0
Private Const pi = 3.14159265359

Private JDE As Double ' Julian Day Ephemeris of phase event
Private E As Double ' Eccentricity anomaly
Private M As Double ' Sun's mean anomaly
Private M1 As Double ' Moon's mean anomaly
Private F As Double ' Moon's argument of latitude
Private O As Double  ' Moon's longitude of ascending node
Private A(15) As Double ' Planetary arguments
Private W As Double ' Added correction for quarter phases

Public Stopping As Boolean
Public TheDate As Date
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

Private Function GetComputerName() As String
   Dim UserName As String * 255
   Dim iky$
   Call GetComputerNameA(UserName, 255)
   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
   iky$ = GetComputerName
   If InStr(iky$, "-") Then
   GetComputerName = Mid$(iky$, 1, InStr(iky$, "-") - 1) + Mid$(iky$, InStr(iky$, "-") + 1)
   End If
End Function

Public Sub DrawMoonPhase(PhaseAngle As Single, MoonBox As PictureBox)
Dim cx As Single
Dim cy As Single
Dim Lat As Single
Dim R As Single
Dim x1 As Single
Dim y1 As Single
Dim x2 As Single
Dim y2 As Single
Dim Sign As Integer
Dim k As Double
k = Atn(1) / 45
MoonBox.ScaleMode = 3
MoonBox.DrawWidth = 2
R = MoonBox.ScaleWidth / 2
cx = R
cy = R
If PhaseAngle >= 360 Then
PhaseAngle = PhaseAngle Mod 360
End If
If PhaseAngle > 180 Then
Sign = 1
PhaseAngle = PhaseAngle - 180
Else
Sign = -1
End If
If PhaseAngle = 180 Then Exit Sub
MoonBox.Cls
For Lat = -90 To 90
x1 = cx + R * Cos(PhaseAngle * k) * Cos(Lat * k)
  y1 = cy - R * Sin(Lat * k)
  x2 = cx + R * Cos(Lat * k) * Sign
  y2 = cy - R * Sin(Lat * k)
MoonBox.Line (x1, y1)-(x2, y2), RGB(12, 12, 12)
Next
End Sub

Public Sub CyclePhases(MoonBox As PictureBox)
Dim PhaseAngle As Integer
Do While Not Stopping
If PhaseAngle = 361 Then PhaseAngle = 0
DoEvents
DrawMoonPhase CSng(PhaseAngle), MoonBox
PhaseAngle = PhaseAngle + 1
Loop
End Sub
Public Function MonthName(ByVal Month As Integer) As String
Select Case Month
Case 1
MonthName = "January"
Case 2
MonthName = "February"
Case 3
MonthName = "March"
Case 4
MonthName = "April"
Case 5
MonthName = "May"
Case 6
MonthName = "June"
Case 7
MonthName = "July"
Case 8
MonthName = "August"
Case 9
MonthName = "September"
Case 10
MonthName = "October"
Case 11
MonthName = "November"
Case 12
MonthName = "December"
End Select
End Function
Public Function MonthNumber(Name As String) As Integer
Select Case Name
Case "January"
MonthNumber = 1
Case "February"
MonthNumber = 2
Case "March"
MonthNumber = 3
Case "April"
MonthNumber = 4
Case "May"
MonthNumber = 5
Case "June"
MonthNumber = 6
Case "July"
MonthNumber = 7
Case "August"
MonthNumber = 8
Case "September"
MonthNumber = 9
Case "October"
MonthNumber = 10
Case "November"
MonthNumber = 11
Case "December"
MonthNumber = 12
End Select
End Function


Public Function TimeZone() As String
Dim TZ As TIME_ZONE_INFORMATION
Dim Bias As Double
Dim temp As Date
Dim temp_str As Date

Const HOUR = 0.04167

'02 of 02 -- SEARCH SAME REM LINE
'------------------------------------------------------------------------------------
'ERROR OF ORGINAL SOURCE -- SUPPOSED TO BE TEMP AS RETURN STRING IN FORM OF --- (GMT)
'OR OPTION HAS THAT IN TEXT STRING AND EXTRA TIME INFO
'BUT TEMP AS DATE -- IS USED THROUGH OUT THIS CODE
'------------------------------------------------------------------------------------
'ALSO MY COMPUTER IS UNABLE TO RETURN INFO ABOUT TIME ZONE -- WIN 10
'------------------------------------------------------------------------------------
'WRAP AROUND SOMETIME RETURN GMT SOMETIME THE VALUE OF temp = CDate(HOUR * Bias)
'------------------------------------------------------------------------------------
'METHOD WORKED OUT
'DON'T DO STRING CONVERT METHOD TO NUMERIC
'AND MEAN ADD ANOTHER FUNCTION HERE
'------------------------------------------------------------------------------------


GetTimeZoneInformation TZ
Bias = TZ.Bias / 60
temp = CDate(HOUR * Bias)
If Bias = 0 Then
TimeZone = "(GMT)"
ElseIf Sgn(Bias) = 1 Then
TimeZone = "(GMT -" & Format(temp, "hh:mm") & ")"
Else
TimeZone = "(GMT +" & Format(temp, "hh:mm") & ")"
End If

End Function



Public Function TimeZone_Numeric() As String

Dim TZ As TIME_ZONE_INFORMATION
Dim Bias As Double
Dim temp As Date
Dim temp_str As Date
Dim TimeZone_Temp_Var As String
Const HOUR = 0.04167

'02 of 02 -- SEARCH SAME REM LINE
'------------------------------------------------------------------------------------
'ERROR OF ORGINAL SOURCE -- SUPPOSED TO BE TEMP AS RETURN STRING IN FORM OF --- (GMT)
'OR OPTION HAS THAT IN TEXT STRING AND EXTRA TIME INFO
'BUT TEMP AS DATE -- IS USED THROUGH OUT THIS CODE
'------------------------------------------------------------------------------------
'ALSO MY COMPUTER IS UNABLE TO RETURN INFO ABOUT TIME ZONE -- WIN 10
'------------------------------------------------------------------------------------
'WRAP AROUND SOMETIME RETURN GMT SOMETIME THE VALUE OF temp = CDate(HOUR * Bias)
'------------------------------------------------------------------------------------
'METHOD WORKED OUT
'DON'T DO STRING CONVERT METHOD TO NUMERIC
'AND MEAN ADD ANOTHER FUNCTION HERE
'------------------------------------------------------------------------------------


GetTimeZoneInformation TZ
Bias = TZ.Bias / 60
temp = CDate(HOUR * Bias)

'---------------  DEBUG TESTER --- temp = TimeValue("01:00:00")

If Bias = 0 Then
'--------------- TimeZone = "(GMT)"
TimeZone_Temp_Var = 0
ElseIf Sgn(Bias) = 1 Then
'--------------- TimeZone = "(GMT -" & Format(temp, "hh:mm") & ")"
TimeZone_Temp_Var = "-"
Else
'--------------- TimeZone = "(GMT +" & Format(temp, "hh:mm") & ")"
TimeZone_Temp_Var = "+"
End If


'--------------------------------------------------------------------------------
'MEHTOD PROBLEM ORIGINAL SOURCE CODE FAULT -- WORK AROUND BUT NOT A FORMATED DATE
'BUT WHEN USED IN A DATE VAR WILL APEEAR
'--------------------------------------------------------------------------------
TimeZone_Numeric = Val(TimeZone_Temp_Var & Format(temp, "hh:mm"))


End Function

Public Function WeekdayName(Weekday As Integer) As String
Select Case Weekday
Case 1
WeekdayName = "Sunday"
Case 2
WeekdayName = "Monday"
Case 3
WeekdayName = "Tuesday"
Case 4
WeekdayName = "Wednesday"
Case 5
WeekdayName = "Thursday"
Case 6
WeekdayName = "Friday"
Case 7
WeekdayName = "Saturday"
End Select
End Function

Public Sub WriteAllPhases()
Dim i As Long
Dim FileNumber As Integer
FileNumber = FreeFile
Open App.Path & "\Phases.txt" For Output As #FileNumber
For i = -22546 To 99898
Print #FileNumber, Format(str(i), "#00000") & " " & _
Format(JulianDaysToUT(MoonPhaseByLunation(i, 0)), "mm/dd/yyyy hh:mm:ss") & "  " & _
Format(JulianDaysToUT(MoonPhaseByLunation(i, 1)), "mm/dd/yyyy hh:mm:ss") & "  " & _
Format(JulianDaysToUT(MoonPhaseByLunation(i, 2)), "mm/dd/yyyy hh:mm:ss") & "  " & _
Format(JulianDaysToUT(MoonPhaseByLunation(i, 3)), "mm/dd/yyyy hh:mm:ss")
Next
Close
End Sub

Public Function Angle(AnyDate As Date) As Single
' AnyDate must already be in UT
' 0 = New, 180 = Full, 360 = New
Angle = Age(AnyDate) * 360 / 29.530589
End Function

Public Function ConvertToUT(Time As Date) As Date
' convert system time to universal time and adjust for DST
Dim TZ As TIME_ZONE_INFORMATION

'01 of 02 -- SEARCH SAME REM LINE
'------------------------------------------------------------------------------------
'ERROR OF ORGINAL SOURCE -- SUPPOSED TO BE TEMP AS RETURN STRING IN FORM OF --- (GMT)
'OR OPTION HAS THAT IN TEXT STRING AND EXTRA TIME INFO
'BUT TEMP AS DATE -- IS USED THROUGH OUT THIS CODE
'------------------------------------------------------------------------------------
'ALSO MY COMPUTER IS UNABLE TO RETURN INFO ABOUT TIME ZONE -- WIN 10
'------------------------------------------------------------------------------------
'WORK AROUND ADD A ISNUMERIC DETECT

'TIMEZONE IS RETURN AS STRING AND NUMERIC IS ADJUST IS MADE AT GMT
'BUT NOT CORRECT TOWARD THE ORGINAL SOURCE CODE METHOD

'------------------------------------------------------------------------------------
'METHOD WORKED OUT
'DON'T DO STRING CONVERT METHOD TO NUMERIC
'AND MEAN ADD ANOTHER FUNCTION HERE
'------------------------------------------------------------------------------------

Dim temp As Date
Dim temp_str As String

Dim Rtn As Long
Rtn = GetTimeZoneInformation(TZ)
If Rtn > TIME_ZONE_ID_UNKNOWN Then

    If Rtn = TIME_ZONE_ID_STANDARD Then
    temp = DateAdd("n", TZ.Bias, Time)
    Else
    temp = DateAdd("n", (TZ.Bias + TZ.DaylightBias), Time)
    End If

End If
ConvertToUT = temp
'Temp = TimeZone
temp = TimeZone_Numeric

'Debug.Print IsNumeric(Mid("00:00:00", 1, 2))
'----------------------------------------
'Debug.Print IsNumeric(TimeZone_Numeric)
'----------------------------------------
'A DATE VAR IS RETRUNED AS NUMERIC = TRUE
'----------------------------------------
'----------------------------------------
'Debug.Print IsNumeric(TimeZone)

'If IsNumeric(IsNumeric(Mid(TimeZone, 1, 2))) Then
'    temp = TimeZone
'Else
'    temp_str = TimeZone
'End If

End Function

Public Function IsLeapYear(TheYear As Integer) As Boolean
Dim Remainder As Integer
Remainder = TheYear Mod 4
    If Remainder = 0 Then
    Remainder = TheYear Mod 100
        If Remainder = 0 Then
        Remainder = TheYear Mod 400
            If Remainder = 0 Then
            IsLeapYear = True
            Else
            IsLeapYear = False
            End If
        Else
        IsLeapYear = True
        End If
    Else
    IsLeapYear = False
    End If
End Function

Public Function IsLeapYear2(Year As Integer) As Boolean
Dim x As Date
' this is the shortcut method
On Local Error GoTo Err1
x = CDate("02/29/" & Year)
IsLeapYear2 = True
Exit Function
Err1:
' type mismatch date error
IsLeapYear2 = False
End Function
Public Function MoonDescription(AnyDate As Date) As String
Dim Phase(1 To 3) As Date
Dim i As Integer
' since the actual phases are exact instants in time
' the moon's description will always be in between two
If Not DateInLimits(AnyDate) Then Exit Function
For i = 1 To 3
Phase(i) = UTtoLocal(JulianDaysToUT(MoonPhaseByLunation(Lunation(AnyDate), i)))
Next
If AnyDate >= Phase(3) Then
MoonPhaseDate = Phase(3)
MoonDescription = "Waning Crescent"
' then Last Quarter
ElseIf AnyDate >= Phase(2) Then
MoonPhaseDate = Phase(2)
MoonDescription = "Waning Gibbous"
' then Full Moon
ElseIf AnyDate >= Phase(1) Then
MoonPhaseDate = Phase(1)
MoonDescription = "Waxing Gibbous"
' then First Quarter
Else
MoonPhaseDate = Phase(3)
MoonDescription = "Waxing Crescent"
' then New Moon
End If


'Wheres One

End Function

Public Function NextLastQuarter(AnyDate As Date) As Date
Dim temp As Date
temp = JulianDaysToUT(MoonPhaseByLunation(Lunation(AnyDate), 3))
If temp > AnyDate Then
NextLastQuarter = temp
Else
NextLastQuarter = JulianDaysToUT(MoonPhaseByLunation(Lunation(AnyDate) + 1, 3))
End If
End Function
Public Function NextFirstQuarter(AnyDate As Date) As Date
Dim temp As Date
temp = JulianDaysToUT(MoonPhaseByLunation(Lunation(AnyDate), 1))
If temp > AnyDate Then
NextFirstQuarter = temp
Else
NextFirstQuarter = JulianDaysToUT(MoonPhaseByLunation(Lunation(AnyDate) + 1, 1))
End If
End Function

Public Function PreviousNewMoon(AnyDate As Date) As Date
PreviousNewMoon = JulianDaysToUT(MoonPhaseByLunation(Lunation(AnyDate), 0))
End Function
Public Function NextNewMoon(AnyDate As Date) As Date
Dim temp As Date
temp = JulianDaysToUT(MoonPhaseByLunation(Lunation(AnyDate), 0))
If temp > AnyDate Then
NextNewMoon = temp
Else
NextNewMoon = JulianDaysToUT(MoonPhaseByLunation(Lunation(AnyDate) + 1, 0))
End If
End Function
Public Function PreviousFirstQuarter(AnyDate As Date) As Date
Dim temp As Date
temp = JulianDaysToUT(MoonPhaseByLunation(Lunation(AnyDate), 1))
If temp < AnyDate Then
PreviousFirstQuarter = temp
Else
PreviousFirstQuarter = JulianDaysToUT(MoonPhaseByLunation(Lunation(AnyDate) - 1, 1))
End If
End Function
Public Function PreviousLastQuarter(AnyDate As Date) As Date
Dim temp As Date
temp = JulianDaysToUT(MoonPhaseByLunation(Lunation(AnyDate), 3))
If temp < AnyDate Then
PreviousLastQuarter = temp
Else
PreviousLastQuarter = JulianDaysToUT(MoonPhaseByLunation(Lunation(AnyDate) - 1, 3))
End If
End Function
Public Function PreviousFullMoon(AnyDate As Date) As Date
Dim temp As Date
temp = UTtoLocal(JulianDaysToUT(MoonPhaseByLunation(Lunation(AnyDate), 2)))
If temp < AnyDate Then
PreviousFullMoon = temp
Else
PreviousFullMoon = UTtoLocal(JulianDaysToUT(MoonPhaseByLunation(Lunation(AnyDate) - 1, 2)))
End If
End Function
Public Function NextFullMoon(AnyDate As Date) As Date
Dim temp As Date
temp = JulianDaysToUT(MoonPhaseByLunation(Lunation(AnyDate), 2))
If temp > AnyDate Then
NextFullMoon = temp
Else
NextFullMoon = JulianDaysToUT(MoonPhaseByLunation(Lunation(AnyDate) + 1, 2))
End If
End Function
Public Function UTtoLocal(Time As Date) As Date
' convert from universal time to system time and adjust for DST
Dim TZ As TIME_ZONE_INFORMATION
Dim temp As Date
Dim Rtn As Long
Rtn = GetTimeZoneInformation(TZ)
If Rtn > TIME_ZONE_ID_UNKNOWN Then
    If Rtn = TIME_ZONE_ID_STANDARD Then
    temp = DateAdd("n", -TZ.Bias, Time)
    Else
    temp = DateAdd("n", -(TZ.Bias + TZ.DaylightBias), Time)
    End If
End If

UTtoLocal = temp
End Function
Public Function DateInLimits(AnyDate As Date) As Boolean
On Local Error GoTo VbDateError
' little tricky because 01/01/0100 23:59:59 is
' numerically less than 01/01/0100 00:00:01 in VB
' due to their being negative after converted to double
' the +- integer part represents the day but the
' absolute value of the decimal part is the
' fraction of the day...see the Julian Day conversions
' Also, in this program only, must use 01/17/0100 04:14:49
' as a lower limit to avoid bad dates in intermediate functions
If CDbl(AnyDate) > -657418.176956019 And _
   CDbl(AnyDate) <= 2958465.99999 And _
   CDbl(AnyDate) <> -657418 Then
    If CDbl(AnyDate) <= -115869 Or _
       CDbl(AnyDate) > -115859.99999 Then
' no such date as "10/5/1582" to "10/14/1582 23:59:59"
    DateInLimits = True
    End If
End If
Exit Function
VbDateError:
DateInLimits = False
' type mismatch from bad Vb date
End Function

Public Function JulianDaysToUT(ByVal JD As Double) As Date
' Convert Julian Days to Universal Time.
' Julian Days are not the same as the Julian Calendar.
' The Julian period began Jan 1st, 4713 B.C. 12:00:00 with year 1
' The Julian period repeats after 7980 years (3268 A.D.)but for
' conversion purposes we will pretend it continues
' This will only work for dates >= 01/01/0100 00:00:01(VB limit)
Dim TheDay As Double
Dim Remainder As Double
If Not JulianWithinLimits(JD) Then Exit Function
' correct for skip from 10/4/1582 to 10/15/1582
If JD >= 2299160.5 Then
TheDay = JD - 2415018.5
Else
TheDay = JD - 2415028.5
End If
 ' handle positive and negative VB date values
If JD >= 2415020.5 Then ' positive VB dates > 1900
JulianDaysToUT = CDate(TheDay)
Else ' 100 A.D. to 1900
Remainder = CDbl(CDate(TheDay)) - Fix(CDate(TheDay))
    If Remainder <> 0 Then
    JulianDaysToUT = CDate(Fix(TheDay - 3) + (1 - Remainder))
    Else
    JulianDaysToUT = CDate(Fix(TheDay))
    End If
End If
End Function
Public Function JulianWithinLimits(JD As Double) As Boolean
' betwen the dates 01/01/0100 00:00:01 and
' 12/31/9999 23:59:59  These are the limits of the
' VB Date data type with one second added to the
' beginning to avoid overflow during calculations
' Cannot check for Oct 1582 error here
If JD >= 1757593.50001 And JD <= 5373484.49998843 Then
JulianWithinLimits = True
End If
End Function

Public Function UTtoJulianDays(ByVal AnyDate As Date) As Double
Dim TheDay As Double
Dim Remainder As Double

If Sgn(CDbl(AnyDate)) = -1 Then
Remainder = CDbl(AnyDate) - Fix(CDbl(AnyDate))
    If CDbl(AnyDate) - Remainder < CDbl(CDate("10/15/1582")) Then
    TheDay = Fix(CDbl(AnyDate)) + 2415028
    Else
    TheDay = Fix(CDbl(AnyDate)) + 2415018
    End If
    UTtoJulianDays = TheDay + 0.5 - Remainder
    Else
UTtoJulianDays = CDbl(AnyDate) + 2415018.5
End If
End Function

Public Function DaysInMonth(AnyDate As Date) As Integer
' returns number of days in a month
Dim Remainder As Integer
Dim TheMonth As Integer
Dim TheYear As Integer
TheMonth = Month(AnyDate)
TheYear = Year(AnyDate)
If TheMonth = 4 Or TheMonth = 6 Or TheMonth = 9 Or TheMonth = 11 Then
DaysInMonth = 30
Exit Function
ElseIf TheMonth = 2 Then
Remainder = TheYear Mod 4
    If Remainder = 0 Then
    Remainder = TheYear Mod 100
        If Remainder = 0 Then
        Remainder = TheYear Mod 400
            If Remainder = 0 Then
            DaysInMonth = 29
            Exit Function
            Else
            DaysInMonth = 28
            Exit Function
            End If
        Else
        DaysInMonth = 29
        Exit Function
        End If
    Else
    DaysInMonth = 28
    Exit Function
    End If
Else
DaysInMonth = 31
Exit Function
End If
End Function

Public Function NextLeapYear(AnyDate As Date) As Integer
Dim i As Integer
For i = 1 To 4
If DaysInMonth(CDate("02/" & Year(AnyDate) + i)) = 29 Then
NextLeapYear = Year(AnyDate) + i
Exit Function
End If
Next
End Function



Public Function DaylightSavingsTime(TheDate As Date) As Boolean
' returns True if a given date is within DST limits
Dim Oct31 As String
If Month(TheDate) > 10 Or Month(TheDate) < 4 Then
DaylightSavingsTime = False
Exit Function
End If
    
If Month(TheDate) < 10 And Month(TheDate) > 4 Then
DaylightSavingsTime = True
Exit Function
End If
   
If Month(TheDate) = 4 Then
    If Day(TheDate) < Weekday(TheDate) Then
    DaylightSavingsTime = False
    Exit Function
    Else
    DaylightSavingsTime = True
    Exit Function
    End If
End If
    
If Month(TheDate) = 10 Then
Oct31 = "10/31/" & Year(TheDate)

    If (Weekday(CDate(Oct31)) >= (Weekday(TheDate))) And (Day(TheDate) > 24) Then
    DaylightSavingsTime = False
    Exit Function
    Else
    DaylightSavingsTime = True
    Exit Function
    End If

End If
End Function




Public Function Illum(AnyDate As Date) As Single
' Percent of face illuminated
Dim Phase As Double
Dim AnyDate_Var As Date
Dim AnyDate_Offset As Date

'----
'Moon Phases 2017 – Lunar Calendar
'https://www.timeanddate.com/moon/phases/
'----
If Val(Replace(ATidalX.Label10.Caption, "%", "")) >= 70.8 Then
    If IsIDE = True And 1 = 2 Then
        Stop
    End If
End If

'SORT MOON CALIBRATE_PROPER
'ShiftPressed = True Debug Helper
'--------------------------------
If ShiftPressed_True_Debug_Helper_SORT_MOON_CALIBRATE_PROPER = False Then
If IsIDE = True And GetAsyncKeyState(16) < -300 Then Stop
End If

ShiftPressed_True_Debug_Helper_SORT_MOON_CALIBRATE_PROPER = True

AnyDate_Offset = TimeSerial(8, 9, 440)
Phase = Age(AnyDate + AnyDate_Offset)


Illum = (1 - Cos((2 * pi * Phase) / LPERIOD)) / 2 * 100
End Function


Public Function Lunation(ByVal A_Date As Date) As Long
' Number of mooncycles since 12/18/1922 12:20:09
' First, estimate the lunation using the average
' lunar period (LPERIOD)
Dim JDate As Double
Dim temp As Long 'trial value of lunation
Dim Test As Double ' lower limit of result
Dim Test2 As Double ' upper limit for interpolation
Dim Diff As Long
Diff = CDbl(A_Date) - EPOCH
temp = Diff / LPERIOD
JDate = UTtoJulianDays(A_Date) + (A_SECOND / 2)
' Add a half-second to account for the dropped
' fraction of a second in VB dates

' Now interpolate until correct,
' most cases will not need it.
Do
'THIS DO EVENTS NEED TO GO BECAUSE NOT SHUT DOWN AT UNLOAD
'DoEvents
Test = MoonPhaseByLunation(temp, 0)
Test2 = MoonPhaseByLunation(temp + 1, 0)
    If Test > JDate Then
    temp = temp - 1
    ElseIf Test2 < JDate Then
    temp = temp + 1
    End If
Loop Until Test <= JDate And Test2 > JDate
Lunation = temp
End Function


Public Function Fmod(ByVal x As Double, ByVal Y As Double) As Double
' returns the decimal remainder of division
Dim temp As Long
temp = x \ Y
temp = temp * Y
Fmod = Sgn(x) * Abs(x - temp)
End Function
Public Function MoonPhaseByLunation(ByVal TheLunation As Long, ByVal Phase As Integer) As Double
Dim k As Double
Dim t As Double
Dim i As Integer
' Phase: 0 = NewMoon, 1 = First Qrtr., 2 = Full, 3 = Last Qrtr.
' Lunation: number of complete lunar cycles since
' 12/18/1922 12:20:09 PM (lunation 0)
k = TheLunation - 953 + (Phase / 4)
' Calculate phase of the moon per Meeus, Ch. 47.
' Phase: 0 = NewMoon, 1 = First Qrtr., 2 = Full, 3 = Last Qrtr.
t = k / 1236.85
JDE = 2451550.09765 + (29.530588853 * k) + t * t * (0.0001337 + t * (-0.00000015 + 0.00000000073 * t))

' these are correction parameters used below
E = 1 + t * (-0.002516 + -0.0000074 * t)
M = 2.5534 + 29.10535669 * k + t * t * (-0.0000218 + -0.00000011 * t)
M1 = 201.5643 + 385.81693528 * k + t * t * (0.0107438 + t * (0.00001239 + -0.000000058 * t))
F = 160.7108 + 390.67050274 * k + t * t * (-0.0016341 * t * (-0.00000227 + 0.000000011 * t))
O = 124.7746 - 1.5637558 * k + t * t * (0.0020691 + 0.00000215 * t)
'planetary arguments

A(0) = 0  ' unused
A(1) = 299.77 + 0.107408 * k - 0.009173 * t * t
A(2) = 251.88 + 0.016321 * k
A(3) = 251.83 + 26.651886 * k
A(4) = 349.42 + 36.412478 * k
A(5) = 84.66 + 18.206239 * k
A(6) = 141.74 + 53.303771 * k
A(7) = 207.14 + 2.453732 * k
A(8) = 154.84 + 7.30686 * k
A(9) = 34.52 + 27.261239 * k
A(10) = 207.19 + 0.121824 * k
A(11) = 291.34 + 1.844379 * k
A(12) = 161.72 + 24.198154 * k
A(13) = 239.56 + 25.513099 * k
A(14) = 331.55 + 3.592518 * k

' all into radians except for E
M = Radians(M)
M1 = Radians(M1)
F = Radians(F)
O = Radians(O)

'all those planetary arguments, too
For i = 1 To 14
A(i) = Radians(A(i))
Next

' apply parameters to the JDE.
Adjust Phase
MoonPhaseByLunation = JDE
End Function

Public Function Age(AnyDate As Date) As Double
' days between previous new moon and AnyDate
Dim temp As Double
temp = MoonPhaseByLunation(Lunation(AnyDate), 0)
Age = UTtoJulianDays(AnyDate) - temp
End Function
Private Sub Adjust(Phase As Integer)
On Local Error GoTo Default

Select Case Phase

Case 0 ' New Moon
JDE = JDE - 0.4072 * Sin(M1) + 0.17241 * E * Sin(M) _
    + 0.01608 * Sin(2 * M1) + 0.01039 * Sin(2 * F) _
    + 0.00739 * E * Sin(M1 - M) - 0.00514 * E * Sin(M1 + M) _
    + 0.00208 * E * E * Sin(2 * M) - 0.00111 * Sin(M1 - 2 * F) _
    - 0.00057 * Sin(M1 + 2 * F) + 0.00056 * E * Sin(2 * M1 + M) _
    - 0.00042 * Sin(3 * M1) + 0.00042 * E * Sin(M + 2 * F) _
    + 0.00038 * E * Sin(M - 2 * F) - 0.00024 * E * Sin(2 * M1 - M) _
    - 0.00017 * Sin(O) - 0.00007 * Sin(M1 + 2 * M) _
    + 0.00004 * Sin(2 * M1 - 2 * F) + 0.00004 * Sin(3 * M) _
    + 0.00003 * Sin(M1 + M - 2 * F) + 0.00003 * Sin(2 * M1 + 2 * F) _
    - 0.00003 * Sin(M1 + M + 2 * F) + 0.00003 * Sin(M1 - M + 2 * F) _
    - 0.00002 * Sin(M1 - M - 2 * F) - 0.00002 * Sin(3 * M1 + M) _
    + 0.00002 * Sin(4 * M1)

Case 2 ' Full Moon
JDE = JDE - 0.40614 * Sin(M1) _
+ 0.17302 * E * Sin(M) + 0.01614 * Sin(2 * M1) _
+ 0.01043 * Sin(2 * F) + 0.00734 * E * Sin(M1 - M) _
- 0.00515 * E * Sin(M1 + M) + 0.00209 * E * E * Sin(2 * M) _
- 0.00111 * Sin(M1 - 2 * F) - 0.00057 * Sin(M1 + 2 * F) _
+ 0.00056 * E * Sin(2 * M1 + M) - 0.00042 * Sin(3 * M1) _
+ 0.00042 * E * Sin(M + 2 * F) + 0.00038 * E * Sin(M - 2 * F) _
- 0.00024 * E * Sin(2 * M1 - M) - 0.00017 * Sin(O) _
- 0.00007 * Sin(M1 + 2 * M) + 0.00004 * Sin(2 * M1 - 2 * F) _
+ 0.00004 * Sin(3 * M) + 0.00003 * Sin(M1 + M - 2 * F) _
+ 0.00003 * Sin(2 * M1 + 2 * F) - 0.00003 * Sin(M1 + M + 2 * F) _
+ 0.00003 * Sin(M1 - M + 2 * F) - 0.00002 * Sin(M1 - M - 2 * F) _
- 0.00002 * Sin(3 * M1 + M) + 0.00002 * Sin(4 * M1)

Case 1, 3 ' First Quarter, Last Quarter
JDE = JDE - 0.62801 * Sin(M1) _
+ 0.17172 * E * Sin(M) - 0.01183 * E * Sin(M1 + M) _
+ 0.00862 * Sin(2 * M1) + 0.00804 * Sin(2 * F) _
+ 0.00454 * E * Sin(M1 - M) + 0.00204 * E * E * Sin(2 * M) _
- 0.0018 * Sin(M1 - 2 * F) - 0.0007 * Sin(M1 + 2 * F) _
- 0.0004 * Sin(3 * M1) - 0.00034 * E * Sin(2 * M1 - M) _
+ 0.00032 * E * Sin(M + 2 * F) + 0.00032 * E * Sin(M - 2 * F) _
- 0.00028 * E * E * Sin(M1 + 2 * M) + 0.00027 * E * Sin(2 * M1 + M) _
- 0.00017 * Sin(O) - 0.00005 * Sin(M1 - M - 2 * F) _
+ 0.00004 * Sin(2 * M1 + 2 * F) - 0.00004 * Sin(M1 + M + 2 * F) _
+ 0.00004 * Sin(M1 - 2 * M) + 0.00003 * Sin(M1 + M - 2 * F) _
+ 0.00003 * Sin(3 * M) + 0.00002 * Sin(2 * M1 - 2 * F) _
+ 0.00002 * Sin(M1 - M + 2 * F) - 0.00002 * Sin(3 * M1 + M) _

' further adjustment for first and last quarters
W = 0.00306 - 0.00038 * E * Cos(M) + 0.00026 * Cos(M1) _
- 0.00002 * Cos(M1 - M) + 0.00002 * Cos(M1 + M) _
+ 0.00002 * Cos(2 * F)

' subtract for last, add for first
If Phase = 3 Then W = -W
JDE = JDE + W
End Select

' now there are some final correction to everything
JDE = JDE + 0.000325 * Sin(A(1)) + 0.000165 * Sin(A(2)) _
+ 0.000164 * Sin(A(3)) + 0.000126 * Sin(A(4)) _
+ 0.00011 * Sin(A(5)) + 0.000062 * Sin(A(6)) _
+ 0.00006 * Sin(A(7)) + 0.000056 * Sin(A(8)) _
+ 0.000047 * Sin(A(9)) + 0.000042 * Sin(A(10)) _
+ 0.00004 * Sin(A(11)) + 0.000037 * Sin(A(12)) _
+ 0.000035 * Sin(A(13)) + 0.000023 * Sin(A(14))

Exit Sub
Default:
'Debug.Print "The Moon has exploded!"
End Sub

Private Function Radians(ByVal x As Double) As Double
'convert x to radians
Dim temp As Double
temp = Fmod(x, 360) ' normalize the angle
Radians = temp * 1.74532925199433E-02
End Function



