Attribute VB_Name = "MoonRise"
'
'   This module contains a user defined spreadsheet function called
'   sunevent(year, month, day, glong, glat, tz,event) that returns the
'   time of day of sunrise, sunset, or the beginning and end one of
'   three kinds of twilight
'
'   The method used here is adapted from Montenbruck and Pfleger's
'   Astronomy on the Personal Computer, 3rd Ed, section 3.8
'
'   the arguments for the function are as follows...
'   year, month, day - your date in zone time
'   glong - your longitude in degrees, west negative
'   glat - your latitude in degrees, north positive
'   tz - your time zone in decimal hours, west or 'behind' Greenwich negative
'   event - a code integer representing the event you want as follows
'       positive integers mean rising events
'       negative integers mean setting events
'       1 = sunrise                 -1  = sunset
'       2 = begin civil twilight    -2  = end civil twilight
'       3 = begin nautical twi      -3  = end nautical twi
'       4 = begin astro twi         -4  = end astro twi
'
'   the results are returned as a variant with either a time of day
'   in zone time or a string reporting an 'event not occuring' condition
'   event not occuring can be one of the following
'       .....    always below horizon, so no rise or set
'       *****    always above horizon, so no rise or set
'       -----    the particular rise or set event does not occur on that day
'
'   The function will produce meaningful results at all latitudes
'   but there will be a small range of latitudes around 67.43 degrees North or South
'   when the function might indicate a sunrise very close to noon (or a sunset
'   very soon after noon) where in fact the Sun is below the horizon all day
'   this behaviour relates to the approximate Sun position formulas in use
'
'   As always, the sunrise / set times relate to an earth which is smooth
'   and has no large obstructions on the horizon - you might get a close
'   approximation to this at sea but rarely on land. Accuracy more than 1 min
'   of time is not worth striving for - atmospheric refraction alone can
'   alter observed rise times by minutes
'
'   The module also defines a series of 'named funtions' based on sunevent()
'   as follows
'   astrotwilightstarts(date, tz, glong, glat)
'   nauticaltwilightstarts(date, tz, glong, glat)
'   civiltwilightstarts(date, tz, glong, glat)
'   sunrise(date, tz, glong, glat)
'   sunset(date, tz, glong, glat)
'   civiltwilightends(date, tz, glong, glat)
'   nauticaltwilightends(date, tz, glong, glat)
'   astrotwilightends(date, tz, glong, glat)
'
'   these functions  take a date in Excel date format and return times in
'   Excel time format (ie a day fraction) if there is an event on the day
'   If there isn't an event on the day, then you get #VALUE!
'   These functions lend themselves to plotting sunrise and set times on
'   charts - just multiply output in 'number' format by 24 to get the decimal
'   hours for the event
'
Private Function mjd(year As Integer, month As Integer, day As Integer) As Double
'
'   takes the year, month and day as a Gregorian calendar date
'   and returns the modified julian day number
'
    Dim a As Double
    Dim b As Double
    
    If (month <= 2) Then
        month = month + 12
        year = year - 1
        End If
    b = Fix(year / 400) - Fix(year / 100) + Fix(year / 4)
    a = 365# * year - 679004#
    mjd = a + b + Fix(30.6001 * (month + 1)) + day
End Function


Private Function frac(x As Double) As Double
'
'  returns the fractional part of x as used in minimoon and minisun
'
    Dim a As Double
    a = x - Fix(x)
    frac = a
End Function

Private Function range(x As Double) As Double
'
'   returns an angle in degrees in the range 0 to 360
'   used to condition the arguments for the Sun's orbit
'   in function minisun below
'
    Dim a As Double
    Dim b As Double
    b = x / 360
    a = 360 * (b - Fix(b))
    If (a < 0) Then
        a = a + 360
        End If
    range = a
End Function

Private Function hrsmin(t As Double) As String
'
'   takes a time as a decimal number of hours between 0 and 23.9999...
'   and returns a string with the time in hhmm format
'
    Dim hour As Double
    Dim min As Double
    hour = Fix(t)
    min = Fix((t - hour) * 60 + 0.5)
    hrsmin = Format(hour, "00") + Format(min, "00")
End Function


Private Function lmst(mjd As Double, glong As Double) As Double
'
'  Takes the mjd and the longitude (west negative) and then returns
'  the local sidereal time in hours. Im using Meeus formula 11.4
'  instead of messing about with UTo and so on
'
    Dim lst As Double
    Dim t As Double
    Dim d As Double
    d = mjd - 51544.5
    t = d / 36525#
    lst = range(280.46061837 + 360.98564736629 * d + 0.000387933 * t * t - t * t * t / 38710000)
    lmst = lst / 15# + glong / 15
End Function
    
Private Sub minisun(t As Double, ra As Double, dec As Double)
'
'   takes t (julian centuries since J2000.0) and empty variables ra and dec
'   sets ra and dec to the value of the Sun coordinates at t
'
'   positions claimed to be within 1 arc min by Montenbruck and Pfleger
'
    Dim p2 As Double
    Dim coseps As Double
    Dim sineps As Double
    Dim L As Double
    Dim M As Double
    Dim DL As Double
    Dim SL As Double
    Dim x As Double
    Dim Y As Double
    Dim Z As Double
    Dim RHO As Double
    p2 = 6.283185307
    coseps = 0.91748
    sineps = 0.39778

    M = p2 * frac(0.993133 + 99.997361 * t)
    DL = 6893# * Sin(M) + 72# * Sin(2 * M)
    L = p2 * frac(0.7859453 + M / p2 + (6191.2 * t + DL) / 1296000)
    SL = Sin(L)
    x = Cos(L)
    Y = coseps * SL
    Z = sineps * SL
    RHO = Sqr(1 - Z * Z)
    dec = (360# / p2) * Atn(Z / RHO)
    ra = (48# / p2) * Atn(Y / (x + RHO))
    If ra < 0 Then ra = ra + 24
End Sub


Private Sub quad(ym As Double, yz As Double, yp As Double, nz As Integer, z1 As Double, z2 As Double, xe As Double, ye As Double)
'
'  finds the parabola throuh the three points (-1,ym), (0,yz), (1, yp)
'  and sets the coordinates of the max/min (if any) xe, ye
'  the values of x where the parabola crosses zero (z1, z2)
'  and the nz number of roots (0, 1 or 2) within the interval [-1, 1]
'
    Dim a As Double
    Dim b As Double
    Dim c As Double
    Dim dis As Double
    Dim dx As Double

    nz = 0
    a = 0.5 * (ym + yp) - yz
    b = 0.5 * (yp - ym)
    c = yz
    xe = -b / (2 * a)
    ye = (a * xe + b) * xe + c
    dis = b * b - 4# * a * c
    If (dis > 0) Then
        dx = 0.5 * Sqr(dis) / Abs(a)
        z1 = xe - dx
        z2 = xe + dx
        If (Abs(z1) <= 1#) Then nz = nz + 1
        If (Abs(z2) <= 1#) Then nz = nz + 1
        If (z1 < -1#) Then z1 = z2
        End If
End Sub

Private Function SinAltSun(mjd0 As Double, hour As Double, glong As Double, cglat As Double, sglat As Double) As Double
'
'  this rather mickey mouse function takes a lot of
'  arguments and then returns the sine of the altitude of
'  the object labelled by iobj. iobj = 1 is moon, iobj = 2 is sun
'
    Dim mjd As Double
    Dim t As Double
    Dim ra As Double
    Dim dec As Double
    Dim tau As Double
    Dim salt As Double
    Dim rads As Double
    rads = 0.0174532925
    mjd = mjd0 + hour / 24#
    t = (mjd - 51544.5) / 36525#
    Call minisun(t, ra, dec)
    ' hour angle of object
    tau = 15# * (lmst(mjd, glong) - ra)
    ' sin(alt) of object using the conversion formulas
    salt = sglat * Sin(rads * dec) + cglat * Cos(rads * dec) * Cos(rads * tau)
    SinAltSun = salt
End Function

'
'   Worksheet functions below....
'
    
Function sunevent(year As Integer, month As Integer, day As Integer, tz As Double, glong As Double, glat As Double, EventType As Integer) As String
'
'   This is the function that does most of the work
'
    Dim sglong As Double, sglat As Double, cglat As Double, ddate As Double, ym As Double
    Dim yz As Double, above As Integer, utrise As Double, utset As Double
    Dim yp As Double, nz As Integer, rise As Integer, sett As Integer, j As Integer
    Dim hour As Double, z1 As Double, z2 As Double, rads As Double, xe As Double, ye As Double
    Dim AlwaysUp As String, AlwaysDown As String, OutString As String, NoEvent As String
    Dim sinho(5) As Double
    rads = 0.0174532925
    AlwaysUp = "****"
    AlwaysDown = "...."
    NoEvent = "----"
    
'
'   Set up the array with the 4 values of sinho needed for the 4
'   kinds of sun event
'
    sinho(1) = Sin(rads * -0.833)     'sunset upper limb simple refraction
    sinho(2) = Sin(rads * -6#)        'civil twi
    sinho(3) = Sin(rads * -12#)       'nautical twi
    sinho(4) = Sin(rads * -18#)       'astro twi
    sglat = Sin(rads * glat)
    cglat = Cos(rads * glat)
    ddate = mjd(year, month, day) - tz / 24
'
'   main loop takes each value of sinho in turn and finds the rise/set
'   events associated with that altitude of the Sun
'
    j = Abs(EventType)
        nz = 0
        z1 = 0
        z2 = 0
        xe = 0
        ye = 0
        rise = 0
        sett = 0
        above = 0
        hour = 1#
        ym = SinAltSun(ddate, hour - 1#, glong, cglat, sglat) - sinho(j)
        If (ym > 0#) Then above = 1
        '
        '  the while loop finds the sin(alt) for sets of three consecutive
        '  hours, and then tests for a single zero crossing in the interval
        '  or for two zero crossings in an interval or for a grazing event
        '  The flags rise and sett are set accordingly
        '
        Do While (hour < 25 And (sett = 0 Or rise = 0))
            yz = SinAltSun(ddate, hour, glong, cglat, sglat) - sinho(j)
            yp = SinAltSun(ddate, hour + 1#, glong, cglat, sglat) - sinho(j)
            Call quad(ym, yz, yp, nz, z1, z2, xe, ye)
            ' case when one event is found in the interval
            If (nz = 1) Then
                If (ym < 0#) Then
                    utrise = hour + z1
                    rise = 1
                Else
                    utset = hour + z1
                    sett = 1
                End If
            End If ' end of nz = 1 case
            '
            '   case where two events are found in this interval
            '   (rare but whole reason we are not using simple iteration)
            '
            If (nz = 2) Then
                If (ye < 0#) Then
                    utrise = hour + z2
                    utset = hour + z1
                Else
                    utrise = hour + z1
                    utset = hour + z2
                End If
            rise = 1
            sett = 1
            End If
            '
            '   set up the next search interval
            '
            ym = yp
            hour = hour + 2#

        Loop ' end of while loop
            '
            ' now search has completed, we compile the string to pass back
            ' to the user. The string depends on several combinations
            ' of the above flag (always above or always below) and the rise
            ' and sett flags
            '
        If (rise = 1 Or sett = 1) Then
            If (rise = 1) Then
                If (EventType > 0) Then OutString = hrsmin(utrise)
            Else
                If (EventType > 0) Then OutString = NoEvent
            End If
            If (sett = 1) Then
                If (EventType < 0) Then OutString = hrsmin(utset)
            Else
                If (EventType < 0) Then OutString = NoEvent
            End If
        Else
            If (above = 1) Then
                OutString = AlwaysUp
            Else
                OutString = AlwaysDown
            End If
        End If
        sunevent = OutString
End Function

Function sunrise(ddate As Date, tz As Double, glong As Double, glat As Double) As Double
'
'   simple way of calling sunevent() using the Excel date format
'   returns just the sunrise time or NULL if no event
'   I used the day(), month() and year() functions in excel to allow
'   portability to the MAC (different date serial numbers)
'
    Dim EventTime As Double, hour As Double, minfrac As Double
    Dim out As String
    out = sunevent(year(ddate), month(ddate), day(ddate), tz, glong, glat, 1)
    If (out = "...." Or out = "----" Or out = "****") Then
        EventTime = Null
    Else
        hour = Fix(Val(out) / 100)
        minfrac = (Val(out) - 100 * hour) / 60
        hour = hour + minfrac
        hour = hour / 24
        EventTime = hour
    End If
    sunrise = EventTime
End Function

Function sunset(ddate As Date, tz As Double, glong As Double, glat As Double) As Double
'
'   simple way of calling sunevent() using the Excel date format
'   returns just the sunset time or ****, ...., ---- as approptiate in a string
'   I used the day(), month() and year() functions in excel to allow
'   portability to the MAC (different date serial number base)
'
    Dim EventTime As Double, hour As Double, minfrac As Double
    Dim out As String
    out = sunevent(year(ddate), month(ddate), day(ddate), tz, glong, glat, -1)
    If (out = "...." Or out = "----" Or out = "****") Then
        EventTime = Null
    Else
        hour = Fix(Val(out) / 100)
        minfrac = (Val(out) - 100 * hour) / 60
        hour = hour + minfrac
        hour = hour / 24
        EventTime = hour
    End If
    sunset = EventTime
End Function

Function CivilTwilightStarts(ddate As Date, tz As Double, glong As Double, glat As Double) As Double
'
'   simple way of calling sunevent() using the Excel date format
'   returns just the start of civil twilight time or ****, ...., ---- as approptiate
'   I used the day(), month() and year() functions in excel to allow
'   portability to the MAC (different date serial numbers)
'
    Dim EventTime As Double, hour As Double, minfrac As Double
    Dim out As String
    out = sunevent(year(ddate), month(ddate), day(ddate), tz, glong, glat, 2)
    If (out = "...." Or out = "----" Or out = "****") Then
        EventTime = Null
    Else
        hour = Fix(Val(out) / 100)
        minfrac = (Val(out) - 100 * hour) / 60
        hour = hour + minfrac
        hour = hour / 24
        EventTime = hour
    End If
    CivilTwilightStarts = EventTime
End Function

Function CivilTwilightEnds(ddate As Date, tz As Double, glong As Double, glat As Double) As Double
'
    Dim EventTime As Double, hour As Double, minfrac As Double
    Dim out As String
    out = sunevent(year(ddate), month(ddate), day(ddate), tz, glong, glat, -2)
    If (out = "...." Or out = "----" Or out = "****") Then
        EventTime = Null
    Else
        hour = Fix(Val(out) / 100)
        minfrac = (Val(out) - 100 * hour) / 60
        hour = hour + minfrac
        hour = hour / 24
        EventTime = hour
    End If
    CivilTwilightEnds = EventTime
    End Function

Function NauticalTwilightStarts(ddate As Date, tz As Double, glong As Double, glat As Double) As Double
'
    Dim EventTime As Double, hour As Double, minfrac As Double
    Dim out As String
    out = sunevent(year(ddate), month(ddate), day(ddate), tz, glong, glat, 3)
    If (out = "...." Or out = "----" Or out = "****") Then
        EventTime = Null
    Else
        hour = Fix(Val(out) / 100)
        minfrac = (Val(out) - 100 * hour) / 60
        hour = hour + minfrac
        hour = hour / 24
        EventTime = hour
    End If
    NauticalTwilightStarts = EventTime
End Function

Function NauticalTwilightEnds(ddate As Date, tz As Double, glong As Double, glat As Double) As Double
'
    Dim EventTime As Double, hour As Double, minfrac As Double
    Dim out As String
    out = sunevent(year(ddate), month(ddate), day(ddate), tz, glong, glat, -3)
    If (out = "...." Or out = "----" Or out = "****") Then
        EventTime = Null
    Else
        hour = Fix(Val(out) / 100)
        minfrac = (Val(out) - 100 * hour) / 60
        hour = hour + minfrac
        hour = hour / 24
        EventTime = hour
    End If
    NauticalTwilightEnds = EventTime
End Function

Function AstroTwilightStarts(ddate As Date, tz As Double, glong As Double, glat As Double) As Double
'
    Dim EventTime As Double, hour As Double, minfrac As Double
    Dim out As String
    out = sunevent(year(ddate), month(ddate), day(ddate), tz, glong, glat, 4)
    If (out = "...." Or out = "----" Or out = "****") Then
        EventTime = Null
    Else
        hour = Fix(Val(out) / 100)
        minfrac = (Val(out) - 100 * hour) / 60
        hour = hour + minfrac
        hour = hour / 24
        EventTime = hour
    End If
    AstroTwilightStarts = EventTime
End Function

Function AstroTwilightEnds(ddate As Date, tz As Double, glong As Double, glat As Double) As Double
'
    Dim EventTime As Double, hour As Double, minfrac As Double
    Dim out As String
    out = sunevent(year(ddate), month(ddate), day(ddate), tz, glong, glat, -4)
    If (out = "...." Or out = "----" Or out = "****") Then
        EventTime = Null
    Else
        hour = Fix(Val(out) / 100)
        minfrac = (Val(out) - 100 * hour) / 60
        hour = hour + minfrac
        hour = hour / 24
        EventTime = hour
    End If
    AstroTwilightEnds = EventTime
End Function





