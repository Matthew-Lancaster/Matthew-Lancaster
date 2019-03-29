VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSunRiseSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 ' ------ Class clsSunRiseSet
' Version 2.01

Option Explicit

' -- The following properties are exposed:
'Sunrise (r) - Sunrise time
'Sunset (r) - Sunset time
'SolarNoon (r) - Solar noon
'
'CityCount (r) - Number of cities
'CityName (r) - Name of city, by index
'City (w) - Sets the longitude/latitude/timezone based off a city
'           name or city index
'
'TimeZone (r/w) - Current Timezone
'DaySavings (r/w) - Daylight savings time in effect
'Longitude (r/w) - Longitude to calculate for
'Latitude (r/w) - Latitude to calculate for
'
'DateDay (r/w) - Date to calculate for
'
'
' ------ Class clsSunRiseSet
' Version 2.01
' -- The following method is exposed
'CalculateSun - Calculate sunrise, sunset and solar noon
'
'
' Scott Seligman
'   http://www.scottandmichelle.net/scott/
' Based off of
'   http://www.srrb.noaa.gov/highlights/sunrise/gen.html

Private Type typeMonth
    Name As String
    NumDays As Long
End Type

Private Type typeCity
    Name As String
    longitude As Double
    latitude As Double
    TimeZone As Long
End Type

Private m_cNumberCities As Long
Private m_Cities() As typeCity
    
Private m_monthList(0 To 11) As typeMonth
Private m_monthLeap(0 To 11) As typeMonth

Private m_nTimeZone As Long
Private m_bDaySavings As Boolean
Private m_nLongitude As Double
Private m_nLatitude As Double
Private m_dateSel As Date

Private m_dateSunrise As Date
Private m_dateSunset As Date
Private m_dateNoon As Date

Public Property Get Sunrise() As Date
    Sunrise = m_dateSunrise
End Property

Public Property Get Sunset() As Date
    Sunset = m_dateSunset
End Property

Public Property Get SolarNoon() As Date
    SolarNoon = m_dateNoon
End Property

Public Property Get CityCount() As Long
    CityCount = m_cNumberCities + 1
End Property

Public Property Get CityName(nCity As Long) As String
    If nCity < 0 Or nCity > m_cNumberCities Then
        CityName = "(Error)"
    Else
        CityName = m_Cities(nCity).Name
    End If
End Property

Public Property Let City(City)
    Dim nCity As Long
    Dim bFound As Boolean
    
    If VarType(City) = vbString Then
        For nCity = 0 To m_cNumberCities
            If Trim(LCase(City)) = _
                Trim(LCase(m_Cities(nCity).Name)) Then
                bFound = True
                Exit For
            End If
        Next
        If Not bFound Then
            nCity = -1
        End If
    Else
        If IsNumeric(City) Then
            nCity = City
        Else
            nCity = -1
        End If
    End If
    
    If nCity < 0 Or nCity > m_cNumberCities Then
        m_nTimeZone = 0
        m_bDaySavings = False
        m_nLongitude = 0
        m_nLatitude = 0
    Else
        m_nTimeZone = m_Cities(nCity).TimeZone
        m_bDaySavings = False
        m_nLongitude = m_Cities(nCity).longitude
        m_nLatitude = m_Cities(nCity).latitude
    End If
    
End Property

Public Property Let TimeZone(nNew As Long)
    m_nTimeZone = nNew
End Property

Public Property Get TimeZone() As Long
    TimeZone = m_nTimeZone
End Property

Public Property Let DaySavings(bNew As Boolean)
    m_bDaySavings = bNew
End Property

Public Property Get DaySavings() As Boolean
    DaySavings = m_bDaySavings
End Property

Public Property Let longitude(nNew As Double)
    m_nLongitude = nNew
End Property

Public Property Get longitude() As Double
    longitude = m_nLongitude
End Property

Public Property Let latitude(nNew As Double)
    m_nLatitude = nNew
End Property

Public Property Get latitude() As Double
    latitude = m_nLatitude
End Property
    
Public Property Let DateDay(dateNew As Date)
    m_dateSel = dateNew
End Property
    
Public Property Get DateDay() As Date
    DateDay = m_dateSel
End Property


Private Function IsLeapYear(nYear As Long) As Boolean
    If (nYear Mod 4 = 0 And nYear Mod _
        100 <> 0) Or nYear Mod 400 = 0 Then
        IsLeapYear = True
    Else
        IsLeapYear = False
    End If
End Function

Private Function RadToDeg(angleRad As Double) As Double
    RadToDeg = 180 * angleRad / 3.1415926535
End Function

Private Function DegToRad(angleDeg As Double) As Double
    DegToRad = 3.1415926535 * angleDeg / 180
End Function

Private Function acos(X As Double) As Double
    On Error Resume Next
    acos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
End Function

Private Sub InitMonths()

    m_monthList(0).Name = "January": m_monthList(0).NumDays = 31
    m_monthList(1).Name = "February": m_monthList(1).NumDays = 28
    m_monthList(2).Name = "March": m_monthList(2).NumDays = 31
    m_monthList(3).Name = "April": m_monthList(3).NumDays = 30
    m_monthList(4).Name = "May": m_monthList(4).NumDays = 31
    m_monthList(5).Name = "June": m_monthList(5).NumDays = 30
    m_monthList(6).Name = "July": m_monthList(6).NumDays = 31
    m_monthList(7).Name = "August": m_monthList(7).NumDays = 31
    m_monthList(8).Name = "September": m_monthList(8).NumDays = 30
    m_monthList(9).Name = "October": m_monthList(9).NumDays = 31
    m_monthList(10).Name = "November": m_monthList(10).NumDays = 30
    m_monthList(11).Name = "DEcember": m_monthList(11).NumDays = 31
    
    m_monthLeap(0).Name = "January": m_monthLeap(0).NumDays = 31
    m_monthLeap(1).Name = "February": m_monthLeap(1).NumDays = 29
    m_monthLeap(2).Name = "March": m_monthLeap(2).NumDays = 31
    m_monthLeap(3).Name = "April": m_monthLeap(3).NumDays = 30
    m_monthLeap(4).Name = "May": m_monthLeap(4).NumDays = 31
    m_monthLeap(5).Name = "June": m_monthLeap(5).NumDays = 30
    m_monthLeap(6).Name = "July": m_monthLeap(6).NumDays = 31
    m_monthLeap(7).Name = "August": m_monthLeap(7).NumDays = 31
    m_monthLeap(8).Name = "September": m_monthLeap(8).NumDays = 30
    m_monthLeap(9).Name = "October": m_monthLeap(9).NumDays = 31
    m_monthLeap(10).Name = "November": m_monthLeap(10).NumDays = 30
    m_monthLeap(11).Name = "DEcember": m_monthLeap(11).NumDays = 31

End Sub

Private Sub AddCity(sCity As String, nLatitude As Double, _
    nLongitude As Double, nZone As Long)


    m_cNumberCities = m_cNumberCities + 1
    If m_cNumberCities > UBound(m_Cities) Then
        ReDim Preserve m_Cities(UBound(m_Cities) + 10)
    End If

    m_Cities(m_cNumberCities).Name = sCity
    m_Cities(m_cNumberCities).latitude = nLatitude
    m_Cities(m_cNumberCities).longitude = nLongitude
    m_Cities(m_cNumberCities).TimeZone = nZone

End Sub

Private Sub InitCities()

    m_cNumberCities = -1
    ReDim m_Cities(0)

    AddCity "Albuquerque, NM", 35.0833, 106.65, 7
    AddCity "Anchorage, AK", 61.217, 149.9, 9
    AddCity "Atlanta, GA", 33.733, 84.383, 5
    AddCity "Austin, TX", 30.283, 97.733, 6
    AddCity "Birmingham, AL", 33.521, 86.8025, 6
    AddCity "Bismarck, ND", 46.817, 100.783, 6
    AddCity "Boston, MA", 42.35, 71.05, 5
    AddCity "Boulder, CO", 40.125, 105.237, 7
    AddCity "Chicago, IL", 41.85, 87.65, 6
    AddCity "Dallas, TX", 32.46, 96.47, 6
    AddCity "Denver, CO", 39.733, 104.983, 7
    AddCity "Detroit, MI", 42.333, 83.05, 5
    AddCity "Honolulu, HI", 21.3, 157.85, 10
    AddCity "Houston, TX", 29.75, 95.35, 6
    AddCity "Indianapolis, IN", 39.767, 86.15, 5
    AddCity "Jackson, MS", 32.283, 90.183, 6
    AddCity "Kansas City, MO", 39.083, 94.567, 6
    AddCity "Los Angeles, CA", 34.05, 118.233, 8
    AddCity "Menomonee Falls, WI", 43.11, 88.1, 6
    AddCity "Miami, FL", 25.767, 80.183, 5
    AddCity "Minneapolis, MN", 44.967, 93.25, 6
    AddCity "New Orleans, LA", 29.95, 90.067, 6
    AddCity "New York City, NY", 40.7167, 74.0167, 5
    AddCity "Oklahoma City, OK", 35.483, 97.533, 6
    AddCity "Philadelphia, PA", 39.95, 75.15, 5
    AddCity "Phoenix, AZ", 33.433, 112.067, 7
    AddCity "Pittsburgh, PA", 40.433, 79.9833, 5
    AddCity "Portland, ME", 43.666, 70.283, 5
    AddCity "Portland, OR", 45.517, 122.65, 8
    AddCity "Raleigh, NC", 35.783, 78.65, 5
    AddCity "Richmond, VA", 37.5667, 77.45, 5
    AddCity "Saint Louis, MO", 38.6167, 90.1833, 6
    AddCity "San Diego, CA", 32.7667, 117.2167, 8
    AddCity "San Francisco, CA", 37.7667, 122.4167, 8
    AddCity "Seattle, WA", 47.6, 122.3167, 8
    AddCity "Washington DC", 38.8833, 77.0333, 5
    AddCity "Beijing, China", 39.9167, -116.4167, -8
    AddCity "Berlin, Germany", 52.33, -13.3, -1
    AddCity "Bombay, India", 18.9333, -72.8333, -5.5
    AddCity "Buenos Aires, Argentina", -34.6, 58.45, 3
    AddCity "Cairo, Egypt", 30.1, -31.3667, -2
    AddCity "Cape Town, South Africa", -33.9167, -18.3667, -2
    AddCity "Caracas, Venezuela", 10.5, 66.9333, 4
    AddCity "Helsinki, Finland", 60.1667, -24.9667, -2
    AddCity "Hong Kong, China", 22.25, -114.1667, -8
    AddCity "Jerusalem, Israel", 31.7833, -35.2333, -2
    AddCity "London, England", 51.5, 0.1667, 0
    AddCity "Mexico City, Mexico", 19.4, 99.15, 6
    AddCity "Moscow, Russia", 55.75, -37.5833, -3
    AddCity "New Delhi, India", 28.6, -77.2, -5.5
    AddCity "Ottawa, Canada", 45.41667, 75.7, 5
    AddCity "Paris, France", 48.8667, -2.667, -1
    AddCity "Rio de Janeiro, Brazil", -22.9, 43.2333, 3
    AddCity "Riyadh, Saudi Arabia", 24.633, -46.71667, -3
    AddCity "Rome, Italy", 41.9, -12.4833, -1
    AddCity "Sydney, Australia", -33.8667, -151.2167, -10
    AddCity "Tokyo, Japan", 35.7, -139.7667, -9
    AddCity "Zurich, Switzerland", 47.3833, -8.5333, -1

End Sub

Private Sub Class_Initialize()

    InitMonths
    InitCities
    m_dateSel = Now

End Sub

Private Function calcDayOfYear(ByVal nmn As Long, ByVal ndy As Long, _
    bLeapYear As Boolean) As Long

    calcDayOfYear = Int((275 * nmn) / 9) - IIf(bLeapYear, 1, 2) _
        * Int((nmn + 9) / 12) + ndy - 30

End Function

Private Function calcJD(ByVal nYear As Long, ByVal nMonth As Long, _
    ByVal nDay As Long) As Double

    If nMonth <= 2 Then
        nYear = nYear - 1
        nMonth = nMonth + 12
    End If
    Dim A As Long
    Dim B As Long
    
    A = Int(nYear / 100)
    B = 2 - A + Int(A / 4)
    
    calcJD = Int(365.25 * (nYear + 4716)) + Int(30.6001 * _
        (nMonth + 1)) + nDay + B - 1524.5

End Function

Private Function calcTimeJulianCent(ByVal njd As Double) As Double

    calcTimeJulianCent = (njd - 2451545#) / 36525#
    
End Function

Private Function calcJDFromJulianCent(ByVal nt As Double) As Double

    calcJDFromJulianCent = nt * 36525# + 2451545#

End Function

Private Function calcGeomMeanLongSun(ByVal nt As Double) As Double

    Dim nLO As Double
    nLO = 280.46646 + nt * (36000.76983 + 0.0003032 * nt)
    Do While nLO > 360#
        nLO = nLO - 360#
    Loop
    Do While nLO < 0
        nLO = nLO + 360#
    Loop
    calcGeomMeanLongSun = nLO

End Function

Private Function calcGeomMeanAnomalySun(ByVal nt As Double) As Double

    calcGeomMeanAnomalySun = 357.52911 + nt * (35999.05029 - _
        0.0001537 * nt)

End Function

Private Function calcEccentricityEarthOrbit(ByVal nt As Double) _
    As Double

    calcEccentricityEarthOrbit = 0.016708634 - nt * _
        (0.000042037 + 0.0000001267 * nt)

End Function

Private Function calcSunEqOfCenter(ByVal nt As Double) As Double

    Dim nm As Double
    Dim nmrad As Double
    Dim nsinm As Double
    Dim nsin2m As Double
    Dim nsin3m As Double

    nm = calcGeomMeanAnomalySun(nt)
    
    nmrad = DegToRad(nm)
    nsinm = Sin(nmrad)
    nsin2m = Sin(nmrad + nmrad)
    nsin3m = Sin(nmrad + nmrad + nmrad)
    
    calcSunEqOfCenter = nsinm * (1.914602 - nt * _
        (0.004817 + 0.000014 * nt)) + nsin2m * _
        (0.019993 - 0.000101 * nt) + nsin3m * 0.000289

End Function

Private Function calcSunTrueLong(ByVal nt As Double) As Double

    Dim n10 As Double
    Dim nc As Double
    
    n10 = calcGeomMeanLongSun(nt)
    nc = calcSunEqOfCenter(nt)
    
    calcSunTrueLong = n10 + nc

End Function

Private Function calcSunApparentLong(ByVal nt As Double) As Double
    Dim no As Double
    Dim nomega As Double
    
    no = calcSunTrueLong(nt)
    nomega = 125.04 - 1934.136 * nt
    calcSunApparentLong = no - 0.00569 - 0.00478 * _
        Sin(DegToRad(nomega))
    
End Function

Private Function calcMeanObliquityOfEcliptic(ByVal nt As Double) _
    As Double

    Dim nseconds As Double
    
    nseconds = 21.448 - nt * (46.815 + nt * _
        (0.00059 - nt * (0.001813)))
    calcMeanObliquityOfEcliptic = 23# + (26# + _
        (nseconds / 60#)) / 60#

End Function

Private Function calcObliquityCorrection(ByVal nt As Double) As Double

    Dim ne0 As Double
    ne0 = calcMeanObliquityOfEcliptic(nt)
    
    Dim nomega As Double
    nomega = 125.04 - 1934.136 * nt
    calcObliquityCorrection = ne0 + 0.00256 * Cos(DegToRad(nomega))

End Function

Private Function calcSunRtAscension(nt As Double) As Double

    Dim ne As Double
    Dim nlambda As Double
    Dim ntananum As Double
    Dim ntanadenom As Double
    
    ne = calcObliquityCorrection(nt)
    nlambda = calcSunApparentLong(nt)

    ntananum = (Cos(DegToRad(ne)) * Sin(DegToRad(nlambda)))
    ntanadenom = (Cos(DegToRad(nlambda)))
    
    calcSunRtAscension = RadToDeg(atan2(ntananum, ntanadenom))
    
End Function

Private Function atan2(ByVal Y As Double, ByVal X As Double) As Double

    If X > 0 Then
        atan2 = Atn(Y / X)
    ElseIf X < 0 Then
        atan2 = Atn(Y / X) + 3.1415926535
    Else
        atan2 = 3.1415926535 / 2 * Sgn(Y)
    End If

End Function

Private Function asin(ByVal X As Double) As Double
    asin = Atn(X / Sqr(-X * X + 1))
End Function

Private Function calcSunDeclination(ByVal nt As Double) As Double

    Dim ne As Double
    Dim nlambda As Double
    Dim nsint As Double
    
    ne = calcObliquityCorrection(nt)
    nlambda = calcSunApparentLong(nt)
    
    nsint = Sin(DegToRad(ne)) * Sin(DegToRad(nlambda))
    calcSunDeclination = RadToDeg(asin(nsint))
    
End Function

Private Function calcEquationOfTime(ByVal nt As Double) As Double

    Dim nepsilon As Double
    Dim nl0 As Double
    Dim ne As Double
    Dim nm As Double
    Dim ny As Double
    Dim nsin2l0 As Double
    Dim nsinm As Double
    Dim ncos2l0 As Double
    Dim nsin4l0 As Double
    Dim nsin2m As Double
    Dim nEtime As Double
    
    nepsilon = calcObliquityCorrection(nt)
    nl0 = calcGeomMeanLongSun(nt)
    ne = calcEccentricityEarthOrbit(nt)
    nm = calcGeomMeanAnomalySun(nt)
    
    ny = Math.Tan(DegToRad(nepsilon) / 2#)
    ny = ny * ny
    
    nsin2l0 = Sin(2# * DegToRad(nl0))
    nsinm = Sin(DegToRad(nm))
    ncos2l0 = Cos(2# * DegToRad(nl0))
    nsin4l0 = Sin(4# * DegToRad(nl0))
    nsin2m = Sin(2# * DegToRad(nm))
    
    nEtime = ny * nsin2l0 - 2# * ne * nsinm + 4# * ne * _
        ny * nsinm * ncos2l0 - 0.5 * ny * ny * nsin4l0 - _
        1.25 * ne * ne * nsin2m
    
    calcEquationOfTime = RadToDeg(nEtime) * 4#
End Function

Private Function calcHourAngleSunrise(ByVal nlat As Double, _
    ByVal nsolarDec As Double) As Double

    Dim nlatRad As Double
    Dim nsdRad As Double
    Dim nHAarg As Double
    Dim nHA As Double

    nlatRad = DegToRad(nlat)
    nsdRad = DegToRad(nsolarDec)

    nHAarg = (Cos(DegToRad(90.833)) / (Cos(nlatRad) * _
        Cos(nsdRad)) - Tan(nlatRad) * Tan(nsdRad))

    Dim nTemp As Double
    nTemp = Cos(DegToRad(90.833)) / (Cos(nlatRad) * _
        Cos(nsdRad)) - Tan(nlatRad) * Tan(nsdRad)
    If Abs(nTemp) > 1 Then
        nHA = -999
    Else
        nHA = (acos(nTemp))
    End If

    calcHourAngleSunrise = nHA

End Function

Private Function calcHourAngleSunset(ByVal nlat As Double, _
    ByVal nsolarDec As Double) As Double

    Dim nlatRad As Double
    Dim nsdRad As Double
    Dim nHAarg As Double
    Dim nHA As Double

    nlatRad = DegToRad(nlat)
    nsdRad = DegToRad(nsolarDec)
    
    nHAarg = (Cos(DegToRad(90.833)) / (Cos(nlatRad) * _
        Cos(nsdRad)) - Tan(nlatRad) * Tan(nsdRad))
    
    Dim nTemp As Double
    nTemp = Cos(DegToRad(90.833)) / (Cos(nlatRad) * _
        Cos(nsdRad)) - Tan(nlatRad) * Tan(nsdRad)
    If Abs(nTemp) > 1 Then
        nHA = 999
    Else
        nHA = (acos(nTemp))
    End If
    
    calcHourAngleSunset = -nHA

End Function

Private Function calcSunriseUTC(ByVal njd As Double, _
    ByVal nLatitude As Double, ByVal nLongitude As Double) As Double

    Dim nt As Double
    Dim neqTime As Double
    Dim nsolarDec As Double
    Dim nhourAngle As Double
    
    Dim ndelta As Double
    Dim ntimeDiff As Double
    Dim ntimeUTC As Double
    
    nt = calcTimeJulianCent(njd)

    neqTime = calcEquationOfTime(nt)
    nsolarDec = calcSunDeclination(nt)
    nhourAngle = calcHourAngleSunrise(nLatitude, nsolarDec)
    If nhourAngle = -999 Then
        calcSunriseUTC = -999
        Exit Function
    End If

    ndelta = nLongitude - RadToDeg(nhourAngle)
    ntimeDiff = 4 * ndelta
    ntimeUTC = 720 + ntimeDiff - neqTime

    Dim nnewt As Double
    nnewt = calcTimeJulianCent(calcJDFromJulianCent(nt) + _
        ntimeUTC / 1440#)
    neqTime = calcEquationOfTime(nnewt)
    nsolarDec = calcSunDeclination(nnewt)
    nhourAngle = calcHourAngleSunrise(nLatitude, nsolarDec)
    If nhourAngle = -999 Then
        calcSunriseUTC = -999
        Exit Function
    End If
    ndelta = nLongitude - RadToDeg(nhourAngle)
    ntimeDiff = 4 * ndelta
    ntimeUTC = 720 + ntimeDiff - neqTime

    calcSunriseUTC = ntimeUTC
    
End Function

Private Function calcSolNoonUTC(ByVal nt As Double, _
    ByVal nLongitude As Double) As Double

    Dim nnewt As Double
    Dim neqTime As Double
    Dim nsolarNoonDec As Double
    Dim nsolNoonUTC As Double
    
    nnewt = calcTimeJulianCent(calcJDFromJulianCent(nt) + _
        0.5 + nLongitude / 360#)
    neqTime = calcEquationOfTime(nt)
    nsolarNoonDec = calcSunDeclination(nt)
    nsolNoonUTC = 720 + (nLongitude * 4) - neqTime
    
    calcSolNoonUTC = nsolNoonUTC

End Function

Private Function calcSunsetUTC(ByVal njd As Double, _
    ByVal nLatitude As Double, ByVal nLongitude As _
    Double) As Double

    Dim neqTime As Double
    Dim nsolarDec As Double
    Dim nhourAngle As Double
    
    Dim ndelta As Double
    Dim ntimeDiff As Double
    Dim ntimeUTC As Double
    Dim nnewt As Double
    Dim nt As Double

    nt = calcTimeJulianCent(njd)

    neqTime = calcEquationOfTime(nt)
    nsolarDec = calcSunDeclination(nt)
    nhourAngle = calcHourAngleSunset(nLatitude, nsolarDec)
    If nhourAngle = -999 Then
        calcSunsetUTC = -999
        Exit Function
    End If

    ndelta = nLongitude - RadToDeg(nhourAngle)
    ntimeDiff = 4 * ndelta
    ntimeUTC = 720 + ntimeDiff - neqTime
    
    nnewt = calcTimeJulianCent(calcJDFromJulianCent(nt) + _
        ntimeUTC / 1440#)
    neqTime = calcEquationOfTime(nnewt)
    nsolarDec = calcSunDeclination(nnewt)
    nhourAngle = calcHourAngleSunset(nLatitude, nsolarDec)
    If nhourAngle = -999 Then
        calcSunsetUTC = -999
        Exit Function
    End If
    
    ndelta = nLongitude - RadToDeg(nhourAngle)
    ntimeDiff = 4 * ndelta
    ntimeUTC = 720 + ntimeDiff - neqTime
    
    calcSunsetUTC = ntimeUTC

End Function

Private Function findRecentSunrise(ByVal njd As Double, _
    ByVal nLatitude As Double, ByVal nLongitude _
    As Double) As Double

    Dim njulianday As Double
    njulianday = njd
    Dim nBail As Long
    
    Dim ntime As Double
    ntime = calcSunriseUTC(njulianday, nLatitude, nLongitude)
    Do While ntime = -999 And nBail < 367
        nBail = nBail + 1
        njulianday = njulianday - 1
        ntime = calcSunriseUTC(njulianday, nLatitude, nLongitude)
    Loop
    
    findRecentSunrise = njulianday

End Function

Private Function findRecentSunset(ByVal njd As Double, _
    ByVal nLatitude As Double, ByVal nLongitude _
    As Double) As Double

    Dim njulianday As Double
    Dim ntime As Double
    Dim nBail As Long
    
    njulianday = njd

    ntime = calcSunsetUTC(njulianday, nLatitude, nLongitude)
    Do While ntime = -999 And nBail < 367
        nBail = nBail + 1
        njulianday = njulianday - 1
        ntime = calcSunsetUTC(njulianday, nLatitude, nLongitude)
    Loop
    
    findRecentSunset = njulianday

End Function

Private Function findNextSunrise(ByVal njd As Double, ByVal _
    nLatitude As Double, ByVal nLongitude As Double) As Double
    Dim njulianday As Double
    Dim ntime As Double
    Dim nBail As Long

    njulianday = njd

    ntime = calcSunriseUTC(njulianday, nLatitude, nLongitude)
    Do While ntime = -999 And nBail < 367
        nBail = nBail + 1
        njulianday = njulianday + 1
        ntime = calcSunriseUTC(njulianday, nLatitude, nLongitude)
    Loop
    
    findNextSunrise = njulianday

End Function

Private Function findNextSunset(ByVal njd As Double, ByVal _
    nLatitude As Double, ByVal nLongitude As Double) As Double

    Dim njulianday As Double
    Dim ntime As Double
    Dim nBail As Long
    njulianday = njd
    ntime = calcSunsetUTC(njulianday, nLatitude, nLongitude)
    Do While ntime = -999 And nBail < 367
        nBail = nBail + 1
        njulianday = njulianday + 1
        ntime = calcSunsetUTC(njulianday, nLatitude, nLongitude)
    Loop
    
    findNextSunset = njulianday

End Function

Public Function CalculateSun()
    Dim nLatitude As Double
    Dim nLongitude As Double
    nLatitude = m_nLatitude
    nLongitude = m_nLongitude

    If nLatitude >= -90 And nLatitude < -89.5 Then
        nLatitude = -89.5
    End If
    If nLatitude <= 90 And nLatitude > 89.8 Then
        nLatitude = 89.8
    End If

    Dim njd As Double
    Dim ndoy As Double
    Dim nt As Double
    
    njd = calcJD(Year(m_dateSel), Month(m_dateSel), Day(m_dateSel))
    ndoy = calcDayOfYear(Month(m_dateSel), Day(m_dateSel), _
        IsLeapYear(Year(m_dateSel)))
    nt = calcTimeJulianCent(njd)
    
    Dim nAlpha As Double
    Dim nTheta As Double
    Dim nEtime As Double
    
    nAlpha = calcSunRtAscension(nt)
    nTheta = calcSunDeclination(nt)
    nEtime = calcEquationOfTime(nt)
    
    Dim neqTime As Double
    Dim nsolarDec As Double
    
    neqTime = nEtime
    nsolarDec = nTheta
    
    'Calculate sunrise
    Dim bNoSunrise As Boolean
    bNoSunrise = False
    
    Dim nRiseTimeGMT As Double
    nRiseTimeGMT = calcSunriseUTC(njd, nLatitude, nLongitude)
    
    If nRiseTimeGMT = -999 Then
        bNoSunrise = True
    End If

    'Calculate sunset
    Dim bNoSunset As Boolean
    bNoSunset = False
    Dim nSetTimeGMT As Double
    nSetTimeGMT = calcSunsetUTC(njd, nLatitude, nLongitude)
    If nSetTimeGMT = -999 Then
        bNoSunset = True
    End If
    
    Dim ndaySavings As Double
    Dim nZone As Double
    
    ndaySavings = IIf(m_bDaySavings, 60, 0)
    nZone = m_nTimeZone
    If nZone > 12 Or nZone < -12.5 Then
        nZone = 0
    End If
    
    If Not bNoSunrise Then
        Dim nriseTimeLST As Double
        nriseTimeLST = nRiseTimeGMT - (60 * nZone) + ndaySavings
        
        m_dateSunrise = DateAdd("s", nriseTimeLST * 60, _
            Int(m_dateSel))
    End If

    If Not bNoSunset Then
        Dim nsetTimeLST As Double
        nsetTimeLST = nSetTimeGMT - (60 * nZone) + ndaySavings
        
        m_dateSunset = DateAdd("s", nsetTimeLST * 60, Int(m_dateSel))
    End If
    
    'Calculate solar noon for this date
    Dim nsolNoonGMT As Double
    Dim nsolNoonLST As Double
    nsolNoonGMT = calcSolNoonUTC(nt, nLongitude)
    nsolNoonLST = nsolNoonGMT - (60 * nZone) + ndaySavings
    
    m_dateNoon = DateAdd("s", nsolNoonLST * 60, Int(m_dateSel))

    Dim nnewjd As Double
    Dim nnewtime As Double

    If bNoSunrise Then
        If ((nLatitude > 66.4) And (ndoy > 79) And _
            (ndoy < 267)) Or ((nLatitude < -66.4) And _
            ((ndoy < 83) Or (ndoy > 236))) Then

            nnewjd = findRecentSunrise(njd, nLatitude, nLongitude)
            nnewtime = calcSunriseUTC(nnewjd, nLatitude, _
                nLongitude) - (60 * nZone) + ndaySavings

            If nnewtime > 1440 Then
                nnewtime = nnewtime - 1440
                nnewjd = nnewjd + 1
            End If
            If nnewtime < 0 Then
                nnewtime = nnewtime + 1440
                nnewjd = nnewjd - 1
            End If
            
            m_dateSunrise = DateAdd("s", nnewtime * 60, _
                Int(m_dateSel))
            m_dateSunrise = DateAdd("d", nnewjd - njd, m_dateSunrise)
        
        ElseIf ((nLatitude > 66.4) And ((ndoy < 83) Or _
            (ndoy > 263))) Or ((nLatitude < -66.4) And _
            (ndoy > 79) And (ndoy < 267)) Then

            nnewjd = findNextSunrise(njd, nLatitude, nLongitude)
            nnewtime = calcSunriseUTC(nnewjd, nLatitude, _
                nLongitude) - (60 * nZone) + ndaySavings
            If nnewtime > 1440 Then
                nnewtime = nnewtime - 1440
                nnewjd = nnewjd + 1
            End If
            If nnewtime < 0 Then
                nnewtime = nnewtime + 1440
                nnewjd = nnewjd - 1
            End If
            
            m_dateSunrise = DateAdd("s", nnewtime * 60, _
                Int(m_dateSel))
            m_dateSunrise = DateAdd("d", nnewjd - njd, m_dateSunrise)
        
        End If
    End If
    
    If bNoSunset Then
        If (((nLatitude > 66.4) And (ndoy > 79) And _
            (ndoy < 267)) Or ((nLatitude < -66.4) And _
            ((ndoy < 83) Or (ndoy > 263)))) Then

            nnewjd = findNextSunset(njd, nLatitude, nLongitude)
            nnewtime = calcSunsetUTC(nnewjd, nLatitude, _
                nLongitude) - (60 * nZone) + ndaySavings
            If nnewtime > 1440 Then
                nnewtime = nnewtime - 1440
                nnewjd = nnewjd + 1
            End If
            If nnewtime < 0 Then
                nnewtime = nnewtime + 1440
                nnewjd = nnewjd - 1
            End If

            m_dateSunset = DateAdd("s", nnewtime * 60, Int(m_dateSel))
            m_dateSunset = DateAdd("d", nnewjd - njd, m_dateSunset)

        ElseIf (((nLatitude > 66.4) And ((ndoy < 83) Or _
            (ndoy > 263))) Or ((nLatitude < -66.4) And _
            (ndoy > 79) And (ndoy < 267))) Then

            nnewjd = findRecentSunset(njd, nLatitude, nLongitude)
            nnewtime = calcSunsetUTC(nnewjd, nLatitude, _
                nLongitude) - (60 * nZone) + ndaySavings
            If nnewtime > 1440 Then
                nnewtime = nnewtime - 1440
                nnewjd = nnewjd + 1
            End If
            If nnewtime < 0 Then
                nnewtime = nnewtime + 1440
                nnewjd = nnewjd - 1
            End If

            m_dateSunset = DateAdd("s", nnewtime * 60, Int(m_dateSel))
            m_dateSunset = DateAdd("d", nnewjd - njd, m_dateSunset)

        End If
    End If
End Function

' ------ End of class clsSunRiseSet





