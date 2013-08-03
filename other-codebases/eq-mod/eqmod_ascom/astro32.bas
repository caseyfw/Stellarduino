Attribute VB_Name = "Astronomy_Funcs"
'---------------------------------------------------------------------
'
'   ===========
'   ASTRO32.BAS
'   ===========
'
' Interface declarations for the Astronomy Library. Drop this into
' any VB project to get access to the astronomical support functions
' in astro32.dll. For the latest copy of astro32.dll, contact the
' author at the address below.
'
' Routines in astronomy DLL have been taken from various open source
' and freeware applications as well as original code by the author.
' Astro32.dll and this VB module are freely usable in any software
' project. The author assumes no responsibilities for bugs, etc.
'
' Written:  18-Jul-96   Robert B. Denny <rdenny@dc3.com>
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' 18-Jul-96 rbd     Initial edit (yes, 1996!)
' 19-Jul-98 rbd     1.2 of astro32.dll, add deltat()
' 20-Jul-98 rbd     Change comments on now_lst, change interface to take
'                   longitude - West
' 10-Aug-98 rbd     tz_name now returns 0/1 indicaing whether DST is in effect
' -----------------------------------------------------------------------------
Option Explicit


' You know what this is!
Public Const PI = 3.14159265358979
' Ratio of from synodic (solar) to sidereal (stellar) rate
Public Const SIDRATE = 0.9972695677
' Seconds per Sidereal day
Public Const SPD = 86400#
Public Const SPSD = 86164.0905
'
' Modified Julian Date (MJD) calculations. The epoch for MJD is
'
Public Const MJD0 = 2415020#    ' MJD Julian epoch (JD = MJD + MJD0)
Public Const J2000 = 36525#     ' MJD for 2000 (2451545.0 - MJD0)
'
' Date formatting preferences for fmt_mjd() and scn_date()
'
Public Const DATE_YMD = 0
Public Const DATE_MDY = 1
Public Const DATE_DMY = 2

Public r(2, 2) As Double
Const DEG_RAD As Double = 0.0174532925
Const SEC_RAD As Double = DEG_RAD / 3600
Const RAD_DEG As Double = 57.2957795
Const HRS_RAD As Double = 0.2617993881
Const RAD_HRS As Double = 3.81971863


'
' Timezone name preferences for tz_name()
'
Public Const DATE_UTCTZ = 3
Public Const DATE_LOCALTZ = 4

'
' =================
' LIBRARY FUNCTIONS
' =================
'
' NOTES:
'
' (1) For whatever reason, the authors of the original C functions chose
'     to pass back and forth via parameters only for most functions.
'
' (2) The descriptive comments below were lifted straight out of the C
'     functions. Some variables are listed with the C dereferening '*'.
'     Note that these are passed ByRef in the declarations, then forget
'     about the '*'.
'
' (3) Modified Julian Dates (number of days elapsed since 1900 jan 0.5,)
'     are used for most times. Several functions are provided for converting
'     between mjd and other time systems (C runtime, VB, Win32).
'
'
' given latitude (n+, radians), lat, altitude (up+, radians), alt, and
' azimuth (angle around to the east from north+, radians),
' return hour angle (radians), ha, and declination (radians), dec.
'
Declare Sub aa_hadec Lib "astro32" (ByVal lat As Double, ByVal Alt As Double, ByVal Az As Double, ByRef ha As Double, ByRef DEC As Double)
'
' given a date in months, mn, days, dy, years, yr,
' return the modified Julian date (number of days elapsed since 1900 jan 0.5),
' *mjd.
'
Declare Sub cal_mjd Lib "astro32" (ByVal mn As Long, ByVal dy As Double, ByVal yr As Long, ByRef mjd As Double)
'
' given the difference in two RA's, in rads, return their difference,
'   accounting for wrap at 2*PI. caller need *not* first force it into the
'   range 0..2*PI.
'
Declare Function delra Lib "astro32" (ByVal dRA As Double) As Double
'
' given the modified Julian date, mjd, find delta-T (TT-UTC)
'
Declare Function delta_t Lib "astro32" Alias "deltat" (ByVal mjd As Double) As Double
'
' Format a date string into buf, given a modified julian date and the
' selected format (m/d/y, etc.). Typically mm/dd.ddd/yyyy (note the
' fractional days).
'
Declare Sub fmt_mjd Lib "astro32" (ByVal buf As String, ByVal mjd As Double, ByVal pref As Long)
'
' format the Double (e.g., mjd, lst) in sexagesimal format into buf[].
' w is the number of spaces for the whole part.
' fracbase is the number of pieces a whole is to broken into; valid options:
'  360000: <w>:mm:ss.ss
'  36000:  <w>:mm:ss.s
'  3600:   <w>:mm:ss
'  600:    <w>:mm.m
'  60:     <w>:mm
'
Declare Sub fmt_sexa Lib "astro32" (ByVal buf As String, ByVal val As Double, ByVal w As Long, ByVal fracbase As Long)
'
' given a modified julian date, mjd, and a greenwich mean siderial time, gst,
' return universally coordinated time, *utc.
'
Declare Sub gst_utc Lib "astro32" (ByVal mjd As Double, ByVal gst As Double, ByRef utc As Double)
'
' given latitude (n+, radians), lat, hour angle (radians), ha, and declination
'   (radians), dec, return altitude (up+, radians), alt, and azimuth (angle
'   round to the east from north+, radians),
'
Declare Sub hadec_aa Lib "astro32" (ByVal lat As Double, ByVal ha As Double, ByVal DEC As Double, ByRef Alt As Double, ByRef Az As Double)
'
' Convert "MM/DD/YY" to VB Date
'
Declare Function mdy_vb Lib "astro32" (ByVal mdy As String) As Date
'
' return the Modified Julian Date of the epoch 2000
'
Declare Function mjd_2000 Lib "astro32" () As Double
'
' given the modified Julian date, mjd, return the calendar date in months, *mn,
' days, *dy, and years, *yr.
'
Declare Sub mjd_cal Lib "astro32" (ByVal mjd As Double, ByRef mn As Long, ByRef dy As Double, ByRef yr As Long)
'
' given an mjd, truncate it to the beginning of the whole day
'
Declare Function mjd_day Lib "astro32" (ByVal jd As Double) As Double
'
' given an mjd, set *dow to 0..6 according to which day of the week it falls
' on (0=sunday). return 0 if ok else -1 if can't figure it out.
'
Declare Function mjd_dow Lib "astro32" (ByVal mjd As Double, ByRef dow As Long) As Long
'
' given a mjd, return the the number of days in the month.
'
Declare Sub mjd_dpm Lib "astro32" (ByVal mjd As Double, ByRef ndays As Long)
'
' given an mjd, return the number of hours past midnight of the
' whole day
'
Declare Function mjd_hr Lib "astro32" (ByVal jd As Double) As Double
'
' Return the Visual Basic Date given a Modified Julian Date
'
Declare Function mjd_vb Lib "astro32" (ByVal mjd As Double) As Date
'
' given a mjd, return the year as a double.
'
Declare Sub mjd_year Lib "astro32" (ByVal mjd As Double, ByRef yr As Double)
'
' Return the current Local Apparent Sidereal Time (LST) from the clock and longitude (rad, - west)
'
Declare Function now_lst Lib "astro32" (ByVal lng As Double) As Double
'
' Return the current Modified Julian Date derived from the system clock
'
Declare Function now_mjd Lib "astro32" () As Double
'
' given the modified JD, mjd, correct, IN PLACE, the right ascension *ra
' and declination *dec (both in radians) for nutation.
'
Declare Sub nut_eq Lib "astro32" (ByVal mjd As Double, ByRef RA As Double, ByRef DEC As Double)
'
' given the modified JD, mjd, find the nutation in obliquity, *deps, and
' the nutation in longitude, *dpsi, each in radians.
'
Declare Sub nut Lib "astro32" Alias "nutation" (ByVal mjd As Double, ByRef deps As Double, ByRef dpsi As Double)
'
' given the modified Julian date, mjd, find the mean obliquity of the
' ecliptic, *eps, in radians.
'
Declare Sub obliq Lib "astro32" Alias "obliquity" (ByVal mjd As Double, ByRef eps As Double)
'
' insure 0 <= *v < r. Used to range angles and times
'
Declare Sub range Lib "astro32" (ByRef v As Double, ByVal r As Double)
'
' correct the true altitude, ta, for refraction to the apparent altitude, aa,
' each in radians, given the local atmospheric pressure, pr, in mbars, and
' the temperature, tr, in degrees C.
'
Declare Sub refract Lib "astro32" (ByVal pr As Double, ByVal tr As Double, ByVal ta As Double, ByRef aa As Double)
'
' crack a floating date string, bp, of the form X/Y/Z determined by the
'   DATE_DATE_FORMAT preference into its components. allow the day to be a
'   floating point number. A lone component is always a year if it contains
'   a decimal point or pref is MDY or DMY and it is not a reasonable month
'   or day value, respectively. Leave any unspecified component unchanged.
'   ( actually, the slashes may be anything but digits or a decimal point)
'   'pref' indicates the format of the date (DATE_xxx).
'
Declare Function scn_date Lib "astro32" (ByVal dtstr As String, ByRef m As Long, ByRef d As Double, ByRef Y As Long, ByVal pref As Long)
'
' scan a sexagesimal string and update a double. the string, bp, is of the form
'   H:M:S. a negative value may be indicated by a '-' char before any
'   component. All components may be integral or real. In addition to ':' the
'   separator may also be '/' or ';' or ',' or '-'.
' any components not specified in bp[] are copied from old's in 'o'.
'   eg:  ::10  only changes S
'        10    only changes H
'        10:0  changes H and M
'
Declare Function scn_sexa Lib "astro32" (ByVal o As Double, ByVal sexa As String) As Double
'
' round a time in days, *t, to the nearest second, IN PLACE.
'
Declare Sub rnd_second Lib "astro32" (ByRef t As Double)
'
' Fill buffer with the name of the current timezone, given a preference
' flag, pref, (DATE_UTCTZ = always "UTC", DATE_LOCALTZ = e.g., "PDT")
' Returns 0/1 indicating whether DST is currently in effect.
'
Declare Function tz_name Lib "astro32" (ByVal buf As String, ByVal pref As Long) As Long
'
' correct the apparent altitude, aa, for refraction to the true altitude, ta,
' each in radians, given the local atmospheric pressure, pr, in mbars, and
' the temperature, tr, in degrees C.
'
Declare Sub unrefract Lib "astro32" (ByVal pr As Double, ByVal tr As Double, ByVal aa As Double, ByRef ta As Double)
'
' given a modified julian DATE, mjd, and a universally coordinated time, utc,
' return greenwich mean siderial time, *gst.
' NOTE: mjd must be at the beginning of the day!
'
Declare Sub utc_gst Lib "astro32" (ByVal mjd As Double, ByVal utc As Double, ByRef gst As Double)
'
' Return the current UTC offset (+ = West) in seconds
'
Declare Function utc_offs Lib "astro32" () As Long
'
' Return a Modified Julian Date given a Visual Basic Date
'
Declare Function vb_mjd Lib "astro32" (ByVal d As Date) As Double
'
' given a decimal year, return mjd
'
Declare Sub year_mjd Lib "astro32" (ByVal Y As Double, ByRef mjd As Double)

'---------------------------------------------------------------------
'   Degrees to Radians
'---------------------------------------------------------------------
Public Function degrad(d As Double) As Double
    degrad = (d * PI) / 180#
End Function

'---------------------------------------------------------------------
'   Radians to Degrees
'---------------------------------------------------------------------
Public Function raddeg(r As Double) As Double
    raddeg = (r * 180#) / PI
End Function

'---------------------------------------------------------------------
'   Hours to Degrees
'---------------------------------------------------------------------
Public Function hrdeg(h As Double)
    hrdeg = h * 15#
End Function

'---------------------------------------------------------------------
'   Degrees to Hours
'---------------------------------------------------------------------
Public Function deghr(d As Double) As Double
    deghr = d / 15#
End Function

'---------------------------------------------------------------------
'   Hours to Radians
'---------------------------------------------------------------------
Public Function hrrad(h As Double) As Double
    hrrad = degrad(hrdeg(h))
End Function

'---------------------------------------------------------------------
'   Radians to Hours
'---------------------------------------------------------------------
Public Function radhr(r As Double) As Double
    radhr = deghr(raddeg(r))
End Function

Public Sub Precess(ByRef RA As Double, ByRef DEC As Double, equinox1 As Double, equinox2 As Double)

' INPUT - OUTPUT:
'      RA - Input right ascension (scalar or vector) in DEGREES
'      DEC - Input declination in DEGREES (scalar or vector)
'
'      The input RA and DEC are modified by PRECESS to give the
'      values after precession.
'
' RESTRICTIONS:
'       Accuracy of precession decreases for declination values near 90
'       degrees.  PRECESS should not be used more than 2.5 centuries from
'       2000 on the FK5 system (1950.0 on the FK4 system).
'
' EXAMPLES:
'       (1) The Pole Star has J2000.0 coordinates (2h, 31m, 46.3s,
'               89d 15' 50.6"); compute its coordinates at J1985.0
'
'       IDL> precess, ten(2,31,46.3)*15, ten(89,15,50.6), 2000, 1985, /PRINT
'
'               ====> 2h 16m 22.73s, 89d 11' 47.3"
'
'       (2) Precess the B1950 coordinates of Eps Ind (RA = 21h 59m,33.053s,
'       DEC = (-56d, 59', 33.053") to equinox B1975.
'
'       IDL> ra = ten(21, 59, 33.053)*15
'       IDL> dec = ten(-56, 59, 33.053)
'       IDL> precess, ra, dec ,1950, 1975, /fk4
'
' PROCEDURE:
'       Algorithm from Computational Spherical Astronomy by Taff (1983),
'       p. 24. (FK4). FK5 constants from "Astronomical Almanac Explanatory
'       Supplement 1992, page 104 Table 3.211.1.
'
' PROCEDURE CALLED:
'       Function PREMAT - computes precession matrix
'
    Dim X As Double
    Dim Y As Double
    Dim z As Double
    Dim x1 As Double
    Dim y1 As Double
    Dim z1 As Double
    Dim tmp As Double
    Dim ra_rad As Double
    Dim dec_rad As Double
    Dim ra_in As Double
    Dim dec_in As Double

    ra_in = RA

    ' switch to radians
    ra_rad = RA * 15 * DEG_RAD
    dec_rad = DEC * DEG_RAD
    tmp = RA * 15

    X = Cos(ra_rad) * Cos(dec_rad)
    Y = Sin(ra_rad) * Cos(dec_rad)
    z = Sin(dec_rad)

    ' initailaise the precession matrix from Equinox1 to Equinox2
    Call premat(equinox1, equinox2)

    x1 = r(0, 0) * X + r(1, 0) * Y + r(2, 0) * z
    y1 = r(0, 1) * X + r(1, 1) * Y + r(2, 1) * z
    z1 = r(0, 2) * X + r(1, 2) * Y + r(2, 2) * z
    
    ra_rad = Atn(y1 / x1)
    dec_rad = ArcSin(z1)
    
    ' apply nuation
    Call nut_eq(now_mjd(), ra_rad, dec_rad)
    
    DEC = dec_rad / DEG_RAD
    RA = ra_rad / DEG_RAD
    
    If x1 < 0 Then RA = RA + 180
    If y1 < 0 And x1 > 0 Then RA = RA + 360
    
    RA = RA / 15
   
    If RA > 24 Then
        RA = RA - 24
    End If


End Sub

Public Function ArcSin(X As Double) As Double
    ArcSin = Atn(X / Sqr(-X * X + 1))
End Function

Public Function ArcCos(X As Double) As Double
    ArcCos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
End Function

Public Sub premat(equinox1 As Double, equinox2 As Double)
    
    Dim t As Double
    Dim st As Double
    Dim a As Double
    Dim b As Double
    Dim c As Double
    Dim sina As Double
    Dim sinb As Double
    Dim sinc As Double
    Dim cosa As Double
    Dim cosb As Double
    Dim cosc As Double

     t = 0.001 * (equinox2 - equinox1)
     st = 0.001 * (equinox1 - 2000)
    
    'Compute 3 rotation angles
    a = SEC_RAD * t * (23062.181 + st * (139.656 + 0.0139 * st) + t * (30.188 - 0.344 * st + 17.998 * t))
    b = SEC_RAD * t * t * (79.28 + 0.41 * st + 0.205 * t) + a
    c = SEC_RAD * t * (20043.109 - st * (85.33 + 0.217 * st) + t * (-42.665 - 0.217 * st - 41.833 * t))
    
    sina = Sin(a)
    sinb = Sin(b)
    sinc = Sin(c)
    cosa = Cos(a)
    cosb = Cos(b)
    cosc = Cos(c)
    
    r(0, 0) = cosa * cosb * cosc - sina * sinb
    r(0, 1) = sina * cosb + cosa * sinb * cosc
    r(0, 2) = cosa * sinc
    r(1, 0) = -cosa * sinb - sina * cosb * cosc
    r(1, 1) = cosa * cosb - sina * sinb * cosc
    r(1, 2) = -sina * sinc
    r(2, 0) = -cosb * sinc
    r(2, 1) = -sinb * sinc
    r(2, 2) = cosc
    
End Sub


