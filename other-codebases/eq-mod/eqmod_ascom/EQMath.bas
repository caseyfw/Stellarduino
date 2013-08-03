Attribute VB_Name = "EQMath"
'---------------------------------------------------------------------
' Copyright © 2006 Raymund Sarmiento
'
' Permission is hereby granted to use this Software for any purpose
' including combining with commercial products, creating derivative
' works, and redistribution of source or binary code, without
' limitation or consideration. Any redistributed copies of this
' Software must include the above Copyright Notice.
'
' THIS SOFTWARE IS PROVIDED "AS IS". THE AUTHOR OF THIS CODE MAKES NO
' WARRANTIES REGARDING THIS SOFTWARE, EXPRESS OR IMPLIED, AS TO ITS
' SUITABILITY OR FITNESS FOR A PARTICULAR PURPOSE.
'---------------------------------------------------------------------
'
' EQMATH.bas - Math functions for EQMOD ASCOM RADECALTAZ computations
'
'
' Written:  07-Oct-06   Raymund Sarmiento
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' 04-Nov-06 rcs     Initial edit for EQ Mount Driver Function Prototype
' 20-Nov-06 rcs     wrote a new function for now_lst that will generate millisecond
'                   granularity
' 21-Nov-06 rcs     Append RA GOTO Compensation to minimize discrepancy
' 19-Mar-07 rcs     Initial Edit for Three star alignment
' 05-Apr-07 rcs     Add MAXSYNC
' 08-Apr-07 rcs     N-star implementation
'---------------------------------------------------------------------
'
'
'  SYNOPSIS:
'
'  This is a demonstration of a EQ6/ATLAS/EQG direct stepper motor control access
'  using the EQCONTRL.DLL driver code.
'
'  File EQCONTROL.bas contains all the function prototypes of all subroutines
'  encoded in the EQCONTRL.dll
'
'  The EQ6CONTRL.DLL simplifies execution of the Mount controller board stepper
'  commands.
'
'  The mount circuitry needs to be modified for this test program to work.
'  Circuit details can be found at http://www.freewebs.com/eq6mod/
'

'  DISCLAIMER:

'  You can use the information on this site COMPLETELY AT YOUR OWN RISK.
'  The modification steps and other information on this site is provided
'  to you "AS IS" and WITHOUT WARRANTY OF ANY KIND, express, statutory,
'  implied or otherwise, including without limitation any warranty of
'  merchantability or fitness for any particular or intended purpose.
'  In no event the author will  be liable for any direct, indirect,
'  punitive, special, incidental or consequential damages or loss of any
'  kind whether or not the author  has been advised of the possibility
'  of such loss.

'  WARNING:

'  Circuit modifications implemented on your setup could invalidate
'  any warranty that you may have with your product. Use this
'  information at your own risk. The modifications involve direct
'  access to the stepper motor controls of your mount. Any "mis-control"
'  or "mis-command"  / "invalid parameter" or "garbage" data sent to the
'  mount could accidentally activate the stepper motors and allow it to
'  rotate "freely" damaging any equipment connected to your mount.
'  It is also possible that any garbage or invalid data sent to the mount
'  could cause its firmware to generate mis-steps pulse sequences to the
'  motors causing it to overheat. Make sure that you perform the
'  modifications and testing while there is no physical "load" or
'  dangling wires on your mount. Be sure to disconnect the power once
'  this event happens or if you notice any unusual sound coming from
'  the motor assembly.
'
'  CREDITS:
'
'  Portions of the information on this code should be attributed
'  to Mr. John Archbold from his initial observations and analysis
'  of the interface circuits and of the ASCII data stream between
'  the Hand Controller (HC) and the Go To Controller.
'

Option Explicit

Public Const DEG_RAD As Double = 0.0174532925
Public Const RAD_DEG As Double = 57.2957795
Public Const HRS_RAD As Double = 0.2617993881
Public Const RAD_HRS As Double = 3.81971863

Public Const SID_RATE As Double = 15.041067
Public Const SOL_RATE As Double = 15
Public Const LUN_RATE As Double = 14.511415

Public Const gEMUL_RATE As Double = 20.98                   ' 0.2 * 9024000/( (23*60*60)+(56*60)+4)
                                                            ' 0.2 = 200ms
                                                
Public Const gEMUL_RATE2 As Double = 104.730403903004         ' (9024000/86164.0905)
                                                            
                                                            ' 104.73040390300411747513310083625
                                                            
Public Const gARCSECSTEP As Double = 0.144                  ' .144 arcesconds / step

' Iterative GOTO Constants
'Public Const NUM_SLEW_RETRIES As Long = 5                   ' Iterative MAX retries
Public Const gRA_Allowed_diff As Double = 10                ' Iterative Slew minimum difference


' Home Position of the mount (pointing at NCP/SCP)

Public Const RAEncoder_Home_pos As Double = &H800000        ' Start at 0 Hour
Public Const DECEncoder_Home_pos As Double = &HA26C80 ' Start at 90 Degree position

Public Const gRAEncoder_Zero_pos As Double = &H800000       ' ENCODER 0 Hour initial position
Public Const gDECEncoder_Zero_pos As Double = &H800000      ' ENCODER 0 Degree Initial position

Public Const gDefault_step As Double = 9024000              ' Total Encoder count (EQ5/6)



'Public Const EQ_MAXSYNC As Double = &H111700

' Public Const EQ_MAXSYNC_Const As Double = &H88B80                 ' Allow a 45 degree discrepancy


Public Const EQ_MAXSYNC_Const As Double = &H113640                 ' Allow a 45 degree discrepancy

'------------------------------------------------------------------------------------------------

' Define all Global Variables


Public gXshift  As Double
Public gYshift As Double
Public gXmouse As Double
Public gYmouse As Double


Public gEQ_MAXSYNC As Double                                ' Max Sync Diff
Public gSiderealRate As Double                              ' Sidereal rate arcsecs/sec
Public gMount_Ver As Double                                 ' Mount Version

Public gRA_LastRate As Double                               ' Last PEC Rate
Public gpl_interval As Integer                              ' Pulseguide Interval

Public eqres As Double
Public gTot_step As Double                                  ' Total Common RA-Encoder Steps
Public gTot_RA As Double                                    ' Total RA Encoder Steps
Public gTot_DEC As Double                                   ' Total DEC Encoder Steps
Public gRAWormSteps As Double                               ' Steps per RA worm revolution
Public gRAWormPeriod As Double                              ' Period of RA worm revolution
Public gDECWormSteps As Double                              ' Steps per DEC worm revolution
Public gDECWormPeriod As Double                             ' Period of DEC worm revolution

Public gLatitude As Double                                  ' Site Latitude
Public gLongitude As Double                                 ' Site Longitude
Public gElevation As Double                                 ' Site Elevation
Public gHemisphere As Long

Public gDECEncoder_Home_pos As Double                       ' DEC HomePos - Varies with different mounts

Public gRA_Encoder As Double                                ' RA Current Polled RA Encoder value
Public gDec_Encoder As Double                               ' DEC Current Polled Encoder value
Public gRA_Hours As Double                                  ' RA Encoder to Hour position
Public gDec_Degrees As Double                               ' DEC Encoder to Degree position Ranged to -90 to 90
Public gDec_DegNoAdjust As Double                           ' DEC Encoder to actual degree position
Public gRAStatus As Double                                  ' RA Polled Motor Status
Public gRAStatus_slew As Boolean                            ' RA motor tracking poll status
Public gDECStatus As Double                                 ' DEC Polloed motor status
 
Public gRA_Limit_East As Double                             ' RA Limit at East Side
Public gRA_Limit_West As Double                             ' RA Limit at West Side
 
Public gRA1Star As Double                                   ' Initial RA Alignment adjustment
Public gDEC1Star As Double                                  ' Initial DEC Alignment adjustment
 
Public gRASync01 As Double                                  ' Initial RA sync adjustment
Public gDECSync01 As Double                                 ' Initial DEC sync adjustment

Public gRA As Double
Public gDec As Double
Public gAlt As Double
Public gAz As Double
Public gha As Double
Public gSOP As Double

Public gPort As String
Public gBaud As Long
Public gTimeout As Long
Public gRetry As Long

Public gTrackingStatus As Long
Public gSlewStatus As Boolean

Public gRAMoveAxis_Rate As Double
Public gDECMoveAxis_Rate As Double


' Added for emulated Stepper Counters
Public gEmulRA As Double
Public gEmulDEC As Double
Public gEmulOneShot As Boolean
Public gEmulNudge As Boolean

Public gCurrent_time As Double
Public gLast_time As Double
Public gEmulRA_Init As Double

Public Enum PierSide2
    pierUnknown2 = -1
    PierEast2 = 0
    PierWest2 = 1
End Enum

Public gRAEncoderPolarHomeGoto As Long
Public gDECEncoderPolarHomeGoto As Long
Public gRAEncoderUNPark As Long
Public gDECEncoderUNPark As Long
Public gRAEncoderPark As Long
Public gDECEncoderPark As Long
Public gRAEncoderlastpos As Long
Public gDECEncoderlastpos As Long
Public gEQparkstatus As Long

Public gEQRAPulseDuration As Long
Public gEQDECPulseDuration As Long
Public gEQPulsetimerflag As Boolean

Public gEQTimeDelta As Double


Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

' Public variables for Custom Tracking rates

Public gDeclinationRate As Double
Public gRightAscensionRate As Double


' Public Variables for Spiral Slew

Public gSPIRAL_JUMP As Long
Public gDeclination_Start As Double
Public gRightAscension_Start As Double
Public gDeclination_Dir As Double
Public gRightAscension_Dir As Double
Public gDeclination_Len As Long
Public gRightAscension_Len As Long

Public gSpiral_AxisFlag As Double



' Public variables for debugging

Public gAffine1 As Double
Public gAffine2 As Double
Public gAffine3 As Double

Public gTaki1 As Double
Public gTaki2 As Double
Public gTaki3 As Double


'Pulseguide Indicators

Public Const gMAX_plotpoints As Integer = 100
Public gMAX_RAlevel As Integer
Public gMAX_DEClevel As Integer
Public gPlot_ra_pos As Integer
Public gPlot_dec_pos As Integer
Public gplot_ra_cur As Double
Public gPlot_dec_cur As Double
Public gRAHeight As Double
Public gDecHeight As Double

' Polar Alignment Variables

Public gPolarAlign_RA As Double
Public gPolarAlign_DEC As Double

Public Declare Sub GetSystemTime Lib "kernel32" _
   (lpSystemTime As SYSTEMTIME)



Public Function Get_EncoderHours(encOffset0 As Double, encoderval As Double, Tot_enc As Double, hmspr As Long) As Double

Dim i As Double

    ' Compute in Hours the encoder value based on 0 position value (RAOffset0)
    ' and Total 360 degree rotation microstep count (Tot_Enc

    If encoderval > encOffset0 Then
        i = ((encoderval - encOffset0) / Tot_enc) * 24
        i = 24 - i
    Else
        i = ((encOffset0 - encoderval) / Tot_enc) * 24
    End If
    
    If hmspr = 0 Then
        Get_EncoderHours = Range24(i + 6#)       ' Set to true Hours which is perpendicula to RA Axis
    Else
        Get_EncoderHours = Range24((24 - i) + 6#)
    End If

End Function

Public Function Get_EncoderfromHours(encOffset0 As Double, hourval As Double, Tot_enc As Double, hmspr As Long) As Long
    
    hourval = Range24(hourval - 6#)         ' Re-normalize from a perpendicular position
    If hmspr = 0 Then
        If (hourval < 12) Then
            Get_EncoderfromHours = encOffset0 - ((hourval / 24) * Tot_enc)
        Else
            Get_EncoderfromHours = (((24 - hourval) / 24) * Tot_enc) + encOffset0
        End If
    Else
        If (hourval < 12) Then
           Get_EncoderfromHours = ((hourval / 24) * Tot_enc) + encOffset0
        Else
            Get_EncoderfromHours = encOffset0 - (((24 - hourval) / 24) * Tot_enc)
        End If
    End If

End Function

Public Function Get_EncoderfromDegrees(encOffset0 As Double, degval As Double, Tot_enc As Double, Pier As Double, hmspr As Long) As Long

    If hmspr = 1 Then degval = 360 - degval
    If (degval > 180) And (Pier = 0) Then
        Get_EncoderfromDegrees = encOffset0 - (((360 - degval) / 360) * Tot_enc)
    Else
        Get_EncoderfromDegrees = ((degval / 360) * Tot_enc) + encOffset0
    End If

End Function


Public Function Get_EncoderDegrees(encOffset0 As Double, encoderval As Double, Tot_enc As Double, hmspr As Long) As Double

Dim i As Double

    ' Compute in Hours the encoder value based on 0 position value (EncOffset0)
    ' and Total 360 degree rotation microstep count (Tot_Enc

    If encoderval > encOffset0 Then
        i = ((encoderval - encOffset0) / Tot_enc) * 360
         Else
        i = ((encOffset0 - encoderval) / Tot_enc) * 360
        i = 360 - i
    End If

    If hmspr = 0 Then
        Get_EncoderDegrees = Range360(i)
    Else
        Get_EncoderDegrees = Range360(360 - i)
    End If
End Function

' Function that will ensure that the DEC value will be between -90 to 90
' Even if it is set at the other side of the pier

Public Function Range_DEC(decdegrees As Double) As Double

    If (decdegrees >= 270) And (decdegrees <= 360) Then
        Range_DEC = decdegrees - 360
        Exit Function
    End If
    
    If (decdegrees >= 180) And (decdegrees < 270) Then
        Range_DEC = 180 - decdegrees
        Exit Function
    End If
    
    If (decdegrees >= 90) And (decdegrees < 180) Then
        Range_DEC = 180 - decdegrees
        Exit Function
    End If
    
    Range_DEC = decdegrees

End Function



Public Function Get_RAEncoderfromRA(ra_in_hours As Double, dec_in_degrees As Double, pLongitude As Double, encOffset0 As Double, Tot_enc As Double, hmspr As Long) As Long

Dim i As Double
Dim j As Double

    i = ra_in_hours - EQnow_lst(pLongitude * DEG_RAD)
    
    If hmspr = 0 Then
        If (dec_in_degrees > 90) And (dec_in_degrees <= 270) Then i = i - 12#
    Else
        If (dec_in_degrees > 90) And (dec_in_degrees <= 270) Then i = i + 12#
    End If

    i = Range24(i)
    
    Get_RAEncoderfromRA = Get_EncoderfromHours(encOffset0, i, Tot_enc, hmspr)
   
End Function

Public Function Get_RAEncoderfromAltAz(Alt_in_deg As Double, Az_in_deg As Double, pLongitude As Double, pLatitude As Double, encOffset0 As Double, Tot_enc As Double, hmspr As Long) As Long

Dim i As Double
Dim ttha As Double
Dim ttdec As Double

    aa_hadec (pLatitude * DEG_RAD), (Alt_in_deg * DEG_RAD), ((360# - Az_in_deg) * DEG_RAD), ttha, ttdec
    i = (ttha * RAD_HRS)
    i = Range24(i)
    Get_RAEncoderfromAltAz = Get_EncoderfromHours(encOffset0, i, Tot_enc, hmspr)
   
End Function

Public Function Get_DECEncoderfromAltAz(Alt_in_deg As Double, Az_in_deg As Double, pLongitude As Double, pLatitude As Double, encOffset0 As Double, Tot_enc As Double, Pier As Double, hmspr As Long) As Long

Dim i As Double
Dim ttha As Double
Dim ttdec As Double

    aa_hadec (pLatitude * DEG_RAD), (Alt_in_deg * DEG_RAD), ((360# - Az_in_deg) * DEG_RAD), ttha, ttdec
    i = ttdec * RAD_DEG ' tDec was in Radians
    If Pier = 1 Then i = 180 - i
    Get_DECEncoderfromAltAz = Get_EncoderfromDegrees(encOffset0, i, Tot_enc, Pier, hmspr)
   
End Function

Public Function Get_DECEncoderfromDEC(dec_in_degrees As Double, Pier As Double, encOffset0 As Double, Tot_enc As Double, hmspr As Long) As Long

Dim i As Double

    i = dec_in_degrees
    If Pier = 1 Then i = 180 - i
    Get_DECEncoderfromDEC = Get_EncoderfromDegrees(encOffset0, i, Tot_enc, Pier, hmspr)
   
End Function

Public Function printhex(inpval As Double) As String

    printhex = " " & Hex$((inpval And &HF00000) / 1048576 And &HF) + Hex$((inpval And &HF0000) / 65536 And &HF) + Hex$((inpval And &HF000) / 4096 And &HF) + Hex$((inpval And &HF00) / 256 And &HF) + Hex$((inpval And &HF0) / 16 And &HF) + Hex$(inpval And &HF)

End Function

Public Function FmtSexa(ByVal N As Double, ShowPlus As Boolean) As String
    Dim sg As String
    Dim us As String
    Dim ms As String
    Dim ss As String
    Dim u As Long
    Dim m As Long
    Dim fmt
    
    sg = "+"                                ' Assume positive
    If N < 0 Then                           ' Check neg.
        N = -N                              ' Make pos.
        sg = "-"                            ' Remember sign
    End If

    m = Fix(N)                              ' Units (deg or hr)
    us = Format$(m, "00")

    N = (N - m) * 60#
    m = Fix(N)                              ' Minutes
    ms = Format$(m, "00")

    N = (N - m) * 60#
    m = Fix(N)                              ' Minutes
    ss = Format$(m, "00")

    FmtSexa = us & ":" & ms & ":" & ss
    If ShowPlus Or (sg = "-") Then FmtSexa = sg & FmtSexa
    
End Function
Public Function EQnow_lst(plong As Double) As Double

    Dim typTime As SYSTEMTIME
    Dim eps As Double
    Dim lst As Double
    Dim deps As Double
    Dim dpsi As Double
    Dim mjd As Double
    
'    mjd = vb_mjd(CDbl(Now) + gGPSTimeDelta)

    GetSystemTime typTime
    mjd = vb_mjd(CDbl(gEQTimeDelta + Now + (typTime.wMilliseconds / 86400000)))
    Call utc_gst(mjd_day(mjd), mjd_hr(mjd), lst)
    lst = lst + radhr(plong)
    Call obliq(mjd, eps)
    Call nut(mjd, deps, dpsi)
    lst = lst + radhr(dpsi * Cos(eps + deps))
    Call range(lst, 24#)

    EQnow_lst = lst
'    EQnow_lst = now_lst(plong)

End Function


Public Function EQnow_lst_norange() As Double

    Dim typTime As SYSTEMTIME
    Dim mjd As Double
    Dim MTMP As Double

    GetSystemTime typTime
    mjd = (typTime.wMinute * 60) + (typTime.wSecond) + (typTime.wMilliseconds / 1000)
    MTMP = (typTime.wHour)
    MTMP = MTMP * 3600
    mjd = mjd + MTMP + (typTime.wDay * 86400)
 
    EQnow_lst_norange = mjd

End Function


Public Function EQnow_lst_time(plong As Double, ptime As Double) As Double

    Dim eps As Double
    Dim lst As Double
    Dim deps As Double
    Dim dpsi As Double
    Dim mjd As Double
    
    mjd = vb_mjd(ptime)
    Call utc_gst(mjd_day(mjd), mjd_hr(mjd), lst)
    lst = lst + radhr(plong)
    Call obliq(mjd, eps)
    Call nut(mjd, deps, dpsi)
    lst = lst + radhr(dpsi * Cos(eps + deps))
    Call range(lst, 24#)

    EQnow_lst_time = lst

End Function

Public Function SOP_Physical(vha As Double) As PierSide2
Dim ha As Double
    
    ha = RangeHA(vha - 6#)
    SOP_Physical = IIf(ha >= 0, PierEast2, PierWest2)
 
End Function

Public Function SOP_Pointing(ByVal DEC As Double) As PierSide2
    
    If DEC <= 90 Or DEC >= 270 Then
        SOP_Pointing = PierEast2
    Else
        SOP_Pointing = PierWest2
    End If
 
End Function
Public Function SOP_RA(vRA As Double, pLongitude As Double) As PierSide2
Dim i As Double

    i = vRA - EQnow_lst(pLongitude * DEG_RAD)
    i = RangeHA(i - 6#)
    SOP_RA = IIf(i < 0, PierEast2, PierWest2)

End Function

Public Function Range24(ByVal vha As Double)
    
    While vha < 0#
        vha = vha + 24#
    Wend
    While vha >= 24#
        vha = vha - 24#
    Wend
    
    Range24 = vha
    
End Function

Public Function Range360(ByVal vdeg As Double)
    
    While vdeg < 0#
        vdeg = vdeg + 360#
    Wend
    While vdeg >= 360#
        vdeg = vdeg - 360#
    Wend
    
    Range360 = vdeg
    
End Function
Public Function Range90(ByVal vdeg As Double)
    
    While vdeg < -90#
        vdeg = vdeg + 360#
    Wend
    While vdeg >= 360#
        vdeg = vdeg - 90#
    Wend
    
    Range90 = vdeg
    
End Function

Public Function RangeHA(ByVal ha As Double)
    
    While ha < -12#
        ha = ha + 24#
    Wend
    While ha >= 12#
        ha = ha - 24#
    Wend
    
    RangeHA = ha
    
End Function

Public Function GetSlowdown(ByVal deltaval As Double) As Double
Dim i As Double
     
     i = deltaval - 80000
     If i < 0 Then i = deltaval * 0.5
     GetSlowdown = i

End Function

Public Function Delta_RA_Map(ByVal RAENCODER As Double) As Double
    
    Delta_RA_Map = RAENCODER + gRA1Star + gRASync01

End Function

Public Function Delta_DEC_Map(ByVal DecEncoder As Double) As Double

    Delta_DEC_Map = DecEncoder + gDEC1Star + gDECSync01

End Function


Public Function Delta_Matrix_Map(ByVal RA As Double, ByVal DEC As Double) As Coordt
Dim i As Integer
Dim obtmp As Coord
Dim obtmp2 As Coord

    If (RA >= &H1000000) Or (DEC >= &H1000000) Then
        Delta_Matrix_Map.x = RA
        Delta_Matrix_Map.Y = DEC
        Delta_Matrix_Map.z = 1
        Delta_Matrix_Map.F = 0
        Exit Function
    End If
    
    obtmp.x = RA
    obtmp.Y = DEC
    obtmp.z = 1
    
    ' re transform based on the nearest 3 stars
    i = EQ_UpdateTaki(RA, DEC)
    
    obtmp2 = EQ_plTaki(obtmp)

    Delta_Matrix_Map.x = obtmp2.x
    Delta_Matrix_Map.Y = obtmp2.Y
    Delta_Matrix_Map.z = 1
    Delta_Matrix_Map.F = i
    
End Function


Public Function Delta_Matrix_Reverse_Map(ByVal RA As Double, ByVal DEC As Double) As Coordt

Dim i As Integer
Dim obtmp As Coord
Dim obtmp2 As Coord

    If (RA >= &H1000000) Or (DEC >= &H1000000) Then
        Delta_Matrix_Reverse_Map.x = RA
        Delta_Matrix_Reverse_Map.Y = DEC
        Delta_Matrix_Reverse_Map.z = 1
        Delta_Matrix_Reverse_Map.F = 0
        Exit Function
    End If
    
    obtmp.x = RA + gRASync01
    obtmp.Y = DEC + gDECSync01
    obtmp.z = 1
    
    ' re transform using the 3 nearest stars
    i = EQ_UpdateAffine(obtmp.x, obtmp.Y)
    obtmp2 = EQ_plAffine(obtmp)
    
    Delta_Matrix_Reverse_Map.x = obtmp2.x
    Delta_Matrix_Reverse_Map.Y = obtmp2.Y
    Delta_Matrix_Reverse_Map.z = 1
    Delta_Matrix_Reverse_Map.F = i
    
    gSelectStar = 0
    
End Function


Public Function DeltaSync_Matrix_Map(ByVal RA As Double, ByVal DEC As Double) As Coordt
Dim i As Long

    If (RA >= &H1000000) Or (DEC >= &H1000000) Then GoTo HandleError

    i = GetNearest(RA, DEC)
    If i <> -1 Then
        gSelectStar = i
        DeltaSync_Matrix_Map.x = RA + (ct_Points(i).x - my_Points(i).x) + gRASync01
        DeltaSync_Matrix_Map.Y = DEC + (ct_Points(i).Y - my_Points(i).Y) + gDECSync01
        DeltaSync_Matrix_Map.z = 1
        DeltaSync_Matrix_Map.F = 0
    Else
HandleError:
        DeltaSync_Matrix_Map.x = RA
        DeltaSync_Matrix_Map.Y = DEC
        DeltaSync_Matrix_Map.z = 0
        DeltaSync_Matrix_Map.F = 0
    End If
End Function


Public Function DeltaSyncReverse_Matrix_Map(ByVal RA As Double, ByVal DEC As Double) As Coordt
Dim i As Long

    If (RA >= &H1000000) Or (DEC >= &H1000000) Or gAlignmentStars_count = 0 Then GoTo HandleError

    i = GetNearest(RA, DEC)
    
    If i <> -1 Then
        gSelectStar = i
        DeltaSyncReverse_Matrix_Map.x = RA - (ct_Points(i).x - my_Points(i).x)
        DeltaSyncReverse_Matrix_Map.Y = DEC - (ct_Points(i).Y - my_Points(i).Y)
        DeltaSyncReverse_Matrix_Map.z = 1
        DeltaSyncReverse_Matrix_Map.F = 0
    Else
HandleError:
        DeltaSyncReverse_Matrix_Map.x = RA
        DeltaSyncReverse_Matrix_Map.Y = DEC
        DeltaSyncReverse_Matrix_Map.z = 1
        DeltaSyncReverse_Matrix_Map.F = 0
    End If

End Function
Public Function GetQuadrant(ByRef tmpcoord As Coord) As Integer
Dim ret As Integer
    
    If tmpcoord.x >= 0 Then
        If tmpcoord.Y >= 0 Then
            ret = 0
        Else
            ret = 1
        End If
    Else
        If tmpcoord.Y >= 0 Then
            ret = 2
        Else
            ret = 3
        End If
    End If
    
    GetQuadrant = ret

End Function


Public Function GetNearest(ByVal RA As Double, ByVal DEC As Double) As Integer
Dim i As Integer
Dim tmpcoord As Coord
Dim tmpcoord2 As Coord
Dim datholder(1 To MAX_STARS) As Double
Dim datholder2(1 To MAX_STARS) As Integer
Dim Count As Integer
    
    tmpcoord.x = RA
    tmpcoord.Y = DEC
    tmpcoord = EQ_sp2Cs(tmpcoord)

    Count = 0
    
    For i = 1 To gAlignmentStars_count
        
        tmpcoord2 = my_PointsC(i)
        Select Case gPointFilter
            
            Case 0
                ' all points
        
            Case 1
                ' only consider points on this side of the meridian
                If tmpcoord2.Y * tmpcoord.Y < 0 Then
                    GoTo NextPoint
                End If
                
            Case 2
                ' local quadrant
                If GetQuadrant(tmpcoord) <> GetQuadrant(tmpcoord2) Then
                    GoTo NextPoint
                End If
                
        End Select
        
        Count = Count + 1
        If HC.CheckLocalPier.Value = 1 Then
            ' calculate polar distance
            datholder(Count) = (my_Points(i).x - RA) ^ 2 + (my_Points(i).Y - DEC) ^ 2
        Else
            ' calculate cartesian disatnce
            datholder(Count) = (tmpcoord2.x - tmpcoord.x) ^ 2 + (tmpcoord2.Y - tmpcoord.Y) ^ 2
        End If
        
        datholder2(Count) = i

NextPoint:
    Next i

    If Count = 0 Then
        GetNearest = -1
    Else
    '    i = EQ_FindLowest(datholder(), 1, gAlignmentStars_count)
        i = EQ_FindLowest(datholder(), 1, Count)
        If i = -1 Then
            GetNearest = -1
        Else
            GetNearest = datholder2(i)
        End If
    End If

End Function


'Public Function Delta_RA_Map_encoder(ByVal RAENCODER As Double) As Double'
'
'    Delta_RA_Map_encoder = RAENCODER - gRASync01 - gRA1Star
'
'End Function

'Public Function Delta_DEC_Map_encoder(ByVal DECENCODER As Double) As Double
'
'   Delta_DEC_Map_encoder = DECENCODER - gDECSync01 - gDEC1Star
'
'End Function


