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

' Iterative GOTO Constants

Public Const NUM_SLEW_RETRIES As Long = 5                   ' Iterative MAX retries
Public Const gRA_Allowed_diff As Double = 10                ' Iterative Slew minimum difference
Public Const gRA_Compensate = 40                            ' Least RA discrepancy Compensation


' Home Position of the mount (pointing at NCP/SCP)

Public Const RAEncoder_Home_pos As Double = &H800000        ' Start at 0 Hour
Public Const DECEncoder_Home_pos As Double = &HA26C80       ' Start at 90 Degree position

Public Const gRAEncoder_Zero_pos As Double = &H800000       ' ENCODER 0 Hour initial position
Public Const gDECEncoder_Zero_pos As Double = &H800000      ' ENCODER 0 Degree Initial position
  
Public Const gTot_step As Double = 9024000                  ' Total Encoder count


'------------------- EQCONTRL.DLL Constants -----------------------

Public Const EQ_OK As Double = &H0
Public Const EQ_COMNOTOPEN As Double = &H1
Public Const EQ_COMTIMEOUT As Double = &H3
Public Const EQ_MOTORBUSY As Double = &H10
Public Const EQ_NOTINITIALIZED As Double = &HC8
Public Const EQ_INVALIDCOORDINATE As Double = &H1000000



'------------------------------------------------------------------------------------------------

' Define all Global Variables

Public eqres As Double
Public gTot_RA As Double                                    ' Total RA Encoder Steps
Public gTot_DEC As Double                                   ' Total DEC Encoder Steps

Public gLatitude As Double                                  ' Site Latitude
Public gLongitude As Double                                 ' Site Longitude
Public gElevation As Double                                 ' Site Elevation
Public gHemisphere As Long

Public gRA_Encoder As Double                                ' RA Current Polled RA Encoder value
Public gDec_Encoder As Double                               ' DEC Current Polled Encoder value
Public gRA_Hours As Double                                  ' RA Encoder to Hour position
Public gDec_Degrees As Double                               ' DEC Encoder to Degree position Ranged to -90 to 90
Public gDec_DegNoAdjust As Double                           ' DEC Encoder to actual degree position
Public gRAStatus As Double                                  ' RA Polled Motor Status
Public gRAStatus_slew As Boolean                            ' RA motor tracking poll status
Public gDECStatus As Double                                 ' DEC Polloed motor status
 
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

Public gTargetRA As Double
Public gTargetDec As Double


Public gPort As String
Public gBaud As Long
Public gTimeout As Long
Public gRetry As Long

Public gTrackingStatus As Long
Public gSlewCount As Long
Public gSlewStatus As Boolean

Public Enum PierSide2
    pierUnknown2 = -1
    PierEast2 = 0
    PierWest2 = 1
End Enum

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

Public Function FmtSexa(ByVal n As Double, ShowPlus As Boolean) As String
    Dim sg As String
    Dim us As String
    Dim ms As String
    Dim ss As String
    Dim u As Long
    Dim m As Long
    Dim fmt
    

    sg = "+"                                ' Assume positive
    If n < 0 Then                           ' Check neg.
        n = -n                              ' Make pos.
        sg = "-"                            ' Remember sign
    End If

    m = Fix(n)                              ' Units (deg or hr)
    us = Format$(m, "00")

    n = (n - m) * 60#
    m = Fix(n)                              ' Minutes
    ms = Format$(m, "00")

    n = (n - m) * 60#
    m = Fix(n)                              ' Minutes
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
    mjd = vb_mjd(CDbl(Now + (typTime.wMilliseconds / 86400000)))
    Call utc_gst(mjd_day(mjd), mjd_hr(mjd), lst)
    lst = lst + radhr(plong)
    Call obliq(mjd, eps)
    Call nut(mjd, deps, dpsi)
    lst = lst + radhr(dpsi * Cos(eps + deps))
    Call range(lst, 24#)

    EQnow_lst = lst
'    EQnow_lst = now_lst(plong)

End Function


Public Function SOP_RAHours(vha As Double) As PierSide2

    Dim ha As Double
    
        ha = RangeHA(vha - 6#)
        SOP_RAHours = IIf(ha >= 0, PierEast2, PierWest2)
 
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

Public Function Delta_DEC_Map(ByVal DECENCODER As Double) As Double

    Delta_DEC_Map = DECENCODER + gDEC1Star + gDECSync01

End Function

Public Function Delta_RA_Map_encoder(ByVal RAENCODER As Double) As Double

    Delta_RA_Map_encoder = RAENCODER - gRASync01 - gRA1Star

End Function

Public Function Delta_DEC_Map_encoder(ByVal DECENCODER As Double) As Double

    Delta_DEC_Map_encoder = DECENCODER - gDECSync01 - gDEC1Star

End Function


