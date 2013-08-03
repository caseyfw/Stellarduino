Attribute VB_Name = "Alignment"
Option Explicit
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
' Common.bas - Common functions for EQMOD ASCOM driver
'
'
' Written:  07-Oct-06   Raymund Sarmiento
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' 12-Feb-07 sander  Created file, copied contents from common.bas
'                   including new datastructure for alignment data
' 19-Mar-07 rcs     Initial Edit for Three star alignment
' 08-Apr-07 rcs     N-star implementation
' 14-Jul-07         Use 1star even before 3 star is activated
'---------------------------------------------------------------------
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



Public Const MAX_STARS As Integer = 1000
Public Const MAX_COMBINATION As Integer = 32767
Public Const MAX_COMBINATION_COUNT As Integer = 50


Public gThreeStarEnable As Boolean
Public gSelectStar As Long
Public gMaxCombinationCount As Integer
Public gLoadAPresetOnUnpark As Integer
Public gSaveAPresetOnPark As Integer
Public gSaveAPresetOnAppend As Integer
    
Public ProximityRa As Long
Public ProximityDec As Long

Public gRA_GOTO As Double
Public gDEC_GOTO As Double


Public Type AlignmentData
    OrigTargetRA As Double
    OrigTargetDEC As Double
    TargetRA    As Double
    TargetDEC   As Double
    EncoderRA   As Double
    EncoderDEC  As Double
    AlignTime   As Date
End Type

Public Enum AlignmentType
    Onestar = 1
    ThreeStar = 3
    multistar = 99
End Enum

Public gAlignmentStars_count As Integer

Public AlignmentStars(MAX_STARS) As AlignmentData

Public ct_Points(1 To MAX_STARS) As Coord  'Catalog Points
Public my_Points(1 To MAX_STARS) As Coord   'My Measured Points
Public ct_PointsC(1 To MAX_STARS) As Coord   'Catalog Points (Cartesian)
Public my_PointsC(1 To MAX_STARS) As Coord   'My Measured Points (Cartesian)



Public Sub EQ_NPointDelete(ByVal Index As Long)
Dim i As Long
    
    If Index <> gAlignmentStars_count Then
        ' first or middle element, move elements one spot
        For i = Index To gAlignmentStars_count - 1
            AlignmentStars(i) = AlignmentStars(i + 1)
        Next i
    End If
    gAlignmentStars_count = gAlignmentStars_count - 1

End Sub

Public Sub CalcPromximityLimits(ByVal range As Integer)
    ProximityRa = range * gTot_RA / 360
    ProximityDec = range * gTot_DEC / 360
End Sub

Public Sub EQ_NPointAppend(ByVal RightAscension As Double, ByVal Declination As Double, ByVal pLongitude As Double, ByVal pHemisphere As Long)

Dim tRA As Double
Dim tha As Double
Dim tPier As Double
Dim vRA As Double
Dim vDEC As Double

Dim deltaRA As Double
Dim deltadec As Double

Dim curalign As Integer
Dim i As Integer
Dim Count As Integer
Dim ERa As Long
Dim EDec As Long
Dim RA_Hours As Double
Dim flipped As Boolean

    If gSlewStatus = True Then
        HC.Add_Message (oLangDll.GetLangString(5027))
        Exit Sub
    End If
    
    HC.EncoderTimer.Enabled = False
    
    curalign = gAlignmentStars_count + 1

    ' build alignment record
    ERa = EQGetMotorValues(0)
    EDec = EQGetMotorValues(1)
    vRA = RightAscension
    vDEC = Declination
    
    ' look at current position and detemrine if flipped
    RA_Hours = Get_EncoderHours(gRAEncoder_Zero_pos, CDbl(ERa), gTot_RA, gHemisphere)
    If RA_Hours > 12 Then
        ' Yes we're currently flipped!
        flipped = True
    Else
        flipped = False
    End If
    
    tha = RangeHA(vRA - EQnow_lst(pLongitude * DEG_RAD))
    If tha < 0 Then
        If flipped Then
            If gHemisphere = 0 Then
                tPier = 0
            Else
                tPier = 1
            End If
            tRA = vRA
        Else
            If gHemisphere = 0 Then
                tPier = 1
            Else
                tPier = 0
            End If
            tRA = Range24(vRA - 12)
       End If
    Else
        If flipped Then
            If gHemisphere = 0 Then
                tPier = 1
            Else
                tPier = 0
            End If
            tRA = Range24(vRA - 12)
        Else
            If gHemisphere = 0 Then
                tPier = 0
            Else
                tPier = 1
            End If
            tRA = vRA
        End If
    End If

    'Compute for Sync RA/DEC Encoder Values
    With AlignmentStars(curalign)
        .OrigTargetDEC = Declination
        .OrigTargetRA = RightAscension
        .TargetRA = Get_RAEncoderfromRA(tRA, 0, pLongitude, gRAEncoder_Zero_pos, gTot_RA, pHemisphere)
        .TargetDEC = Get_DECEncoderfromDEC(vDEC, tPier, gDECEncoder_Zero_pos, gTot_DEC, pHemisphere)
        .EncoderRA = ERa
        .EncoderDEC = EDec
        .AlignTime = Now
        
        deltaRA = .TargetRA - .EncoderRA
        deltadec = .TargetDEC - .EncoderDEC
        
    End With
    
    HC.EncoderTimer.Enabled = True
    
    If (Abs(deltaRA) < gEQ_MAXSYNC) And (Abs(deltadec) < gEQ_MAXSYNC) Then
        
        ' Use this data also for next sync until a three star is achieved
        gRA1Star = deltaRA
        gDEC1Star = deltadec
        
        If curalign < 3 Then
            HC.Add_Message (str(curalign) & " " & oLangDll.GetLangString(6009))
            gAlignmentStars_count = gAlignmentStars_count + 1
        Else
            If curalign = 3 Then
                gAlignmentStars_count = 3
                Call SendtoMatrix
            Else
                ' add new point
                Count = 1
                ' copy points to temp array
                For i = 1 To curalign - 1
                    deltaRA = Abs(AlignmentStars(i).EncoderRA - ERa)
                    deltadec = Abs(AlignmentStars(i).EncoderDEC - EDec)
                    If deltaRA > ProximityRa Or deltadec > ProximityDec Then
                            ' point is far enough away from the new point - so keep it
                            AlignmentStars(Count) = AlignmentStars(i)
                            Count = Count + 1
                    Else
'                        HC.Add_Message ("Old Point too close " & CStr(deltaRA) & " " & CStr(deltadec) & " " & CStr(ProximityDec))
                    End If
                Next i
                
                AlignmentStars(Count) = AlignmentStars(curalign)
                curalign = Count
                gAlignmentStars_count = curalign
                
                Call SendtoMatrix

                StarEditform.RefreshDisplay = True
            
            End If
        End If
    Else
        ' sync is too large!
        HC.Add_Message (oLangDll.GetLangString(6004))
        HC.Add_Message ("Target  RA=" & FmtSexa(gRA, False))
        HC.Add_Message ("Sync    RA=" & FmtSexa(RightAscension, False))
        HC.Add_Message ("Target DEC=" & FmtSexa(gDec, True))
        HC.Add_Message ("Sync   DEC=" & FmtSexa(Declination, True))
    End If
    
    If gSaveAPresetOnAppend = 1 Then
        ' don't write emtpy list!
        If (gAlignmentStars_count > 0) Then
            'idx = GetPresetIdx
            Call SaveAlignmentStars(GetPresetIdx, "")
        End If
    End If
    
End Sub

Public Sub SendtoMatrix()

Dim i As Integer
    
    For i = 1 To gAlignmentStars_count
       ct_Points(i).x = AlignmentStars(i).TargetRA
       ct_Points(i).Y = AlignmentStars(i).TargetDEC
       ct_Points(i).z = 1
       ct_PointsC(i) = EQ_sp2Cs(ct_Points(i))
       my_Points(i).x = AlignmentStars(i).EncoderRA
       my_Points(i).Y = AlignmentStars(i).EncoderDEC
       my_Points(i).z = 1
       my_PointsC(i) = EQ_sp2Cs(my_Points(i))
    Next i

    'Activate Matrix here
    Call ActivateMatrix

End Sub

Public Sub ActivateMatrix()

Dim i As Integer

    ' assume false - will set true later if 3 stars active
    gThreeStarEnable = False
    
    HC.EncoderTimer.Enabled = False
    If HC.PolarEnable.Value = 1 Then
        If gAlignmentStars_count >= 3 Then
            i = EQ_AssembleMatrix_Taki(0, 0, ct_PointsC(1), ct_PointsC(2), ct_PointsC(3), my_PointsC(1), my_PointsC(2), my_PointsC(3))
            i = EQ_AssembleMatrix_Affine(0, 0, my_PointsC(1), my_PointsC(2), my_PointsC(3), ct_PointsC(1), ct_PointsC(2), ct_PointsC(3))
            gThreeStarEnable = True
        End If
    Else
        If gAlignmentStars_count >= 3 Then
            i = EQ_AssembleMatrix_Taki(0, 0, ct_PointsC(1), ct_PointsC(2), ct_PointsC(3), my_PointsC(1), my_PointsC(2), my_PointsC(3))
            i = EQ_AssembleMatrix_Affine(0, 0, my_PointsC(1), my_PointsC(2), my_PointsC(3), ct_PointsC(1), ct_PointsC(2), ct_PointsC(3))
            gThreeStarEnable = True
        End If
    End If
    HC.EncoderTimer.Enabled = True

End Sub



'''''''''''''''''''''''''
' Alignment preset stuff
'''''''''''''''''''''''''

Public Sub SaveAlignmentStars(preset As Integer, presetName As String)
Dim Index As Integer
Dim DataStr As String
Dim tmp As String
Dim key As String
Dim Alignini As String

    ' set up a file path for the align.ini file
    Alignini = HC.oPersist.GetIniPath & "\ALIGN.ini"

    key = "[alignment_preset" & CStr(preset) & "]"
    
    If presetName = "" Then
        ' get existing name
        presetName = HC.oPersist.ReadIniValueEx("NAME", key, Alignini)
    End If
    
    ' delete existing section
    Call HC.oPersist.DeleteSection(key, Alignini)
    
    ' write new data
    HC.oPersist.WriteIniValueEx "STAR_COUNT", CStr(gAlignmentStars_count), key, Alignini
    HC.oPersist.WriteIniValueEx "NAME", presetName, key, Alignini

    For Index = 1 To gAlignmentStars_count
        tmp = "Star" + CStr(Index)
        With AlignmentStars(Index)
            DataStr = CStr(.AlignTime) + ";" + CStr(.OrigTargetRA) + ";" + CStr(.OrigTargetDEC) + ";" + CStr(.TargetRA) + ";" + CStr(.TargetDEC) + ";" + CStr(.EncoderRA) + ";" + CStr(.EncoderDEC) + ";"
            HC.oPersist.WriteIniValueEx tmp, DataStr, key, Alignini
        End With
    Next Index
End Sub

Public Function LoadAlignmentPreset(preset As Integer) As Boolean
Dim Count As Integer
Dim tmptxt As String
Dim tmptxt2 As String
Dim VarStr As String
Dim pos As Integer
Dim Index As Integer
Dim ValidCount As Integer
Dim MaxCount As Integer
Dim NewData As AlignmentData
Dim key As String
Dim Alignini As String
Dim ret As Boolean

    ret = False

    ' set up a file path for the align.ini file
    Alignini = HC.oPersist.GetIniPath & "\ALIGN.ini"

    key = "[alignment_preset" & CStr(preset) & "]"

    tmptxt = HC.oPersist.ReadIniValueEx("STAR_COUNT", key, Alignini)
    If tmptxt <> "" Then
        MaxCount = val(tmptxt)
        If MaxCount > MAX_STARS Then
            MaxCount = MAX_STARS
        End If
    Else
        MaxCount = 0
    End If

    On Error GoTo DecodeError
    If MaxCount <> 0 Then
        ValidCount = 0
        For Index = 1 To MaxCount
            VarStr = "Star" + CStr(Index)
            tmptxt = HC.oPersist.ReadIniValueEx(VarStr, key, Alignini)
            If tmptxt <> "" Then
                pos = InStr(tmptxt, ";")
                If pos = 0 Then GoTo DecodeError
                tmptxt2 = Left$(tmptxt, pos - 1)
                tmptxt = Right$(tmptxt, Len(tmptxt) - pos)
                NewData.AlignTime = tmptxt2
                    
                pos = InStr(tmptxt, ";")
                If pos = 0 Then GoTo DecodeError
                tmptxt2 = Left$(tmptxt, pos - 1)
                tmptxt = Right$(tmptxt, Len(tmptxt) - pos)
                NewData.OrigTargetRA = CDbl(tmptxt2)
                
                pos = InStr(tmptxt, ";")
                If pos = 0 Then GoTo DecodeError
                tmptxt2 = Left$(tmptxt, pos - 1)
                tmptxt = Right$(tmptxt, Len(tmptxt) - pos)
                NewData.OrigTargetDEC = CDbl(tmptxt2)
                
                pos = InStr(tmptxt, ";")
                If pos = 0 Then GoTo DecodeError
                tmptxt2 = Left$(tmptxt, pos - 1)
                tmptxt = Right$(tmptxt, Len(tmptxt) - pos)
                NewData.TargetRA = CDbl(tmptxt2)
                
                pos = InStr(tmptxt, ";")
                If pos = 0 Then GoTo DecodeError
                tmptxt2 = Left$(tmptxt, pos - 1)
                tmptxt = Right$(tmptxt, Len(tmptxt) - pos)
                NewData.TargetDEC = CDbl(tmptxt2)
                
                pos = InStr(tmptxt, ";")
                If pos = 0 Then GoTo DecodeError
                tmptxt2 = Left$(tmptxt, pos - 1)
                tmptxt = Right$(tmptxt, Len(tmptxt) - pos)
                NewData.EncoderRA = CDbl(tmptxt2)
            
                pos = InStr(tmptxt, ";")
                If pos = 0 Then GoTo DecodeError
                tmptxt2 = Left$(tmptxt, pos - 1)
                NewData.EncoderDEC = CDbl(tmptxt2)
                
                ' all data read ok - copy to alignment stars
                AlignmentStars(Index) = NewData
                ValidCount = ValidCount + 1
            Else
                GoTo DecodeError
            End If
        Next Index
    
DecodeError:
        On Error Resume Next
        gAlignmentStars_count = ValidCount
        
        ' send to matrix will initialise the catalog and measured points arrays
        Call SendtoMatrix
        
        ret = True
    End If
 
    LoadAlignmentPreset = ret
End Function

Public Sub SavePresetIdx(idx As Integer)
Dim Index As Integer
Dim tmptxt As String
Dim Alignini As String

    ' set up a file path for the align.ini file
    Alignini = HC.oPersist.GetIniPath & "\ALIGN.ini"
    tmptxt = CStr(idx)
    Call HC.oPersist.WriteIniValueEx("active_preset", tmptxt, "[default]", Alignini)
    
End Sub

Public Function GetPresetIdx() As Integer
Dim Index As Integer
Dim tmptxt As String
Dim Alignini As String

    ' set up a file path for the align.ini file
    Alignini = HC.oPersist.GetIniPath & "\ALIGN.ini"
    tmptxt = HC.oPersist.ReadIniValueEx("active_preset", "[default]", Alignini)
    
    If tmptxt = "" Then
        ' ini file entry doesn't exist so create one
        tmptxt = "0"
        Call HC.oPersist.WriteIniValueEx("active_preset", tmptxt, "[default]", Alignini)
    End If
    
    GetPresetIdx = val(tmptxt)
    
    If GetPresetIdx > 10 Then
        GetPresetIdx = 0
    End If
    
End Function

Public Sub ReadParkOptions()
Dim keyStr As String
Dim tmptxt As String
Dim Alignini As String

    keyStr = "[default]"
    ' set up a file path for the align.ini file
    Alignini = HC.oPersist.GetIniPath & "\ALIGN.ini"
    tmptxt = HC.oPersist.ReadIniValueEx("LOAD_APRESET_ON_UNPARK", keyStr, Alignini)
    If tmptxt = "" Then
        ' create a preset place holder
        Call HC.oPersist.WriteIniValueEx("LOAD_APRESET_ON_UNPARK", "0", keyStr, Alignini)
        gLoadAPresetOnUnpark = 0
    Else
        gLoadAPresetOnUnpark = val(tmptxt)
    End If
    
    tmptxt = HC.oPersist.ReadIniValueEx("SAVE_APRESET_ON_UNPARK", keyStr, Alignini)
    If tmptxt = "" Then
        ' create a preset place holder
        Call HC.oPersist.WriteIniValueEx("SAVE_APRESET_ON_UNPARK", "0", keyStr, Alignini)
        gSaveAPresetOnPark = 0
    Else
        gSaveAPresetOnPark = val(tmptxt)
    End If
    
    tmptxt = HC.oPersist.ReadIniValueEx("SAVE_APRESET_ON_APPEND", keyStr, Alignini)
    If tmptxt = "" Then
        ' create a preset place holder
        Call HC.oPersist.WriteIniValueEx("SAVE_APRESET_ON_APPEND", "0", keyStr, Alignini)
        gSaveAPresetOnAppend = 0
    Else
        gSaveAPresetOnAppend = val(tmptxt)
    End If
End Sub
Public Sub WriteParkOptions()
Dim keyStr As String
Dim tmptxt As String
Dim Alignini As String

    keyStr = "[default]"
    ' set up a file path for the align.ini file
    Alignini = HC.oPersist.GetIniPath & "\ALIGN.ini"
    Call HC.oPersist.WriteIniValueEx("LOAD_APRESET_ON_UNPARK", CStr(gLoadAPresetOnUnpark), keyStr, Alignini)
    Call HC.oPersist.WriteIniValueEx("SAVE_APRESET_ON_UNPARK", CStr(gSaveAPresetOnPark), keyStr, Alignini)
    Call HC.oPersist.WriteIniValueEx("SAVE_APRESET_ON_APPEND", CStr(gSaveAPresetOnAppend), keyStr, Alignini)
End Sub

Public Sub AligmentStarsPark()
Dim idx As Integer
    ReadParkOptions
    If gSaveAPresetOnPark = 1 Then
        ' don't write emtpy list!
        If (gAlignmentStars_count > 0) Then
            idx = GetPresetIdx
            Call SaveAlignmentStars(idx, "")
        End If
    End If
End Sub
Public Sub AlignmentStarsUnpark()
Dim idx As Integer
    ReadParkOptions
    ' if load on unpark selected
    If gLoadAPresetOnUnpark = 1 Then
        ' read curent preset index from ini file
        idx = GetPresetIdx
        ' load the preset data
        Call LoadAlignmentPreset(idx)
    End If
End Sub

