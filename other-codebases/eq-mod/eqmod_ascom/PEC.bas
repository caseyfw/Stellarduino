Attribute VB_Name = "PEC"
'---------------------------------------------------------------------
' Copyright © 2008 EQMOD Development Team
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
' PEC.bas - Periodic Error Correction functions for EQMOD ASCOM Driver
'
'
' Written:  12-Oct-07   Chris Shillito
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
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
'  The mount circuitry needs to be modified for this program to work.
'  Circuit details can be found at http://sourceforge.net/projects/eq-mod/
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

Option Explicit

Type CapRecord
    time As Double
    MotorPos As Double
    DeltaPos As Double
    DeltaTime As Double
    rate As Double
    pe As Double
    peSmoothed As Double
    peInc As Double
End Type
    

Type PECCapDef
    StartTime As Date
    Period As Double
    Steps As Double
    idx As Integer
    FileName As String
    CapureData() As CapRecord
End Type

Type PECData
    time As Double
    PEPosition As Double
    PECPosition As Double
    RawPosn As Double
    signal As Double
    PErate As Double
    PECrate As Double
    cycle As Integer
End Type

Type PECDefinition
    PECCurve() As PECData
    PECCurveTmp() As PECData
    Period As Double
    Steps As Double
    MaxPe As Double
    MinPe As Double
    FileName As String
    CurrIdx As Integer
End Type

Type PECFileData
    time As Double
    Position As Double
    pe As Double
    cycle As Integer
End Type


Private PECCap As PECCapDef
Public PECDef1 As PECDefinition
Public gLastPE As Double
Public gPEC_Enabled As Boolean
Public gUsePEC As Boolean
Public gPEC_Gain As Double                   ' current gain setting
Public gPEC_Capture_Cycles As Integer
Public gPEC_filter_lowpass As Integer
Public gPEC_mag As Integer
Public gPEC_PhaseAdjust As Integer          ' current phase adjustment (samples)
Public gPEC_TimeStampFiles As Integer
Public gPEC_DynamicRateAdjust As Integer
Public gPEC_FileDir As String
Public gPEC_trace As Integer
Public gPEC_AutoApply As Integer
Public gPEC_Debug As Integer

Private PEC_File As String                  ' path and name of PEC file
Private threshold As Double                 ' minimum correction PEC will make
Private phaseshift As Double                ' current phase shift (steps)
Private gMaxRateAdjust As Double            ' Maximum correction PEC is allowed to make
Private MaxRate As Double                   ' Fastset rate allowed
Private MinRate As Double                   ' slowest rate allowed
Private SID_RATE_NORTH As Double            ' 15.041067        ' arcsecs/sec  (60*60*360) / ((23*60*60)+(56*60)+4)
Private SID_RATE_SOUTH As Double            ' -15.041067       ' arcsecs/sec

Type PlaybackTimerStatic
    PecResyncCount As Integer
    CurrRate As Double                  ' current rate
    Firsttime As Boolean                ' oneshot flag
    newpos As Single
    oldpos As Single
    timerflag As Boolean                ' timer interlock
    ringcounter As Long
    StartRingCounter As Double
    LastRingCounter As Double
    StartTime As Double
    lasttime As Double
    RateSumExpected As Double
    RateSumActual As Double
    TraceIdx As Long
    strPlayback As String
End Type

Type CaptureTimerStatic
    State As Integer
    timerflag As Boolean                ' timer interlock
    ringcounter As Long
    StartRingCounter As Double
    LastRingCounter As Double
    StartTime As Double
    lasttime As Double
    pe As Double
    yoffset As Single
    lastx As Single
    lasty As Single
    PenToggle As Boolean
    InvertCapture As Integer
    strCapture As String
End Type

Private CaptureTimer As CaptureTimerStatic
Private PlaybackTimer As PlaybackTimerStatic

Private TraceFileNum As Integer

Const ARCSECS_PER_360DEGREES = 1296000      ' 360*60*60

Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub PEC_LoPassScroll_Change()
    gPEC_filter_lowpass = PECConfigFrm.HScroll1.Value
    If gPEC_filter_lowpass < 9 Then
        PECConfigFrm.Label3.Caption = oLangDll.GetLangString(6116)
        gPEC_filter_lowpass = 0
    Else
        PECConfigFrm.Label3.Caption = CStr(gPEC_filter_lowpass)
    End If

End Sub

Public Sub PEC_MagScroll_Change()
    gPEC_mag = PECConfigFrm.HScroll2.Value
    If gPEC_mag = 0 Then
        PECConfigFrm.Label2.Caption = oLangDll.GetLangString(6116)
    Else
        PECConfigFrm.Label2.Caption = CStr(gPEC_mag)
    End If
End Sub

Public Sub PEC_PhaseScroll_Change()
Dim adj As Double
    PECConfigFrm.Label45.Caption = CStr(Int(360 * (PECConfigFrm.PhaseScroll.Value / gRAWormPeriod))) & " deg."
    gPEC_PhaseAdjust = PECConfigFrm.PhaseScroll.Value
    phaseshift = PECConfigFrm.PhaseScroll.Value * (gRAWormSteps / gRAWormPeriod)
    If PECConfigFrm.PhaseScroll.Enabled Then
        PlaybackTimer.ringcounter = EQGetMotorValues(0)
        PECDef1.CurrIdx = GetIdx(PECDef1)
    End If

End Sub

Public Sub PECMode_click()
    Dim key As String
    Dim Ini As String

    Ini = HC.oPersist.GetIniPath & "\EQMOD.ini"
    key = "[pec]"
     
    Call HC.oPersist.WriteIniValueEx("DYNAMIC_RATE_ADJUST", CStr(gPEC_DynamicRateAdjust), key, Ini)

End Sub


Public Sub PEC_Initialise()
    PlaybackTimer.Firsttime = True
    HC.CmdPecSave.Enabled = False
    PECConfigFrm.GainScroll.Enabled = False
    PECConfigFrm.PhaseScroll.Enabled = False
    
    ReDim PECDef1.PECCurve(gRAWormPeriod)
    ReDim PECDef1.PECCurveTmp(gRAWormPeriod)
    PECDef1.Period = gRAWormPeriod
    PECDef1.Steps = gRAWormSteps
    
    SID_RATE_NORTH = SID_RATE
    SID_RATE_SOUTH = -1 * SID_RATE  ' gSiderealRate
    Call PEC_ReadParams
    
    If gHemisphere Then
        MaxRate = SID_RATE_SOUTH + gMaxRateAdjust
        MinRate = SID_RATE_SOUTH - gMaxRateAdjust
    Else
        MaxRate = SID_RATE_NORTH + gMaxRateAdjust
        MinRate = SID_RATE_NORTH - gMaxRateAdjust
    End If
    
    If (gTot_RA <> 0) Then
        If (gRAWormPeriod > 0) Then
            PECConfigFrm.PhaseScroll.max = gRAWormPeriod
            If Import(PECDef1) <> True Then KillPec
         End If
    End If
    
    Call PEC_DrawAxis(HC.plot)
    Call PEC_DrawAxis(HC.PlotCap)
    
    Call PEC_UpdateControls
      
    PlaybackTimer.strPlayback = oLangDll.GetLangString(6117)

End Sub

Public Sub PEC_Timestamp()
    Dim key As String
    Dim Ini As String
    Dim pos As Double
    Dim temp As String

    Ini = HC.oPersist.GetIniPath & "\EQMOD.ini"
    key = "[pec]"
    pos = EQGetMotorValues(0)
    temp = Now
    Call HC.oPersist.WriteIniValueEx("SYNCPOS", CStr(pos), key, Ini)
    Call HC.oPersist.WriteIniValueEx("SYNCTIME", temp, key, Ini)
    Call HC.oPersist.WriteIniValueEx("STAR_DEC", CStr(gDec), key, Ini)
    Call HC.oPersist.WriteIniValueEx("STAR_RA", CStr(gRA), key, Ini)
End Sub

Public Sub PEC_StartTracking()
    Dim rate As Double

    HC.PECTimer.Enabled = False
    
    gPEC_Enabled = True
    HC.TrackingFrame.Caption = oLangDll.GetLangString(121) & " " & oLangDll.GetLangString(188)
    PlaybackTimer.CurrRate = 0
    If gTrackingStatus = 0 Then
        gTrackingStatus = 1
    End If
    Call PEC_PlotCurve(PECDef1)
    PlaybackTimer.Firsttime = True
    PlaybackTimer.timerflag = False
    PlaybackTimer.TraceIdx = 0
    HC.PECTimer.Interval = 1000
    HC.PECTimer.Enabled = True
    
    On Error Resume Next
    TraceFileNum = FreeFile
    Close TraceFileNum
    If gPEC_trace = 1 Then
        Close TraceFileNum
        Open HC.oPersist.GetIniPath() + "\pectrace_" & CStr(gPEC_DynamicRateAdjust) & ".txt" For Output As TraceFileNum
'        Print TraceFileNum, "Idx WormIndex Motor StepsMoved NextRate ElapsedTime CurrentRate RateAchived RateSumExpected RateSumActual RateError OverallRate"
        Print #TraceFileNum, "Idx WormIdx Motor StepsMoved NextRate OverallRate elapesedtime dt TimerInterval RateError MeasuredRate"
    End If

End Sub

Public Sub PEC_StopTracking()
    gPEC_Enabled = False
    PlaybackTimer.strPlayback = oLangDll.GetLangString(6117)
    gRA_LastRate = 0
    Close TraceFileNum
End Sub

Public Sub PEC_Unload()
    Call PEC_WriteParams
    Close TraceFileNum
End Sub

Public Sub PEC_GainScroll_Change()
    On Error Resume Next
    gPEC_Gain = CDbl(PECConfigFrm.GainScroll.Value) / 10
    PECConfigFrm.Label43.Caption = "x" & CStr(gPEC_Gain)
    If PECConfigFrm.GainScroll.Enabled Then
        If CalcRates(PECDef1) Then KillPec
        PEC_WriteParams
    End If
End Sub
Public Sub PEC_Clear()
    PEC_File = ""
    PEC_WriteParams
    KillPec

End Sub

Public Sub PEC_OnUse()
    If HC.CheckPEC.Value Then
        HC.CmdTrack(1).Visible = True
        HC.CmdTrack(0).Visible = False
        HC.CommandPecPlay.Picture = LoadResPicture(109, vbResBitmap)
    Else
        HC.CmdTrack(1).Visible = False
        HC.CmdTrack(0).Visible = True
        HC.CommandPecPlay.Picture = LoadResPicture(108, vbResBitmap)
        PEC_StopTracking
    End If
    
    If gTrackingStatus = 1 Then
        EQStartSidereal
    End If
End Sub
Public Function PEC_SetGain(sGain As String) As Boolean
    Dim dGain As Double
    dGain = val(sGain)
    If (dGain >= PECConfigFrm.GainScroll.min And dGain <= PECConfigFrm.GainScroll.max) Then
        PECConfigFrm.GainScroll.Value = dGain
        PEC_SetGain = True
    Else
        PEC_SetGain = False
    End If
End Function

Public Function PEC_SetPhase(sPhase As String) As Boolean
    Dim dPhase As Double
    dPhase = val(sPhase)
    If (dPhase >= PECConfigFrm.PhaseScroll.min And dPhase <= PECConfigFrm.PhaseScroll.max) Then
        PECConfigFrm.PhaseScroll.Value = dPhase
        PEC_SetPhase = True
    Else
        PEC_SetPhase = False
    End If
End Function

Public Sub PEC_Load()
    FileDlg.filter = "*.txt*"
    FileDlg.Show (1)
    If FileDlg.FileName <> "" Then PEC_LoadFile FileDlg.FileName
End Sub

Public Function PEC_LoadFile(FileName As String) As Boolean
    PEC_File = FileName
        
    PECDef1.FileName = FileName
    If Import(PECDef1) Then
        PEC_WriteParams
        HC.Change_Display (3)
        PEC_LoadFile = True
    Else
        KillPec
        PEC_LoadFile = False
    End If

End Function

Public Sub PEC_Save()
Dim i As Integer
    FileDlg.filter = "*.txt*"
    FileDlg.Show (1)
    If FileDlg.FileName <> "" Then
        PEC_File = FileDlg.FileName
        ' force a .txt extension
        i = InStr(PEC_File, ".")
        If i <> 0 Then
            PEC_File = Left$(PEC_File, i - 1)
        End If
        PEC_File = PEC_File & ".txt"
        PECDef1.FileName = PEC_File
        Call Export(PECDef1, PECConfigFrm.PhaseScroll.Value)
        PECConfigFrm.PhaseScroll.Value = 0
        Call PEC_WriteParams
        If Import(PECDef1) = True Then
            Call PEC_PlotCurve(PECDef1)
        Else
            KillPec
        End If
    End If
End Sub

Public Function PEC_SaveFile(FileName As String, PECDef As PECDefinition) As Boolean
    PECDef.FileName = FileName
    If Export(PECDef, 0) Then
        PEC_SaveFile = True
    Else
        PEC_SaveFile = False
    End If
End Function

Public Sub PEC_Timer()
    Dim rate As Double
    Dim x As Integer
    Dim timenow As Double
    Dim TimeSlip As Double
    Dim RateSumError As Double
    Dim StepsMoved As Long
    Dim RateError As Double
    Dim OverallRate As Double
    Dim MeasuredRate As Double
    
    Dim elapsedtime As Double
    Dim curr As Double 'current time
    Dim dt As Double 'delta time
    Dim TimerInterval As Double
    
    On Error Resume Next
    
    If Not PlaybackTimer.timerflag Then
        PlaybackTimer.timerflag = True
        
        If PlaybackTimer.Firsttime Then
            PlaybackTimer.lasttime = GetTickCount()
            PlaybackTimer.StartTime = PlaybackTimer.lasttime
            PlaybackTimer.ringcounter = EQGetMotorValues(0)
            PlaybackTimer.LastRingCounter = PlaybackTimer.ringcounter
            PlaybackTimer.StartRingCounter = PlaybackTimer.ringcounter
        
            ' force immediate rate update
            PECDef1.CurrIdx = GetIdx(PECDef1)
            rate = PECDef1.PECCurve(PECDef1.CurrIdx).PECrate
            If gPEC_Enabled And gTrackingStatus = 1 Then
                Call PEC_MoveAxis(0, rate)
            End If
            
            HC.plot.DrawMode = 7
            PlaybackTimer.newpos = PECDef1.CurrIdx * HC.plot.ScaleWidth / PECDef1.Period
            HC.plot.Line (PlaybackTimer.newpos, 0)-(PlaybackTimer.newpos, HC.plot.ScaleHeight), vbRed
            PlaybackTimer.oldpos = PlaybackTimer.newpos

            
            PlaybackTimer.RateSumActual = 0
            PlaybackTimer.RateSumExpected = rate
            PlaybackTimer.CurrRate = rate
            
            PlaybackTimer.Firsttime = False
        Else
            curr = GetTickCount() ' read current system time
            'determine the diff between times
            elapsedtime = Abs(CDbl(curr - PlaybackTimer.StartTime)) / 1000
            
            dt = Abs(CDbl(curr - PlaybackTimer.lasttime)) / 1000 'determine the diff between times
            PlaybackTimer.lasttime = curr
                                      
'            If gTrackingStatus <> 0 Then
                'only maintain pe trace updates if we're tracking
                                               
                ' only apply rate changes if we're tracking at sidreal and PEC is on
                If gPEC_Enabled And gTrackingStatus = 1 Then
                                       
                    PlaybackTimer.ringcounter = EQGetMotorValues(0)
                    StepsMoved = PlaybackTimer.ringcounter - PlaybackTimer.LastRingCounter
                    
                    PECDef1.CurrIdx = PECDef1.CurrIdx + 1
                    If PECDef1.CurrIdx >= PECDef1.Period Then
                        PECDef1.CurrIdx = 0
                    End If
                    
                    PlaybackTimer.PecResyncCount = PlaybackTimer.PecResyncCount + 1
                    If PlaybackTimer.PecResyncCount >= PECDef1.Period Then
                        PECDef1.CurrIdx = GetIdx(PECDef1)
                        TimerInterval = 1000
                        PlaybackTimer.StartTime = GetTickCount()
                        PlaybackTimer.StartRingCounter = PlaybackTimer.ringcounter
                        PlaybackTimer.RateSumExpected = 0
                        PlaybackTimer.RateSumActual = 0
                        RateError = 0
                        Call PEC_PlotCurve(PECDef1)
                    Else
                        TimeSlip = elapsedtime - PlaybackTimer.PecResyncCount
                        MeasuredRate = (StepsMoved / dt) * (1296000 / CDbl(gTot_RA))
                        PlaybackTimer.RateSumActual = PlaybackTimer.RateSumActual + MeasuredRate
                        RateError = PlaybackTimer.RateSumExpected - PlaybackTimer.RateSumActual
                        
                        If TimeSlip > 0 Then
                            TimerInterval = 1000 - (TimeSlip * 1000)
                            If TimerInterval < 100 Then TimerInterval = 100
                        Else
                            TimerInterval = 1000
                        End If
                    End If
                    HC.PECTimer.Interval = TimerInterval
                    
                    PlaybackTimer.LastRingCounter = PlaybackTimer.ringcounter
                    
                    
                    ' Get next rate to apply.
                    rate = PECDef1.PECCurve(PECDef1.CurrIdx).PECrate
                    PlaybackTimer.RateSumExpected = PlaybackTimer.RateSumExpected + rate
                    If rate <> PlaybackTimer.CurrRate Then
                        ' apply the min/max limits - just in case there's
                        ' an error in the rate calculations this prevents'
                        ' the mount from ever slewing wildly!
                        If rate > MaxRate Then
                            rate = MaxRate
                        Else
                            If rate < MinRate Then
                                rate = MinRate
                            End If
                        End If
                        
                        If gHemisphere = 0 Then
                            PlaybackTimer.strPlayback = oLangDll.GetLangString(6118) & " " & FormatNumber(rate - SID_RATE, 3)
                        Else
                            PlaybackTimer.strPlayback = oLangDll.GetLangString(6118) & " " & FormatNumber(-rate - SID_RATE, 3)
                        End If
                        
                        Select Case gPEC_DynamicRateAdjust
                            Case 1
                                Call PEC_MoveAxis(0, rate + RateError)
                            Case 0
                                Call PEC_MoveAxis(0, rate)
                        End Select
                        PlaybackTimer.CurrRate = rate
                    End If
                Else
                    PlaybackTimer.strPlayback = oLangDll.GetLangString(6117)
                    PlaybackTimer.ringcounter = gEmulRA
                    PECDef1.CurrIdx = GetIdx(PECDef1)
                End If
                
                HC.plot.DrawMode = 7
                PlaybackTimer.newpos = PECDef1.CurrIdx * HC.plot.ScaleWidth / PECDef1.Period
                If CInt(PlaybackTimer.oldpos) <> CInt(PlaybackTimer.newpos) Then
                    HC.plot.Line (PlaybackTimer.oldpos, 0)-(PlaybackTimer.oldpos, HC.plot.ScaleHeight), vbRed
                    HC.plot.Line (PlaybackTimer.newpos, 0)-(PlaybackTimer.newpos, HC.plot.ScaleHeight), vbRed
                    PlaybackTimer.oldpos = PlaybackTimer.newpos
                End If
                
 '           End If
            
            If gPEC_trace = 1 Then
                OverallRate = ((PlaybackTimer.ringcounter - PlaybackTimer.StartRingCounter) * 1296000 / gTot_RA) / elapsedtime
                Print #TraceFileNum, CStr(PlaybackTimer.TraceIdx) & " " & CStr(PECDef1.CurrIdx) & " " & CStr(PlaybackTimer.ringcounter) & " " & CStr(StepsMoved) & " " & CStr(PlaybackTimer.CurrRate) & " " & CStr(OverallRate) & " " & CStr(elapsedtime) & " " & CStr(dt) & " " & CStr(CInt(TimerInterval)) & " " & CStr(RateError) & " " & CStr(MeasuredRate)
             End If
                        
        End If
        
        PlaybackTimer.timerflag = False
    Else
        Print #TraceFileNum, "TimerOverflow"
    End If
    PlaybackTimer.TraceIdx = PlaybackTimer.TraceIdx + 1

End Sub
Public Sub PEC_CaptureTimer()
    Dim timenow As Double
    Dim TimeSlip As Double
    Dim StepsMoved As Long
    Dim elapsedtime As Double
    Dim curr As Double 'current time
    Dim dt As Double 'delta time
    Dim TimerInterval As Double
    Dim x As Single
    Dim Y As Single
    
    
    On Error Resume Next
    
    If Not CaptureTimer.timerflag Then
        CaptureTimer.timerflag = True
        
        Select Case CaptureTimer.State
        
            Case 0
                ' initialise
                gPEC_Capture_Cycles = PECConfigFrm.ComboPecCap.ItemData(PECConfigFrm.ComboPecCap.ListIndex)
                ReDim PECCap.CapureData(gRAWormPeriod * gPEC_Capture_Cycles)
                PECCap.Period = gRAWormPeriod
                PECCap.Steps = gRAWormSteps
                PECCap.idx = 0
                
                CaptureTimer.lasttime = GetTickCount()
                CaptureTimer.StartTime = CaptureTimer.lasttime
                CaptureTimer.ringcounter = EQGetMotorValues(0)
                CaptureTimer.LastRingCounter = CaptureTimer.ringcounter
                CaptureTimer.StartRingCounter = CaptureTimer.ringcounter
                CaptureTimer.pe = 0
                CaptureTimer.lastx = 0
                CaptureTimer.lasty = HC.PlotCap.ScaleHeight / 2
                CaptureTimer.yoffset = 0
                HC.PlotCap.Cls
                HC.PlotCap.DrawMode = 13
                CaptureTimer.State = 1
                
            Case 1
                ' capture
                If gTrackingStatus = 1 Then
                    CaptureTimer.strCapture = oLangDll.GetLangString(6119) & " " & CStr(PECCap.idx) & "/" & CStr(UBound(PECCap.CapureData()))
                    curr = GetTickCount() ' read current system time
                    'determine the diff between times
                    elapsedtime = Abs(CDbl(curr - CaptureTimer.StartTime)) / 1000
                    dt = Abs(CDbl(curr - CaptureTimer.lasttime)) / 1000 'determine the diff between times
                    CaptureTimer.lasttime = curr
                
                    With PECCap.CapureData(PECCap.idx)
                        .MotorPos = EQGetMotorValues(0)
                        .DeltaPos = .MotorPos - CaptureTimer.LastRingCounter
                        .DeltaTime = dt
                        CaptureTimer.LastRingCounter = .MotorPos
                        .time = elapsedtime
                        .rate = (.DeltaPos / .DeltaTime) * (1296000 / CDbl(gTot_RA))
                        If CaptureTimer.InvertCapture <> 0 Then
                            .peInc = .rate - gSiderealRate
                        Else
                            .peInc = gSiderealRate - .rate
                        End If
                        CaptureTimer.pe = CaptureTimer.pe + .peInc
                        .pe = CaptureTimer.pe
                        x = PECCap.idx Mod PECCap.Period
                        If x = 0 Then
                            CaptureTimer.yoffset = CaptureTimer.pe * HC.PlotCap.ScaleHeight / 180
                            CaptureTimer.lasty = HC.PlotCap.ScaleHeight / 2
                            CaptureTimer.lastx = 0
                            If CaptureTimer.PenToggle Then
                                HC.PlotCap.ForeColor = vbGreen
                                CaptureTimer.PenToggle = False
                            Else
                                HC.PlotCap.ForeColor = vbRed
                                CaptureTimer.PenToggle = True
                            End If
                        End If
                        x = x * HC.PlotCap.ScaleWidth / PECCap.Period
                        If CInt(x) <> CInt(CaptureTimer.lastx) Then
                            Y = HC.PlotCap.ScaleHeight / 2 - (CaptureTimer.pe * HC.PlotCap.ScaleHeight / 180) + CaptureTimer.yoffset
                            HC.PlotCap.Line (x + 1, HC.PlotCap.ScaleHeight / 2)-(x + 1, Y) ', vbRed
                            HC.PlotCap.Line (x, 0)-(x, HC.PlotCap.ScaleHeight), vbBlack
                            HC.PlotCap.Line (CaptureTimer.lastx, CaptureTimer.lasty)-(x, Y) ', vbMagenta
                            CaptureTimer.lastx = x
                            CaptureTimer.lasty = Y
                        End If
                        Call PEC_DrawAxis(HC.PlotCap)
                    End With
                    
                    PECCap.idx = PECCap.idx + 1
                    If PECCap.idx < UBound(PECCap.CapureData()) Then
                        TimeSlip = elapsedtime - PECCap.idx
                        If TimeSlip > 0 Then
                            TimerInterval = 1000 - (TimeSlip * 1000)
                            If TimerInterval < 100 Then TimerInterval = 100
                        Else
                            TimerInterval = 1000
                        End If
                        HC.PECCapTimer.Interval = TimerInterval
                    Else
                        ' capture complete
                        CaptureTimer.State = 2
                        HC.PECCapTimer.Enabled = False
                        HC.CheckCapPec.Value = 0
                        CaptureTimer.strCapture = ""
                    End If
                Else
                    ' kill capture if not tracking
                    HC.CheckCapPec.Value = 0
                End If
                
            Case 2
                ' capture complete
                HC.PECCapTimer.Enabled = False
                CaptureTimer.strCapture = ""
                
            Case Else
                HC.CheckCapPec.Value = 0
        
        End Select
        
        CaptureTimer.timerflag = False
    End If

End Sub

Public Sub PEC_UpdateControls()
    If gPEC_Gain * 10 > PECConfigFrm.GainScroll.max Then gPEC_Gain = 1
    PECConfigFrm.GainScroll.Value = gPEC_Gain * 10
    Call PEC_GainScroll_Change
    PECConfigFrm.PhaseScroll.Value = gPEC_PhaseAdjust
    Call PEC_PhaseScroll_Change
End Sub

Public Sub PEC_ReadParams()

    Dim tmptxt As String
    Dim i As Integer
    Dim key As String
    Dim Ini As String

    Ini = HC.oPersist.GetIniPath & "\EQMOD.ini"
    key = "[pec]"
     
    tmptxt = HC.oPersist.ReadIniValueEx("WORKNG_DIR", key, Ini)
    If tmptxt <> "" Then
        gPEC_FileDir = tmptxt
    Else
       ' no value exists - create a default
        gPEC_FileDir = Environ("ProgramFiles") & "\EQMOD\PEC\"
        Call HC.oPersist.WriteIniValueEx("WORKING_DIR", gPEC_FileDir, key, Ini)
    End If
    
    tmptxt = HC.oPersist.ReadIniValueEx("INVERT_CAPTURE", key, Ini)
    If tmptxt <> "" Then
        CaptureTimer.InvertCapture = CInt(tmptxt)
    Else
       ' no value exists - create a default
        CaptureTimer.InvertCapture = 0
        Call HC.oPersist.WriteIniValueEx("INVERT_CAPTURE", CStr(CaptureTimer.InvertCapture), key, Ini)
    End If
     
    tmptxt = HC.oPersist.ReadIniValueEx("TIMESTAMP_FILES", key, Ini)
    If tmptxt <> "" Then
        gPEC_TimeStampFiles = CInt(tmptxt)
    Else
       ' no value exists - create a default
        gPEC_TimeStampFiles = 0
        Call HC.oPersist.WriteIniValueEx("TIMESTAMP_FILES", CStr(gPEC_TimeStampFiles), key, Ini)
    End If
     
     
    tmptxt = HC.oPersist.ReadIniValueEx("FILTER_LOPASS", key, Ini)
    If tmptxt <> "" Then
        gPEC_filter_lowpass = CInt(tmptxt)
        If gPEC_filter_lowpass < 10 Then gPEC_filter_lowpass = 10
    Else
       ' no value exists - create a default
        gPEC_filter_lowpass = 30
        Call HC.oPersist.WriteIniValueEx("FILTER_LOPASS", CStr(gPEC_filter_lowpass), key, Ini)
    End If
    
    tmptxt = HC.oPersist.ReadIniValueEx("FILTER_MAG", key, Ini)
    If tmptxt <> "" Then
        gPEC_mag = CInt(tmptxt)
    Else
       ' no value exists - create a default
        gPEC_mag = 10
        Call HC.oPersist.WriteIniValueEx("FILTER_MAG", CStr(gPEC_mag), key, Ini)
    End If
     
    tmptxt = HC.oPersist.ReadIniValueEx("CAPTURE_CYCLES", key, Ini)
    If tmptxt <> "" Then
        gPEC_Capture_Cycles = CInt(tmptxt)
    Else
       ' no value exists - create a default
        gPEC_Capture_Cycles = 5
        Call HC.oPersist.WriteIniValueEx("CAPTURE_CYCLES", CStr(gPEC_Capture_Cycles), key, Ini)
    End If
     
    tmptxt = HC.oPersist.ReadIniValueEx("AUTO_APPLY", key, Ini)
    If tmptxt <> "" Then
        If CInt(tmptxt) = 1 Then
            gPEC_AutoApply = 1
        Else
            gPEC_AutoApply = 0
        End If
    Else
       ' no value exists - create a default
        gPEC_AutoApply = 1
        Call HC.oPersist.WriteIniValueEx("AUTO_APPLY", CStr(gPEC_AutoApply), key, Ini)
    End If
     
     
     
'    tmptxt = HC.oPersist.ReadIniValueEx("FILTER_HIPASS", key, Ini)
'    If tmptxt <> "" Then
'        filter_hipass = CInt(tmptxt)
'    Else
'       ' no value exists - create a default
'        filter_hipass = 1000#
'        Call HC.oPersist.WriteIniValueEx("FILTER_HIPASS", CStr(filter_hipass), key, Ini)
'    End If
         
    tmptxt = HC.oPersist.ReadIniValueEx("THRESHOLD", key, Ini)
    If tmptxt <> "" Then
        threshold = CDbl(tmptxt)
    Else
       ' no value exists - create a default
        threshold = 0#
        Call HC.oPersist.WriteIniValueEx("THRESHOLD", CStr(threshold), key, Ini)
    End If
    
    tmptxt = HC.oPersist.ReadIniValueEx("GAIN", key, Ini)
    If tmptxt <> "" Then
       gPEC_Gain = CDbl(tmptxt)
    Else
       ' no value exists - create a default
       gPEC_Gain = 1#
       Call HC.oPersist.WriteIniValueEx("GAIN", CStr(gPEC_Gain), key, Ini)
    End If
   
    PEC_File = HC.oPersist.ReadIniValueEx("PEC_FILE", key, Ini)
    PECDef1.FileName = PEC_File
    If PECDef1.FileName = "" Then
'       PEC_File = "pec.txt"
       ' no value exists - create a default
       Call HC.oPersist.WriteIniValueEx("PEC_FILE", "", key, Ini)
    End If
    
    tmptxt = HC.oPersist.ReadIniValueEx("PHASE_SHIFT", key, Ini)
    If tmptxt <> "" Then
       gPEC_PhaseAdjust = CInt(tmptxt)
    Else
       ' no value exists - create a default
       Call HC.oPersist.WriteIniValueEx("PHASE_SHIFT", "0", key, Ini)
       gPEC_PhaseAdjust = 0
    End If
    
    tmptxt = HC.oPersist.ReadIniValueEx("MAX_RATEADJUST", key, Ini)
    If tmptxt <> "" Then
        gMaxRateAdjust = CDbl(tmptxt)
        If gMaxRateAdjust < 3 Then
            ' fix to increse previous default of 1
            gMaxRateAdjust = 3
            Call HC.oPersist.WriteIniValueEx("MAX_RATEDAJUST", "3", key, Ini)
        End If
    Else
       ' no value exists - create a default
       Call HC.oPersist.WriteIniValueEx("MAX_RATEDAJUST", "3", key, Ini)
       gMaxRateAdjust = 3
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("DYNAMIC_RATE_ADJUST", key, Ini)
    If tmptxt <> "" Then
        gPEC_DynamicRateAdjust = val(tmptxt)
    Else
        Call HC.oPersist.WriteIniValueEx("DYNAMIC_RATE_ADJUST", "0", key, Ini)
        gPEC_DynamicRateAdjust = 0
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("DEBUG", key, Ini)
    If tmptxt = "" Then
        gPEC_Debug = 0
        Call HC.oPersist.WriteIniValueEx("DEBUG", "0", key, Ini)
    Else
        gPEC_Debug = val(tmptxt)
    End If
    
    Select Case gPEC_Debug
        Case 1
            PECConfigFrm.CheckTracePec.Visible = True
            PECConfigFrm.PECMethodCombo.Visible = True
        Case Else
            PECConfigFrm.CheckTracePec.Visible = False
            PECConfigFrm.PECMethodCombo.Visible = False
    End Select

End Sub
Public Sub PEC_WriteParams()
    Dim key As String
    Dim Ini As String

    Ini = HC.oPersist.GetIniPath & "\EQMOD.ini"
    key = "[pec]"
    Call HC.oPersist.WriteIniValueEx("THRESHOLD", CStr(threshold), key, Ini)
    Call HC.oPersist.WriteIniValueEx("GAIN", CStr(gPEC_Gain), key, Ini)
    Call HC.oPersist.WriteIniValueEx("PEC_FILE", PECDef1.FileName, key, Ini)
    Call HC.oPersist.WriteIniValueEx("PHASE_SHIFT", CStr(PECConfigFrm.PhaseScroll.Value), key, Ini)
    Call HC.oPersist.WriteIniValueEx("CAPTURE_CYCLES", CStr(gPEC_Capture_Cycles), key, Ini)
    Call HC.oPersist.WriteIniValueEx("FILTER_LOPASS", CStr(gPEC_filter_lowpass), key, Ini)
    Call HC.oPersist.WriteIniValueEx("FILTER_MAG", CStr(gPEC_mag), key, Ini)
    Call HC.oPersist.WriteIniValueEx("TIMESTAMP_FILES", CStr(gPEC_TimeStampFiles), key, Ini)
    Call HC.oPersist.WriteIniValueEx("AUTO_APPLY", CStr(gPEC_AutoApply), key, Ini)
    Call HC.oPersist.WriteIniValueEx("WORKING_DIR", gPEC_FileDir, key, Ini)

End Sub

Public Function NormalisePosition(ByVal Position As Double, wormsteps As Double) As Double
' Normalisation is intended for raw stepper position which are centered
' around H80000. Once normalised the range will be 0-50132.

    If Position > wormsteps Then
'        While position < gRAEncoder_Zero_pos
'            position = position + gTot_RA
'        Wend
'        position = position - gRAEncoder_Zero_pos
        NormalisePosition = Position Mod (wormsteps)
    Else
        ' don't take into account the H80000 ofset if the position
        ' is already normalised!
        NormalisePosition = Position
    End If

End Function
Private Function Import(PECDef As PECDefinition) As Boolean
Dim temp1 As String
Dim temp2 As String
Dim lineno As Integer
Dim idx As Integer
Dim pos As Integer
Dim CurveMin As PECData
Dim CurveMax As PECData
Dim ratesum As Double
Dim MotorPos As Double
Dim CycleCount As Integer
Dim error As Boolean
Dim drift As Double
Dim wp As Double
Dim NF1 As Integer

    On Error GoTo ImportError
    
    NF1 = FreeFile
    error = False
    
    If PECDef.FileName = "" Then
        error = True
        GoTo errCheck
    End If
    
    Close #NF1
    Open PECDef.FileName For Input As NF1
    
    lineno = 0
    idx = 0
    While Not EOF(NF1)
        Line Input #NF1, temp1
        If lineno > 0 Then
            If Left$(temp1, 1) = "!" Then
                ' parse parameters
                pos = InStr(temp1, "=")
                If pos <> 0 Then
                    temp2 = Left$(temp1, pos - 1)
                    If temp2 = "!WormPeriod" Then
                        temp1 = Right$(temp1, Len(temp1) - pos)
                        wp = Int(CDbl(temp1) + 0.5)
                        PECDef.Period = wp
                        ReDim PECDef.PECCurve(wp)
                        ReDim PECDef.PECCurveTmp(wp)
                        For idx = 0 To (wp - 1)
                            PECDef.PECCurve(idx).signal = 0
                        Next idx
                        idx = 0
                        ' apply a default if steps per worm isn't in the pec file
                        PECDef.Steps = gRAWormSteps
                    Else
                        If temp2 = "!StepsPerWorm" Then
                            temp1 = Right$(temp1, Len(temp1) - pos)
                            PECDef.Steps = Int(CDbl(temp1) + 0.5)
                        End If
                    End If
                End If
            Else
                If Left$(temp1, 1) <> "#" Then
                    With PECDef.PECCurve(idx)
                        ' replace tabs with spaces
                        temp1 = Replace(temp1, Chr(9), " ")
                        pos = InStr(temp1, " ")
                        If pos <> 0 Then
                            temp2 = Left$(temp1, pos - 1)
                            temp1 = Right$(temp1, Len(temp1) - pos)
                            .time = CDbl(temp2)
                            
                            pos = InStr(temp1, " ")
                            If pos <> 0 Then
                                temp2 = Left$(temp1, pos - 1)
                                temp1 = Right$(temp1, Len(temp1) - pos)
                                If CycleCount = 0 Then
                                    ' store the motor positions for the first cycle
                                    MotorPos = CDbl(temp2)
                                    .RawPosn = MotorPos
                                    .PEPosition = NormalisePosition(Int(MotorPos), PECDef.Steps)
                                End If
                                .signal = (.signal + CDbl(temp1))
                                .cycle = CycleCount + 1
                            End If
                            idx = idx + 1
                            If idx = wp Then
                                CycleCount = CycleCount + 1
                                idx = 0
                            End If
                        End If
                    End With
                End If
            End If
        End If
        lineno = lineno + 1
    Wend

closefile:
    Close #NF1
    If error Then GoTo errCheck
    
    If CycleCount >= 1 Then
        ' average the signal
        For idx = 0 To (PECDef.Period - 1)
            With PECDef.PECCurve(idx)
                .signal = .signal / .cycle
            End With
        Next idx
        
        ' remove any net cycle offset from the PEC curve
        drift = (PECDef.PECCurve(PECDef.Period - 1).signal - PECDef.PECCurve(0).signal) / (PECDef.Period + 1)
        CurveMin.signal = 100
        CurveMax.signal = -100
        For idx = 0 To (PECDef.Period - 1)
            With PECDef.PECCurve(idx)
                .signal = .signal - idx * drift
                If .signal > CurveMax.signal Then
                    CurveMax = PECDef.PECCurve(idx)
                End If
                If .signal < CurveMin.signal Then
                    CurveMin = PECDef.PECCurve(idx)
                End If
            End With
        Next idx
        
        PECDef.CurrIdx = 0
        
        PECDef.MaxPe = CurveMax.signal
        PECDef.MinPe = CurveMin.signal
        Call PEC_PlotCurve(PECDef)
        
        ' caluculate correction rates to be used.
        error = CalcRates(PECDef)
        
    Else
        HC.Add_Message ("PECImport: Insufficient PEC samples!")
        error = True
    End If
    
    GoTo errCheck
    
ImportError:
    Close #NF1
    error = True
    HC.Add_Message ("PECImport: ErrNo." & Err.Number)
    HC.Add_Message (Err.Description)
    Err.Clear
    
errCheck:
    If error Then
        Import = False
    Else
        HC.PECTimer.Enabled = True
        HC.CmdPecSave.Enabled = True
        PECConfigFrm.GainScroll.Enabled = True
        PECConfigFrm.PhaseScroll.Enabled = True
        HC.CheckPEC.Enabled = True
        HC.CheckPEC.Value = 1
        HC.CmdTrack(1).Enabled = True
        Import = True
        ' set the PEC frame caption to show file name
        pos = InStrRev(PEC_File, "\")
        temp1 = Right(PEC_File, Len(PEC_File) - pos)
        HC.Frame9.Caption = oLangDll.GetLangString(19) & " " & temp1
    End If

End Function

Private Sub KillPec()
        HC.Add_Message ("PEC: Disabled")
        HC.Frame9.Caption = oLangDll.GetLangString(19)
        HC.CmdTrack(1).Visible = False
        HC.PECTimer.Enabled = False
        HC.CmdPecSave.Enabled = False
        PECConfigFrm.GainScroll.Enabled = False
        PECConfigFrm.PhaseScroll.Enabled = False
        HC.CheckPEC.Enabled = False
        HC.CheckPEC.Value = 0
        gPEC_Enabled = False
        HC.plot.Cls
End Sub

Private Function Export(PECDef As PECDefinition, phaseshift As Integer) As Boolean

Dim temp1 As String
Dim idx As Integer
Dim pos As Integer
Dim NF1 As Integer

    On Error GoTo exporterr
    Export = True
    NF1 = FreeFile
    
    Close #NF1
    Open PECDef.FileName For Output As NF1
'        Print NF1, Date$ & " " & Time$
    Print #NF1, "# " & HC.MainLabel.Caption
    Print #NF1, "!WormPeriod=" & CStr(PECDef.Period)
    Print #NF1, "!StepsPerWorm=" & CStr(PECDef.Steps)
    Print #NF1, "# time - motor - smoothed PE"
    For idx = 0 To UBound(PECDef.PECCurve) - 1
        ' apply local phase shift
        pos = (idx + phaseshift) Mod PECDef.Period
        Print #NF1, FormatNumber(idx, 0, , , 0) & " " & FormatNumber(PECDef.PECCurve(idx).PEPosition, 0, , , 0) & " " & FormatNumber(PECDef.PECCurve(pos).signal, 4, , , 0)
    Next idx
    GoTo endexport
exporterr:
    Export = False
endexport:
    Close #NF1

End Function

Public Function PEC_Write_Table(Index As Double, Position As Double, signal As Double) As Boolean
Dim i As Integer
    If Index = 0 Then
        ' first element being written - clear out existing data
        ReDim PECDef1.PECCurveTmp(PECDef1.Period)
    End If

    If Index >= UBound(PECDef1.PECCurveTmp) Then
        PEC_Write_Table = False
        Exit Function
    End If
        
    If Position > PECDef1.Steps Then
        PEC_Write_Table = False
        Exit Function
    End If
        
    ' data is only written to the temporary curve
    With PECDef1.PECCurveTmp(Index)
        .time = Index
        .signal = signal
        .PEPosition = Position
        .RawPosn = Position
    End With
    
    If Index = UBound(PECDef1.PECCurveTmp) - 1 Then
        ' when the last index is written we save the file.
        PEC_File = HC.oPersist.GetIniPath & "\PEC.txt"
        PECDef1.FileName = PEC_File
        ' copy temp store to real curve
        For i = 0 To PECDef1.Period - 1
            PECDef1.PECCurve(i) = PECDef1.PECCurveTmp(i)
        Next i
        ' write curve to file
        Export PECDef1, 0
        ' save PEC file name
        Call PEC_WriteParams
        ' load from file
        If Import(PECDef1) = True Then
            ' update display
            Call PEC_PlotCurve(PECDef1)
        Else
            KillPec
        End If
    End If
    PEC_Write_Table = True
            
End Function

Private Sub PEC_DrawAxis(plot As PictureBox)
Dim mid As Integer
    mid = plot.ScaleHeight / 2
    plot.Line (0, mid)-(plot.ScaleWidth, mid), &H80FF&
    plot.Line (0, (mid) - 4)-(0, mid + 2), &H80FF&
    plot.Line (plot.ScaleWidth * 0.25, mid - 2)-(plot.ScaleWidth * 0.25, mid + 2), &H80FF&
    plot.Line (plot.ScaleWidth * 0.5, mid - 2)-(plot.ScaleWidth * 0.5, mid + 2), &H80FF&
    plot.Line (plot.ScaleWidth * 0.75, (mid) - 2)-(plot.ScaleWidth * 0.75, mid + 2), &H80FF&
    plot.Line (plot.ScaleWidth - 1, mid - 2)-(plot.ScaleWidth - 1, mid + 2), &H80FF&
End Sub

Private Sub PEC_PlotCurve(PECDef As PECDefinition)
Dim idx As Integer
Dim oldval, newval As Double
Dim range, hscale As Double
Dim mid As Integer
   
    range = PECDef.MaxPe - PECDef.MinPe
    HC.plot.Cls
    Call PEC_DrawAxis(HC.plot)
    mid = HC.plot.ScaleHeight / 2
    
    hscale = HC.plot.ScaleWidth / PECDef.Period
    If range > 0 Then
        oldval = mid - PECDef.PECCurve(0).signal * 0.8 * HC.plot.ScaleHeight / range
        For idx = 1 To (PECDef.Period - 1)
            newval = mid - PECDef.PECCurve(idx).signal * 0.8 * HC.plot.ScaleHeight / range
            HC.plot.Line (idx * hscale, newval)-((idx - 1) * hscale, oldval), vbMagenta
            oldval = newval
        Next idx
    End If
End Sub

Private Function CalcRates(PECDef As PECDefinition) As Boolean
Dim idx, sec As Integer
Dim ratesum As Double
Dim rate, lastrate, truerate, remainder As Double
Dim newpos As Double
Dim StepsMoved As Double
Dim debugfile As String
Dim debugmark As Integer
Dim i As Integer
Dim NF1 As Integer

    debugmark = 0
    debugfile = HC.oPersist.GetIniPath() + "\pec_rates_" & CStr(PECDef.Period) & ".txt"
    On Error GoTo endsub
    NF1 = FreeFile
    
    Close #NF1
    Open debugfile For Output As NF1
    On Error GoTo handle_error
    
    ' calculate the rate change between each PE curve sample
    ratesum = 0
    For idx = 1 To (PECDef.Period - 1)
        With PECDef.PECCurve(idx)
            .PErate = PECDef.PECCurve(idx - 1).signal - .signal
            ratesum = ratesum + .PErate
        End With
    Next idx
    ' first rate = average of 2nd and last to remove any discontinuities
    PECDef.PECCurve(0).PErate = (PECDef.PECCurve(PECDef.Period - 1).PErate + PECDef.PECCurve(1).PErate) / 2
    ratesum = ratesum + PECDef.PECCurve(0).PErate
 
    ' Apply the current threshhold and gain settings
    ' The threshold setting allows us to reduce the number of rate corrections sent
    ' to the mount.
    ' The gain is just a user 'fiddle' factor ;-)
    debugmark = 1
    ratesum = 0
    lastrate = PECDef.PECCurve(PECDef.Period - 1).PErate
    For idx = 0 To (PECDef.Period - 1)
        rate = PECDef.PECCurve(idx).PErate
        If Abs(lastrate - rate) > threshold Then
            lastrate = PECDef.PECCurve(idx).PErate
        Else
            rate = lastrate
        End If
        rate = rate * gPEC_Gain
        PECDef.PECCurve(idx).PECrate = rate
        ratesum = ratesum + rate
    Next idx
    
    ' The sum of all rate changes over a single cycle should be 0 i.e. sidereal rate
    ' If this isn't the case then adjust them accordingly to ensure there is no net drift
    debugmark = 2
    For idx = 0 To (PECDef.Period - 1)
        PECDef.PECCurve(idx).PECrate = PECDef.PECCurve(idx).PECrate - (ratesum / PECDef.Period)
    Next idx
    
    ' unfortunately the way the mount accepts 'quantised' rate change messages means that we can't have
    ' just any rate we want. So work out what the rate will the mount will actually track at
    ' determine any error and attemp to correct for it in the next sample.
    ' Using this approach we should be able to acheve at worst sidereal - 0.024 arcsecs/sec
    debugmark = 3
    remainder = 0
    For i = 1 To 3
        ratesum = 0
        For idx = 0 To (PECDef.Period - 1)
            With PECDef.PECCurve(idx)
                If gHemisphere = 0 Then
                    rate = SID_RATE_NORTH + .PECrate + remainder
        '            truerate = (9325.46154 / (Int((9325.46154 / RATE) + 0.5)))
                    truerate = (gTrackFactorRA / (Int((gTrackFactorRA / rate) + 0.5)))
                    ' work out the error for next time
                    remainder = rate - truerate
                    .PECrate = truerate - SID_RATE_NORTH
                    ratesum = ratesum + .PECrate
                Else
                    rate = SID_RATE_SOUTH + .PECrate + remainder
        '            truerate = (9325.46154 / (Int((9325.46154 / RATE) - 0.5)))
                    truerate = (gTrackFactorRA / (Int((gTrackFactorRA / rate) - 0.5)))
                    remainder = rate - truerate
                    .PECrate = truerate - SID_RATE_SOUTH
                    ratesum = ratesum + .PECrate
                End If
            End With
        Next idx
    Next i
    
    For idx = 0 To (PECDef.Period - 1)
        If gHemisphere = 0 Then
            PECDef.PECCurve(idx).PECrate = PECDef.PECCurve(idx).PECrate + SID_RATE_NORTH
        Else
            PECDef.PECCurve(idx).PECrate = PECDef.PECCurve(idx).PECrate + SID_RATE_SOUTH
        End If
    Next idx
   
    
    ' now we know what rates the mount will be tracking at we can
    ' calculate the expected motor positions at each sample
    debugmark = 4
    PECDef.PECCurve(0).PECPosition = PECDef.PECCurve(0).PEPosition
    lastrate = 0
    For idx = 1 To (PECDef.Period - 1)
        ' motorpos = lastmotorpos + elapsedtime * gTot_RA /  (ARCSECS_PER_360DEGREES / lastRate)
        rate = PECDef.PECCurve(idx - 1).PECrate
        If rate = 0 Then
            rate = lastrate
        Else
            lastrate = rate
        End If
        
        StepsMoved = gTot_RA / (ARCSECS_PER_360DEGREES / rate)
        newpos = PECDef.PECCurve(idx - 1).PECPosition + StepsMoved
        PECDef.PECCurve(idx).PECPosition = newpos Mod (PECDef.Steps)
    Next idx
    
    ' A lot has gone on here so write out a debug file
    ' for anaysis if things don't work as the should.
    debugmark = 5
    Print #NF1, "Index PE PosRawPE PosPE PosPEC RatePE RatePEC"
    For idx = 0 To (PECDef.Period - 1)
        With PECDef.PECCurve(idx)
            Print #NF1, CStr(idx) & " " & _
                      FormatNumber(.signal, 4) & " " & _
                      FormatNumber(.RawPosn, 0) & " " & _
                      FormatNumber(.PEPosition, 0) & " " & _
                      FormatNumber(.PECPosition, 0) & " " & _
                      FormatNumber(.PErate, 4) & " " & _
                      FormatNumber(.PECrate, 4)
        End With
    Next idx
    Print #NF1, "RateSum=" & FormatNumber(ratesum, 4)
    CalcRates = False
    GoTo endsub
    
handle_error:
    Print #NF1, "ERROR NUMBER=" & Err.Number
    Print #NF1, "ERROR DESCRIPTION=" & Err.Description
    Print #NF1, "CodeTrace=" & debugmark
    Print #NF1, "idx=" & idx
    Print #NF1, "PECDef.PECCurve LBound=" & CStr(LBound(PECDef1.PECCurve))
    Print #NF1, "PECDef.PECCurve UBound=" & CStr(UBound(PECDef1.PECCurve))
    Print #NF1, "Period=" & CStr(PECDef.Period)
    Err.Clear
    CalcRates = True
endsub:
    Close #NF1

End Function
        
Private Function GetIdx(PECDef As PECDefinition) As Integer
Dim MotorPos As Double
Dim curvepos As Double
Dim i, idx As Integer
    ' Determine an appropriate index into the PEC table that gives the
    ' best match for the current motor position. The best match is the
    ' position that is one step in advance of the motor.
    ' Generally this will just result in an incermenting of the index so
    ' a starting position is suplied to speed up the number of comparisons
    ' required.
    
    idx = 0

    ' if this PEC definition is in use
    If PECDef.Period <> 0 Then

        idx = PECDef.CurrIdx
        
        If gPEC_Enabled Then
            ' PEC tacking - sync with PEC calculated motor positions.
            curvepos = PECDef.PECCurve(idx).PECPosition
        Else
            ' sidereal track so use uncorrected positions
            If PECDef.Period <> 0 Then
                curvepos = PECDef.PECCurve(idx).PEPosition
            Else
                idx = 0
                GoTo endfunc
            End If
        End If
        
        i = 0
        If gHemisphere = 0 Then
            ' Northern hemisphere
            ' For northern hemisphere curves the motor positions increase
            ' with increasing index
            MotorPos = NormalisePosition(PlaybackTimer.ringcounter + phaseshift, PECDef.Steps)
            If (MotorPos > curvepos) Then
                While MotorPos > curvepos And i < PECDef.Period
                    ' search forwards till we find a curve position that is
                    ' greater than the motor position
                    idx = (idx + 1) Mod PECDef.Period
                    curvepos = PECDef.PECCurve(idx).PECPosition
                    i = i + 1
                Wend
            Else
                While MotorPos < curvepos And i < PECDef.Period
                    ' search backwards till we find a curve position that is
                    ' less than the motor position
                    idx = idx - 1
                    If idx < 0 Then idx = (PECDef.Period - 1)
                    curvepos = PECDef.PECCurve(idx).PECPosition
                    i = i + 1
                Wend
                ' now increment to next curve position
                ' its best to have the curve in advance of the motor!
                idx = (idx + 1) Mod PECDef.Period
            End If
        Else
            ' Southern hemisphere
            ' For southern hemisphere curves the motor positions decrease
            ' with increasing index
            MotorPos = NormalisePosition(PlaybackTimer.ringcounter - phaseshift, PECDef.Steps)
            If (MotorPos > curvepos) Then
                While MotorPos > curvepos And i < PECDef.Period
                    ' search backwards till we find a curve position that is
                    ' smaller than the motor position
                    idx = idx - 1
                    If idx < 0 Then idx = (PECDef.Period - 1)
                    curvepos = PECDef.PECCurve(idx).PECPosition
                    i = i + 1
                Wend
                ' now increment to next curve position
                ' its best to have the curve in advance of the motor!
                idx = (idx + 1) Mod PECDef.Period
            Else
                ' search forwards till we find a curve position that is
                ' greater than the motor position
                While MotorPos < curvepos And i < PECDef.Period
                    idx = (idx + 1) Mod PECDef.Period
                    curvepos = PECDef.PECCurve(idx).PECPosition
                    i = i + 1
                Wend
            End If
        End If
    End If
endfunc:
    GetIdx = idx
    PlaybackTimer.PecResyncCount = 0
End Function

Public Sub PEC_MoveAxis(axis As Double, rate As Double)

'    If rate <> 0 Then HC.TrackingFrame.Caption = oLangDll.GetLangString(121) & " " & oLangDll.GetLangString(188)
    If axis = 0 Then
        If (rate = 0) And (gDeclinationRate = 0) Then HC.TrackingFrame.Caption = oLangDll.GetLangString(121) & " " & oLangDll.GetLangString(178)
        If gEQRAPulseDuration = 0 Then
            If (gRightAscensionRate * rate) <= 0 Or gTrackingStatus <> 1 Then
                Call StartRA_by_Rate(rate)
            Else
                Call ChangeRA_by_Rate(rate)
            End If
        End If
        gRightAscensionRate = rate
        gTrackingStatus = 1
        gRA_LastRate = rate
    End If
End Sub

Public Sub PEC_StartCapture()
    If gTrackingStatus = 1 Then
'        Call KillPec
        CaptureTimer.State = 0
        HC.PECCapTimer.Enabled = True
    Else
        ' can't capture if not tacking!
        HC.CheckCapPec.Value = 0
    End If
End Sub

Public Sub PEC_StopCapture()
    HC.PECCapTimer.Enabled = False
    If CaptureTimer.State = 2 Then
        ' capture has completed
        Call SaveCaptureData
    Else
        ' capture has been aborted
    End If
    CaptureTimer.State = 0
End Sub

Private Sub SaveCaptureData()

Dim temp1 As String
Dim idx As Integer
Dim PECIdx As Integer
Dim pos As Integer
Dim pe As Double
Dim PEC_Data() As PECFileData
Dim FileName As String
Dim NF1 As Integer

    On Error GoTo exporterr
    
    ' create a capture file for debug
    ReDim PEC_Data(PECCap.Period)
    
    ' clear pe
    For idx = 0 To UBound(PEC_Data) - 1
        PEC_Data(idx).pe = 0
        PEC_Data(idx).cycle = 0
    Next idx

    
    ' linearly regress to remove drifts and apply fft smoothing
    PEC_RegressAndSmooth
               
    If gPEC_TimeStampFiles = 1 Then
        FileName = gPEC_FileDir & "pecapture_" & GetTimeStamp & "_EQMOD.txt"
    Else
        FileName = gPEC_FileDir & "pecapture_EQMOD.txt"
    End If
    
    NF1 = FreeFile
    Close #NF1
    Open FileName For Output As NF1
    
    ' output a perecorder format type file of the raw data
    Print #NF1, "# " & HC.MainLabel.Caption
    Print #NF1, "# AUTO-PEC"
    Print #NF1, "# RA  = " & CStr(gRA)
    Print #NF1, "# DEC = " & CStr(gDec)
    If gAscomCompatibility.AllowPulseGuide Then
        Print #NF1, "# PulseGuide, Rate=" & CStr(HC.HScrollRARate.Value * 0.1)
    Else
        Print #NF1, "# ST-4 Guide, Rate=" & HC.RAGuideRateList.Text
    End If
    Print #NF1, "!WormPeriod=" & CStr(PECCap.Period)
    Print #NF1, "!StepsPerWorm=" & CStr(PECCap.Steps)
    Print #NF1, "#Time MotorPosition PE"
               
    ' Average signals to get a single cycle error signal
    For idx = 0 To PECCap.idx - 1 'UBound(PECCap.CapureData) - 1
        With PECCap.CapureData(idx)
            PECIdx = idx Mod PECCap.Period
            PEC_Data(PECIdx).Position = NormalisePosition(.MotorPos, PECCap.Steps)
            ' ignore first and last 120 samples as fft filter may exagerate their data
            If idx > 120 And idx < PECCap.idx - 120 Then
'              PEC_signal(PECIdx) = PEC_signal(PECIdx) + .peSmoothed / (gPEC_Capture_Cycles)
                pe = PEC_Data(PECIdx).pe * PEC_Data(PECIdx).cycle + .peSmoothed
                PEC_Data(PECIdx).cycle = PEC_Data(PECIdx).cycle + 1
                PEC_Data(PECIdx).pe = pe / PEC_Data(PECIdx).cycle
            End If
            Print #NF1, FormatNumber(.time, 3, , , 0) & " " & CStr(.MotorPos) & " " & FormatNumber(.pe, 4, , , 0)
        End With
    Next idx
    Close #NF1
   
       
    ' generate pec file
    If gPEC_TimeStampFiles = 1 Then
        FileName = gPEC_FileDir & "pec_" & GetTimeStamp & ".txt"
    Else
        FileName = gPEC_FileDir & "pec.txt"
    End If
    NF1 = FreeFile
    Close #NF1
    Open FileName For Output As NF1
    Print #NF1, "# " & HC.MainLabel.Caption
    Print #NF1, "!WormPeriod=" & CStr(PECCap.Period)
    Print #NF1, "!StepsPerWorm=" & CStr(PECCap.Steps)
    Print #NF1, "# time - motor - smoothed PE"
    For idx = 0 To UBound(PEC_Data) - 1
        pe = PEC_Data(idx).pe
        Print #NF1, FormatNumber(idx, 0, , , 0) & " " & FormatNumber(PEC_Data(idx).Position, 0, , , 0) & " " & FormatNumber(pe, 4, , , 0)
    Next idx
    Close #NF1

    ' load pec
    If gPEC_AutoApply = 1 Then
        PEC_LoadFile FileName
        ' set pec gain to x1
        PECConfigFrm.GainScroll.Value = 10
        ' set phase shift to 0
        PECConfigFrm.PhaseScroll.Value = 0
    End If
    GoTo endexport

exporterr:
    Close #NF1
endexport:

End Sub


Private Sub PEC_RegressAndSmooth()

Dim xy_sum As Double
Dim xx_sum As Double
Dim x_sum As Double
Dim y_sum As Double
Dim tmp As Double
Dim i As Long
Dim RawDataSize As Double
Dim slope As Double
Dim intercept As Double
Dim NF1 As Integer

    On Error GoTo endsub
    
    xy_sum = 0
    xx_sum = 0
    y_sum = 0
    x_sum = 0
    RawDataSize = UBound(PECCap.CapureData)
    For i = 0 To PECCap.idx - 1
        xy_sum = xy_sum + (i * PECCap.CapureData(i).pe)
        x_sum = x_sum + i
        y_sum = y_sum + PECCap.CapureData(i).pe
        xx_sum = xx_sum + (i * i)
    Next i
    
    ' Calculate slope of linear regression line
    slope = ((RawDataSize * xy_sum) - (x_sum * y_sum)) / ((RawDataSize * xx_sum) - (x_sum * x_sum))

    ' Calculate intercept of linear regression line
    intercept = (y_sum - (slope * x_sum)) / RawDataSize
    
    ' initalise fft
    Call FFT_Initialise(4096, 1)
    
    ' remove slope from data and store in fft time domain
    For i = 0 To RawDataSize - 1
        tmp = PECCap.CapureData(i).pe - (slope * i) - intercept
'        PECCap.CapureData(i).pe = tmp
        ' add to time domain
        Call FFT_SetSample(CInt(i), tmp)
    Next i
    
    ' generate frequency domain
    Call FFT_ForwardFFTComplex
    Call FFT_NormaliseMag
    
    ' filter out anything with a relative magnitude of 10% or less, and anything with a period < 33 sec or period > 1.5*worm period
'    Call FFT_ApplyFilter(0.03, 1 / (2 * PECCap.Period), 10)
    Call FFT_ApplyFilter(1 / CDbl(gPEC_filter_lowpass), 1 / (1.5 * PECCap.Period), CDbl(gPEC_mag))
    
    ' generate new time domain
    FFT_InverseFFTComplex
    
    If gPEC_Debug = 1 Then
        ' create a capture file for debug
        PECCap.FileName = gPEC_FileDir & "PECCapture_" & GetTimeStamp & ".txt"
        NF1 = FreeFile
        Close #NF1
        Open PECCap.FileName For Output As NF1
        Print #NF1, "# " & HC.MainLabel.Caption
        Print #NF1, "!WormPeriod=" & CStr(PECCap.Period)
        Print #NF1, "!StepsPerWorm=" & CStr(PECCap.Steps)
        Print #NF1, "# RA  = " & CStr(gRA)
        Print #NF1, "# DEC = " & CStr(gDec)
        If gAscomCompatibility.AllowPulseGuide Then
            Print #NF1, "# PulseGuide, Rate=" & CStr(HC.HScrollRARate.Value * 0.1)
        Else
            Print #NF1, "# ST-4 Guide, Rate=" & HC.RAGuideRateList.Text
        End If
        Print #NF1, "#Idx Time DeltaTime MotorPos DeltaPos Rate DeltaPE RawPE SmothedPE"
    End If
    
    ' store as smoothed capture data.
    For i = 0 To RawDataSize - 1
        With PECCap.CapureData(i)
            .peSmoothed = FFT_GetSample(CInt(i))
            If gPEC_Debug = 1 Then
                Print #NF1, FormatNumber(i, 0, , , 0) & " " & FormatNumber(.time, 3, , , 0) & " " & FormatNumber(.DeltaTime, 3, , , 0) & " " & CStr(.MotorPos) & " " & CStr(.DeltaPos) & " " & FormatNumber(.rate, 4, , , 0) & " " & FormatNumber(.peInc, 4, , , 0) & " " & FormatNumber(.pe, 4, , , 0) & " " & FormatNumber(.peSmoothed, 4, , , 0)
            End If
        End With
    Next i
    
    Close #NF1

    FFT_Free

endsub:
End Sub

Private Function GetTimeStamp() As String

    GetTimeStamp = Date$ & time$
    GetTimeStamp = Replace(GetTimeStamp, ":", "")
    GetTimeStamp = Replace(GetTimeStamp, "\", "")
    GetTimeStamp = Replace(GetTimeStamp, "/", "")
    GetTimeStamp = Replace(GetTimeStamp, " ", "")
    GetTimeStamp = Replace(GetTimeStamp, "-", "")

End Function

Public Sub PEC_DispalyUpdate(ByRef plot As PictureBox)
    plot.Cls
    plot.FontSize = 8
    plot.Print oLangDll.GetLangString(191) & " = " & CStr(gPEC_Gain)
    plot.FontSize = 12
    plot.Print PlaybackTimer.strPlayback
    plot.Print CaptureTimer.strCapture

End Sub

