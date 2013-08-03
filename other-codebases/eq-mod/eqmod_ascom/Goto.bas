Attribute VB_Name = "Goto"
Option Explicit

Type GOTO_PARAMS
    RA_currentencoder As Double
    RA_Direction As Integer
    RA_targetencoder As Double
    RA_SlewActive As Integer
    DEC_currentencoder As Double
    DEC_Direction As Integer
    DEC_targetencoder As Double
    DEC_SlewActive As Integer
    rate As Integer
    SuperSafeMode As Integer
End Type

Public gGotoParams As GOTO_PARAMS
Public gGotoRate As Integer
Public gDisbleFlipGotoReset As Integer
Public gCWUP As Boolean
Public gMaxSlewCount As Integer
Public gSlewCount As Long
Public gFRSlewCount As Integer
Public gGotoResolution As Integer
Public gTargetRA As Double
Public gTargetDec As Double
Public gRAGotoRes As Double    ' Iterative Slew minimum difference in arcsecs
Public gDECGotoRes As Double   ' Iterative Slew minimum difference in arcsecs
Public gRA_Compensate As Long  ' Least RA discrepancy Compensation
Public gRAMeridianWest As Double
Public gRAMeridianEast As Double


'Routine to Slew the mount to target location
Public Sub radecAsyncSlew(ByVal GotoRate As Integer)

    HC.EncoderTimer.Enabled = False
    With gGotoParams
        Call CalcEncoderTargets
        .rate = GotoRate
        If gCWUP Then
            gSupressHorizonLimits = True
            ' a counterweights up slew has been requested
            If RALimitsActive() = False Then
                ' Limits are off so play safe and slew RA and DEC independently
                If gRA_Hours > 12 Then
                    ' we're currently in a counterweights up position
                    If .RA_currentencoder > RAEncoder_Home_pos Then
                        ' single axis slew to nearest limit position
                        ' followed by dual axis slew to target limit
                        ' followed by single axis slew to target ra
                        .SuperSafeMode = 3
                        Call StartSlew(gRAMeridianWest, .DEC_currentencoder, .RA_currentencoder, .DEC_currentencoder)
                    Else
                        ' single axis slew to nearest limit position
                        ' followed by dual axis slew to target limit
                        ' followed by single axis slew to target ra
                        .SuperSafeMode = 3
                        Call StartSlew(gRAMeridianEast, .DEC_currentencoder, .RA_currentencoder, .DEC_currentencoder)
                    End If
                Else
                    ' we're currently in a counterweights down position
                    If .RA_targetencoder > RAEncoder_Home_pos Then
                        ' dual axis slew to limit position followed by ra only slew to target
                        .SuperSafeMode = 1
                        Call StartSlew(gRAMeridianWest, .DEC_targetencoder, .RA_currentencoder, .DEC_currentencoder)
                    Else
                        ' dual axis slew to limit position followed by ra only slew to target
                        .SuperSafeMode = 1
                        Call StartSlew(gRAMeridianEast, .DEC_targetencoder, .RA_currentencoder, .DEC_currentencoder)
                    End If
                End If
            Else
                ' Limits are active so allow simulatenous RA/DEC movement
                .SuperSafeMode = 0
                Call StartSlew(.RA_targetencoder, .DEC_targetencoder, .RA_currentencoder, .DEC_currentencoder)
            End If
        Else
            ' we're currently in a counterweights up position
            If RALimitsActive() = False Then
                ' Limits are off
                If .RA_currentencoder > gRAMeridianWest Then
                    'Slew in RA to limit position - then complete move as dual axis slew
                    .SuperSafeMode = 1
                    gSupressHorizonLimits = True
                    Call StartSlew(gRAMeridianWest, .DEC_currentencoder, .RA_currentencoder, .DEC_currentencoder)
                Else
                    If .RA_currentencoder < gRAMeridianEast Then
                        'Slew in RA to limit position - then complete move as dual axis slew
                        .SuperSafeMode = 1
                        gSupressHorizonLimits = True
                        Call StartSlew(gRAMeridianEast, .DEC_currentencoder, .RA_currentencoder, .DEC_currentencoder)
                    Else
                        ' standard slew - simulatanous RA and DEc movement
                        .SuperSafeMode = 0
                        Call StartSlew(.RA_targetencoder, .DEC_targetencoder, .RA_currentencoder, .DEC_currentencoder)
                    End If
                End If
            Else
                ' Limits are enabled
                If .RA_currentencoder > gRA_Limit_West Then
                    'Slew in RA to limit position - then complete move as dual axis slew
                    .SuperSafeMode = 1
                    gSupressHorizonLimits = True
                    Call StartSlew(gRA_Limit_West, .DEC_currentencoder, .RA_currentencoder, .DEC_currentencoder)
                Else
                    If .RA_currentencoder < gRA_Limit_East Then
                        'Slew in RA to limit position - then complete move as dual axis slew
                        .SuperSafeMode = 1
                        gSupressHorizonLimits = True
                        Call StartSlew(gRA_Limit_East, .DEC_currentencoder, .RA_currentencoder, .DEC_currentencoder)
                    Else
                        ' standard slew - simulatanous RA and DEc movement
                        .SuperSafeMode = 0
                        Call StartSlew(.RA_targetencoder, .DEC_targetencoder, .RA_currentencoder, .DEC_currentencoder)
                    End If
                End If
            End If
        End If
    End With
    HC.EncoderTimer.Enabled = True

End Sub

Public Sub CalcEncoderTargets()
Dim targetRAEncoder As Double
Dim targetDECEncoder As Double
Dim currentRAEncoder As Double
Dim currentDECEncoder As Double
Dim tmpcoord As Coordt
Dim DeltaRAStep As Long
Dim DeltaDECStep As Long
Dim RASlowdown As Long
Dim DECSlowdown As Long
Dim tRA As Double
Dim tha As Double
Dim tPier As Double

    On Error GoTo endradecslew

    gSlewStatus = False
   
    'stop the motors
    PEC_StopTracking
    eqres = EQ_MotorStop(0)
    eqres = EQ_MotorStop(1)
      
    'Wait for motor stop , Need to add timeout routines here
    Do
        eqres = EQ_GetMotorStatus(0)
        If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo SL01
    Loop While (eqres And EQ_MOTORBUSY) <> 0

SL01:
    Do
        eqres = EQ_GetMotorStatus(1)
        If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo SL02
    Loop While (eqres And EQ_MOTORBUSY) <> 0

SL02:
 
    ' read current
    currentRAEncoder = EQGetMotorValues(0)
    currentDECEncoder = EQGetMotorValues(1)

    tha = RangeHA(gTargetRA - EQnow_lst(gLongitude * DEG_RAD))
    If tha < 0 Then
        If gCWUP Then
            If gHemisphere = 0 Then
                tPier = 0
            Else
                tPier = 1
            End If
            tRA = gTargetRA
        Else
            If gHemisphere = 0 Then
                tPier = 1
            Else
                tPier = 0
            End If
            tRA = Range24(gTargetRA - 12)
       End If
    Else
        If gCWUP Then
            If gHemisphere = 0 Then
                tPier = 1
            Else
                tPier = 0
            End If
            tRA = Range24(gTargetRA - 12)
        Else
            If gHemisphere = 0 Then
                tPier = 0
            Else
                tPier = 1
            End If
            tRA = gTargetRA
        End If
    End If

    'Compute for Target RA/DEC Encoder
    targetRAEncoder = Get_RAEncoderfromRA(tRA, 0, gLongitude, gRAEncoder_Zero_pos, gTot_RA, gHemisphere)
    targetDECEncoder = Get_DECEncoderfromDEC(gTargetDec, tPier, gDECEncoder_Zero_pos, gTot_DEC, gHemisphere)
         
    If gCWUP Then
        HC.Add_Message "Goto: CW-UP slew requested"
        ' if RA limits are active
        If HC.ChkEnableLimits.Value = 1 And gRA_Limit_East <> 0 And gRA_Limit_West <> 0 Then
            ' check that the target position is within limits
            If gHemisphere = 0 Then
                If targetRAEncoder < gRA_Limit_East Or targetRAEncoder > gRA_Limit_West Then
                    ' target position is outside limits
                    gCWUP = False
                End If
            Else
                If targetRAEncoder > gRA_Limit_East Or targetRAEncoder < gRA_Limit_West Then
                    ' target position is outside limits
                    gCWUP = False
                End If
            End If
            
            ' if target position is outside limits
            If gCWUP = False Then
                HC.Add_Message "Goto: RA Limits prevent CW-UP slew"
                'then abandon Counter Weights up Slew and recalculate for a standard slew.
                If tha < 0 Then
                    If gHemisphere = 0 Then
                        tPier = 1
                    Else
                        tPier = 0
                    End If
                    tRA = Range24(gTargetRA - 12)
                Else
                    If gHemisphere = 0 Then
                        tPier = 0
                    Else
                        tPier = 1
                    End If
                    tRA = gTargetRA
                End If
                targetRAEncoder = Get_RAEncoderfromRA(tRA, 0, gLongitude, gRAEncoder_Zero_pos, gTot_RA, gHemisphere)
                targetDECEncoder = Get_DECEncoderfromDEC(gTargetDec, tPier, gDECEncoder_Zero_pos, gTot_DEC, gHemisphere)
            End If
        End If
    End If
             
    If gThreeStarEnable = False Then
       gSelectStar = 0
       currentRAEncoder = Delta_RA_Map(currentRAEncoder)
       currentDECEncoder = Delta_DEC_Map(currentDECEncoder)
    Else
       ' Transform target using model
       Select Case gAlignmentMode
           Case 2
             ' n-star+nearest
             tmpcoord = DeltaSyncReverse_Matrix_Map(targetRAEncoder - gRASync01, targetDECEncoder - gDECSync01)
           Case 1
               ' n-star
               tmpcoord = Delta_Matrix_Map(targetRAEncoder - gRASync01, targetDECEncoder - gDECSync01)
           Case Else
               ' nearest
               tmpcoord = Delta_Matrix_Map(targetRAEncoder - gRASync01, targetDECEncoder - gDECSync01)
               
               If tmpcoord.F = 0 Then
                   tmpcoord = DeltaSyncReverse_Matrix_Map(targetRAEncoder - gRASync01, targetDECEncoder - gDECSync01)
               End If
       End Select
       targetRAEncoder = tmpcoord.X
       targetDECEncoder = tmpcoord.Y
    End If
         
    'Execute the actual slew
    gGotoParams.RA_targetencoder = targetRAEncoder
    gGotoParams.RA_currentencoder = currentRAEncoder
    gGotoParams.DEC_targetencoder = targetDECEncoder
    gGotoParams.DEC_currentencoder = currentDECEncoder
    HC.Add_Message "Goto: " & FmtSexa(gTargetRA, False) & " " & FmtSexa(gTargetDec, True)
'    HC.Add_Message "Goto: RaEnc=" & CStr(currentRAEncoder) & " Target=" & CStr(targetRAEncoder)
'    HC.Add_Message "Goto: DecEnc=" & CStr(currentDECEncoder) & " Target=" & CStr(targetDECEncoder)
    
endradecslew:


 End Sub
Public Sub StartSlew(ByVal targetRAEncoder As Double, ByVal targetDECEncoder As Double, ByVal currentRAEncoder As Double, ByVal currentDECEncoder As Double)
Dim DeltaRAStep As Long
Dim DeltaDECStep As Long
    
    On Error GoTo endradecslew
    
    ' calculate relative amount to move
    DeltaRAStep = Abs(targetRAEncoder - currentRAEncoder)
    DeltaDECStep = Abs(targetDECEncoder - currentDECEncoder)
               
    If DeltaRAStep <> 0 Then
        ' Compensate for the smallest discrepancy after the final slew
        If gTrackingStatus > 0 Then
           If targetRAEncoder > currentRAEncoder Then
               If gHemisphere = 0 Then
                   DeltaRAStep = DeltaRAStep + gRA_Compensate
               Else
                   DeltaRAStep = DeltaRAStep - gRA_Compensate
               End If
           Else
               If gHemisphere = 0 Then
                   DeltaRAStep = DeltaRAStep - gRA_Compensate
               Else
                   DeltaRAStep = DeltaRAStep + gRA_Compensate
               End If
           End If
           If DeltaRAStep < 0 Then DeltaRAStep = 0
        End If
        
        If targetRAEncoder > currentRAEncoder Then
            gGotoParams.RA_Direction = 0
            Select Case gGotoParams.rate
                Case 0
                    ' let mount decide on slew rate
                    gGotoParams.RA_SlewActive = 0
                    eqres = EQStartMoveMotor(0, 0, 0, DeltaRAStep, GetSlowdown(DeltaRAStep))
                Case Else
                    gGotoParams.RA_SlewActive = 1
                    eqres = EQ_Slew(0, 0, 0, CLng(gGotoParams.rate))
            End Select
        Else
            gGotoParams.RA_Direction = 1
            Select Case gGotoParams.rate
                Case 0
                    gGotoParams.RA_SlewActive = 0
                    eqres = EQStartMoveMotor(0, 0, 1, DeltaRAStep, GetSlowdown(DeltaRAStep))
                Case Else
                    gGotoParams.RA_SlewActive = 1
                    eqres = EQ_Slew(0, 0, 1, CLng(gGotoParams.rate))
            End Select
        End If
    End If
     
    If DeltaDECStep <> 0 Then
        If targetDECEncoder > currentDECEncoder Then
            gGotoParams.DEC_Direction = 0
            Select Case gGotoParams.rate
                Case 0
                    ' let mount decide on slew rate
                    gGotoParams.DEC_SlewActive = 0
                    eqres = EQStartMoveMotor(1, 0, 0, DeltaDECStep, GetSlowdown(DeltaDECStep))
                Case Else
                    gGotoParams.DEC_SlewActive = 1
                    eqres = EQ_Slew(1, 0, 0, CLng(gGotoParams.rate))
            End Select
        Else
            gGotoParams.DEC_Direction = 1
            Select Case gGotoParams.rate
                Case 0
                    ' let mount decide on slew rate
                    gGotoParams.DEC_SlewActive = 0
                    eqres = EQStartMoveMotor(1, 0, 1, DeltaDECStep, GetSlowdown(DeltaDECStep))
                Case Else
                    gGotoParams.DEC_SlewActive = 1
                    eqres = EQ_Slew(1, 0, 1, CLng(gGotoParams.rate))
            End Select
        End If
    End If

     ' Activate Asynchronous Slew Monitoring Routine
     gRAStatus = EQ_MOTORBUSY
     gDECStatus = EQ_MOTORBUSY
     gRAStatus_slew = False
     
endradecslew:
     gSlewStatus = True

End Sub

' called from the encoder timer to supervise active gotos
Public Sub ManageGoto()
Dim tRA As Double
Dim tha As Double
Dim ra_diff As Double
Dim dec_diff As Double
    
    ''''''''''''''''''''''''''''''''''''''''''''''
    ' Fixed rate slew
    ''''''''''''''''''''''''''''''''''''''''''''''
    If gGotoParams.RA_SlewActive = 1 Or gGotoParams.DEC_SlewActive = 1 Then
        ' Handle as fixed rate slew
        If gGotoParams.RA_SlewActive Then
            If gGotoParams.RA_Direction = 0 Then
                If gRA_Encoder >= gGotoParams.RA_targetencoder Then
                    eqres = EQ_MotorStop(0)
                    Do
                        eqres = EQ_GetMotorStatus(0)
                        If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo MG1
                    Loop While (eqres And EQ_MOTORBUSY) <> 0
MG1:
                    gGotoParams.RA_SlewActive = 0
                    eqres = EQ_StartRATrack(0, gHemisphere, gHemisphere)
                End If
            Else
                If gRA_Encoder <= gGotoParams.RA_targetencoder Then
                    eqres = EQ_MotorStop(0)
                    Do
                        eqres = EQ_GetMotorStatus(0)
                        If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo MG2
                    Loop While (eqres And EQ_MOTORBUSY) <> 0
MG2:
                    gGotoParams.RA_SlewActive = 0
                    eqres = EQ_StartRATrack(0, gHemisphere, gHemisphere)
                End If
            End If
        End If
    
        If gGotoParams.DEC_SlewActive Then
            If gGotoParams.DEC_Direction = 0 Then
                If gDec_Encoder >= gGotoParams.DEC_targetencoder Then
                    eqres = EQ_MotorStop(1)
                    Do
                        eqres = EQ_GetMotorStatus(1)
                        If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo MG3
                    Loop While (eqres And EQ_MOTORBUSY) <> 0
MG3:
                    gGotoParams.DEC_SlewActive = 0
                End If
            Else
                If gDec_Encoder <= gGotoParams.DEC_targetencoder Then
                    eqres = EQ_MotorStop(1)
                    Do
                        eqres = EQ_GetMotorStatus(1)
                        If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo MG4
                    Loop While (eqres And EQ_MOTORBUSY) <> 0
MG4:
                    gGotoParams.DEC_SlewActive = 0
                End If
            End If
        End If
    
   
        If gGotoParams.RA_SlewActive = 0 And gGotoParams.DEC_SlewActive = 0 Then
            
            Select Case gGotoParams.SuperSafeMode
                Case 0
                    ' rough fixed rate slew complete
                    Call CalcEncoderTargets
                    ra_diff = Abs(gGotoParams.RA_targetencoder - gRA_Encoder)
                    dec_diff = Abs(gGotoParams.DEC_targetencoder - gDec_Encoder)
                    HC.Add_Message "Goto: FRSlew complete ra_diff=" & CStr(ra_diff) & " dec_diff=" & CStr(dec_diff)
                    If (ra_diff < gTot_RA / 360) And (dec_diff < gTot_DEC / 540) Then
                        ' initiate a standard itterative goto if within a 3/4 of a degree.
                        gGotoParams.rate = 0
                        Call StartSlew(gGotoParams.RA_targetencoder, gGotoParams.DEC_targetencoder, gGotoParams.RA_currentencoder, gGotoParams.DEC_currentencoder)
                    Else
                        ' Do another rough slew.
                        HC.Add_Message "Goto: FRSlew"
                        gFRSlewCount = gFRSlewCount + 1
                        If gFRSlewCount >= 5 Then
                            'if we can't get close after 5 attempts then abandon the FR slew
                            'and use the full speed iterative slew
                            gFRSlewCount = 0
                            gGotoParams.rate = 0
                        End If
                        Call StartSlew(gGotoParams.RA_targetencoder, gGotoParams.DEC_targetencoder, gGotoParams.RA_currentencoder, gGotoParams.DEC_currentencoder)
                    End If
        
                Case 1
                    ' move to RA target
                    Call CalcEncoderTargets
                    gGotoParams.SuperSafeMode = 0
                    Call StartSlew(gGotoParams.RA_targetencoder, gGotoParams.DEC_targetencoder, gGotoParams.RA_currentencoder, gGotoParams.DEC_currentencoder)
               
'                Case 2
'                    ' we're at a limit about to go to target
'                    Call CalcEncoderTargets
'                    gGotoParams.SuperSafeMode = 0
'                    Call StartSlew(gGotoParams.RA_targetencoder, gGotoParams.DEC_targetencoder, gGotoParams.RA_currentencoder, gGotoParams.DEC_currentencoder)
                    
                Case 3
                    ' were at a limit position
                    If gGotoParams.RA_targetencoder > RAEncoder_Home_pos Then
                        ' dual axis slew to limit position nearest to target
                        gGotoParams.SuperSafeMode = 1
                        If RALimitsActive() = False Then
                            Call StartSlew(gRAMeridianWest, gGotoParams.DEC_targetencoder, gEmulRA, gEmulDEC)
                        Else
                            Call StartSlew(gRA_Limit_West, gGotoParams.DEC_targetencoder, gEmulRA, gEmulDEC)
                        End If
                    Else
                        ' dual axis slew to limit position nearest to target
                        gGotoParams.SuperSafeMode = 1
                        If RALimitsActive() = False Then
                            Call StartSlew(gRAMeridianEast, gGotoParams.DEC_targetencoder, gEmulRA, gEmulDEC)
                        Else
                            Call StartSlew(gRA_Limit_East, gGotoParams.DEC_targetencoder, gEmulRA, gEmulDEC)
                        End If
                    End If
                    
            End Select
        End If
        Exit Sub
    
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''
    ' Iterative slew - variable rate
    ''''''''''''''''''''''''''''''''''''''''''''''
    If (gRAStatus And EQ_MOTORBUSY) = 0 Then
        'At This point RA motor has completed the slew
        gRAStatus_slew = True
        If (gDECStatus And EQ_MOTORBUSY) <> 0 Then
            ' The DEC motor is still moving so start sidereal tracking to hold position in RA
            eqres = EQ_StartRATrack(0, gHemisphere, gHemisphere)
        End If
    End If

    If (gDECStatus And EQ_MOTORBUSY) = 0 And gRAStatus_slew Then
        'DEC and RA motors have finished slewing at this point
        'We need to check if a new slew is needed to reduce the any difference
        'Caused by the Movement of the earth during the slew process

        Select Case gGotoParams.SuperSafeMode
            Case 0
                ' decrement the slew retry count
                gSlewCount = gSlewCount - 1

                ' calculate the difference (arcsec)  between target and current coords
                ra_diff = 3600 * Abs(gRA - gTargetRA)
                dec_diff = 3600 * Abs(gDec - gTargetDec)
                
                If (gSlewCount > 0) And (gTrackingStatus > 0) Then  ' Retry only if tracking is enabled
                    ' aim to get within the goto resolution (default = 10 steps)
                    If gGotoResolution > 0 And ra_diff <= gRAGotoRes And dec_diff <= gDECGotoRes Then
                        GoTo slewcomplete
                    Else
                       'Re Execute a new RA-Only slew here
                        Call CalcEncoderTargets
                        gGotoParams.rate = 0
                        Call StartSlew(gGotoParams.RA_targetencoder, gGotoParams.DEC_targetencoder, gGotoParams.RA_currentencoder, gGotoParams.DEC_currentencoder)
                    End If
                Else
                    GoTo slewcomplete
                End If
            
            Case 1
                ' move to target
                gGotoParams.SuperSafeMode = 0
                Call CalcEncoderTargets
                gGotoParams.rate = 0
                'kick of an iterative slew to get us accurately to target RA
                Call StartSlew(gGotoParams.RA_targetencoder, gGotoParams.DEC_targetencoder, gGotoParams.RA_currentencoder, gGotoParams.DEC_currentencoder)
           
'            Case 2
'                ' At a limit about to slew to target
'                Call CalcEncoderTargets
'                gGotoParams.SuperSafeMode = 0
 '               Call StartSlew(gGotoParams.RA_targetencoder, gGotoParams.DEC_targetencoder, gEmulRA, gEmulDEC)
        
            Case 3
                ' we are at a limit position
                If gGotoParams.RA_targetencoder > RAEncoder_Home_pos Then
                    ' dual axis slew to limit position nearest to target
                    gGotoParams.SuperSafeMode = 1
                    If RALimitsActive() = False Then
                        Call StartSlew(gRAMeridianWest, gGotoParams.DEC_targetencoder, gEmulRA, gEmulDEC)
                    Else
                        Call StartSlew(gRA_Limit_West, gGotoParams.DEC_targetencoder, gEmulRA, gEmulDEC)
                    End If
                Else
                    ' dual axis slew to limit position nearest to target
                    gGotoParams.SuperSafeMode = 1
                    If RALimitsActive() = False Then
                        Call StartSlew(gRAMeridianEast, gGotoParams.DEC_targetencoder, gEmulRA, gEmulDEC)
                    Else
                        Call StartSlew(gRA_Limit_West, gGotoParams.DEC_targetencoder, gEmulRA, gEmulDEC)
                    End If
                End If
        
        
        End Select
    End If
    Exit Sub

slewcomplete:
    gSlewStatus = False
    gRAStatus_slew = False
    gSupressHorizonLimits = False

    ' slew may have terminated early if parked
    If gEQparkstatus <> 1 Then
        ' we've reached the desired target coords - resume tracking.
        Select Case gTrackingStatus
            Case 0, 1
                EQStartSidereal
            Case 2, 3, 4
                RestartTracking
        End Select
        
        HC.Add_Message (oLangDll.GetLangString(5018) & " " & FmtSexa(gRA, False) & " " & FmtSexa(gDec, True))
        HC.Add_Message ("Goto: SlewItereations=" & CStr(gMaxSlewCount - gSlewCount))
        HC.Add_Message ("Goto: " & "RaDiff=" & Format$(str(ra_diff), "000.00") & " DecDiff=" & Format$(str(dec_diff), "000.00"))
        
        ' goto complete
        Call EQ_Beep(6)
    End If
    
    If gDisbleFlipGotoReset = 0 Then
        HC.ChkForceFlip.Value = 0
    End If
 
End Sub

Public Sub writeGotoRate()
    HC.oPersist.WriteIniValue "GOTO_RATE", CStr(gGotoRate)
End Sub
Public Sub readGotoRate()
Dim tmptxt As String

    On Error Resume Next
    tmptxt = HC.oPersist.ReadIniValue("GOTO_RATE")
    If tmptxt <> "" Then
        gGotoRate = val(tmptxt)
    Else
        gGotoRate = 0
        Call writeCustomRa
    End If
'    gGotoRate = 0
    If gGotoRate = 0 Then
        HC.HScrollSlewLimit.Value = HC.HScrollSlewLimit.min
    Else
        HC.HScrollSlewLimit.Value = gGotoRate
    End If
    gParkParams.rate = gGotoRate

End Sub

Public Sub readFlipGoto()
Dim tmptxt As String

    On Error Resume Next
    tmptxt = HC.oPersist.ReadIniValue("DISABLE_FLIPGOTO_RESET")
    If tmptxt <> "" Then
        gDisbleFlipGotoReset = val(tmptxt)
    Else
        HC.oPersist.WriteIniValue "DISABLE_FLIPGOTO_RESET", "0"
        gDisbleFlipGotoReset = 0
    End If

End Sub
