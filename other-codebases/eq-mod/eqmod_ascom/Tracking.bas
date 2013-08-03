Attribute VB_Name = "Tracking"
Option Explicit

Public gCustomTrackingOffsetRA As Integer
Public gCustomTrackingOffsetDEC As Integer
Public gTrackFactorRA As Double
Public gTrackFactorDEC As Double
Public g_RAAxisRates As Rates           ' rates available for MoveAxis
Public g_DECAxisRates As Rates          ' rates available for MoveAxis
Public g_TrackingRates As TrackingRates ' Collection of supported drive rates
Public gCustomTrackFile As String
Public gCustomTrackName As String

Type TrackRecord_def
    time_mjd As Double
    DeltaRa As Double
    DeltaDec As Double
    RaRate As Double
    DecRate As Double
    DecDir  As Integer
    RAJ2000 As Double
    DECJ2000 As Double
    RaRateRaw As Double
    DECRateRaw As Double
    UseRate As Boolean
End Type

' main control structure for custom tracking
Type TrackCtrl_def
    FileFormat As Integer
    Precess As Boolean
    Waypoint As Boolean
    AdjustRA As Double
    AdjustDEC As Double
    TrackIdx As Integer
    TrackingChangesEnabled As Boolean
    TrackSchedule() As TrackRecord_def
End Type

Type RaDecCoords
    RA As Double
    DEC As Double
End Type

Dim TrackCtrl As TrackCtrl_def


' Start RA motor based on an input rate of arcsec per Second

Public Sub StartRA_by_Rate(ByVal RA_RATE As Double)

Dim i As Double
Dim j As Double
Dim k As Double
Dim m As Double

    k = 0
    m = 1
    i = Abs(RA_RATE)
    
    If gMount_Ver > &H301 Then
        If i > 1000 Then
            k = 1
            m = EQGP(0, 10003)
        End If
    Else
        If i > 3000 Then
            k = 1
            m = EQGP(0, 10003)
        End If
    End If

    HC.Add_Message (oLangDll.GetLangString(117) & " " & str(m) & " , " & str(RA_RATE) & " arcsec/sec")

    eqres = EQ_MotorStop(0)          ' Stop RA Motor
    If eqres <> EQ_OK Then
        GoTo RARateEndhome1
    End If
    
    Do
       eqres = EQ_GetMotorStatus(0)
       If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then
            GoTo RARateEndhome1
       End If
    Loop While (eqres And EQ_MOTORBUSY) <> 0
    

    If RA_RATE = 0 Then
        gSlewStatus = False
        gRAStatus_slew = False
        eqres = EQ_MotorStop(0)
        gRAMoveAxis_Rate = 0
        Exit Sub
    End If

    i = RA_RATE
    j = Abs(i)              'Get the absolute value for parameter passing
    
    If gMount_Ver = &H301 Then
      If (j > 1350) And (j <= 3000) Then
        If j < 2175 Then
            j = 1350
        Else
            j = 3001
            k = 1
            m = EQGP(0, 10003)
        End If
      End If
    End If
    
    gRAMoveAxis_Rate = k    'Save Speed Settings
    
    HC.Add_FileMessage ("StartRARate=" & FormatNumber(RA_RATE, 5))
'    j = Int((m * 9325.46154 / j) + 0.5) + 30000 'Compute for the rate
    j = Int((m * gTrackFactorRA / j) + 0.5) + 30000 'Compute for the rate

    If i >= 0 Then
        eqres = EQ_SetCustomTrackRate(0, 1, j, k, gHemisphere, 0)
    Else
        eqres = EQ_SetCustomTrackRate(0, 1, j, k, gHemisphere, 1)
    End If

RARateEndhome1:

End Sub

' Change RA motor rate based on an input rate of arcsec per Second

Public Sub ChangeRA_by_Rate(ByVal rate As Double)

Dim j As Double
Dim k As Double
Dim m As Double
Dim dir As Long
Dim init As Long

    If rate >= 0 Then
        dir = 0
    Else
        dir = 1
    End If
    
    If rate = 0 Then
        ' rate = 0 so stop motors
        gSlewStatus = False
        eqres = EQ_MotorStop(0)
        gRAStatus_slew = False
        gRAMoveAxis_Rate = 0
        Exit Sub
    End If
    
    k = 0   ' Assume low speed
    m = 1   ' Speed multiplier = 1
    
    init = 0
    j = Abs(rate)
    
    If gMount_Ver > &H301 Then
       ' if above high speed theshold
        If j > 1000 Then
            k = 1               ' HIGH SPEED
            m = EQGP(0, 10003)  ' GET HIGH SPEED MULTIPLIER
        End If
    Else
        ' who knows what Mon is up to here - a special for his mount perhaps?
        If gMount_Ver = &H301 Then
            If (j > 1350) And (j <= 3000) Then
                If j < 2175 Then
                    j = 1350
                Else
                    j = 3001
                    k = 1
                    m = EQGP(0, 10003)
                End If
            End If
        End If
        ' if above high speed theshold
         If j > 3000 Then
             k = 1               ' HIGH SPEED
             m = EQGP(0, 10003)  ' GET HIGH SPEED MULTIPLIER
         End If
    End If
    
    HC.Add_FileMessage ("ChangeRARate=" & FormatNumber(rate, 5))

    ' if there's a switch between high/low speed or if operating at high speed
    ' we ned to do additional initialisation
    If k <> 0 Or k <> gRAMoveAxis_Rate Then init = 1
    
    If init = 1 Then
        ' Stop Motor
        HC.Add_FileMessage ("Direction or High/Low speed change")
        eqres = EQ_MotorStop(0)
        If eqres <> EQ_OK Then GoTo RARateEndhome2
        
        ' wait for motor to stop
        Do
          eqres = EQ_GetMotorStatus(0)
          If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then
               GoTo RARateEndhome2
          End If
        Loop While (eqres And EQ_MOTORBUSY) <> 0
        'force initialisation
    End If
    
    gRAMoveAxis_Rate = k
    
     'Compute for the rate
'    j = Int((m * 9325.46154 / j) + 0.5) + 30000
    j = Int((m * gTrackFactorRA / j) + 0.5) + 30000
    
    eqres = EQ_SetCustomTrackRate(0, init, j, k, gHemisphere, dir)
    HC.Add_FileMessage ("EQ_SetCustomTrackRate=0," & CStr(init) & "," & CStr(j) & "," & CStr(k) & "," & CStr(gHemisphere) & "," & CStr(dir))
    HC.Add_Message (oLangDll.GetLangString(117) & "=" & str(rate) & " arcsec/sec" & "," & CStr(eqres))
    
RARateEndhome2:
        
End Sub


' Start DEC motor based on an input rate of arcsec per Second

Public Sub StartDEC_by_Rate(ByVal DEC_RATE As Double)

Dim i As Double
Dim j As Double
Dim k As Double
Dim m As Double

    k = 0
    m = 1
    i = Abs(DEC_RATE)
    
    If gMount_Ver > &H301 Then
        If i > 1000 Then
            k = 1
            m = EQGP(1, 10003)
        End If
    Else
        If i > 3000 Then
            k = 1
            m = EQGP(1, 10003)
        End If
    End If
    
    
    HC.Add_Message (oLangDll.GetLangString(118) & " " & str(m) & " , " & str(DEC_RATE) & " arcsec/sec")
    
    eqres = EQ_MotorStop(1)          ' Stop RA Motor
    If eqres <> EQ_OK Then
        GoTo DECRateEndhome1
    End If
    
    Do
       eqres = EQ_GetMotorStatus(1)
       If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then
            GoTo DECRateEndhome1
       End If
    Loop While (eqres And EQ_MOTORBUSY) <> 0

    If DEC_RATE = 0 Then
        gSlewStatus = False
        gRAStatus_slew = False
        eqres = EQ_MotorStop(1)
        gDECMoveAxis_Rate = 0
        Exit Sub
    End If
    
    i = DEC_RATE
    j = Abs(i)              'Get the absolute value for parameter passing
    
  
    If gMount_Ver = &H301 Then
      If (j > 1350) And (j <= 3000) Then
        If j < 2175 Then
            j = 1350
        Else
            j = 3001
            k = 1
            m = EQGP(1, 10003)
        End If
      End If
    End If
    
    
    gDECMoveAxis_Rate = k    'Save Speed Settings
    
    HC.Add_FileMessage ("StartDecRate=" & FormatNumber(DEC_RATE, 5))
'    j = Int((m * 9325.46154 / j) + 0.5) + 30000 'Compute for the rate
    j = Int((m * gTrackFactorDEC / j) + 0.5) + 30000 'Compute for the rate

    If i >= 0 Then
        eqres = EQ_SetCustomTrackRate(1, 1, j, k, gHemisphere, 0)
    Else
        eqres = EQ_SetCustomTrackRate(1, 1, j, k, gHemisphere, 1)
    End If
    
DECRateEndhome1:

End Sub



' Change DEC motor rate based on an input rate of arcsec per Second

Public Sub ChangeDEC_by_Rate(ByVal rate As Double)

Dim j As Double
Dim k As Double
Dim m As Double
Dim dir As Long
Dim init As Long

    If rate >= 0 Then
        dir = 0
    Else
        dir = 1
    End If
    
    If rate = 0 Then
        ' rate = 0 so stop motors
        gSlewStatus = False
        eqres = EQ_MotorStop(1)
'        gRAStatus_slew = False
        gDECMoveAxis_Rate = 0
        Exit Sub
    End If
    
    k = 0   ' Assume low speed
    m = 1   ' Speed multiplier = 1
    init = 0
    j = Abs(rate)
    
    If gMount_Ver > &H301 Then
       ' if above high speed theshold
        If j > 1000 Then
            k = 1               ' HIGH SPEED
            m = EQGP(1, 10003)  ' GET HIGH SPEED MULTIPLIER
        End If
    Else
        ' who knows what Mon is up to here - a special for his mount perhaps?
        If gMount_Ver = &H301 Then
            If (j > 1350) And (j <= 3000) Then
                If j < 2175 Then
                    j = 1350
                Else
                    j = 3001
                    k = 1
                    m = EQGP(1, 10003)
                End If
            End If
        End If
        ' if above high speed theshold
         If j > 3000 Then
             k = 1               ' HIGH SPEED
             m = EQGP(1, 10003)  ' GET HIGH SPEED MULTIPLIER
         End If
    End If
    
    HC.Add_FileMessage ("ChangeDECRate=" & FormatNumber(rate, 5))

    ' if there's a switch between high/low speed or if operating at high speed
    ' we need to do additional initialisation
    If k <> 0 Or k <> gDECMoveAxis_Rate Then init = 1
    
    If init = 1 Then
        ' Stop Motor
        HC.Add_FileMessage ("Direction or High/Low speed change")
        eqres = EQ_MotorStop(1)
        If eqres <> EQ_OK Then GoTo DECRateEndhome2
        
        ' wait for motor to stop
        Do
          eqres = EQ_GetMotorStatus(1)
          If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then
               GoTo DECRateEndhome2
          End If
        Loop While (eqres And EQ_MOTORBUSY) <> 0
        'force initialisation
    End If
    
    
    gDECMoveAxis_Rate = k
    
     'Compute for the rate
    j = Int((m * gTrackFactorDEC / j) + 0.5) + 30000
'    j = Int((m * 9325.46154 / j) + 0.5) + 30000
    
    eqres = EQ_SetCustomTrackRate(1, init, j, k, gHemisphere, dir)
    HC.Add_FileMessage ("EQ_SetCustomTrackRate=1," & CStr(init) & "," & CStr(j) & "," & CStr(k) & "," & CStr(gHemisphere) & "," & CStr(dir))
    HC.Add_Message (oLangDll.GetLangString(118) & "=" & str(rate) & " arcsec/sec" & "," & CStr(eqres))
    
DECRateEndhome2:

End Sub


Public Sub EQMoveAxis(axis As Double, rate As Double)

Dim j As Double
Dim current_rate As Double

    If rate <> 0 Then
        HC.TrackingFrame.Caption = oLangDll.GetLangString(121) & " " & oLangDll.GetLangString(189)
    End If

    j = rate * 3600 ' Convert to Arcseconds

    If axis = 0 Then
    
        If rate = 0 And (gDeclinationRate = 0) Then
            HC.TrackingFrame.Caption = oLangDll.GetLangString(121) & " " & oLangDll.GetLangString(178)
        End If
    
        
        If gHemisphere = 1 Then
            j = -1 * j
            current_rate = gRightAscensionRate * -1
        Else
            current_rate = gRightAscensionRate
        End If
        
        ' check for change of direction
        If (current_rate * j) <= 0 Then
            Call StartRA_by_Rate(j)
        Else
            Call ChangeRA_by_Rate(j)
        End If
  
        gRightAscensionRate = j
        If rate <> 0 Then gTrackingStatus = 4
    
    End If
    
    If axis = 1 Then
    
        If rate = 0 And (gRightAscensionRate = 0) Then
            HC.TrackingFrame.Caption = oLangDll.GetLangString(121) & " " & oLangDll.GetLangString(178)
        End If
        
' Mon seems to have included the code below for the Move South/Move North requirements of satelite tracking
' However ASCOM requires that a positive rate always moces the axis clockwise so this code is well iffy!
'        j = j * -1
'
'        If gHemisphere = 0 Then
'            If (gDec_DegNoAdjust > 90) And (gDec_DegNoAdjust <= 270) Then j = j * -1
'        Else
'            If (gDec_DegNoAdjust <= 90) Or (gDec_DegNoAdjust > 270) Then j = j * -1
'        End If
    
        ' check for change of direction
        If (gDeclinationRate * j) <= 0 Then
            Call StartDEC_by_Rate(j)
        Else
            Call ChangeDEC_by_Rate(j)
        End If
    
        gDeclinationRate = j
        If rate <> 0 Then gTrackingStatus = 4
    
    End If
    
End Sub

Public Sub CustomMoveAxis(axis As Double, rate As Double, init As Boolean, RateName As String)

Dim j As Double

    If rate <> 0 Then
        HC.TrackingFrame.Caption = oLangDll.GetLangString(121) & " " & RateName
    End If

    j = rate

    If axis = 0 Then
        If rate = 0 And (gDeclinationRate = 0) Then
            HC.TrackingFrame.Caption = oLangDll.GetLangString(121) & " " & oLangDll.GetLangString(178)
        End If
        If init = True Then
            Call StartRA_by_Rate(j)
        Else
            If j <> gRightAscensionRate Then
                Call ChangeRA_by_Rate(j)
            End If
        End If
        gRightAscensionRate = j
        gTrackingStatus = 4
    End If
    
    If axis = 1 Then
        If rate = 0 And (gRightAscensionRate = 0) Then
            HC.TrackingFrame.Caption = oLangDll.GetLangString(121) & " " & oLangDll.GetLangString(178)
        End If
        If init = True Then
            Call StartDEC_by_Rate(j)
        Else
            If j <> gDeclinationRate Then
                Call ChangeDEC_by_Rate(j)
            End If
        End If
        gDeclinationRate = j
        gTrackingStatus = 4
    End If
    
End Sub


Public Sub Start_CustomTracking2()
    If gEQparkstatus <> 0 Then
        HC.Add_Message (oLangDll.GetLangString(5013))
        Exit Sub
    End If
    
    gRA_LastRate = 0
    If gPEC_Enabled Then
       PEC_StopTracking
    End If
    EQ_Beep (13)
    Call Start_CustomTracking

End Sub


Public Sub Start_CustomTracking()
    Dim i As Double
    Dim j As Double
    
    On Error GoTo handlerr

    If gCustomTrackFile = "" Then
    
        TrackCtrl.TrackingChangesEnabled = False

        
        i = CDbl(HC.raCustom)
        j = CDbl(HC.decCustom)
        If gHemisphere = 1 Then
            i = -1 * i
        End If
    
        If (Abs(i) > 12000) Or (Abs(j) > 12000) Then
            GoTo handlerr
        End If
    
        HC.Add_Message (oLangDll.GetLangString(5040) & Format$(str(i), "000.00") & " DEC:" & Format$(str(j), "000.00") & " arcsec/sec")

        Call CustomMoveAxis(0, i, True, oLangDll.GetLangString(189))
        Call CustomMoveAxis(1, j, True, oLangDll.GetLangString(189))
    Else
        ' custom track file is assigned
        TrackCtrl.TrackIdx = GetTrackFileIdx(1, True)
        If TrackCtrl.TrackIdx <> -1 Then
            If TrackCtrl.Waypoint Then
                Call GetTrackTarget(i, j)
                TrackCtrl.AdjustRA = gRA - i
                TrackCtrl.AdjustDEC = gDec - j
            Else
                TrackCtrl.AdjustRA = 0
                TrackCtrl.AdjustDEC = 0
            End If
            i = TrackCtrl.TrackSchedule(TrackCtrl.TrackIdx).RaRate
            j = TrackCtrl.TrackSchedule(TrackCtrl.TrackIdx).DecRate
            HC.decCustom.Text = FormatNumber(j, 5)
            If gHemisphere = 1 Then
                HC.raCustom.Text = FormatNumber(-1 * i, 5)
            Else
                HC.raCustom.Text = FormatNumber(i, 5)
            End If
            Call CustomMoveAxis(0, i, True, gCustomTrackName)
            Call CustomMoveAxis(1, j, True, gCustomTrackName)
        Else
        End If
        TrackCtrl.TrackingChangesEnabled = True
        HC.CustomTrackTimer.Enabled = True
    End If
    Exit Sub
        
handlerr:
'   HC.Add_Message (oLangDll.GetLangString(5039))
    Call emergency_stop
        
End Sub

Public Sub Restore_CustomTracking()
Dim rate As Double
Dim RA As Double
Dim DEC As Double
  
    If gTrackingStatus = 4 Then
        If gCustomTrackFile = "" Then
            TrackCtrl.TrackingChangesEnabled = False
            Call CustomMoveAxis(0, gRightAscensionRate, True, oLangDll.GetLangString(189))
            Call CustomMoveAxis(1, gDeclinationRate, True, oLangDll.GetLangString(189))
        Else
            TrackCtrl.TrackIdx = GetTrackFileIdx(1, False)
            If TrackCtrl.TrackIdx <> -1 Then
            
                If TrackCtrl.Waypoint Then
                    Call GetTrackTarget(RA, DEC)
                    TrackCtrl.AdjustRA = gRA - RA
                    TrackCtrl.AdjustDEC = gDec - DEC
                Else
                    TrackCtrl.AdjustRA = 0
                    TrackCtrl.AdjustDEC = 0
                End If
                
                rate = TrackCtrl.TrackSchedule(TrackCtrl.TrackIdx).RaRate
                If gHemisphere = 1 Then
                    HC.raCustom.Text = FormatNumber(-1 * rate, 5)
                Else
                    HC.raCustom.Text = FormatNumber(rate, 5)
                End If
                Call CustomMoveAxis(0, rate, True, gCustomTrackName)
                
                rate = TrackCtrl.TrackSchedule(TrackCtrl.TrackIdx).DecRate
                HC.decCustom.Text = FormatNumber(rate, 5)
                Call CustomMoveAxis(1, rate, True, gCustomTrackName)
            Else
                Call CustomMoveAxis(0, gRightAscensionRate, True, oLangDll.GetLangString(189))
                Call CustomMoveAxis(1, gDeclinationRate, True, oLangDll.GetLangString(189))
                HC.raCustom.Text = FormatNumber(gRightAscensionRate, 5)
                HC.decCustom.Text = FormatNumber(gDeclinationRate, 5)
            End If
            TrackCtrl.TrackingChangesEnabled = True
        End If
    End If
        
End Sub

Public Sub EQStartSidereal2()
    If gEQparkstatus <> 0 Then
        ' no tracking if parked!
        HC.Add_Message (oLangDll.GetLangString(5013))
    Else
        Call EQStartSidereal
        EQ_Beep (10)
    End If
End Sub



Public Sub EQStartSidereal()
    gRA_LastRate = 0
                
    If gEQparkstatus <> 0 Then
        ' no tracking if parked!
        HC.Add_Message (oLangDll.GetLangString(5013))
    Else
        ' Stop DEC motor
        eqres = EQ_MotorStop(1)
        gDeclinationRate = 0
            
        ' start RA motor at sidereal
        eqres = EQ_StartRATrack(0, gHemisphere, gHemisphere)
        gRAMoveAxis_Rate = 0
        gTrackingStatus = 1
        gRightAscensionRate = SID_RATE
        
        If HC.CheckPEC.Value = 1 Then
            ' track using PEC
            PEC_StartTracking
        Else
            ' Set Caption
            HC.TrackingFrame.Caption = oLangDll.GetLangString(121) & " " & oLangDll.GetLangString(122)
            HC.Add_Message (oLangDll.GetLangString(5014))
        End If
    End If

End Sub

Public Sub StopTrackingUpdates()
    Select Case gTrackingStatus
        Case 1
            Call PEC_StopTracking
        Case 2
        Case 3
        Case 4
            TrackCtrl.TrackingChangesEnabled = False
        Case Else
    End Select
End Sub


Public Sub RestartTracking()
    
    gRAMoveAxis_Rate = 0
    
    Select Case gTrackingStatus
        Case 1
           EQStartSidereal
        Case 2
           Start_Lunar
        Case 3
           Start_Solar
        Case 4
           Call Restore_CustomTracking
        Case Else
            ' not tracking
            eqres = EQ_MotorStop(0)
            eqres = EQ_MotorStop(1)
    End Select

End Sub

Public Sub writeCustomRa()
    HC.oPersist.WriteIniValue "CUSTOM_RA", HC.raCustom.Text
    HC.oPersist.WriteIniValue "CUSTOM_DEC", HC.decCustom.Text
    HC.oPersist.WriteIniValue "CUSTOM_TRACKFILE", HC.LabelTrackFile.ToolTipText
End Sub

Public Sub readCustomRa()
Dim tmptxt As String

    tmptxt = HC.oPersist.ReadIniValue("CUSTOM_RA")
    If tmptxt <> "" Then
        HC.raCustom.Text = tmptxt
    Else
        HC.raCustom.Text = CStr(15.041067)
        Call writeCustomRa
    End If
    
    tmptxt = HC.oPersist.ReadIniValue("CUSTOM_DEC")
    If tmptxt <> "" Then
        HC.decCustom.Text = tmptxt
    Else
        HC.decCustom.Text = "0"
        Call writeCustomRa
    End If
    
    ' reload custom track file
    tmptxt = HC.oPersist.ReadIniValue("CUSTOM_TRACKFILE")
    If tmptxt <> "" Then
        If Track_LoadFile(tmptxt) = True Then
            gCustomTrackFile = tmptxt
            HC.LabelTrackFile.Caption = StripPath(gCustomTrackFile)
            HC.LabelTrackFile.ToolTipText = gCustomTrackFile
            HC.CmdTrack(4).ToolTipText = gCustomTrackName
            ' check data is current
            TrackCtrl.TrackIdx = GetTrackFileIdx(1, True)
        Else
            gCustomTrackFile = ""
            HC.LabelTrackFile.Caption = ""
            HC.LabelTrackFile.ToolTipText = ""
            HC.CmdTrack(4).ToolTipText = oLangDll.GetLangString(189)
            Call writeCustomRa
        End If
    Else
        gCustomTrackFile = ""
        HC.LabelTrackFile.Caption = ""
        HC.LabelTrackFile.ToolTipText = ""
    End If
    
    If gCustomTrackFile <> "" Then
        HC.CmdTrack(4).Picture = LoadResPicture(111, vbResBitmap)
    Else
        HC.CmdTrack(4).Picture = LoadResPicture(110, vbResBitmap)
    End If
    
    
End Sub
Public Sub readSiderealRate()
Dim tmptxt As String
    On Error GoTo readerr

    tmptxt = HC.oPersist.ReadIniValue("SIDEREAL_RATE")
    If tmptxt <> "" Then
        gSiderealRate = CDbl(tmptxt)
    Else
readerr:
        gSiderealRate = 15.041067
        Call writeSiderealRate
    End If

End Sub

Public Sub writeSiderealRate()
    HC.oPersist.WriteIniValue "SIDEREAL_RATE", CStr(gSiderealRate)
End Sub

Public Sub LoadTrackingRates()
    FileDlg.filter = "*.txt*"
    If gCustomTrackFile <> "" Then
        FileDlg.lastdir = GetPath(gCustomTrackFile)
        FileDlg.notfirst = True
    End If
    FileDlg.Show (1)
    If FileDlg.FileName <> "" Then
        If Track_LoadFile(FileDlg.FileName) = True Then
            gCustomTrackFile = FileDlg.FileName
            HC.LabelTrackFile.Caption = FileDlg.filename2
            HC.LabelTrackFile.ToolTipText = FileDlg.FileName
            If gCustomTrackName = "" Then
                gCustomTrackName = FileDlg.filename2
            End If
            HC.CmdTrack(4).ToolTipText = gCustomTrackName
            ' check data is current
            TrackCtrl.TrackIdx = GetTrackFileIdx(1, True)
        Else
            gCustomTrackFile = ""
            HC.LabelTrackFile.Caption = ""
            HC.LabelTrackFile.ToolTipText = ""
            HC.CmdTrack(4).ToolTipText = oLangDll.GetLangString(189)
        End If
    End If
            
    If gCustomTrackFile <> "" Then
        HC.CmdTrack(4).Picture = LoadResPicture(111, vbResBitmap)
    Else
        HC.CmdTrack(4).Picture = LoadResPicture(110, vbResBitmap)
        HC.CmdTrack(4).ToolTipText = oLangDll.GetLangString(189)
    End If
    Call writeCustomRa

End Sub

Public Function Track_LoadFile(FileName As String) As Boolean
Dim temp1 As String
Dim temp2() As String
Dim lineno As Integer
Dim idx As Integer
Dim NF1 As Integer
Dim NF2 As Integer
Dim month As Double
Dim year As Double
Dim day As Double
Dim hour As Double
Dim minute As Double
Dim second As Double
Dim RA As Double
Dim DEC As Double
Dim Lastra As Double
Dim Lastdec As Double
Dim mjd As Double
Dim Lastmjd As Double
Dim DecEncoder As Double
Dim LastDecEncoder As Double
Dim LastRaRate As Double
Dim LastDecRate As Double
Dim RaRate As Double
Dim DecRate As Double
Dim deltat As Double
Dim Format As Integer
Dim TrackNew As TrackRecord_def
Dim params() As String
Dim timestr As String
Dim typTime As SYSTEMTIME
Dim mjdNow As Double

    On Error GoTo ImportError
    
    Track_LoadFile = False
    
    If FileName = "" Then
        GoTo ImportError
    End If
    
    
    NF2 = FreeFile
    Close #NF2
    temp1 = HC.oPersist.GetIniPath & "\CustomRateDebug.txt"
    Open temp1 For Output As #NF2
    Print #NF2, "MJD RaDelta RaRate DecDelta DecRate DecDir"
    
    NF1 = FreeFile
    Close #NF1
    Open FileName For Input As #NF1
    
    lineno = 0
    idx = 0
    Lastra = 0
    Lastdec = 0
    Lastmjd = 0
    TrackCtrl.FileFormat = 0
    TrackCtrl.Precess = True
    TrackCtrl.Waypoint = False
    ReDim TrackCtrl.TrackSchedule(1)
    
    gCustomTrackName = StripPath(FileName)
    
    While Not EOF(NF1)
        Line Input #NF1, temp1
        If temp1 <> "" Then
            Select Case Left(temp1, 1)
            
                Case "#"
                    ' comment
                    
                Case "!"
                     ' parameter
                     params = Split(temp1, "=")
                     Select Case params(0)
                        Case "!Format", "!Format "
                            Select Case params(1)
                                Case " MPC", "MPC"
                                    TrackCtrl.FileFormat = 1
                                Case "MPC2"
                                    TrackCtrl.FileFormat = 2
                                Case "JPL"
                                    TrackCtrl.FileFormat = 3
                                Case "JPL2"
                                    TrackCtrl.FileFormat = 4
                                Case Else
                                    TrackCtrl.FileFormat = 0
                            End Select
                        Case "!Name"
                            gCustomTrackName = params(1)
                        Case "!Precess"
                            Select Case params(1)
                                Case "1"
                                    TrackCtrl.Precess = True
                                Case "0"
                                    TrackCtrl.Precess = False
                                Case Else
                                    TrackCtrl.Precess = False
                            End Select
                     
                        Case "!Waypoints"
                            Select Case params(1)
                                Case "1"
                                    TrackCtrl.Waypoint = True
                                Case "0"
                                    TrackCtrl.Waypoint = False
                                Case Else
                                    TrackCtrl.Waypoint = False
                            End Select
                            
                        Case "!End"
                            GoTo ParseEnd
                     
                     End Select
                
                Case Else
                    Select Case TrackCtrl.FileFormat
                        Case 0
                            
                        Case 1, 2
                            'mpc
                            
                            ' strip out multiple spaces
                            Do While (InStr(temp1, "  "))
                                temp1 = Replace(temp1, "  ", " ")
                            Loop
                                                       
                            temp2 = Split(temp1, " ")
                            
                            year = val(temp2(0))
                            month = val(temp2(1))
                            day = val(temp2(2))
                            hour = val(Left(temp2(3), 2))
                            minute = val(mid(temp2(3), 3, 2))
                            second = val(mid(temp2(3), 5, 2))
'                            day = day + (hour * 3600 + minute * 60 + second) / 86400
                            Call cal_mjd(month, day, year, mjd)
                            ' convert to "julian seconds"
                            mjd = mjd * 86400 + (hour * 3600 + minute * 60 + second)
                            
                            ' calculates current julian date in seconds
                            GetSystemTime typTime
                            day = CDbl(typTime.wDay)
                            Call cal_mjd(typTime.wMonth, day, typTime.wYear, mjdNow)
                            mjdNow = mjdNow * 86400 + (CDbl(typTime.wHour) * 3600 + CDbl(typTime.wMinute) * 60 + CDbl(typTime.wSecond))

                            If mjd < mjdNow Then
                                ' data is earlier then now so keep line number reset to 0
                                lineno = 0
                            End If

                            ' get RA in seconds (of time)
                            RA = val(temp2(4)) * 3600 + val(temp2(5)) * 60 + val(temp2(6))
                            
                            ' get DEC in seconds (angle)
                            DEC = val(temp2(7))
                            If DEC < 0 Then
                                DEC = DEC * 3600 - val(temp2(8)) * 60 - val(temp2(9))
                            Else
                                DEC = DEC * 3600 + val(temp2(8)) * 60 + val(temp2(9))
                            End If
                            
                            DecEncoder = EncoderFromDec(DEC / 3600, RA / 3600)
                                                  
                            RaRate = val(temp2(15))
                            DecRate = val(temp2(16))
                                                                                                   
                            If lineno > 0 Then
                                With TrackNew
'                                    DeltaT = (mjd - Lastmjd) * 86400
                                    deltat = (mjd - Lastmjd)
                                    ' calc change in RA (seconds of time)
                                    .DeltaRa = RA - Lastra
                                    ' calc change in DEC (seconds of angle)
                                    .DeltaDec = DEC - Lastdec
                                    ' Establish DEC direction
                                    If DecEncoder > LastDecEncoder Then
                                        .DecDir = 0
                                    Else
                                        .DecDir = 1
                                    End If
                                    
                                    .time_mjd = Lastmjd
                                    .RAJ2000 = Lastra
                                    .DECJ2000 = Lastdec
                                    If TrackCtrl.FileFormat = 2 Then
                                        ' high precision - calculated
                                        ' Convert from seconds to arcseconds
                                        .RaRate = .DeltaRa * 15 / deltat
                                        .DecRate = .DeltaDec / deltat
                                    Else
                                        ' lower precision - read rates direct from file
                                        .RaRate = LastRaRate
                                        .DecRate = LastDecRate
                                    End If
                                    .RaRateRaw = .RaRate
                                    .DECRateRaw = .DecRate
                                    
                                    ' increase in tracking rate will decrease RA
                                    ' so need to subtract
                                    .RaRate = SID_RATE - .RaRate
                                    If gHemisphere = 1 Then
                                        ' for some reason dll doesn't seem to sort out
                                        ' southern hemisphere movement so we must make it negative
                                        .RaRate = .RaRate * -1
                                    End If
                                    
                                    .DecRate = Abs(.DecRate)
                                    If .DecDir = 1 Then
                                         .DecRate = -1 * .DecRate
                                    End If
                                    
                                    ' add new record
                                    ReDim Preserve TrackCtrl.TrackSchedule(lineno)
                                    TrackCtrl.TrackSchedule(lineno) = TrackNew
                                    Print #NF2, CStr(.time_mjd) & " " & FormatNumber(.DeltaRa, 5) & " " & FormatNumber(.RaRate, 5) & " " & FormatNumber(.DeltaDec, 5) & " " & FormatNumber(.DecRate, 5) & " " & FormatNumber(.DecDir, 0)
                                End With
                            End If
                            lineno = lineno + 1
                            Lastra = RA
                            Lastdec = DEC
                            Lastmjd = mjd
                            LastDecEncoder = DecEncoder
                            LastRaRate = RaRate
                            LastDecRate = DecRate
                            
                            If mjd > (mjdNow + 86400) Then
                                'we've loaded 24 hours of data - should be enough
                                'if it isn't then user can always reload when current set runs out
                                GoTo ParseEnd
                            End If
                        
                        Case 3
                            'JPL
                            
                            temp1 = Replace(temp1, "     ", " ? ")
                            ' strip out multiple spaces
                            Do While (InStr(temp1, "  "))
                                temp1 = Replace(temp1, "  ", " ")
                            Loop
                            temp2 = Split(Trim(temp1), " ")
                            
                            year = val(Left(temp2(0), 4))
                            Select Case mid(temp2(0), 6, 3)
                                Case "Jan"
                                    month = 1
                                Case "Feb"
                                    month = 2
                                Case "Mar"
                                    month = 3
                                Case "Apr"
                                    month = 4
                                Case "May"
                                    month = 5
                                Case "Jun"
                                    month = 6
                                Case "Jul"
                                    month = 7
                                Case "Aug"
                                    month = 8
                                Case "Sep"
                                    month = 9
                                Case "Oct"
                                    month = 10
                                Case "Nov"
                                    month = 11
                                Case "Dec"
                                    month = 12
                            End Select
                            
                            day = val(Right(temp2(0), 2))
                            Select Case Len(temp2(1))
                                Case 5
                                    hour = val(Left(temp2(1), 2))
                                    minute = val(mid(temp2(1), 4, 2))
                                    second = 0
                                Case 8
                                    hour = val(Left(temp2(1), 2))
                                    minute = val(mid(temp2(1), 4, 2))
                                    second = val(Right(temp2(1), 2))
                                Case 12
                                    hour = val(Left(temp2(1), 2))
                                    minute = val(mid(temp2(1), 4, 2))
                                    second = val(Right(temp2(1), 6))
                                Case Else
                                    GoTo ImportError
                            End Select

'                            day = day + (hour * 3600 + minute * 60 + second) / 86400
                            Call cal_mjd(month, day, year, mjd)
                            ' convert to "julian seconds"
                            mjd = mjd * 86400 + (hour * 3600 + minute * 60 + second)

                            ' calculates current julian date in seconds
                            GetSystemTime typTime
                            day = CDbl(typTime.wDay)
                            Call cal_mjd(typTime.wMonth, day, typTime.wYear, mjdNow)
                            mjdNow = mjdNow * 86400 + (CDbl(typTime.wHour) * 3600 + CDbl(typTime.wMinute) * 60 + CDbl(typTime.wSecond))

                            If mjd < mjdNow Then
                                ' data is earlier then now so keep line number reset to 0
                                lineno = 0
                            End If


                            RA = val(temp2(3)) * 3600 + val(temp2(4)) * 60 + val(temp2(5))
                            DEC = val(temp2(6))
                                                       
                            If DEC < 0 Then
                                DEC = DEC * 3600 - val(temp2(7)) * 60 - val(temp2(8))
                            Else
                                DEC = DEC * 3600 + val(temp2(7)) * 60 + val(temp2(8))
                            End If
                            
                            DecEncoder = EncoderFromDec(DEC / 3600, RA / 3600)
                                                  
'                            ' d(RA*cos(Dec))/dt  arcsec/hour
'                            RaRate = val(temp2(9)) / Cos(RA * DEG_RAD / 3600)
'                            RaRate = RaRate / 3600
'                            ' d(DEC)/dt arcsec/hour
'                            DecRate = val(temp2(10)) / 3600
                                                                                                   
                            If lineno > 0 Then
                                With TrackNew
'                                    DeltaT = (mjd - Lastmjd) * 86400
                                    deltat = (mjd - Lastmjd)
                                    .DeltaRa = RA - Lastra
                                    .DeltaDec = DEC - Lastdec
                                    If DecEncoder > LastDecEncoder Then
                                        .DecDir = 0
                                    Else
                                        .DecDir = 1
                                    End If
                                    
                                    .time_mjd = Lastmjd
                                    .RAJ2000 = Lastra
                                    .DECJ2000 = Lastdec
                                    If TrackCtrl.FileFormat = 3 Then
                                        .RaRate = .DeltaRa * 15 / deltat
                                        .DecRate = .DeltaDec / deltat
                                    Else
'                                        .RaRate = LastRaRate
'                                        .DecRate = LastDecRate
                                    End If
                                    .RaRateRaw = .RaRate
                                    .DECRateRaw = .DecRate
                                    
                                    ' increase in tracking rate will decrease RA
                                    ' so need to subtract
                                    .RaRate = SID_RATE - .RaRate
                                    If gHemisphere = 1 Then
                                        ' for some reason dll doesn't seem to sort out
                                        ' southern hemisphere movement so we must make it negative
                                        .RaRate = .RaRate * -1
                                    End If
                                    
                                    .DecRate = Abs(.DecRate)
                                    If .DecDir = 1 Then
                                         .DecRate = -1 * .DecRate
                                    End If
                                    
                                    ' add new record
                                    ReDim Preserve TrackCtrl.TrackSchedule(lineno)
                                    TrackCtrl.TrackSchedule(lineno) = TrackNew
                                    Print #NF2, CStr(.time_mjd) & " " & FormatNumber(.DeltaRa, 5) & " " & FormatNumber(.RaRate, 5) & " " & FormatNumber(.DeltaDec, 5) & " " & FormatNumber(.DecRate, 5) & " " & FormatNumber(.DecDir, 0)
                                End With
                            End If
                            lineno = lineno + 1
                            Lastra = RA
                            Lastdec = DEC
                            Lastmjd = mjd
                            LastDecEncoder = DecEncoder
                            LastRaRate = RaRate
                            LastDecRate = DecRate
                            
                            If mjd > (mjdNow + 86400) Then
                                'we've loaded 24 hours of data - should be enough
                                'if it isn't then user can always reload when current set runs out
                                GoTo ParseEnd
                            End If
                    End Select
            End Select
        End If
    Wend
ParseEnd:
    If lineno >= 2 Then
        Track_LoadFile = True
    Else
         gCustomTrackName = ""
        If Format = 0 Then
            HC.Add_Message ("Tracking File Error: Missing Header")
        Else
            HC.Add_Message ("Tracking File Error: Insufficient Data")
        End If
    End If
    Close #NF1
    Close #NF2
    Exit Function
    
ImportError:
    HC.Add_Message ("Tracking File Error")
    Close #NF1
    Close #NF2
    gCustomTrackName = ""
 
End Function

Public Sub TrackTimer()
    Dim rate As Double
    Dim idx As Integer
    Dim RA As Double
    Dim DEC As Double
    
    If gTrackingStatus = 4 Then
        If TrackCtrl.TrackingChangesEnabled = True Then
            If TrackCtrl.TrackIdx <> -1 Then
                idx = GetTrackFileIdx(TrackCtrl.TrackIdx, False)
                If idx <> -1 Then
                    If idx <> TrackCtrl.TrackIdx Then
                        TrackCtrl.TrackIdx = idx
                        rate = TrackCtrl.TrackSchedule(TrackCtrl.TrackIdx).RaRate
                        If gHemisphere = 1 Then
                            HC.raCustom.Text = FormatNumber(-1 * rate, 5)
                        Else
                            HC.raCustom.Text = FormatNumber(rate, 5)
                        End If
                        If rate <> gRightAscensionRate Then
                            Call CustomMoveAxis(0, rate, False, gCustomTrackName)
                        End If
                        rate = TrackCtrl.TrackSchedule(TrackCtrl.TrackIdx).DecRate
                        If rate <> gDeclinationRate Then
                            Call CustomMoveAxis(1, rate, False, gCustomTrackName)
                        End If
                        HC.decCustom.Text = FormatNumber(gDeclinationRate, 5)
                        If TrackCtrl.Waypoint = True Then
                            ' perform waypoint correction
                            If GetTrackTarget(RA, DEC) = True Then
                                Call goto_TrackTarget(RA + TrackCtrl.AdjustRA, DEC + TrackCtrl.AdjustDEC, True)
                            End If
                        End If
                    End If
                Else
                                
                End If
            End If
        End If
    End If
  
End Sub

Public Function GetTrackFileIdx(StartIdx As Integer, Alert As Boolean) As Integer
Dim i As Integer
Dim typTime As SYSTEMTIME
Dim mjd As Double
Dim day As Double
On Error GoTo HandleError
    
    GetSystemTime typTime

    GetTrackFileIdx = -1
    day = CDbl(typTime.wDay)
'    day = CDbl(typTime.wDay) + (CDbl(typTime.wHour) * 3600 + CDbl(typTime.wMinute) * 60 + CDbl(typTime.wSecond)) / 86400
    Call cal_mjd(typTime.wMonth, day, typTime.wYear, mjd)
    ' calc elasped 'julian' seconds
    mjd = mjd * 86400 + (CDbl(typTime.wHour) * 3600 + CDbl(typTime.wMinute) * 60 + CDbl(typTime.wSecond))

    If StartIdx = 0 Then StartIdx = 1

    ' search forwards through data
    For i = StartIdx To UBound(TrackCtrl.TrackSchedule())
        If TrackCtrl.TrackSchedule(i).time_mjd > mjd Then
            GetTrackFileIdx = i - 1
            Exit Function
        End If
    Next i
    
    ' data set is out of date - try reloading new data set from file
    Track_LoadFile (gCustomTrackFile)
    ' check through all of data
    For i = 1 To UBound(TrackCtrl.TrackSchedule())
        If TrackCtrl.TrackSchedule(i).time_mjd > mjd Then
            GetTrackFileIdx = i - 1
            Exit Function
        End If
    Next i
    
    ' file has no useful data - use last rate we know about
    GetTrackFileIdx = i - 1
    ' turn icon red
    HC.CmdTrack(4).Picture = LoadResPicture(112, vbResBitmap)
    ' send out warning message
    If Alert Then
        HC.Add_Message "Tracking file is out of date!" & vbCrLf & "Using last known rate."
    End If
    Exit Function

HandleError:
    GetTrackFileIdx = -1
    HC.CmdTrack(4).Picture = LoadResPicture(112, vbResBitmap)

End Function

Private Function GetPosnIdx(StartIdx As Integer, Alert As Boolean) As Integer
Dim i As Integer
Dim typTime As SYSTEMTIME
Dim mjd As Double
Dim day As Double
    
    On Error GoTo HandleError
    
    GetSystemTime typTime

    GetPosnIdx = -1
    day = CDbl(typTime.wDay)
'    day = CDbl(typTime.wDay) + (CDbl(typTime.wHour) * 3600 + CDbl(typTime.wMinute) * 60 + CDbl(typTime.wSecond)) / 86400
    Call cal_mjd(typTime.wMonth, day, typTime.wYear, mjd)
    ' calc elasped 'julian' seconds
    mjd = mjd * 86400 + (CDbl(typTime.wHour) * 3600 + CDbl(typTime.wMinute) * 60 + CDbl(typTime.wSecond))

    If StartIdx = 0 Then StartIdx = 1

    For i = StartIdx To UBound(TrackCtrl.TrackSchedule())
        If TrackCtrl.TrackSchedule(i).time_mjd > mjd Then
            GetPosnIdx = i - 1
            Exit Function
        End If
    Next i
    
    ' data set is out of date - try reloading new data set from file
    Track_LoadFile (gCustomTrackFile)
    ' check through all of data
    For i = 1 To UBound(TrackCtrl.TrackSchedule())
        If TrackCtrl.TrackSchedule(i).time_mjd > mjd Then
            GetPosnIdx = i - 1
            Exit Function
        End If
    Next i
    
    If Alert Then
        HC.CmdTrack(4).Picture = LoadResPicture(112, vbResBitmap)
        HC.Add_Message "Tracking data is out of date!"
    End If
    
    Exit Function
    
HandleError:
    If Alert Then
    End If
    GetPosnIdx = -1

End Function

Private Function EncoderFromDec(DEC As Double, RA) As Double
    Dim tPier As Double
    
    If RangeHA(RA - EQnow_lst(gLongitude * DEG_RAD)) < 0 Then
        If gHemisphere = 0 Then
            tPier = 1
        Else
            tPier = 0
        End If
    Else
        If gHemisphere = 0 Then
            tPier = 0
        Else
            tPier = 1
        End If
    End If
    EncoderFromDec = Get_DECEncoderfromDEC(DEC, tPier, gDECEncoder_Zero_pos, gTot_DEC, gHemisphere)

End Function

Public Function goto_TrackTarget(RA As Double, DEC As Double, mute As Boolean)
    Dim idx As Integer
    Dim epochnow As Double
    Dim typTime As SYSTEMTIME
    Dim mjd As Double
    Dim day As Double
    Dim deltat As Double
    
    On Error GoTo HandleError
    
    If gEQparkstatus = 0 Then
        ' slew
        gTargetRA = RA
        gTargetDec = DEC
        HC.Add_Message ("Goto: " & oLangDll.GetLangString(105) & "[ " & FmtSexa(gTargetRA, False) & " ] " & oLangDll.GetLangString(106) & "[ " & FmtSexa(gTargetDec, True) & " ]")
        gSlewCount = gMaxSlewCount   'NUM_SLEW_RETRIES               'Set initial iterative slew count
        Call radecAsyncSlew(gGotoRate)
        If Not mute Then
            EQ_Beep (20)
        End If
    Else
        HC.Add_Message (oLangDll.GetLangString(5000))
    End If
    
HandleError:

End Function

Public Function GetTrackTarget(ByRef RA As Double, ByRef DEC As Double) As Boolean
    Dim idx As Integer
    Dim epochnow As Double
    Dim typTime As SYSTEMTIME
    Dim mjd As Double
    Dim day As Double
    Dim deltat As Double
    
    idx = GetPosnIdx(0, True)
    If idx >= 0 Then
    
        ' get RA,DEC (in seconds)
        RA = TrackCtrl.TrackSchedule(idx).RAJ2000
        DEC = TrackCtrl.TrackSchedule(idx).DECJ2000
    
        ' calculates current julian date in seconds
        GetSystemTime typTime
        day = CDbl(typTime.wDay)
        Call cal_mjd(typTime.wMonth, day, typTime.wYear, mjd)
        mjd = mjd * 86400 + (CDbl(typTime.wHour) * 3600 + CDbl(typTime.wMinute) * 60 + CDbl(typTime.wSecond))
        
        ' establish how many seconds have elapsed since record date/time
        deltat = mjd - TrackCtrl.TrackSchedule(idx).time_mjd
        
        ' compensate for movement
        RA = RA + TrackCtrl.TrackSchedule(idx).RaRateRaw * deltat / 15
        DEC = DEC + TrackCtrl.TrackSchedule(idx).DECRateRaw * deltat
        
        ' convert back into hours
        RA = RA / 3600
        DEC = DEC / 3600
        
        ' adjust to JNOW
        If TrackCtrl.Precess = True Then
            epochnow = 2000 + (now_mjd() - J2000) / 365.25
            Call Precess(RA, DEC, 2000, epochnow)
        End If
        
        GetTrackTarget = True
    Else
        GetTrackTarget = False
    End If

End Function
