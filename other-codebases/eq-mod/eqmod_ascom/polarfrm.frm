VERSION 5.00
Begin VB.Form polarfrm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "PolarScope"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6675
   Icon            =   "polarfrm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   325
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   445
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox statusplot 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   423
      TabIndex        =   5
      Top             =   0
      Width           =   6375
   End
   Begin VB.Frame CmdFrame 
      BackColor       =   &H00000000&
      Height          =   4335
      Left            =   4680
      TabIndex        =   1
      Top             =   360
      Width           =   1695
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         ItemData        =   "polarfrm.frx":0CCA
         Left            =   120
         List            =   "polarfrm.frx":0CD4
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton MoveHomeCmd 
         BackColor       =   &H0095C1CB&
         Height          =   495
         Left            =   960
         Picture         =   "polarfrm.frx":0CEB
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton SetHomeCmd 
         BackColor       =   &H0095C1CB&
         Height          =   495
         Left            =   120
         Picture         =   "polarfrm.frx":1459
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton PSAlignCmd 
         BackColor       =   &H0095C1CB&
         Height          =   375
         Left            =   120
         Picture         =   "polarfrm.frx":1BC7
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.PictureBox polarplot 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   120
      ScaleHeight     =   279
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   279
      TabIndex        =   0
      Top             =   480
      Width           =   4215
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   0
         Top             =   3240
      End
   End
End
Attribute VB_Name = "polarfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim warnexit As Boolean
Dim status As Integer
Dim strStatus As String
Dim ReticuleEpoch As Double
Dim ReticuleD1 As Double
Dim ReticuleD2 As Double

Private Sub Combo1_Click()
    Call HC.oPersist.WriteIniValue("POLAR_RETICULE_START", CStr(Combo1.ListIndex))
End Sub

Private Sub AlignPolarScope()
Dim steps360, steps180, steps90, steps270, max, min As Long
Dim ha As Double
Dim targetRAEncoder As Double
Dim currentRAEncoder As Double
Dim currentDECEncoder As Double
Dim DeltaRAStep As Long
Dim RASlowdown As Long
    
        If gTrackingStatus Then
            ' stop tracking
              eqres = EQ_MotorStop(0)
              eqres = EQ_MotorStop(1)
              gTrackingStatus = 0
        End If
        
        steps360 = EQGetTotal360microstep(0)
        ha = gPolHa
        currentRAEncoder = EQGetMotorValues(0)
        steps180 = steps360 / 2
        steps90 = steps180 / 2
        steps270 = steps180 + steps90
        max = gRAEncoder_Zero_pos + steps180
        min = gRAEncoder_Zero_pos - steps180
        
        Select Case Combo1.ListIndex
            Case 0
                ' starting from 3 o'clock position
                If gHemisphere = 0 Then
                    targetRAEncoder = (currentRAEncoder - steps90) + steps360 * ha / 24
                Else
                    targetRAEncoder = (currentRAEncoder - steps270) + steps360 * ha / 24
                End If
        
            Case 1
                ' starting from 6 o'clock position
                    targetRAEncoder = (currentRAEncoder - steps360) + steps360 * ha / 24
            
            Case 2
                ' starting from 9 o'clock position
                If gHemisphere = 0 Then
                    targetRAEncoder = (currentRAEncoder - steps270) + steps360 * ha / 24
                Else
                    targetRAEncoder = (currentRAEncoder - steps90) + steps360 * ha / 24
                End If
        
            Case 3
                ' starting from 12 o'clock position
                targetRAEncoder = (currentRAEncoder - steps180) + steps360 * ha / 24
        End Select
        
        ' sort out southern hemisphere movement
        If gHemisphere = 1 Then
            DeltaRAStep = currentRAEncoder - targetRAEncoder
            targetRAEncoder = currentRAEncoder + DeltaRAStep
        End If
              
        ' keep values in bounds
        If targetRAEncoder < min Then
            targetRAEncoder = targetRAEncoder + steps360
        Else
            If targetRAEncoder > max Then
                targetRAEncoder = targetRAEncoder - steps360
            End If
        End If
        
        DeltaRAStep = Abs(targetRAEncoder - currentRAEncoder)
        RASlowdown = GetSlowdown(DeltaRAStep)
        
        'Execute the actual slew
        If DeltaRAStep <> 0 Then
            If targetRAEncoder > currentRAEncoder Then
                eqres = EQStartMoveMotor(0, 0, 0, DeltaRAStep, RASlowdown)
            Else
                eqres = EQStartMoveMotor(0, 0, 1, DeltaRAStep, RASlowdown)
            End If
        End If
        

End Sub




Private Sub Combo2_Click()
    Call HC.oPersist.WriteIniValue("POLAR_RETICULE_TYPE", CStr(Combo2.ListIndex))
End Sub

Private Sub PSAlignCmd_Click()
    If status <> 0 Then Exit Sub
    If MsgBox(oLangDll.GetLangString(2406), vbYesNo Or vbDefaultButton2) = vbYes Then
        status = 1
    End If
End Sub

Private Sub Form_Load()
    Dim tmptxt As String
    Dim tmp As Integer

    Call SetText
    Call readPolarHomeGoto
    
    Timer1.Enabled = True
    Call PutWindowOnTop(polarfrm)
    MoveHomeCmd.Enabled = False
    SetHomeCmd.Enabled = False
    PSAlignCmd.Enabled = False
    warnexit = False
    status = 0
    
    tmptxt = HC.oPersist.ReadIniValue("POLAR_RETICULE_TYPE")
    If tmptxt <> "" Then
        Select Case tmptxt
            Case "0"
                Combo2.ListIndex = 0
            Case 1
                Combo2.ListIndex = 1
            Case Else
                Combo1.ListIndex = 0
                Call HC.oPersist.WriteIniValue("POLAR_RETICULE_TYPE", "0")
        End Select
    Else
        Combo1.ListIndex = 0
        Call HC.oPersist.WriteIniValue("POLAR_RETICULE_TYPE", "0")
    End If
    
    tmptxt = HC.oPersist.ReadIniValue("POLAR_RETICULE_EPOCH")
    If tmptxt <> "" Then
        ReticuleEpoch = CDbl(tmptxt)
    Else
        ReticuleEpoch = 2000
        Call HC.oPersist.WriteIniValue("POLAR_RETICULE_EPOCH", "2000")
    End If
    
    tmptxt = HC.oPersist.ReadIniValue("POLAR_RETICULE_D1")
    If tmptxt <> "" Then
        ReticuleD1 = CDbl(tmptxt)
    Else
        ReticuleD1 = 2.67
        Call HC.oPersist.WriteIniValue("POLAR_RETICULE_D1", CStr(ReticuleD1))
    End If
    
    tmptxt = HC.oPersist.ReadIniValue("POLAR_RETICULE_D2")
    If tmptxt <> "" Then
        ReticuleD2 = CDbl(tmptxt)
    Else
        ReticuleD2 = 0.355
        Call HC.oPersist.WriteIniValue("POLAR_RETICULE_D2", CStr(ReticuleD2))
    End If
    
    tmptxt = HC.oPersist.ReadIniValue("POLAR_RETICULE_START")
    If tmptxt <> "" Then
        Select Case tmptxt
            Case "0"
                Combo1.ListIndex = 0
            Case 1
                Combo1.ListIndex = 1
            Case "2"
                Combo1.ListIndex = 2
            Case "3"
                Combo1.ListIndex = 3
            Case Else
                Combo1.ListIndex = 1
                Call HC.oPersist.WriteIniValue("POLAR_RETICULE_START", "1")
        End Select
    Else
        Combo1.ListIndex = 1
        Call HC.oPersist.WriteIniValue("POLAR_RETICULE_START", "1")
    End If
    
    If gEQparkstatus = 0 Then
        If HC.ChkEnableLimits.Value = 0 Then
            SetHomeCmd.Enabled = True
            PSAlignCmd.Enabled = True
            If gRAEncoderPolarHomeGoto <> 0 And gDECEncoderPolarHomeGoto <> 0 Then
                MoveHomeCmd.Enabled = True
            End If
            status = 0
        Else
            status = -1
            warnexit = True
        End If
    Else
        status = -2
    End If
    
    If gPoleStarIdx <> 0 Then Combo2.Visible = False

    Call drawscope
    
End Sub

Private Sub Form_Resize()
Dim width As Double
    polarplot.Top = 25 * polarfrm.ScaleHeight / 325
    width = 257 * (polarfrm.ScaleHeight - polarplot.Top) / 277
    polarplot.width = width
    polarplot.Height = width
    statusplot.width = polarfrm.ScaleWidth
    statusplot.Height = 25 * polarfrm.ScaleHeight / 325
    CmdFrame.Top = 25 * polarfrm.ScaleHeight / 325
    CmdFrame.Left = width + 20
    CmdFrame.Height = width
    Call drawscope
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False

    If warnexit Then
        If HC.ChkEnableLimits.Value = 0 Then
            'limits are now off - they were originally on
            MsgBox (oLangDll.GetLangString(2407))
        End If
    End If
    
End Sub

Private Sub MoveHomeCmd_Click()
Dim currentRAEncoder As Double
Dim currentDECEncoder As Double
Dim DeltaRAStep As Long
Dim DeltaDECStep As Long
Dim tmptxt As String

    On Error Resume Next
    
    If status <> 0 Then Exit Sub
    
    If MsgBox(oLangDll.GetLangString(2406), vbYesNo Or vbDefaultButton2) = vbYes Then

        tmptxt = HC.oPersist.ReadIniValue("POLARHOME_RETICULE_START")
        If tmptxt <> "" Then
           Combo1.ListIndex = val(tmptxt)
        Else
           Combo1.ListIndex = 0
        End If
        
        If gTrackingStatus Then
            ' stop tracking
              eqres = EQ_MotorStop(0)
              eqres = EQ_MotorStop(1)
              gTrackingStatus = 0
        End If
       
        currentRAEncoder = EQGetMotorValues(0)
        currentDECEncoder = EQGetMotorValues(1)
        DeltaRAStep = Abs(currentRAEncoder - gRAEncoderPolarHomeGoto)
        DeltaDECStep = Abs(currentDECEncoder - gDECEncoderPolarHomeGoto)
        
        
        If DeltaRAStep <> 0 Then
            If currentRAEncoder < gRAEncoderPolarHomeGoto Then
                eqres = EQStartMoveMotor(0, 0, 0, DeltaRAStep, GetSlowdown(DeltaRAStep))
            Else
                eqres = EQStartMoveMotor(0, 0, 1, DeltaRAStep, GetSlowdown(DeltaRAStep))
            End If
        End If
        
        If DeltaDECStep <> 0 Then
            If currentDECEncoder < gDECEncoderPolarHomeGoto Then
                eqres = EQStartMoveMotor(1, 0, 0, DeltaDECStep, GetSlowdown(DeltaDECStep))
            Else
                eqres = EQStartMoveMotor(1, 0, 1, DeltaDECStep, GetSlowdown(DeltaDECStep))
            End If
        End If
        
        PSAlignCmd.Enabled = False
        SetHomeCmd.Enabled = False
        MoveHomeCmd.Enabled = False
        status = 3
        
        Call drawscope
     
     End If

End Sub

Private Sub SetHomeCmd_Click()
        
    If MsgBox(oLangDll.GetLangString(2402), vbYesNo Or vbDefaultButton2) = vbYes Then
       
        eqres = EQ_MotorStop(0)          ' Stop RA Motor
        If eqres <> 0 Then
            GoTo ENDDefinePolarHome
        End If
        eqres = EQ_MotorStop(1)          ' Stop DEC Motor
        If eqres <> 0 Then
            GoTo ENDDefinePolarHome
        End If
    
        'Wait until RA motor is stable
        Do
            eqres = EQ_GetMotorStatus(0)
            If eqres = 1 Then
                GoTo ENDDefinePolarHome
             End If
        Loop While (eqres And EQ_MOTORBUSY) <> 0
    
        'Wait until DEC motor is stable
        Do
            eqres = EQ_GetMotorStatus(1)
            If eqres = 1 Then
                GoTo ENDDefinePolarHome
            End If
        Loop While (eqres And EQ_MOTORBUSY) <> 0
        
        'Read Motor Values
        gRAEncoderPolarHomeGoto = EQGetMotorValues(0)
        gDECEncoderPolarHomeGoto = EQGetMotorValues(1)
    
        ' save to ini
        writePolarHomeGoto (Combo1.ListIndex)
    End If
ENDDefinePolarHome:
End Sub

Private Sub Timer1_Timer()

    
    If gEQparkstatus = 0 Then
        If HC.ChkEnableLimits.Value = 0 Then
            
            If status < 0 Then
                status = 0
            End If
            
            
            Select Case status
                Case 0
                    ' do nothing but wait for alignment command
                    If BTN_POLARSCOPEALIGN <> 0 Then
                        If (gEQjbuttons = BTN_POLARSCOPEALIGN) Then
                            gEQjbuttons = 0
                            If PSAlignCmd.Enabled Then
                                PSAlignCmd.Enabled = False
                                SetHomeCmd.Enabled = False
                                MoveHomeCmd.Enabled = False
                                Call EQ_Beep(25)
                                Call AlignPolarScope
                                status = 4
                            End If
                        End If
                    End If
                    
                Case 1
                    ' start alignment
                    PSAlignCmd.Enabled = False
                    SetHomeCmd.Enabled = False
                    MoveHomeCmd.Enabled = False
                    Call EQ_Beep(25)
                    Call AlignPolarScope
                    status = 4

                Case 3
                    ' move to polarscope home
                    If ((EQ_GetMotorStatus(0) And EQ_MOTORBUSY) = 0) And ((EQ_GetMotorStatus(1) And EQ_MOTORBUSY) = 0) Then
                        Call EQ_Beep(24)
                        status = 0
                        PSAlignCmd.Enabled = True
                        SetHomeCmd.Enabled = True
                        If gRAEncoderPolarHomeGoto <> 0 And gDECEncoderPolarHomeGoto <> 0 Then
                            MoveHomeCmd.Enabled = True
                        End If
                    End If
                    
                Case 4
                    ' align polarscope
                    If ((EQ_GetMotorStatus(0) And EQ_MOTORBUSY) = 0) And ((EQ_GetMotorStatus(1) And EQ_MOTORBUSY) = 0) Then
                        ' good idea to track
                        Call EQStartSidereal
                        status = 0
                        Call EQ_Beep(26)
                        PSAlignCmd.Enabled = True
                        SetHomeCmd.Enabled = True
                        If gRAEncoderPolarHomeGoto <> 0 And gDECEncoderPolarHomeGoto <> 0 Then
                            MoveHomeCmd.Enabled = True
                        End If
                    End If
            End Select
            
        Else
            SetHomeCmd.Enabled = False
            PSAlignCmd.Enabled = False
            status = -1
        End If
    Else
        SetHomeCmd.Enabled = False
        PSAlignCmd.Enabled = False
        status = -2
    End If
    
    Call drawscope

End Sub

Private Sub DrawSyntaJ2000()
Dim i As Integer
Dim x1, y1, x2, y2, tmp, tmp2, PolHaRet As Double
Dim StarRad, RetRad, rad1, rad2, rad, centre, offset As Double
Dim DecScale, SmallCircleRad As Double
Dim penwidth As Integer
Dim PenCol As Long
Dim RA As Double
Dim DEC As Double
Dim TmpLst As Double
    
    rad = polarplot.ScaleWidth - 6
    penwidth = rad / 277
    If penwidth <= 0 Then
        penwidth = 1
    End If
    
    centre = polarplot.ScaleWidth / 2
    rad1 = rad / 2
    rad2 = rad / 3
    offset = rad / 9
    RetRad = rad2 + offset / 2
    
    polarplot.Cls
    polarplot.DrawWidth = penwidth
    PenCol = &H202040
    RA = gPoleStarRaJ2000
    DEC = gPoleStarDecJ2000
    If ReticuleEpoch <> 2000 Then
        Call Precess(RA, DEC, 2000, ReticuleEpoch)
    End If
'    DecScale = RetRad / (90 - Abs(gPoleStarDec))
'    StarRad = (90 - Abs(DEC)) * DecScale
    DecScale = RetRad / (90 - Abs(DEC))
    StarRad = (90 - Abs(gPoleStarDec)) * DecScale
    SmallCircleRad = ReticuleD2 * RetRad / ReticuleD1

    ' centre cross hairs
    polarplot.Line (centre, centre - offset)-(centre, centre + offset), vbRed
    polarplot.Line (centre - offset, centre)-(centre + offset, centre), vbRed
    ' inner scale
    polarplot.Circle (centre, centre), rad2, PenCol
    ' outer scale
    polarplot.Circle (centre, centre), rad2 + offset, PenCol
    ' scale graduations
    For i = 0 To 360 Step 15
        tmp = i * PI / 180
       x1 = rad2 * Cos(tmp) + centre
       y1 = rad2 * Sin(tmp) + centre
       x2 = (rad2 + offset) * Cos(tmp) + centre
       y2 = (rad2 + offset) * Sin(tmp) + centre
       polarplot.Line (x1, y1)-(x2, y2), PenCol
    Next i
    
    polarplot.DrawWidth = 2 * penwidth
    polarplot.Line (centre, centre - rad1)-(centre, centre - rad2), PenCol
    polarplot.Line (centre, centre + rad2)-(centre, centre + rad1), PenCol
    polarplot.Line (centre - rad1, centre)-(centre - rad2, centre), PenCol
    polarplot.Line (centre + rad2, centre)-(centre + rad1, centre), PenCol
    polarplot.Circle (centre, centre), rad1, PenCol
     
    TmpLst = EQnow_lst(gLongitude * DEG_RAD)
    PolHaRet = Range24(TmpLst - RA)
    If gHemisphere = 1 Then
        tmp = 90 + (gPolHa * 360 / 24)
        tmp2 = 90 + (PolHaRet * 360 / 24)
    Else
        tmp = 90 - (gPolHa * 360 / 24)
        tmp2 = 90 - (PolHaRet * 360 / 24)
    End If
    tmp = tmp * PI / 180
    tmp2 = tmp2 * PI / 180
    
    polarplot.Circle (centre, centre), RetRad, vbRed
    
    ' Ghost J2000 circle
    x1 = (RetRad) * Cos(tmp2) + centre
    y1 = (RetRad) * Sin(tmp2) + centre
    polarplot.Circle (x1, y1), SmallCircleRad, &H203020
    
    ' JNow circle
    x1 = (RetRad) * Cos(tmp) + centre
    y1 = (RetRad) * Sin(tmp) + centre
    polarplot.Circle (x1, y1), SmallCircleRad, vbRed
    
    x1 = (StarRad) * Cos(tmp) + centre
    y1 = (StarRad) * Sin(tmp) + centre
    polarplot.DrawWidth = 4 * penwidth
    polarplot.Circle (x1, y1), 1, vbWhite

    polarplot.FontSize = 10 * penwidth
    polarplot.FontBold = True
    polarplot.FontName = "Arial"
    polarplot.ForeColor = vbRed
    polarplot.CurrentX = centre - polarplot.TextWidth(HC.CommandPolaris.Caption) / 2
    polarplot.CurrentY = centre - 1 * offset - polarplot.TextHeight("0")
    polarplot.Print (HC.CommandPolaris.Caption)
    
End Sub

Private Sub drawscope()
Dim i As Integer
Dim x1, y1, x2, y2, tmp As Double
Dim StarRad, RetRad, rad1, rad2, rad, centre, offset As Double
Dim DecScale, SmallCircleRad As Double
Dim penwidth As Integer
Dim PenCol As Long

    If gPoleStarIdx <> 0 Then
        Call DrawGeneric
    Else
        Select Case Combo2.ListIndex
            Case 1
                Call DrawSyntaJ2000
            Case Else
                Call DrawGeneric
        End Select
    End If

    Select Case status
        Case 0
            strStatus = oLangDll.GetLangString(2408) & " " & oLangDll.GetLangString(2412)
        Case 1
            strStatus = oLangDll.GetLangString(2408) & " " & oLangDll.GetLangString(2413)
        Case 2
            strStatus = oLangDll.GetLangString(2408) & " " & oLangDll.GetLangString(2410)
        Case 3
            strStatus = oLangDll.GetLangString(2408) & " " & oLangDll.GetLangString(2414)
        Case 4
            strStatus = oLangDll.GetLangString(2408) & " " & oLangDll.GetLangString(2411)
        Case -1
            strStatus = oLangDll.GetLangString(2408) & " " & oLangDll.GetLangString(2405)
        Case -2
            strStatus = oLangDll.GetLangString(2408) & " " & oLangDll.GetLangString(2404)
    End Select
    
    penwidth = (polarplot.ScaleWidth - 6) / 277
    If penwidth <= 0 Then
        penwidth = 1
    End If
    
    statusplot.Cls
    statusplot.FontSize = 10 * penwidth
    statusplot.FontBold = True
    statusplot.FontName = "Arial"
    statusplot.ForeColor = vbRed
    statusplot.CurrentX = 0
    statusplot.CurrentY = 0
    statusplot.Print (strStatus)

End Sub

Private Sub DrawGeneric()
Dim i As Integer
Dim x1, y1, x2, y2, tmp As Double
Dim StarRad, RetRad, rad1, rad2, rad, centre, offset As Double
Dim DecScale, SmallCircleRad As Double
Dim penwidth As Integer
Dim PenCol As Long

    rad = polarplot.ScaleWidth - 6
    penwidth = rad / 277
    If penwidth <= 0 Then
        penwidth = 1
    End If
    
    centre = polarplot.ScaleWidth / 2
    rad1 = rad / 2
    rad2 = rad / 3
    offset = rad / 16
    RetRad = rad2 + offset / 2
    
    polarplot.Cls
    polarplot.DrawWidth = penwidth
    PenCol = vbRed
    StarRad = rad2 + offset / 2
    
    ' centre cross hairs
    polarplot.Line (centre, centre - offset)-(centre, centre + offset), vbRed
    polarplot.Line (centre - offset, centre)-(centre + offset, centre), vbRed
    ' inner scale
    polarplot.Circle (centre, centre), rad2, PenCol
    ' outer scale
    polarplot.Circle (centre, centre), rad2 + offset, PenCol
    ' scale graduations
    For i = 0 To 360 Step 15
        tmp = i * PI / 180
       x1 = rad2 * Cos(tmp) + centre
       y1 = rad2 * Sin(tmp) + centre
       x2 = (rad2 + offset) * Cos(tmp) + centre
       y2 = (rad2 + offset) * Sin(tmp) + centre
       polarplot.Line (x1, y1)-(x2, y2), PenCol
    Next i
    
    
    polarplot.DrawWidth = 2 * penwidth
    polarplot.Line (centre, centre - rad1)-(centre, centre - rad2), PenCol
    polarplot.Line (centre, centre + rad2)-(centre, centre + rad1), PenCol
    polarplot.Line (centre - rad1, centre)-(centre - rad2, centre), PenCol
    polarplot.Line (centre + rad2, centre)-(centre + rad1, centre), PenCol
    polarplot.Circle (centre, centre), rad1, PenCol
    
     
    If gHemisphere = 1 Then
        tmp = 90 + (gPolHa * 360 / 24)
    Else
        tmp = 90 - (gPolHa * 360 / 24)
    End If
    tmp = tmp * PI / 180
    
    x1 = (StarRad) * Cos(tmp) + centre
    y1 = (StarRad) * Sin(tmp) + centre
    polarplot.DrawWidth = 4 * penwidth
    polarplot.Circle (x1, y1), 1, vbWhite

    polarplot.FontSize = 10 * penwidth
    polarplot.FontBold = True
    polarplot.FontName = "Arial"
    polarplot.ForeColor = vbRed
    polarplot.CurrentX = centre - polarplot.TextWidth(HC.CommandPolaris.Caption) / 2
    polarplot.CurrentY = centre - 2 * offset - polarplot.TextHeight("0")
    polarplot.Print (HC.CommandPolaris.Caption)
    
End Sub



Private Sub SetText()
    polarfrm.Caption = oLangDll.GetLangString(2400)
    PSAlignCmd.ToolTipText = oLangDll.GetLangString(2401)
    SetHomeCmd.ToolTipText = oLangDll.GetLangString(2402)
    MoveHomeCmd.ToolTipText = oLangDll.GetLangString(2403)
    Combo1.AddItem (oLangDll.GetLangString(2417))
    Combo1.AddItem (oLangDll.GetLangString(2415))
    Combo1.AddItem (oLangDll.GetLangString(2418))
    Combo1.AddItem (oLangDll.GetLangString(2416))
End Sub

