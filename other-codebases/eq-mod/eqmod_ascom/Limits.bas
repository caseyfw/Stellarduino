Attribute VB_Name = "Limits"
Option Explicit

Public Type LIMIT
    Alt As Double
    Az As Double
    ha As Double
    DEC As Double
End Type

Public Type TLIMIT_STATUS
    LimitDetected As Boolean
    AtLimit As Boolean
    Horizon As Boolean
    RA As Boolean
End Type

Public LimitStatus As TLIMIT_STATUS
Public LimitArray() As LIMIT        ' used for file I/O
Public LimitArray2(360) As LIMIT    ' constructed from LimitArray to allow speedy indexing by azimuth.
Public gHorizonAlgorithm As Integer
Public gLimitSlews As Integer
Public gLimitPark As Integer
Public gAutoFlipAllowed As Boolean
Public gAutoFlipEnabled As Boolean
Public gSupressHorizonLimits As Boolean
Private AutoFlipState As Integer

Public Sub Limits_Init()
Dim str As String
    
    LimitStatus.Horizon = False
    LimitStatus.RA = False
    LimitStatus.LimitDetected = False

    gSupressHorizonLimits = False
    
    ReDim LimitArray(0)
    Call Limits_BuildLimitDef
  
    str = HC.oPersist.ReadIniValue("LIMIT_ENABLE")
    If str <> "" Then
        HC.ChkEnableLimits.Value = val(str)
    Else
        HC.ChkEnableLimits.Value = 1
    End If
    
    str = HC.oPersist.ReadIniValue("LIMIT_FILE")
    If str <> "" Then
        ' got a file to load
        Limits_ReadFile str
    Else
        ' no file assined - set defaults?
    End If
    
    str = HC.oPersist.ReadIniValue("LIMIT_HORIZON_ALGORITHM")
    If str <> "" Then
        gHorizonAlgorithm = val(str)
    Else
        ' default to interpolated
        gHorizonAlgorithm = 0
        Call HC.oPersist.WriteIniValue("LIMIT_HORIZON_ALGORITHM", "0")
    End If
    
    str = HC.oPersist.ReadIniValue("LIMIT_PARK")
    If str <> "" Then
        gLimitPark = val(str)
    Else
        ' default to interpolated
        gLimitPark = 0
        Call HC.oPersist.WriteIniValue("LIMIT_PARK", "0")
    End If
    
    str = HC.oPersist.ReadIniValue("LIMIT_SLEWS")
    If str <> "" Then
        gLimitSlews = val(str)
    Else
        ' default to interpolated
        gLimitSlews = 1
        Call HC.oPersist.WriteIniValue("LIMIT_SLEWS", "1")
    End If

    Call readAutoFlipData
    AutoFlipState = 0

End Sub
Public Sub Limits_Load()
    FileDlg.filter = "*.txt"
    FileDlg.Show (1)
    Limits_ReadFile FileDlg.FileName
    Call HC.oPersist.WriteIniValue("LIMIT_FILE", FileDlg.FileName)
End Sub

Public Sub Limits_ReadFile(FileName As String)
Dim i As Integer
Dim size, pos, redimcount As Integer
Dim temp1, temp2 As String
Dim ha, DEC As Double
    
    ReDim LimitArray(0)
    On Error GoTo fileerr

    If FileName <> "" Then
        Close #1
        Open FileName For Input As #1
        ReDim LimitArray(100)
        size = 0
        redimcount = 0
        While Not EOF(1)
            Line Input #1, temp1
            temp2 = Left$(temp1, 1)
            If temp2 <> "#" And temp2 <> " " Then
                pos = InStr(temp1, " ")
                If pos <> 0 Then
                    temp2 = Left$(temp1, pos - 1)
                    temp1 = Right$(temp1, Len(temp1) - pos)
                    With LimitArray(size)
                        .Az = CDbl(temp2)
                        .Alt = CDbl(temp1)
                        aa_hadec gLatitude * DEG_RAD, .Alt * DEG_RAD, .Az * DEG_RAD, ha, DEC
                        .ha = Range24(ha * RAD_HRS)
                        .DEC = DEC * RAD_DEG
                    End With
                    size = size + 1
                    redimcount = redimcount + 1
                    If redimcount > 90 Then
                        redimcount = 0
                        ReDim Preserve LimitArray(size + 100)
                    End If
                End If
            End If
        Wend
        ReDim Preserve LimitArray(size)
    End If
    Call Limits_BuildLimitDef
    GoTo endsub

fileerr:
    HC.Add_Message ("Error reading limits file")

endsub:
    Close #1
End Sub

Public Sub Limits_Save()
Dim i As Integer
Dim size As Integer
Dim FileName As String
    
    On Error GoTo fileerr

    size = UBound(LimitArray)
    
    FileDlg.filter = "*.txt"
    FileDlg.Show (1)
    FileName = FileDlg.FileName
    If FileDlg.FileName <> "" Then
        
        ' force a .txt extension
        i = InStr(FileName, ".")
        If i <> 0 Then
            FileName = Left$(FileName, i - 1)
        End If
        FileName = FileName & ".txt"
        
        Close #1
        Open FileName For Output As #1
    
        For i = 0 To size - 1
            Print #1, CStr(CInt(LimitArray(i).Az)) & " " & CStr(LimitArray(i).Alt)
        Next i
        Close #1
    End If
   
    GoTo endsub

fileerr:
    HC.Add_Message ("Error writing limits file")

endsub:
End Sub
Public Sub Limits_Add(ByRef lim As LIMIT)
Dim size, i, j As Integer

    On Error GoTo endsub

    size = UBound(LimitArray)
    
    lim.Az = CInt(lim.Az)
    
    i = 0
    While i < size
        If LimitArray(i).Az > lim.Az Then
            GoTo insert
        Else
            If LimitArray(i).Az = lim.Az Then
                LimitArray(i).Alt = lim.Alt
                GoTo endsub
            End If
        End If
        i = i + 1
    Wend
    GoTo Store
insert:
    For j = size To i + 1 Step -1
        LimitArray(j) = LimitArray(j - 1)
    Next j
Store:
    LimitArray(i) = lim
    ReDim Preserve LimitArray(size + 1)
    Call Limits_BuildLimitDef

endsub:
End Sub
Public Sub Limits_DeleteIdx(idx As Integer)
Dim i, size As Integer
    On Error GoTo endsub
    If idx >= 0 Then
        size = UBound(LimitArray)
        For i = idx To size - 2
            LimitArray(i).Alt = LimitArray(i + 1).Alt
            LimitArray(i).Az = LimitArray(i + 1).Az
        Next i
        ReDim Preserve LimitArray(size - 1)
        Call Limits_BuildLimitDef
    End If
endsub:
End Sub


Public Sub Limits_Execute()
Dim i As Integer
Dim size As Integer
Dim a, b As Integer
Dim Alt As Double
Dim dalt, daz As Double

    LimitStatus.Horizon = False
    LimitStatus.RA = False
    LimitStatus.LimitDetected = False
    
    If HC.ChkEnableLimits.Value = 1 Then
        If gEQparkstatus = 0 Then
            If (gSlewStatus = True And gLimitSlews = 1) Or (gSlewStatus = False And gTrackingStatus > 0) Then
                LimitStatus = Limits_Detect()
                If Limits_Detect.LimitDetected Then
                    LimitStatus.AtLimit = True
                    Call emergency_stop
                    HC.Add_Message (oLangDll.GetLangString(5017))
                    If gLimitPark Then
                        ' park using currently selected park mode.
                        Call HC.ApplyParkMode
                    End If
                Else
                    LimitStatus.AtLimit = False
                End If
            Else
                If LimitStatus.AtLimit Then
                    ' Currently in the limit state so look for clear.
                    LimitStatus = Limits_Detect()
                    LimitStatus.AtLimit = Limits_Detect.LimitDetected
                End If
            End If
        Else
            'If unparking, parking or parked, limits don't apply
            LimitStatus.AtLimit = False
        End If
    Else
        'limits not enabled so we can't be at the limit can we!
        LimitStatus.AtLimit = False
    End If
End Sub

Private Function Limits_Detect() As TLIMIT_STATUS

Dim Alt As Double
Dim LimitDetected As Boolean

    Limits_Detect.LimitDetected = False
    Limits_Detect.Horizon = False
    Limits_Detect.RA = False
    
    ' Routine to handle RA LIMIT processing
    If (gRA_Limit_East <> 0) And (gEmulRA < gRAEncoder_Zero_pos) Then
        If (gEmulRA < gRA_Limit_East) Then
            If gAutoFlipEnabled Then
                Select Case AutoFlipState
                    Case 0
                        'we've hit the RA limit so initiate autoflip!
                        gTargetRA = gRA
                        gTargetDec = gDec
                        HC.Add_Message ("CoordSlew: " & oLangDll.GetLangString(105) & "[ " & FmtSexa(gTargetRA, False) & " ] " & oLangDll.GetLangString(106) & "[ " & FmtSexa(gTargetDec, True) & " ]")
                        gSlewCount = gMaxSlewCount  'NUM_SLEW_RETRIES               'Set initial iterative slew count
                        Call EQ_Beep(2)
                        Call radecAsyncSlew(gGotoRate)
                        AutoFlipState = 1
                    Case Else
                End Select
            Else
                Limits_Detect.RA = True
            End If
            GoTo endsub
        Else
            AutoFlipState = 0
        End If
    End If

    If (gRA_Limit_West <> 0) And (gEmulRA > gRAEncoder_Zero_pos) Then
        If (gEmulRA > gRA_Limit_West) Then
            If gAutoFlipEnabled Then
                Select Case AutoFlipState
                    Case 0
                        'we've hit the RA limit so initiate autoflip!
                        gTargetRA = gRA
                        gTargetDec = gDec
                        HC.Add_Message ("CoordSlew: " & oLangDll.GetLangString(105) & "[ " & FmtSexa(gTargetRA, False) & " ] " & oLangDll.GetLangString(106) & "[ " & FmtSexa(gTargetDec, True) & " ]")
                        gSlewCount = gMaxSlewCount   'NUM_SLEW_RETRIES               'Set initial iterative slew count
                        Call EQ_Beep(2)
                        Call radecAsyncSlew(gGotoRate)
                        AutoFlipState = 1
                    Case Else
                End Select
            Else
                Limits_Detect.RA = True
            End If
            GoTo endsub
        Else
            AutoFlipState = 0
        End If
    End If
    
endsub:
    ' get altitude limit for current azimuth
    If gSupressHorizonLimits = False Then
        If gAlt <= LimitArray2(CInt(gAz)).Alt Then
            Limits_Detect.Horizon = True
        End If
    Else
        Limits_Detect.Horizon = False
    End If


    If Limits_Detect.Horizon = True Or Limits_Detect.RA = True Then
        Limits_Detect.LimitDetected = True
    Else
        Limits_Detect.LimitDetected = False
    End If
End Function

Public Function Limits_GetAltLimit(Az As Double) As Double
Dim i As Integer
Dim size As Integer
Dim a, b As Integer
Dim dalt, daz As Double
    
    ' default to absolute horizon
    Limits_GetAltLimit = 0
    
    On Error GoTo endsub
    
    size = UBound(LimitArray)
    
    Select Case size
    
        Case 0
            Limits_GetAltLimit = 0
            
        Case 1
            Limits_GetAltLimit = LimitArray(0).Alt
        
        Case Else
            If size > 0 Then
            
'                If size = 1 Then
'                    ' only one limit
'                    Limits_GetAltLimit = LimitArray(0).alt
'                    GoTo endsub
 '               End If
        
                a = 0
                For i = 0 To size - 1
                    If LimitArray(i).Az > Az Then
                        a = i
                        GoTo found
                    End If
                Next i
found:
                If a = 0 Then
                    b = size - 1
                Else
                    b = a - 1
                End If
                
                Select Case gHorizonAlgorithm
                    Case 0
                        ' interpolated between two points
                        dalt = LimitArray(a).Alt - LimitArray(b).Alt
                        If LimitArray(a).Az > LimitArray(b).Az Then
                            daz = LimitArray(a).Az - LimitArray(b).Az
                        Else
                            daz = (360 - LimitArray(b).Az) + LimitArray(a).Az
                        End If
                        
                        If daz = 0 Then
                            ' two points with the same azimuth so take the lowest altitude
                            If LimitArray(a).Alt > LimitArray(b).Alt Then
                                Limits_GetAltLimit = LimitArray(a).Alt
                            Else
                               Limits_GetAltLimit = LimitArray(b).Alt
                            End If
                        Else
                            If a = 0 Then
                                If Az < LimitArray(a).Az Then
                                    Limits_GetAltLimit = LimitArray(b).Alt + (dalt / daz * (359 - LimitArray(b).Az + Az))
                                Else
                                    Limits_GetAltLimit = LimitArray(b).Alt + (dalt / daz * (Az - LimitArray(b).Az))
                                End If
                            Else
                                Limits_GetAltLimit = LimitArray(b).Alt + ((Az - LimitArray(b).Az) * dalt / daz)
                            End If
                        End If
                    
                    Case 1
                        ' higher value of two points
                        If LimitArray(a).Alt > LimitArray(b).Alt Then
                            Limits_GetAltLimit = LimitArray(a).Alt
                        Else
                           Limits_GetAltLimit = LimitArray(b).Alt
                        End If
                    
                End Select
                
            End If
    End Select

endsub:
End Function

Public Sub Limits_Clear()
    ReDim LimitArray(0)
    Call Limits_BuildLimitDef
    ' remove reference to limits file
    Call HC.oPersist.WriteIniValue("LIMIT_FILE", "")

End Sub

Public Sub Limits_edit()
    LimitEditForm.Show (0)
End Sub

Public Sub Limits_BuildLimitDef()
Dim idx As Integer
Dim ha, DEC As Double
    
    ' Because of the amount of maths involved to determine current limits we
    ' maintain two arrays. LimitArray is a 'sparse' array used to file storage.
    ' From this we construct LimitArray2 which holds limits for every degree of azimuth.
    ' Limit and display code can therefore quickly access limits by using the current
    ' azimuth as an index into LimitArray(2)
    
    For idx = 0 To 359
        With LimitArray2(idx)
            .Alt = Limits_GetAltLimit(CDbl(idx))
            .Az = idx
            aa_hadec gLatitude * DEG_RAD, .Alt * DEG_RAD, .Az * DEG_RAD, ha, DEC
            .ha = Range24(ha * RAD_HRS)
            .DEC = DEC * RAD_DEG
        End With
    Next idx
End Sub


Public Function Limits_TimeToHorizon() As Double
Dim i As Integer
Dim ha, tmp As Double
    ' Establish the time the scaope will take, at sidereal rate, to reach the horizon
    
    ' -1 indicates never reaches horizon
    Limits_TimeToHorizon = -1
    
    On Error GoTo endsub
    
    ' only consider western horizon (stars just don't set in the east!)
    For i = 180 To 359
        ' search for the point where the horizon declination is greater or equal to our scope declination
        ha = LimitArray2(i).ha
        tmp = LimitArray2(i).DEC
        If tmp >= gDec Then
            ' calulate difference between horizon hour angle and scope hour angle
            
            tmp = ha - Range24(EQnow_lst(gLongitude * DEG_RAD) - gRA)
            If tmp < 0 Then
                tmp = 24 + tmp
            End If
            Limits_TimeToHorizon = tmp
            
            GoTo endsub
        End If
    Next i

endsub:
End Function

Public Function Limits_TimeToMeridian() As Double
Dim Steps As Double
Dim rate As Double
    
    ' Establish the time the scope will take, at sidereal rate, to reach the Meridian limit
    
    ' -1 indicates never reaches horizon
    Limits_TimeToMeridian = -1
    
    On Error GoTo endsub
    
    If gRA_Limit_West <> 0 Then
        If (gEmulRA < gRA_Limit_West) Then
            Steps = gRA_Limit_West - gEmulRA
            ' sidereal rate as steps hour
            rate = 3600 * gTot_RA / 86164.0905
            Limits_TimeToMeridian = Steps / rate
        End If
    End If

endsub:
End Function

Public Sub SetRaLimitDefaults()
    Dim tmp As Double

    ' make up some defaults
    tmp = 90.88 * CDbl(gTot_step) / 360
    gRA_Limit_East = gRAEncoder_Zero_pos - CLng(tmp)          ' homepos - 90.88degrees of step
    gRA_Limit_West = gRAEncoder_Zero_pos + CLng(tmp)          ' homepos + 90.88degrees of step

End Sub

Public Sub writeRAlimit()

    HC.oPersist.WriteIniValue "RA_LIMIT_EAST", CStr(gRA_Limit_East)
    HC.oPersist.WriteIniValue "RA_LIMIT_WEST", CStr(gRA_Limit_West)

End Sub

Public Sub readRALimit()

    Dim tmptxt As String
    Dim i As Long
     
    Call SetRaLimitDefaults
    
    tmptxt = HC.oPersist.ReadIniValue("RA_LIMIT_EAST")
    If tmptxt <> "" Then
       gRA_Limit_East = val(tmptxt)
    Else
        Call HC.oPersist.WriteIniValue("RA_LIMIT_EAST", CStr(gRA_Limit_East))
    End If
    
    tmptxt = HC.oPersist.ReadIniValue("RA_LIMIT_WEST")
    If tmptxt <> "" Then
       gRA_Limit_West = val(tmptxt)
    Else
        Call HC.oPersist.WriteIniValue("RA_LIMIT_WEST", CStr(gRA_Limit_West))
    End If
     
    If gRA_Limit_West = gRA_Limit_East And gRA_Limit_West <> 0 Then
        Call SetRaLimitDefaults
        Call writeRAlimit
    End If
End Sub
Public Function OutOfBounds(ByVal pos As Double) As Boolean
    
    OutOfBounds = False
    
    If HC.ChkEnableLimits.Value = 0 Then
        ' no limits
        OutOfBounds = True
        Exit Function
    End If
    
    ' Routine to handle RA LIMIT processing
    If (gRA_Limit_East <> 0) And (pos < gRAEncoder_Zero_pos) Then
        If (pos < gRA_Limit_East) Then
            OutOfBounds = True
            Exit Function
        End If
    End If

    If (gRA_Limit_West <> 0) And (pos > gRAEncoder_Zero_pos) Then
        If (pos > gRA_Limit_West) Then
            OutOfBounds = True
        End If
    End If

End Function

Public Function RALimitsActive() As Boolean
    If HC.ChkEnableLimits.Value = 0 Then
        RALimitsActive = False
    Else
        If gRA_Limit_West = 0 Or gRA_Limit_West = 0 Then
            RALimitsActive = False
        Else
            RALimitsActive = True
        End If
    End If
End Function
