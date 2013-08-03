Attribute VB_Name = "Parking"
Option Explicit

Public Type parkpos
    name As String
    posR  As Long
    posD As Long
End Type

Public gParkParams As GOTO_PARAMS
Public UserParks(10) As parkpos
Public UserUnparks(10) As parkpos

Public Sub ApplyUnParkMode2(mode As Integer)
    ' only ever allow unparks if mount is already parked
    If gEQparkstatus = 1 Then
        Select Case mode
            Case 0
                'Unpark
                 Call Unparkscope
            Case 1
                'Unpark and slew to last position
                UnparkscopeToLastPos
            Case Else
                'Unpark and slew to user defined startup position
                Call UnparkscopeToUserDef(UserUnparks(mode - 1))
        End Select
    End If
End Sub

Public Sub ApplyParkMode2(mode As Integer)
    Select Case mode
        Case 0
            ' Park to Home
            Call ParkHome
        Case 1
            ' Park to current
            Call Park2Current
        Case Else
            'Park to defined
            Call ParktoUserDefine2(UserParks(mode - 1))
    End Select
End Sub

Private Function ParkInit() As Boolean
    HC.ParkTimer.Enabled = False
    HC.Frame15.Caption = oLangDll.GetLangString(146) & " " & oLangDll.GetLangString(50)
    gEQparkstatus = 2
    
    ' Save Alignment if required
    Call AligmentStarsPark
    
    Call StopTrackingUpdates

    ' stop any active slews from completing
    gSlewStatus = False
    gRAStatus_slew = False
    
    ' clear an active flips
    HC.ChkForceFlip.Value = 0
    gCWUP = False
    gGotoParams.SuperSafeMode = 0
    
    ' stop the motors
    eqres = EQ_MotorStop(0)          ' Stop RA Motor
    If eqres <> EQ_OK Then
        ParkInit = False
        Exit Function
    End If
    eqres = EQ_MotorStop(1)          ' Stop DEC Motor
    If eqres <> EQ_OK Then
        ParkInit = False
        Exit Function
    End If

    'Wait until RA motor is stable
    Do
        eqres = EQ_GetMotorStatus(0)
        If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then
             ParkInit = False
             Exit Function
        End If
    Loop While (eqres And EQ_MOTORBUSY) <> 0
    
    'Wait until DEC motor is stable
    Do
        eqres = EQ_GetMotorStatus(1)
        If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then
             ParkInit = False
             Exit Function
        End If
    Loop While (eqres And EQ_MOTORBUSY) <> 0
    
    ' update tracking status
    gTrackingStatus = 0
    HC.TrackingFrame.Caption = oLangDll.GetLangString(121) & " " & oLangDll.GetLangString(178)

    ParkInit = True

End Function


Public Sub ParkHome()
    Dim currentdecpos As Double
    Dim currentrapos As Double

    ' only allow parking if currently unparked
    If gEQparkstatus = 0 Then
    
        If ParkInit() = True Then
            'Read Motor Values
            currentrapos = EQGetMotorValues(0)
            currentdecpos = EQGetMotorValues(1)
            gRAEncoderlastpos = currentrapos
            gDECEncoderlastpos = currentdecpos
            writelastpos   ' Save current position
            
            gRAEncoderUNPark = RAEncoder_Home_pos
            gDECEncoderUNPark = gDECEncoder_Home_pos
            
            Call writeUnpark
            
            Call StartPark(currentrapos, RAEncoder_Home_pos, currentdecpos, gDECEncoder_Home_pos)
            
            HC.Add_Message (oLangDll.GetLangString(5035))
            HC.ParkTimer.Enabled = True
            
            ' No need to wait at this point - return control to main routine
            Call EQ_Beep(5)
            Call writeRAlimit
            Call SetParkCaption
        End If
    End If
Endhome:

End Sub

' ascom park
Public Sub ParktoUserDefine()
    Dim currentdecpos As Double
    Dim currentrapos As Double
    
    ' only allow parking if currently unparked
    If gEQparkstatus = 0 Then
        If ParkInit() = True Then
            'Read Motor Values
            currentrapos = EQGetMotorValues(0)
            currentdecpos = EQGetMotorValues(1)
            gRAEncoderlastpos = currentrapos
            gDECEncoderlastpos = currentdecpos
            Call writelastpos   ' Save current position
                
            Call readpark       ' Read Userdefined Park data
            gRAEncoderUNPark = gRAEncoderPark
            gDECEncoderUNPark = gDECEncoderPark
            
            Call writeUnpark    ' Save Unpark Data
            
            Call StartPark(currentrapos, CDbl(gRAEncoderPark), currentdecpos, CDbl(gDECEncoderPark))
            HC.Add_Message oLangDll.GetLangString(5003)
            HC.ParkTimer.Enabled = True
        
            Call EQ_Beep(5)
            Call writeRAlimit
            Call SetParkCaption
        End If
    End If
Endparkuser:

End Sub

Public Sub ParktoUserDefine2(userpark As parkpos)
    Dim currentdecpos As Double
    Dim currentrapos As Double
    
    ' only allow parking if currently unparked
    If gEQparkstatus = 0 Then
        
        ' don't park if user position is undefined!
        If userpark.posD = 0 Or userpark.posR = 0 Then GoTo Endparkuser2
    
        If ParkInit() = True Then
            'Read Motor Values
            currentrapos = EQGetMotorValues(0)
            currentdecpos = EQGetMotorValues(1)
            gRAEncoderlastpos = currentrapos
            gDECEncoderlastpos = currentdecpos
            
            ' Save current position
            Call writelastpos
                
            'set target for encoders
            gRAEncoderPark = userpark.posR
            gDECEncoderPark = userpark.posD
            
            ' set the unpark position
            gRAEncoderUNPark = gRAEncoderPark
            gDECEncoderUNPark = gDECEncoderPark
            
            ' Save Unpark Data
            Call writeUnpark
            
            Call StartPark(currentrapos, CDbl(gRAEncoderPark), currentdecpos, CDbl(gDECEncoderPark))
            HC.Add_Message oLangDll.GetLangString(5003)
            HC.ParkTimer.Enabled = True
        
            Call EQ_Beep(5)
            Call writeRAlimit
            Call SetParkCaption
        End If
    End If
Endparkuser2:

End Sub
Public Sub Park2Current()

    ' only allow parking if currently unparked or unparking (used for emergency stop)
    If gEQparkstatus <> 1 Then
        
        If ParkInit() = True Then
            gRAEncoderPark = EQGetMotorValues(0)
            gDECEncoderPark = EQGetMotorValues(1)
        
            HC.TrackingFrame.Caption = oLangDll.GetLangString(121) & " " & oLangDll.GetLangString(178)
            HC.Frame15.Caption = oLangDll.GetLangString(146) & " " & oLangDll.GetLangString(177)
        
            ' Save Alignment if required
            Call AligmentStarsPark
        
            gRAEncoderlastpos = gRAEncoderPark
            gDECEncoderlastpos = gDECEncoderPark
            
            ' Save current position
            Call writelastpos
                
            gRAEncoderUNPark = gRAEncoderPark
            gDECEncoderUNPark = gDECEncoderPark
            
            ' Save Unpark Data
            Call writeUnpark
            
            ' set staues as parked
            gEQparkstatus = 1
            
            ' save park status just incase we don't shutdown
            Call writeParkStatus(gEQparkstatus)
            
            Call EQ_Beep(8)
            HC.Add_Message (oLangDll.GetLangString(5003))
            Call writeRAlimit
            Call SetParkCaption
        End If
    End If
ENDParkToCurrent:

End Sub

Public Sub EmergencyStopPark()

    ' only allow parking if currently unparked or unparking (used for emergency stop)
    If gEQparkstatus <> 1 Then
        
        If ParkInit() = True Then
            gRAEncoderPark = EQGetMotorValues(0)
            gDECEncoderPark = EQGetMotorValues(1)
        
            HC.TrackingFrame.Caption = oLangDll.GetLangString(121) & " " & oLangDll.GetLangString(178)
            HC.Frame15.Caption = oLangDll.GetLangString(146) & " " & oLangDll.GetLangString(177)
        
            ' Save Alignment if required
            Call AligmentStarsPark
        
            gRAEncoderlastpos = gRAEncoderPark
            gDECEncoderlastpos = gDECEncoderPark
            
            ' Save current position
            Call writelastpos
                
            gRAEncoderUNPark = gRAEncoderPark
            gDECEncoderUNPark = gDECEncoderPark
            
            ' Save Unpark Data
            Call writeUnpark
            
            ' set staues as parked
            gEQparkstatus = 1
            
            ' save park status just incase we don't shutdown
            Call writeParkStatus(gEQparkstatus)
            
            Call EQ_Beep(7)
            HC.Add_Message (oLangDll.GetLangString(5003))
            Call writeRAlimit
            Call SetParkCaption
        End If
    End If
ENDParkToCurrent:

End Sub



Public Sub StartPark(ByVal currentRa As Double, ByVal TargetRA As Double, ByVal currentDec As Double, ByVal TargetDEC As Double)
Dim i As Long
Dim j As Long
Dim hours As Double

    If gParkParams.SuperSafeMode = 0 Then
        gParkParams.RA_targetencoder = TargetRA
        gParkParams.DEC_targetencoder = TargetDEC
        gParkParams.Rate = gGotoRate
        
        If RALimitsActive() = False Then
            ' Limits are off
            If gRA_Hours > 12 Then
                ' current position is CW up
                If currentRa > RAEncoder_Home_pos Then
                    'Slew in RA only to nearest limit position
                    'then slew in RA/DEfollowed by dual axis slew
                    gParkParams.SuperSafeMode = 4
                    TargetRA = gRAMeridianWest
                    TargetDEC = currentDec
                Else
                    'Slew in RA to limit position - then complete move as dual axis slew
                    gParkParams.SuperSafeMode = 4
                    TargetRA = gRAMeridianEast
                    TargetDEC = currentDec
                End If
            Else
                ' current postion is CW down
                If TargetRA > gRAMeridianWest Then
                    ' dual axis slew to meridian followed by ra slew to target
                    gParkParams.SuperSafeMode = 1
                    TargetRA = gRAMeridianWest
                Else
                    If TargetRA < gRAMeridianEast Then
                        ' dual axis slew to meridian followed by ra slew to target
                        gParkParams.SuperSafeMode = 1
                        TargetRA = gRAMeridianEast
                    End If
                End If
            
            End If
        Else
            ' limits are active
            If OutOfBounds(currentRa) = True Then
                ' current position is outside the limits
                If OutOfBounds(TargetRA) = True Then
                    ' target is out of limits
                    ' first slew in RA to the nearest limit
                    ' then slew in RA/DEC to the RA limit nearest the target
                    ' then slew in RA to target
                    If currentRa > RAEncoder_Home_pos Then
                        TargetRA = gRA_Limit_West
                        TargetDEC = currentDec
                    Else
                        TargetRA = gRA_Limit_East
                        TargetDEC = currentDec
                    End If
                    gParkParams.SuperSafeMode = 3
                Else
                    ' target is in limits
                    ' first slew in RA to the nearest limit
                    ' then slew in RA/DEC to the target
                    If currentRa > RAEncoder_Home_pos Then
                        TargetRA = gRA_Limit_West
                        TargetDEC = currentDec
                    Else
                        TargetRA = gRA_Limit_East
                        TargetDEC = currentDec
                    End If
                    gParkParams.SuperSafeMode = 2
                End If
            Else
                If OutOfBounds(TargetRA) = True Then
                    ' target is out of limits
                    ' slew in RA/DEC to limit nearest the target
                    ' then slew in RA to target
                    If TargetRA > RAEncoder_Home_pos Then
                        TargetRA = gRA_Limit_West
                        TargetDEC = TargetDEC
                    Else
                        TargetRA = gRA_Limit_East
                        TargetDEC = TargetDEC
                    End If
                    gParkParams.SuperSafeMode = 1
                Else
                    ' target is in limits
                    ' then slew in RA/DEC to the target
                End If
            
            End If
        End If
    End If
    i = Abs(currentRa - TargetRA)
    j = Abs(currentDec - TargetDEC)
    
    If i <> 0 Then
        If currentRa < TargetRA Then
            gParkParams.RA_Direction = 0
            Select Case gParkParams.Rate
                Case 0
                    ' let mount decide on slew rate
                    gParkParams.RA_SlewActive = 0
                    eqres = EQStartMoveMotor(0, 0, 0, i, GetSlowdown(i))
                Case Else
                    gParkParams.RA_SlewActive = 1
                    eqres = EQ_Slew(0, 0, 0, CLng(gParkParams.Rate))
            End Select
        Else
            gParkParams.RA_Direction = 1
            Select Case gParkParams.Rate
                Case 0
                    ' let mount decide on slew rate
                    gParkParams.RA_SlewActive = 0
                    eqres = EQStartMoveMotor(0, 0, 1, i, GetSlowdown(i))
                Case Else
                    gParkParams.RA_SlewActive = 1
                    eqres = EQ_Slew(0, 0, 1, CLng(gParkParams.Rate))
            End Select
        End If
    End If
    
    If j <> 0 Then
        If currentDec < TargetDEC Then
            gParkParams.DEC_Direction = 0
            Select Case gParkParams.Rate
                Case 0
                    ' let mount decide on slew rate
                    gParkParams.DEC_SlewActive = 0
                    eqres = EQStartMoveMotor(1, 0, 0, j, GetSlowdown(j))
                Case Else
                    gParkParams.DEC_SlewActive = 1
                    eqres = EQ_Slew(1, 0, 0, CLng(gParkParams.Rate))
            End Select
        Else
            gParkParams.DEC_Direction = 1
            Select Case gParkParams.Rate
                Case 0
                    ' let mount decide on slew rate
                    gParkParams.DEC_SlewActive = 0
                    eqres = EQStartMoveMotor(1, 0, 1, j, GetSlowdown(j))
                Case Else
                    gParkParams.DEC_SlewActive = 1
                    eqres = EQ_Slew(1, 0, 1, CLng(gParkParams.Rate))
            End Select
        End If
    End If

End Sub

Public Sub DefinePark(ByVal StopMotors As Boolean)

    If StopMotors Then
        Call StopTrackingUpdates
        
        eqres = EQ_MotorStop(0)          ' Stop RA Motor
        If eqres <> 0 Then
            GoTo ENDDefinePark
        End If
        eqres = EQ_MotorStop(1)          ' Stop DEC Motor
        If eqres <> 0 Then
            GoTo ENDDefinePark
        End If

        'Wait until RA motor is stable
        Do
            eqres = EQ_GetMotorStatus(0)
            If eqres = 1 Then
                GoTo ENDDefinePark
             End If
        Loop While (eqres And EQ_MOTORBUSY) <> 0
    
        'Wait until DEC motor is stable
        Do
            eqres = EQ_GetMotorStatus(1)
            If eqres = 1 Then
                GoTo ENDDefinePark
            End If
        Loop While (eqres And EQ_MOTORBUSY) <> 0
     End If
    
    'Read Motor Values
    gRAEncoderPark = EQGetMotorValues(0)
    gDECEncoderPark = EQGetMotorValues(1)
    Call writepark

ENDDefinePark:

End Sub

Public Sub DefineUserPark(ByVal StopMotors As Boolean, ByVal Index As Integer, name As String)

    If StopMotors Then
        
        Call StopTrackingUpdates
        
        eqres = EQ_MotorStop(0)          ' Stop RA Motor
        If eqres <> 0 Then GoTo ENDDefineUserPark
        eqres = EQ_MotorStop(1)          ' Stop DEC Motor
        If eqres <> 0 Then GoTo ENDDefineUserPark

        'Wait until RA motor is stable
        Do
            eqres = EQ_GetMotorStatus(0)
            If eqres = 1 Then GoTo ENDDefineUserPark
        Loop While (eqres And EQ_MOTORBUSY) <> 0
    
        'Wait until DEC motor is stable
        Do
            eqres = EQ_GetMotorStatus(1)
            If eqres = 1 Then GoTo ENDDefineUserPark
        Loop While (eqres And EQ_MOTORBUSY) <> 0
     End If
    
    'Read Motor Values
    UserParks(Index).posR = EQGetMotorValues(0)
    UserParks(Index).posD = EQGetMotorValues(1)
    UserParks(Index).name = name
    
    Call writeUserParkPos

ENDDefineUserPark:

End Sub

Public Sub DefineUserUnPark(ByVal StopMotors As Boolean, ByVal Index As Integer, name As String)

    If StopMotors Then
        
        Call StopTrackingUpdates
        
        eqres = EQ_MotorStop(0)          ' Stop RA Motor
        If eqres <> 0 Then GoTo ENDDefineUserUnPark
        eqres = EQ_MotorStop(1)          ' Stop DEC Motor
        If eqres <> 0 Then GoTo ENDDefineUserUnPark

        'Wait until RA motor is stable
        Do
            eqres = EQ_GetMotorStatus(0)
            If eqres = 1 Then GoTo ENDDefineUserUnPark
        Loop While (eqres And EQ_MOTORBUSY) <> 0
    
        'Wait until DEC motor is stable
        Do
            eqres = EQ_GetMotorStatus(1)
            If eqres = 1 Then GoTo ENDDefineUserUnPark
        Loop While (eqres And EQ_MOTORBUSY) <> 0
     End If
    
    'Read Motor Values
    UserUnparks(Index).posR = EQGetMotorValues(0)
    UserUnparks(Index).posD = EQGetMotorValues(1)
    UserUnparks(Index).name = name
    
    Call writeUserParkPos

ENDDefineUserUnPark:

End Sub

Public Sub Unparkscope()

    If EQ_GetMountStatus() = 1 Then     ' Make sure that we unpark only if the mount is online
        
        If gEQparkstatus = 1 Then
    
            HC.TrackingFrame.Caption = oLangDll.GetLangString(121) & " " & oLangDll.GetLangString(178)
    
            ' Load Alignment if required
            Call AlignmentStarsUnpark
            
            'Just make sure motors are not moving
            PEC_StopTracking
            eqres = EQ_MotorStop(0)
            eqres = EQ_MotorStop(1)
        
            ' Restore Encoder values
            Call readUnpark
            eqres = EQSetMotorValues(0, gRAEncoderUNPark)
            eqres = EQSetMotorValues(1, gDECEncoderUNPark)
            
            HC.Add_Message (oLangDll.GetLangString(5036))
            
            ' set status as unparked
            gEQparkstatus = 0
            writeParkStatus gEQparkstatus
            
            HC.Frame15.Caption = oLangDll.GetLangString(146) & " " & oLangDll.GetLangString(179)
            Call SetParkCaption
            EQ_Beep (9)
        Else
            HC.Add_Message (oLangDll.GetLangString(5037))
        End If
    
    End If

End Sub

Public Sub UnparkscopeToLastPos()

    Dim i As Long
    Dim j As Long

    
    If EQ_GetMountStatus() = 1 Then     ' Make sure that we unpark only if the mount is online
    
        If gEQparkstatus = 1 Then
    
            ' Load Alignment if required
            Call AlignmentStarsUnpark
            'Unpark Scope first
    
            'Just make sure motors are not moving
            PEC_StopTracking
            eqres = EQ_MotorStop(0)
            eqres = EQ_MotorStop(1)
        
            ' Restore encoder values
            Call readUnpark
            eqres = EQSetMotorValues(0, gRAEncoderUNPark)
            eqres = EQSetMotorValues(1, gDECEncoderUNPark)
    
            'get last position prior to park command
            Call readlastpos
            
            ' set status as unparking
            gEQparkstatus = 3
            
            ' start slewing
            Call StartPark(CDbl(gRAEncoderUNPark), CDbl(gRAEncoderlastpos), CDbl(gDECEncoderUNPark), CDbl(gDECEncoderlastpos))
            HC.ParkTimer.Enabled = True
            HC.Add_Message (oLangDll.GetLangString(5038))
            HC.TrackingFrame.Caption = oLangDll.GetLangString(121) & " " & oLangDll.GetLangString(178)
        Else
            HC.Add_Message (oLangDll.GetLangString(5037))
        End If
    
    End If
 
End Sub


Public Sub UnparkscopeToUserDef(userpos As parkpos)

    Dim i As Long
    Dim j As Long

    ' don't unpark if user position is undefined!
    If userpos.posD = 0 Or userpos.posR = 0 Then GoTo EndUnparkscopeToUserDef:
    
    If EQ_GetMountStatus() = 1 Then     ' Make sure that we unpark only if the mount is online
        
        If gEQparkstatus = 1 Then
            HC.TrackingFrame.Caption = oLangDll.GetLangString(121) & " " & oLangDll.GetLangString(178)
            
            ' Load Alignment if required
            Call AlignmentStarsUnpark
            'Unpark Scope first
    
            readUnpark
        
            'Just make sure motors are not moving
            PEC_StopTracking
            eqres = EQ_MotorStop(0)
            eqres = EQ_MotorStop(1)
        
            ' Restore encoder values
            eqres = EQSetMotorValues(0, gRAEncoderUNPark)
            eqres = EQSetMotorValues(1, gDECEncoderUNPark)
    
            gEQparkstatus = 3
            
            Call StartPark(CDbl(gRAEncoderUNPark), CDbl(userpos.posR), CDbl(gDECEncoderUNPark), CDbl(userpos.posD))
            HC.ParkTimer.Enabled = True
            HC.Add_Message (oLangDll.GetLangString(5038))
        Else
              HC.Add_Message (oLangDll.GetLangString(5037))
        End If
        
    End If
EndUnparkscopeToUserDef:
End Sub

Public Sub readUnpark()

     Dim tmptxt As String

     tmptxt = HC.oPersist.ReadIniValue("UNPARK_RA")
     If tmptxt <> "" Then
        gRAEncoderUNPark = Val(tmptxt)
     Else
        gRAEncoderUNPark = RAEncoder_Home_pos
     End If
     
     tmptxt = HC.oPersist.ReadIniValue("UNPARK_DEC")
     If tmptxt <> "" Then
        gDECEncoderUNPark = Val(tmptxt)
     Else
        gDECEncoderUNPark = gDECEncoder_Home_pos
     End If
 
End Sub

Public Sub readpark()

     Dim tmptxt As String

     tmptxt = HC.oPersist.ReadIniValue("PARK_RA")
     If tmptxt <> "" Then
        gRAEncoderPark = Val(tmptxt)
     Else
        gRAEncoderPark = RAEncoder_Home_pos
     End If
     
     tmptxt = HC.oPersist.ReadIniValue("PARK_DEC")
     If tmptxt <> "" Then
        gDECEncoderPark = Val(tmptxt)
     Else
        gDECEncoderPark = gDECEncoder_Home_pos
     End If
 
End Sub

Public Function readparkStatus() As Long

     Dim tmptxt As String
     tmptxt = HC.oPersist.ReadIniValue("EQPARKSTATUS")
     If tmptxt = "parked" Then
        readparkStatus = 1
     Else
        readparkStatus = 0
     End If

End Function

Public Sub readParkModes()
    Dim tmptxt As String
 
     tmptxt = HC.oPersist.ReadIniValue("DEFAULT_PARK_MODE")
     If tmptxt <> "" Then
        HC.ComboPark.ListIndex = Val(tmptxt)
     Else
        HC.ComboPark.ListIndex = 0
     End If
     
     tmptxt = HC.oPersist.ReadIniValue("DEFAULT_UNPARK_MODE")
     If tmptxt <> "" Then
        HC.ComboUnPark.ListIndex = Val(tmptxt)
     Else
        HC.ComboUnPark.ListIndex = 0
     End If
     
     Call SetParkCaption

End Sub

Public Sub writeParkMode_park()
     HC.oPersist.WriteIniValue "DEFAULT_PARK_MODE", CStr(HC.ComboPark.ListIndex)
End Sub

Public Sub writeParkMode_unpark()
     HC.oPersist.WriteIniValue "DEFAULT_UNPARK_MODE", CStr(HC.ComboUnPark.ListIndex)
End Sub

Public Sub writeParkStatus(ByVal pval As Long)

    If pval = 1 Then
        ' mount is parked
        HC.oPersist.WriteIniValue "EQPARKSTATUS", CStr("parked")
    Else
        ' mount is unparked or parking
        HC.oPersist.WriteIniValue "EQPARKSTATUS", CStr("unparked")
    End If
    
End Sub


Public Sub writeUnpark()

    HC.oPersist.WriteIniValue "UNPARK_RA", CStr(gRAEncoderUNPark)
    HC.oPersist.WriteIniValue "UNPARK_DEC", CStr(gDECEncoderUNPark)

End Sub

Public Sub writepark()

    HC.oPersist.WriteIniValue "PARK_RA", CStr(gRAEncoderPark)
    HC.oPersist.WriteIniValue "PARK_DEC", CStr(gDECEncoderPark)

End Sub


Public Sub readUserParkPos()

    Dim tmptxt As String
    Dim valstr As String
    Dim Ini As String
    Dim key As String
    Dim Count As Integer
     
     
    ' set up a file path for the align.ini file
    Ini = HC.oPersist.GetIniPath & "\EQMOD.ini"

    For Count = 1 To 10
        key = "[userparkposn]"
        With UserParks(Count)
            .name = oLangDll.GetLangString(2730)
            .posR = 0
            .posD = 0
            valstr = "NAME_" & CStr(Count)
            tmptxt = HC.oPersist.ReadIniValueEx(valstr, key, Ini)
            If tmptxt <> "" Then
                .name = tmptxt
            Else
                If Count = 1 Then
                    .name = oLangDll.GetLangString(148)
                End If
                Call HC.oPersist.WriteIniValueEx(valstr, .name, key, Ini)
            End If
            
            valstr = "RCOUNT_" & CStr(Count)
            tmptxt = HC.oPersist.ReadIniValueEx(valstr, key, Ini)
            If tmptxt <> "" Then
                .posR = Val(tmptxt)
            Else
                If Count = 1 Then
                    tmptxt = HC.oPersist.ReadIniValue("PARK_RA")
                    If tmptxt <> "" Then
                        .posR = tmptxt
                    End If
                End If
                Call HC.oPersist.WriteIniValueEx(valstr, CStr(.posR), key, Ini)
            End If
            
            valstr = "DCOUNT_" & CStr(Count)
            tmptxt = HC.oPersist.ReadIniValueEx(valstr, key, Ini)
            If tmptxt <> "" Then
                .posD = Val(tmptxt)
            Else
                If Count = 1 Then
                    tmptxt = HC.oPersist.ReadIniValue("PARK_DEC")
                    If tmptxt <> "" Then
                        .posD = tmptxt
                    End If
                End If
                Call HC.oPersist.WriteIniValueEx(valstr, CStr(.posD), key, Ini)
            End If
        End With
        
       
        key = "[userunparkposn]"
        With UserUnparks(Count)
            .name = oLangDll.GetLangString(2730)
            .posR = 0
            .posD = 0
            valstr = "NAME_" & CStr(Count)
            tmptxt = HC.oPersist.ReadIniValueEx(valstr, key, Ini)
            If tmptxt <> "" Then
                .name = tmptxt
            Else
                If Count = 1 Then
                    .name = oLangDll.GetLangString(2000)
                End If
                Call HC.oPersist.WriteIniValueEx(valstr, .name, key, Ini)
            End If
            
            valstr = "RCOUNT_" & CStr(Count)
            tmptxt = HC.oPersist.ReadIniValueEx(valstr, key, Ini)
            If tmptxt <> "" Then
                .posR = Val(tmptxt)
            Else
                If Count = 1 Then
                    tmptxt = HC.oPersist.ReadIniValue("UNPARK_RA")
                    If tmptxt <> "" Then
                        .posR = tmptxt
                    End If
                End If
                Call HC.oPersist.WriteIniValueEx(valstr, CStr(.posR), key, Ini)
            End If
            
            valstr = "DCOUNT_" & CStr(Count)
            tmptxt = HC.oPersist.ReadIniValueEx(valstr, key, Ini)
            If tmptxt <> "" Then
                .posD = Val(tmptxt)
            Else
                If Count = 1 Then
                    tmptxt = HC.oPersist.ReadIniValue("UNPARK_DEC")
                    If tmptxt <> "" Then
                        .posD = tmptxt
                    End If
                End If
                Call HC.oPersist.WriteIniValueEx(valstr, CStr(.posD), key, Ini)
            End If
        End With
        
    Next Count
     
End Sub

Public Sub writeUserParkPos()

    Dim tmptxt As String
    Dim valstr As String
    Dim Ini As String
    Dim key As String
    Dim Count As Integer
     
    ' set up a file path for the align.ini file
    Ini = HC.oPersist.GetIniPath & "\EQMOD.ini"

    For Count = 1 To 10
        key = "[userparkposn]"
        With UserParks(Count)
            valstr = "NAME_" & CStr(Count)
            Call HC.oPersist.WriteIniValueEx(valstr, .name, key, Ini)
            valstr = "RCOUNT_" & CStr(Count)
            Call HC.oPersist.WriteIniValueEx(valstr, CStr(.posR), key, Ini)
            valstr = "DCOUNT_" & CStr(Count)
            Call HC.oPersist.WriteIniValueEx(valstr, CStr(.posD), key, Ini)
        End With
        
        key = "[userunparkposn]"
        With UserUnparks(Count)
            valstr = "NAME_" & CStr(Count)
            Call HC.oPersist.WriteIniValueEx(valstr, .name, key, Ini)
            valstr = "RCOUNT_" & CStr(Count)
            Call HC.oPersist.WriteIniValueEx(valstr, CStr(.posR), key, Ini)
            valstr = "DCOUNT_" & CStr(Count)
            Call HC.oPersist.WriteIniValueEx(valstr, CStr(.posD), key, Ini)
        End With
        
    Next Count
     
End Sub

Public Sub SetParkCaption()
    If gEQparkstatus Then
        ' parked - use unpark text
        HC.CommandPark.Caption = HC.ComboUnPark.Text
    Else
        ' unparked - use park text
        HC.CommandPark.Caption = HC.ComboPark.Text
    End If
End Sub

' called from park timer
Public Sub ManagePark()
    Dim currentrapos As Double
    Dim currentdecpos As Double
    Dim i As Long
    Dim j As Long
    
    If gParkParams.RA_SlewActive = 1 Or gParkParams.DEC_SlewActive = 1 Then
    
        If gParkParams.RA_SlewActive Then
            If gParkParams.RA_Direction = 0 Then
                If gRA_Encoder >= gParkParams.RA_targetencoder Then
                    eqres = EQ_MotorStop(0)
                    Do
                        eqres = EQ_GetMotorStatus(0)
                        If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo PT1
                    Loop While (eqres And EQ_MOTORBUSY) <> 0
PT1:
                    gParkParams.RA_SlewActive = 0
                End If
            Else
                If gRA_Encoder <= gParkParams.RA_targetencoder Then
                    eqres = EQ_MotorStop(0)
                    Do
                        eqres = EQ_GetMotorStatus(0)
                        If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo PT2
                    Loop While (eqres And EQ_MOTORBUSY) <> 0
PT2:
                    gParkParams.RA_SlewActive = 0
                End If
            End If
        End If
    
        If gParkParams.DEC_SlewActive Then
            If gParkParams.DEC_Direction = 0 Then
                If gDec_Encoder >= gParkParams.DEC_targetencoder Then
                    eqres = EQ_MotorStop(1)
                    Do
                        eqres = EQ_GetMotorStatus(1)
                        If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo PT3
                    Loop While (eqres And EQ_MOTORBUSY) <> 0
PT3:
                    gParkParams.DEC_SlewActive = 0
                End If
            Else
                If gDec_Encoder <= gParkParams.DEC_targetencoder Then
                    eqres = EQ_MotorStop(1)
                    Do
                        eqres = EQ_GetMotorStatus(1)
                        If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo PT4
                    Loop While (eqres And EQ_MOTORBUSY) <> 0
PT4:
                    gParkParams.DEC_SlewActive = 0
                End If
            End If
        End If
    
        If gParkParams.RA_SlewActive = 0 And gParkParams.DEC_SlewActive = 0 Then
            currentrapos = EQGetMotorValues(0)
            currentdecpos = EQGetMotorValues(1)
            
            i = Abs(currentrapos - gParkParams.RA_targetencoder)
            j = Abs(currentdecpos - gParkParams.DEC_targetencoder)
            
            If i <> 0 Then
                If currentrapos < gParkParams.RA_targetencoder Then
                    eqres = EQStartMoveMotor(0, 0, 0, i, GetSlowdown(i))
                Else
                    eqres = EQStartMoveMotor(0, 0, 1, i, GetSlowdown(i))
                End If
            End If
            
            If j <> 0 Then
                If currentdecpos < gParkParams.DEC_targetencoder Then
                    eqres = EQStartMoveMotor(1, 0, 0, j, GetSlowdown(j))
                Else
                    eqres = EQStartMoveMotor(1, 0, 1, j, GetSlowdown(j))
                End If
            End If
        
        End If
    
        Exit Sub

    End If
    
    
    If ((EQ_GetMotorStatus(0) And EQ_MOTORBUSY) = 0) And ((EQ_GetMotorStatus(1) And EQ_MOTORBUSY) = 0) Then
        gEmulOneShot = True ' update ra/dec with real reads fron the mount
        Select Case gParkParams.SuperSafeMode
        
            Case 0
                HC.ParkTimer.Enabled = False
                Select Case gEQparkstatus
                
                    Case 0
                        ' was unparked
                        
                    Case 1
                        ' was parked
                        gEQparkstatus = 0
                        writeParkStatus gEQparkstatus
                        HC.Frame15.Caption = oLangDll.GetLangString(146) & " " & oLangDll.GetLangString(179)
                        Call readRALimit
                        Call SetParkCaption
                        EQ_Beep (9)
                        
                    Case 2
                        ' was parking
                        ' write the park status - just incase EQMOD crashes before normal shutdown!
                        gEQparkstatus = 1
                        writeParkStatus gEQparkstatus
                        Call EQ_Beep(8)
                        HC.Frame15.Caption = oLangDll.GetLangString(146) & " " & oLangDll.GetLangString(177)
                        
                    Case 3
                        ' was unparking
                        gEQparkstatus = 0
                        writeParkStatus gEQparkstatus
                        HC.Frame15.Caption = oLangDll.GetLangString(146) & " " & oLangDll.GetLangString(179)
                        Call readRALimit
                        Call SetParkCaption
                        EQ_Beep (9)
                        
                 End Select
            
            Case 1
                ' Currently at RA home and target DEC: Now move to RA target
                currentrapos = EQGetMotorValues(0)
                currentdecpos = EQGetMotorValues(1)
                Call StartPark(currentrapos, gParkParams.RA_targetencoder, currentdecpos, gParkParams.DEC_targetencoder)
                gParkParams.SuperSafeMode = 0
            
            Case 2
                ' we're at the RA home/limit position
                currentrapos = EQGetMotorValues(0)
                currentdecpos = EQGetMotorValues(1)
                If OutOfBounds(gParkParams.RA_targetencoder) Then
                    ' RA target is outside limits so first slew in dec
                    Call StartPark(currentrapos, RAEncoder_Home_pos, currentdecpos, gParkParams.DEC_targetencoder)
                    gParkParams.SuperSafeMode = 1
                Else
                    ' RA target is within limits so slew both RA and DEC to target
                    Call StartPark(currentrapos, gParkParams.RA_targetencoder, currentdecpos, gParkParams.DEC_targetencoder)
                    gParkParams.SuperSafeMode = 0
                End If
                
            Case 3
                ' we're at the RA Limit position
                currentrapos = EQGetMotorValues(0)
                currentdecpos = EQGetMotorValues(1)
                If gParkParams.RA_targetencoder > RAEncoder_Home_pos Then
                    ' now move to limit nearest to target
                    Call StartPark(currentrapos, gRA_Limit_West, currentdecpos, gParkParams.DEC_targetencoder)
                    gParkParams.SuperSafeMode = 1
                Else
                    ' now move to limit nearest to target
                    Call StartPark(currentrapos, gRA_Limit_East, currentdecpos, gParkParams.DEC_targetencoder)
                    gParkParams.SuperSafeMode = 1
                End If
        
            Case 4
                ' we're at the RA Limit position
                currentrapos = EQGetMotorValues(0)
                currentdecpos = EQGetMotorValues(1)
                If gParkParams.RA_targetencoder > gRAMeridianWest Then
                    ' now move to limit nearest to target
                    Call StartPark(currentrapos, gRAMeridianWest, currentdecpos, gParkParams.DEC_targetencoder)
                    gParkParams.SuperSafeMode = 1
                Else
                    If gParkParams.RA_targetencoder < gRAMeridianEast Then
                        ' now move to limit nearest to target
                        Call StartPark(currentrapos, gRAMeridianEast, currentdecpos, gParkParams.DEC_targetencoder)
                        gParkParams.SuperSafeMode = 1
                    Else
                        Call StartPark(currentrapos, gParkParams.RA_targetencoder, currentdecpos, gParkParams.DEC_targetencoder)
                        gParkParams.SuperSafeMode = 0
                    End If
                End If
        
        
        End Select
    
    End If

End Sub
