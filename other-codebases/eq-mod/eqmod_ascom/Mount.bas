Attribute VB_Name = "Mount"

Option Explicit

Public Type MountDefn
    TotalSteps As String
    wormsteps As String
    offset As String
End Type

Public gCustomMount As Integer
Public gCustomRA360 As Long
Public gCustomDEC360 As Long
Public gCustomRAWormSteps As Double
Public gCustomDECWormSteps As Double
Public gMountType As Long

Public Function CheckMount(openstat As Double)

Dim i As Long
Dim tmp As Double

    CheckMount = openstat
    gTot_step = EQGetTotal360microstep(0)
    gRAMeridianWest = gRAEncoder_Zero_pos + CDbl(gTot_step) / 4
    gRAMeridianEast = gRAEncoder_Zero_pos - CDbl(gTot_step) / 4
    gDECEncoder_Home_pos = EQGetTotal360microstep(1) / 4 + gDECEncoder_Zero_pos    ' totstep/4 + Homepos
    gEQ_MAXSYNC = gTot_step / 8             ' totalstep /8 = 45 degree field
    
    If openstat = 5 Then
    
        If gCustomMount = 1 Then
            ' custom mount ;-)
            CheckMount = EQ_OK
        Else
            Select Case gTot_step
                    
                Case 9024000:  'TYPE1 Mounts
                        gDECEncoder_Home_pos = DECEncoder_Home_pos '&HA26C80
                        gEQ_MAXSYNC = EQ_MAXSYNC_Const             ' 88B80
                        CheckMount = EQ_OK
    
                Case 4505600:   'TYPE2 Mounts
                        gDECEncoder_Home_pos = &H913000            ' Set Home Pos for this version
                        gEQ_MAXSYNC = &H445C0
                        CheckMount = EQ_OK
    
                Case 4576000:   'TYPE3 Mounts
                        gDECEncoder_Home_pos = &H9174C0            ' totstep/4 + Homepos
                        gEQ_MAXSYNC = &H45838                      ' 80640 * arcsec/step
                        CheckMount = EQ_OK
    
                Case Else
                    ' HC.Add_Message (oLangDll.GetLangString(5033))
                    ' CheckMount = EQ_INVALID  ' Exit as Unknown mount
                    CheckMount = EQ_OK
                    
            End Select
        End If
    End If
    
End Function

Public Sub ReadMountType()
Dim tmptxt As String

    ' default to EQG
    gMountType = EQMOUNT
    gCoordType = 0
    
    If gDllVer < 3.04 Then
        ' early version dlls only support Synta EQG
        Call HC.oPersist.WriteIniValue("MOUNT_TYPE", EQMOUNT)
        Call HC.oPersist.WriteIniValue("COORD_TYPE", 0)
        Exit Sub
    End If

    tmptxt = HC.oPersist.ReadIniValue("MOUNT_TYPE")
    If tmptxt = "" Then
        Call HC.oPersist.WriteIniValue("MOUNT_TYPE", EQMOUNT)
    Else
        gMountType = val(tmptxt)
    End If
    
    tmptxt = HC.oPersist.ReadIniValue("COORD_TYPE")
    If tmptxt = "" Then
        Call HC.oPersist.WriteIniValue("COORD_TYPE", 0)
    Else
        gCoordType = val(tmptxt)
    End If
        
    On Error Resume Next
    
    ' inform dll
    Call EQ_SP(0, 10008, gMountType)
    Call EQ_SP(0, 10009, gCoordType)

End Sub

Public Sub WriteMountType()
    Call HC.oPersist.WriteIniValue("MOUNT_TYPE", CStr(gMountType))
    Call HC.oPersist.WriteIniValue("COORD_TYPE", CStr(gCoordType))
End Sub


Public Sub readCustomMount()

     Dim tmptxt As String
     Dim i As Long
     Dim NF1 As Integer

    NF1 = FreeFile
    
    On Error Resume Next
    Close #NF1
    Open HC.oPersist.GetIniPath() + "\mountparams.txt" For Output As #NF1
    For i = 10000 To 10007
        Print #NF1, "0," & CStr(i) & ":" & CStr(EQGP(0, i))
        Print #NF1, "1," & CStr(i) & ":" & CStr(EQGP(1, i))
    Next i
    Close #NF1

     tmptxt = HC.oPersist.ReadIniValue("CUSTOM_MOUNT")
     If tmptxt <> "" Then
        gCustomMount = val(tmptxt)
     Else
        gCustomMount = 0
        HC.oPersist.WriteIniValue "CUSTOM_MOUNT", "0"
     End If

     tmptxt = HC.oPersist.ReadIniValue("CUSTOM_RA_STEPS_360")
     If tmptxt <> "" Then
        gCustomRA360 = val(tmptxt)
     Else
        gCustomRA360 = 9024000
        HC.oPersist.WriteIniValue "CUSTOM_RA_STEPS_360", "9024000"
     End If
    
     tmptxt = HC.oPersist.ReadIniValue("CUSTOM_DEC_STEPS_360")
     If tmptxt <> "" Then
        gCustomDEC360 = val(tmptxt)
     Else
        gCustomDEC360 = 9024000
        HC.oPersist.WriteIniValue "CUSTOM_DEC_STEPS_360", "9024000"
     End If
    
     tmptxt = HC.oPersist.ReadIniValue("CUSTOM_RA_STEPS_WORM")
     If tmptxt <> "" Then
        gCustomRAWormSteps = val(tmptxt)
     Else
        gCustomRAWormSteps = 50133
        HC.oPersist.WriteIniValue "CUSTOM_RA_STEPS_WORM", "50133"
     End If
     
     tmptxt = HC.oPersist.ReadIniValue("CUSTOM_DEC_STEPS_WORM")
     If tmptxt <> "" Then
        gCustomDECWormSteps = val(tmptxt)
     Else
        gCustomDECWormSteps = 50133
        HC.oPersist.WriteIniValue "CUSTOM_DEC_STEPS_WORM", "50133"
     End If
     
     tmptxt = HC.oPersist.ReadIniValue("CUSTOM_TRACKING_OFFSET_RA")
     If tmptxt <> "" Then
        gCustomTrackingOffsetRA = val(tmptxt)
     Else
        gCustomTrackingOffsetRA = 0
        HC.oPersist.WriteIniValue "CUSTOM_TRACKING_OFFSET_RA", "0"
     End If

     tmptxt = HC.oPersist.ReadIniValue("CUSTOM_TRACKING_OFFSET_DEC")
     If tmptxt <> "" Then
        gCustomTrackingOffsetDEC = val(tmptxt)
     Else
        gCustomTrackingOffsetDEC = 0
        HC.oPersist.WriteIniValue "CUSTOM_TRACKING_OFFSET_DEC", "0"
     End If
     
     Call EQSetOffsets

End Sub
Public Sub writeCustomMount()
    HC.oPersist.WriteIniValue "CUSTOM_MOUNT", CStr(gCustomMount)
    HC.oPersist.WriteIniValue "CUSTOM_RA_STEPS_360", CStr(gCustomRA360)
    HC.oPersist.WriteIniValue "CUSTOM_DEC_STEPS_360", CStr(gCustomDEC360)
    HC.oPersist.WriteIniValue "CUSTOM_RA_STEPS_WORM", CStr(gCustomRAWormSteps)
    HC.oPersist.WriteIniValue "CUSTOM_DEC_STEPS_WORM", CStr(gCustomDECWormSteps)
    HC.oPersist.WriteIniValue "CUSTOM_TRACKING_OFFSET_RA", CStr(gCustomTrackingOffsetRA)
    HC.oPersist.WriteIniValue "CUSTOM_TRACKING_OFFSET_DEC", CStr(gCustomTrackingOffsetDEC)
End Sub

Public Sub readWormSteps()

     Dim tmptxt As String

     tmptxt = HC.oPersist.ReadIniValue("RA_STEPS_PER_WORM")
     If tmptxt <> "" Then
        gRAWormSteps = val(tmptxt)
     Else
        gRAWormSteps = 50133
        HC.oPersist.WriteIniValue "RA_STEPS_PER_WORM", CStr(gRAWormSteps)
     End If
     
     tmptxt = HC.oPersist.ReadIniValue("DEC_STEPS_PER_WORM")
     If tmptxt <> "" Then
        gDECWormSteps = val(tmptxt)
     Else
        gDECWormSteps = 50133
        HC.oPersist.WriteIniValue "DEC_STEPS_PER_WORM", CStr(gDECWormSteps)
     End If

End Sub

' Function name    : EQGetTotal360microstep()
' Description      : Get RA/DEC Motor Total 360 degree microstep counts
' Return type      : Double - Stepper Counter Values
'                     0 - 16777215  Valid Count Values
'                     0x1000000 - Mount Not available
'                     0x3000000 - Invalid Parameter
' Argument         : DOUBLE motor_id
'                     00 - RA Motor
'                     01 - DEC Motor
'
Public Function EQGetTotal360microstep(ByVal motor_id As Long) As Long
Dim ret As Long
    
    If gCustomMount = 1 Then
        Select Case motor_id
            Case 0
                ret = gCustomRA360
            Case 1
                ret = gCustomDEC360
            Case Else
                ret = EQ_GetTotal360microstep(motor_id)
        End Select
    Else
        ret = EQ_GetTotal360microstep(motor_id)
    End If
    EQGetTotal360microstep = ret
End Function
