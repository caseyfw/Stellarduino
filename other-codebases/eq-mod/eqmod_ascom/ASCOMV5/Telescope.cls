VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Telescope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ASCOM Scope Driver Template"
'---------------------------------------------------------------------
' Copyright © 2000-2002 SPACE.com Inc., New York, NY
'
' Permission is hereby granted to use this Software for any purpose
' including combining with commercial products, creating derivative
' works, and redistribution of source or binary code, without
' limitation or consideration. Any redistributed copies of this
' Software must include the above Copyright Notice.
'
' THIS SOFTWARE IS PROVIDED "AS IS". SPACE.COM, INC. MAKES NO
' WARRANTIES REGARDING THIS SOFTWARE, EXPRESS OR IMPLIED, AS TO ITS
' SUITABILITY OR FITNESS FOR A PARTICULAR PURPOSE.
'---------------------------------------------------------------------
'   =============
'   TELESCOPE.CLS
'   =============
'
' Implementation of the ASCOM XXX driver Telescope interface
'
'   **********
'   ** NOTE **  CHANGE THE INFO IN THE PROJECT/MAKE TAB!
'   ** NOTE **  CHANGE THE ERR_SOURCE STRING!
'   ** NOTE **  CHANGE THE TELESCOPE OBJECT PROPERTIES IN OBJECT BROWSER!  <=== EVERYONE FORGETS THIS!!
'   ** NOTE **  CHANGE THE ID AND DESC STRINGS BELOW
'   **********
'
' Written:  27-Jun-00   Robert B. Denny <rdenny@dc3.com>
'
' Edits:
'
'
Option Explicit
Implements ITelescope

'Public Enum AlignmentModes
'    algUnknown = -1
'    algAltAz = 0
'    algPolar = 1
'    algGermanPolar = 2
'End Enum

'Public Enum GuideDirections
'    guideNorth = 0
'    guideSouth = 1
'    guideEast = 2
'    guideWest = 3
'End Enum


'Public Enum EquatorialCoordinateType
'    equother = 0
'    equLocalTopocentric = 1
'    equJ2000 = 2
'    equJ2050 = 3
'    equB1950 = 4
'End Enum

'Public Enum DriveRates
'    driveSidereal = 0
'    driveLunar = 1
'    driveSolar = 2
'    driveKing = 3
'End Enum

'Public Enum AxisID
'    AxisPrimary = 0
'    axisSecondary = 1
'    axisTertiary = 2
'End Enum

'Public Enum PierSide
'    pierUnknown = -1
'    PierEast = 0
'    PierWest = 1
'End Enum


'Private Const id As String = "EQMOD.Telescope"
'Private Const DESC As String = "EQMOD ASCOM Scope Driver"



Private m_Serial As DriverHelper.Serial
Private m_Util As DriverHelper.Util
Private m_Profile As DriverHelper.Profile
Private m_iSettleTime As Integer
Private m_RaRateAdjust As Double
Private m_DecRateAdjust As Double
Private LastGuideNS As GuideDirections



Private Sub Class_Initialize()

    If ClientCount < 2 Then
    
        eqres = EQ_End()
    
        Set m_Serial = New DriverHelper.Serial
        Set m_Util = New DriverHelper.Util
        Set m_Profile = New DriverHelper.Profile
        m_Profile.DeviceType = "Telescope"          ' We're a Telescope driver
        m_iSettleTime = 0                           ' Default 0 slew settle time
        m_Profile.Register ASCOM_id, ASCOM_DESC     ' Self-register
        
        
        ' initialise a rates colection for each axis but don't assign any rates yet
        ' this will be done on connection
        Set g_RAAxisRates = New Rates
        Set g_DECAxisRates = New Rates
        
        gHemisphere = 1
        gTargetRA = EQ_INVALIDCOORDINATE
        gTargetDec = EQ_INVALIDCOORDINATE
        gEQTimeDelta = 0
        
        gLoadAPresetOnUnpark = 0
        gSaveAPresetOnPark = 0

        gTot_step = gDefault_step
        gEQ_MAXSYNC = EQ_MAXSYNC_Const
            
        Call GetDllVer
        
        'Initialize position values - They are actually done also on
        'Driver connect
        Call readAscomCompatibiity
        Call readSiderealRate
        Call ReadSiteValues
        Call ReadSyncMap
        Call ReadAlignMap
        Call readUserParkPos
        Call ReadMountType
        Call ReadProcessPriority
    
        
        gEQparkstatus = 0
        gTrackingStatus = 0                        ' Initially not tracking
        gDeclinationRate = 0
        m_DecRateAdjust = 0
        gRightAscensionRate = 0
        m_RaRateAdjust = 0
        
        Call readRALimit
        Call Limits_Init
        
        ' At least with an initial value
        
        gDECEncoder_Home_pos = DECEncoder_Home_pos
        
        
        gRAStatus = EQ_MOTORBUSY                   ' RA Motor Busy Status
        gDECStatus = EQ_MOTORBUSY                  ' DEC Motor Busy Status
        gSlewStatus = False                        ' Not Slewing status
        gRAStatus_slew = False                     ' Slew to track condidition
        gSlewCount = 0                             ' Goto Iterative Counter
        gFRSlewCount = 0                           ' Goto Iterative Counter
        
        
        ' Initialize these values for polling emulation
        
        gEmulRA = 0
        gEmulDEC = 0
        gEmulOneShot = False
        gEmulNudge = False
        
        
        ReadComPortSettings
        
      
        gTot_RA = gTot_step                        ' Set RA Total Encoder Step Count
        gTot_DEC = gTot_step                       ' Set DEC total Encoder Step Count
        gRAGotoRes = 12960000 / gTot_RA  ' 12960000 = secs in 360 degress * 10
        gDECGotoRes = 12960000 / gTot_DEC  ' 12960000 = secs in 360 degress * 10
 '       gGotoResolution = 10                        ' steps

        gRAWormSteps = 50133
        gDECWormSteps = 50133
        gRAWormPeriod = 480
        gSOP = 0                                    'Set at unknown pier setting
    
    
        HC.HCMessage.Text = ""                           'Set Message Center to BLANK
'        HC.HCTextAlign.Text = ""
     
        HC.DisplayTimer.Enabled = True
    End If
    
    ClientCount = ClientCount + 1
    HC.Add_Message (oLangDll.GetLangString(5134) & CStr(ClientCount - 1))
    
End Sub

Private Sub Class_Terminate()
    
    On Error Resume Next
    
    If ClientCount = 2 Then
    
        HC.DisplayTimer.Enabled = False
    
        If EQ_GetMountStatus() = 1 Then     ' We update only if the mount is online
            ' clients really should have disconnected via ASCOM prior to shutdown
            ' but if just in case they don't save the park status
            writeParkStatus gEQparkstatus
        End If
        
        eqres = EQ_End()
        
        writeratebarstateHC
        
        Unload HC
        Unload Align
        Unload AscomTrace
        Unload ColorPick
        Unload CustomMountDlg
        Unload DefineParkForm
        Unload FileDlg
        Unload GotoDialog
        Unload GPSSetup
        Unload JStickConfigForm
        Unload LimitEditForm
        Unload polarfrm
        Unload ProgressFrm
        Unload Setupfrm
        Unload Slewpad
        Unload SoundsFrm
        ' Unload StarSim
        Unload StarEditform
        Unload PECConfigFrm
        ClientCount = 1
    Else
        ClientCount = ClientCount - 1
        If ClientCount > 0 Then
            HC.Add_Message (oLangDll.GetLangString(5134) & CStr(ClientCount - 1))
        End If
    End If
    
End Sub

'------------------ V2 Compliance Properties

Public Property Get AlignmentMode() As AlignmentModes

    AlignmentMode = algGermanPolar
    
End Property

Public Property Get AtHome() As Boolean
    AtHome = False
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, "GET AtHome :F"
End Property

Public Property Get AtPark() As Boolean

    Select Case gEQparkstatus
        Case 0
            ' if unparked
            AtPark = False
        Case 1
            ' if parked
            AtPark = True
        Case 2
            ' ASCOM has no means of detecting a parking state
            ' However some folks will be closing their roofs based upon the stare of AtPark
            ' So we must respond with a false!
            AtPark = False
    End Select
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, "GET AtPark :" & AtPark
    
End Property

Public Property Get CanSetDeclinationRate() As Boolean
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 3, "GET CanSetDeclination :T"
    CanSetDeclinationRate = True
End Property

Public Property Get CanSetGuideRates() As Boolean
'    If gAscomCompatibility.AllowPulseGuide Then
        If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 3, "GET CanSetGuideRates :T"
        CanSetGuideRates = True
'    Else
'        If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 3, "GET CanSetGuideRates :F"
'        CanSetGuideRates = False
'    End If
End Property

Public Property Get CanSetPierSide() As Boolean
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 3, "GET CanSetPierSide :F"
    CanSetPierSide = False
End Property

Public Property Get CanSetRightAscensionRate() As Boolean
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 3, "GET CanSetRARate :F"
    CanSetRightAscensionRate = True
End Property

Public Property Get CanSlewAltAz() As Boolean
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 3, "GET CanSlewAltAz :F"
    CanSlewAltAz = False
End Property

Public Property Get CanSlewAltAzAsync() As Boolean
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 3, "GET CanSlewAltAzAsync :F"
    CanSlewAltAzAsync = False
End Property

Public Property Get CanSyncAltAz() As Boolean
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 3, "GET CanSyncAltAz :F"
    CanSyncAltAz = False
End Property

Public Property Get DriverVersion() As String
    DriverVersion = App.Major & "." & App.Minor
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, "GET DriverVersion :" & DriverVersion
End Property

Public Property Get EquatorialSystem() As EquatorialCoordinateType
    Select Case gAscomCompatibility.Epoch
        Case 0
            EquatorialSystem = equOther
        Case 1
            EquatorialSystem = equLocalTopocentric
        Case 2
            EquatorialSystem = equJ2000
        Case 3
            EquatorialSystem = equJ2050
        Case 4
            EquatorialSystem = equB1950
        Case Else
            EquatorialSystem = equOther
    End Select
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, "GET EquatorialCoordinateType :" & EquatorialSystem
End Property

Public Property Get GuideRateDeclination() As Double
    If gAscomCompatibility.AllowPulseGuide Then
        ' movement rate offset in degress/sec
        GuideRateDeclination = (HC.HScrollDecRate.Value * 0.1 * SID_RATE) / 3600
        If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 4, "GET GuideRateDEC :" & CStr(GuideRateDeclination)
    Else
        'RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Property Get GuideRateDeclination" & MSG_NOT_IMPLEMENTED
        Select Case HC.DECGuideRateList.ListIndex
            Case 1
                GuideRateDeclination = (0.5 * SID_RATE) / 3600
            Case 2
                GuideRateDeclination = (0.75 * SID_RATE) / 3600
            Case 3
                GuideRateDeclination = (SID_RATE) / 3600
            Case 4
                RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Property Get GuideRateDeclination" & MSG_NOT_IMPLEMENTED
            Case Else
                GuideRateDeclination = (0.25 * SID_RATE) / 3600
        End Select
        If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 4, "GET GuideRateDEC :" & CStr(GuideRateDeclination)
    End If
    
End Property

Public Property Let GuideRateDeclination(ByVal newval As Double)
    If gAscomCompatibility.AllowPulseGuide Then
        ' convert from degrees per sec to a factor of sidereal
        If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 4, "LET GuideRateDEC(" & newval & ")"
        newval = newval * 3600 / (0.1 * SID_RATE)
        If newval < HC.HScrollDecRate.min Then
            newval = HC.HScrollDecRate.min
        Else
            If newval > HC.HScrollDecRate.max Then
                newval = HC.HScrollDecRate.max
            End If
        End If
        HC.HScrollDecRate.Value = CInt(newval)
    Else
        If HC.DECGuideRateList.ListIndex = 2 Then
            If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 4, "LET GuideRateDEC(" & newval & ") :NOT_SUPPORTED"
            RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Property Let GuideRateDeclination" & MSG_NOT_IMPLEMENTED
        Else
            newval = newval * 3600 / SID_RATE
            If newval > 0.75 Then
                HC.DECGuideRateList.ListIndex = 3
            Else
                If newval > 0.5 Then
                    HC.DECGuideRateList.ListIndex = 2
                Else
                    If newval > 0.25 Then
                        HC.DECGuideRateList.ListIndex = 1
                    Else
                        HC.DECGuideRateList.ListIndex = 0
                    End If
                End If
            End If
            If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 4, "LET GuideRateDEC(" & newval & ")"
        End If
    End If

End Property

Public Property Get GuideRateRightAscension() As Double
    If gAscomCompatibility.AllowPulseGuide Then
        ' movement rate offset in degrees/sec
        GuideRateRightAscension = (HC.HScrollRARate.Value * 0.1 * SID_RATE) / 3600
        If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 4, "GET GuideRateRA :" & CStr(GuideRateRightAscension)
    Else
        ' RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Property Get GuideRateRightAscension" & MSG_NOT_IMPLEMENTED
        Select Case HC.RAGuideRateList.ListIndex
            Case 1
                GuideRateRightAscension = (0.5 * SID_RATE) / 3600
            Case 2
                GuideRateRightAscension = (0.75 * SID_RATE) / 3600
            Case 3
                GuideRateRightAscension = (SID_RATE) / 3600
            Case 4
                RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Property Get GuideRateRightAscension" & MSG_NOT_IMPLEMENTED
            Case Else
                GuideRateRightAscension = (0.25 * SID_RATE) / 3600
        End Select
        If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 4, "GET GuideRateRA :" & CStr(GuideRateRightAscension)
    End If
    
End Property

Public Property Let GuideRateRightAscension(ByVal newval As Double)
    ' We can't support properly beacuse the ASCOM spec does not distinquish between ST4 and Pulseguiding
    ' and states that this property relates to both - crazy!
    If gAscomCompatibility.AllowPulseGuide Then
        If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 4, "LET GuideRateRA(" & newval & ")"
        newval = newval * 3600 / (0.1 * SID_RATE)
        If newval < HC.HScrollRARate.min Then
            newval = HC.HScrollRARate.min
        Else
            If newval > HC.HScrollRARate.max Then
                newval = HC.HScrollRARate.max
            End If
        End If
        HC.HScrollRARate.Value = CInt(newval)
    Else
        If HC.RAGuideRateList.ListIndex = 4 Then
            If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 4, "LET GuideRateRA(" & newval & ") :NOT_SUPPORTED"
            RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Property Let GuideRateRightAscension" & MSG_NOT_IMPLEMENTED
        Else
            newval = newval * 3600 / SID_RATE
            If newval > 0.75 Then
                HC.RAGuideRateList.ListIndex = 3
            Else
                If newval > 0.5 Then
                    HC.RAGuideRateList.ListIndex = 2
                Else
                    If newval > 0.25 Then
                        HC.RAGuideRateList.ListIndex = 1
                    Else
                        HC.RAGuideRateList.ListIndex = 0
                    End If
                End If
            End If
            If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 4, "LET GuideRateRA(" & newval & ")"
        End If
    End If
End Property

Public Property Get InterfaceVersion() As Integer
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, "GET InterfaceVersion=2"
    InterfaceVersion = 2
End Property

Public Property Let SideOfPier(ByVal newval As PierSide)
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, "LET SideOfPier(" & newval & ") :NOT_SUPPORTED"
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Property Let SideOfPier" & MSG_NOT_IMPLEMENTED
End Property
Public Property Get SideOfPier() As PierSide

    ' Pointing Side of Pier
    SideOfPier = SOP_Pointing(gDec_DegNoAdjust)
    
    ' Physical Side of Pier
'    SideOfPier = SOP_Physical(gRA_Hours)
    
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, "GET SideOfPier :" & SideOfPier
End Property

Public Property Let TrackingRate(ByVal newval As DriveRates)
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 5, "LET TrackingRate :" & newval & ")"
    If newval <> driveSidereal Then
        RaiseError SCODE_PROP_RANGE_ERROR, ERR_SOURCE, "Property Let TrackingRate" & MSG_PROP_RANGE_ERROR
    End If
End Property

Public Property Get TrackingRate() As DriveRates
    TrackingRate = driveSidereal
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 5, ("GET TrackingRate :" & TrackingRate)
End Property

Public Property Let TrackingRates(ByVal newval As Object)
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 5, "LET TrackingRates :NOT_SUPPORTED"
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Property Let TrackingRates" & MSG_NOT_IMPLEMENTED
End Property

Public Property Get TrackingRates() As Object
    Set TrackingRates = g_TrackingRates
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 5, ("GET Tracking Rates")
End Property

Public Function AxisRates(axis As TelescopeAxes) As Object
    
    Select Case axis
        Case axisPrimary
            Set AxisRates = g_RAAxisRates
        Case axisSecondary
            Set AxisRates = g_DECAxisRates
        Case Else
            Set AxisRates = New Rates
    End Select
    
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 6, ("AxisRates(" & axis & ")")
End Function

Public Function CanMoveAxis(axis As TelescopeAxes) As Boolean
    
    Select Case axis
    
        Case axisPrimary, axisSecondary
            If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 3, ("CanMoveAxis(" & axis & ") :True")
            CanMoveAxis = True
        
        Case Else
            If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 3, ("CanMoveAxis(" & axis & ") :False")
            CanMoveAxis = False
    
    End Select
    
End Function

Public Function DestinationSideOfPier(ByVal destRa As Double, destDec As Double) As Long
    Dim ha As Double
    
    ha = EQnow_lst(gLongitude * DEG_RAD) - destRa
    range ha, 24#
    
    If ha < 12# Then
        DestinationSideOfPier = pierEast 'looking West
    Else
        DestinationSideOfPier = pierWest 'looking East
    End If
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("DestinationSideOfPier(" & destRa & "," & destDec & ")")
End Function

Public Sub MoveAxis(axis As TelescopeAxes, rate As Double)
    
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 6, ("COMMAND MoveAxis(" & axis & "," & rate & ")")
    
    If gEQparkstatus <> 0 Then
        ' no move axis if parked or parking!
        RaiseError SCODE_SCOPE_PARKED, ERR_SOURCE, "Method MoveAxis() " & MSG_SCOPE_PARKED
        Exit Sub
    End If

    Select Case axis
        Case axisPrimary
            If RateIsInRange(rate, g_RAAxisRates) Then
                If Not gSlewStatus Then
                    Call EQMoveAxis(0, rate)
                End If
            Else
                RaiseError SCODE_VAL_OUTOFRANGE, ERR_SOURCE, "Method MoveAxis() " & MSG_VAL_OUTOFRANGE
            End If
        Case axisSecondary
            If RateIsInRange(rate, g_DECAxisRates) Then
                If Not gSlewStatus Then
                    Call EQMoveAxis(1, rate)
                End If
            Else
                RaiseError SCODE_VAL_OUTOFRANGE, ERR_SOURCE, "Method MoveAxis() " & MSG_VAL_OUTOFRANGE
            End If
        Case Else
            RaiseError SCODE_VAL_OUTOFRANGE, ERR_SOURCE, "Method MoveAxis() " & MSG_VAL_OUTOFRANGE
    End Select
        
End Sub

Public Sub SlewToAltAz(ByVal Azimuth As Double, ByVal altitude As Double)
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 6, ("SlewToAltAz(" & Azimuth & "," & altitude & ")")
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Method SlewToAltAz()" & MSG_NOT_IMPLEMENTED
End Sub

Public Sub SlewToAltAzAsync(ByVal Azimuth As Double, ByVal altitude As Double)
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 6, ("SlewToAltAzAsync(" & Azimuth & "," & altitude & ")")
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Method SlewToAltAzAsync()" & MSG_NOT_IMPLEMENTED
End Sub
Public Sub SyncToAltAz(ByVal Azimuth As Double, ByVal altitude As Double)
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 7, ("SyncToAltAz(" & Azimuth & "," & altitude & ")")
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Method SyncToAltAz()" & MSG_NOT_IMPLEMENTED
End Sub


'-------------------------- V1 Properties Starts here -----------------

Public Property Get DoesRefraction() As Boolean
    DoesRefraction = False
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("GET DoesRfraction :F")
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Property Get DoesRefraction" & MSG_NOT_IMPLEMENTED
End Property

Public Property Let DoesRefraction(ByVal newval As Boolean)
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("SET DoesRfraction :NOT SUPPORTED")
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Property Get DoesRefraction" & MSG_NOT_IMPLEMENTED
End Property

Public Property Get altitude() As Double
    altitude = gAlt
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 2, ("GET altitude :" & gAlt)
End Property

Public Property Get ApertureDiameter() As Double
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("GET ApertureDiameter :NOT_SUPPORTED")
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Property ApertureDiameter" & MSG_NOT_IMPLEMENTED
End Property

Public Property Get ApertureArea() As Double
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("GET ApertureArea :NOT_SUPPORTED")
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Property ApertureDiameter" & MSG_NOT_IMPLEMENTED
End Property

Public Property Get Azimuth() As Double
    Azimuth = gAz
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 2, ("GET Azimuth :" & gAz)
End Property

Public Property Get CanFindHome() As Boolean
    CanFindHome = False
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 3, ("GET CanFindHome :F")
End Property

Public Property Get CanPark() As Boolean
    CanPark = True
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 3, ("GET CanPark :T")
End Property

Public Property Get CanSetPark() As Boolean
    CanSetPark = True
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 3, ("GET CanSetPark :T")
End Property

Public Property Get CanSetTracking() As Boolean
    CanSetTracking = True
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 3, ("GET CanSetTracking :T")
End Property

Public Property Get CanSlew() As Boolean
    CanSlew = True
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 3, ("GET CanSlew :T")
End Property

Public Property Get CanSlewAsync() As Boolean
    CanSlewAsync = True
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 3, ("GET CanSlewAsync :T")
End Property

Public Property Get CanSync() As Boolean
    CanSync = True
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 3, ("GET CanSync :T")
End Property

Public Property Get CanUnpark() As Boolean
    CanUnpark = True
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 3, ("GET CanUnpark :T")
End Property
Public Property Get CanPulseGuide() As Boolean
    If gAscomCompatibility.AllowPulseGuide Then
        CanPulseGuide = True
        If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 3, ("GET CanPulseGuide :T")
    Else
        CanPulseGuide = False
        If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 3, ("GET CanPulseGuide :F")
    End If
End Property

Public Property Get Connected() As Boolean
  If EQ_GetMountStatus() <> 1 Then
      Connected = False
  Else
    If HC.EncoderReadErrCount >= 5 Then
        ' comms has bit problems so report as disconnected
        Connected = False
    Else
        Connected = True
    End If
  End If
  If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 1, ("GET Connected :" & Connected)
End Property

Public Property Let Connected(ByVal newval As Boolean)
  
Dim icast As Long
Dim MaxRate As Double
Dim strtmp As String
Dim i As Integer

    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 1, ("LET Connected(" & newval & ")")
    If newval Then
    
        gpl_interval = 50
        
        Call readRASyncCheckVal ' RA Sync Auto
        Call readPulseguidepwidth ' Read Pulseguide interval
    
'        Call ReadSiteValues
'        Call ReadSyncMap
'        Call ReadAlignMap
    
        If EQ_GetMountStatus() = 0 Then
            ' if mount hasn't been connected yet then do stuff!
        
            gMount_Ver = EQ_GetMountVersion()
            Call readCustomMount
            Call readDriftVal ' Read the Drift offset value
            
            gRA_LastRate = 0
            gCurrent_time = 0
            gLast_time = 0
            gEmulRA_Init = 0
            
            gAlignmentStars_count = 0
        
            gEQRAPulseDuration = 0
            gEQDECPulseDuration = 0
            gEQPulsetimerflag = True
            gThreeStarEnable = False
            gSelectStar = 0
            gRA_GOTO = 0
            gDEC_GOTO = 0
            
            gRAMoveAxis_Rate = 0
            gDECMoveAxis_Rate = 0
            
            HC.Show
            
            ' prefix the com port to ensure com10+ works
            strtmp = "\\.\" & gPort
            
            gInitResult = EQ_Init(strtmp, gBaud, gTimeout, gRetry)
            
            ' Make sure we have the right board
            eqres = CheckMount(gInitResult)
        
            If eqres = EQ_OK And EQ_GetMountStatus() = 1 Then
                Call readRALimit
                Call readCustomMount
                eqres = EQGetTotal360microstep(0)
                If eqres < &H1000000 Then
                    gTot_RA = eqres
                    gRAGotoRes = gGotoResolution * 1296000 / gTot_RA  ' 1296000 = seconds per 360 degrees
                    gRAWormSteps = EQGP(0, 10006)
                    HC.Add_Message CStr(gRAWormSteps) & " RAWormSteps read"
                    If gCustomMount = 0 Then
                        Select Case gRAWormSteps
                            Case 50133
                            Case 66844
                            Case 35200
                            Case 31288
                            Case Else
'                                ' mount isn't returning known values - read from ini instead
'                                Call readWormSteps
                        End Select
                    End If
                    gRAWormPeriod = Int((SPSD * gRAWormSteps / gTot_RA) + 0.5)
                End If
                
                eqres = EQGetTotal360microstep(1)
                If eqres < &H1000000 Then
                    gTot_DEC = eqres
                    gDECGotoRes = gGotoResolution * 1296000 / gTot_DEC  ' 1296000 = seconds per 360 degrees
                    gDECWormSteps = EQGP(1, 10006)
                    HC.Add_Message CStr(gDECWormSteps) & " DECWormSteps read"
                    If gCustomMount = 0 Then
                        Select Case gDECWormSteps
                            Case 50133
                            Case 66844
                            Case 35200
                            Case 31288
                            Case Else
'                                ' mount isn't returning known values - read from ini instead
'                                Call readWormSteps
                        End Select
                    End If
                End If
                
                If gDllVer < 3.04 Then
                    gTrackFactorRA = CDbl(EQGP(0, 10004)) * SID_RATE
                    gTrackFactorDEC = CDbl(EQGP(0, 10005)) * SID_RATE
                Else
                    gTrackFactorRA = CDbl(EQGP(0, 10004)) * SID_RATE
                    gTrackFactorDEC = CDbl(EQGP(1, 10004)) * SID_RATE
                End If
                
'                i = CDbl(EQGP(0, 10002)) / (CDbl(EQGP(0, 10001)) / 86164)
'                gTrackFactorRA = i * SID_RATE
'                i = CDbl(EQGP(1, 10002)) / (CDbl(EQGP(1, 10001)) / 86164)
'                gTrackFactorDEC = i * SID_RATE
                
                'Make sure motors are not running
                eqres = EQ_MotorStop(0)     ' Stop RA Motor
                eqres = EQ_MotorStop(1)     ' Stop DEC Motor
                
                HC.TrackingFrame.Caption = oLangDll.GetLangString(121) & " " & oLangDll.GetLangString(178)
                
                'Get state of at least one of the motors
                
                eqres = EQ_GetMotorStatus(0)
                
                ' If its an error then Initialize it
                
                If eqres = EQ_NOTINITIALIZED Then
                    icast = gDECEncoder_Home_pos ' Typecast
                    eqres = EQ_InitMotors(RAEncoder_Home_pos, icast)
                End If
                
                
                ' set up rates collection
                MaxRate = SID_RATE * 800 / 3600
                g_RAAxisRates.Add MaxRate, 0#
                g_DECAxisRates.Add MaxRate, 0#

                            
                'Make sure we get the latest data from the registry
               
               
                HC.Add_Message (oLangDll.GetLangString(5132) & " " & gPort & ":" & str(gBaud))
                HC.Add_Message (oLangDll.GetLangString(5133) & " " & printhex(EQ_GetMountVersion()) & " DLL Version:" & printhex(EQ_DriverVersion()))
                HC.Add_Message "Using " & CStr(gRAWormSteps) & "RAWormSteps"
                HC.EncoderTimer.Enabled = True
                HC.EncoderTimerFlag = True
                gEQPulsetimerflag = True
                HC.Pulseguide_Timer.Enabled = False     'Enabled only during pulseguide session
                
                Call readParkModes
                Call readAlignProximity

                gEQparkstatus = readparkStatus()
    
                If gEQparkstatus = 1 Then
                    ' currently parked
                    HC.Frame15.Caption = oLangDll.GetLangString(146) & " " & oLangDll.GetLangString(177)
                    ' Read Park position
                    Call readUnpark
                    ' Preset the Encoder values to Park position
                    eqres = EQSetMotorValues(0, gRAEncoderUNPark)
                    eqres = EQSetMotorValues(1, gDECEncoderUNPark)
                Else
                    HC.Frame15.Caption = oLangDll.GetLangString(146) & " " & oLangDll.GetLangString(179)
                End If
                Call SetParkCaption

                Call readportrate ' Read Autoguider port settings from registry and send to mount
                Call PEC_Initialise   ' only initialise PEc when we've defaults for worm

            Else
                HC.EncoderTimer.Enabled = False
                HC.EncoderTimerFlag = False
                eqres = EQ_End()
                HC.Add_Message (oLangDll.GetLangString(5135) & " " & gPort & ":" & str(gBaud))
            
            End If
        End If
    Else
        If ClientCount <= 2 Then
    
           If EQ_GetMountStatus() = 1 Then     ' We update only if the mount is online
               writeParkStatus gEQparkstatus
           End If
       
           HC.EncoderTimer.Enabled = False
           HC.EncoderTimerFlag = False
           
           'Save alignment and Sync data if scope is parked
           'Otherwise an re-alignment / re-Sync process has to be made on restart
           
    '       If gEQparkstatus = 1 Then
    '           Call WriteSyncMap
    '           Call WriteAlignMap
    '       End If
       
           eqres = EQ_End()
           
           
        End If
       
    End If

End Property

Public Property Get Declination() As Double
    Declination = gDec
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 2, ("GET Declination :" & gDec)
End Property

Public Property Get DeclinationRate() As Double
    DeclinationRate = gDeclinationRate
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 5, ("GET DeclinationRate :" & gDeclinationRate)
End Property

Public Property Let DeclinationRate(ByVal newval As Double)

   If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 5, ("SET DeclinationRate(" & newval & ")")
    m_DecRateAdjust = newval
   ' don't action this if we're parked!
    If gEQparkstatus = 0 Then
    
        ' if we're already tracking then apply the new rate.
        If gTrackingStatus <> 0 Then
            If (gDeclinationRate * newval) <= 0 Then
                Call StartDEC_by_Rate(newval)
            Else
                Call ChangeDEC_by_Rate(newval)
            End If
            
            gTrackingStatus = 4
            ' Custom tracking!
            HC.TrackingFrame.Caption = oLangDll.GetLangString(121) & " " & oLangDll.GetLangString(189)
        End If
        gDeclinationRate = newval

    End If
End Property

Public Property Get Description() As String
    Description = "EQMOD ASCOM Driver"
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("GET Destription :" & Description)
End Property

Public Property Get DriverInfo() As String
    '
    ' Use the Project/Properties sheet, Make tab, to set these
    ' items. That way they will show in the Version tab of the
    ' Explorer property sheet, and the exact same data will
    ' show in Telescope.DriverInfo.
    '
    DriverInfo = "EQASCOM " & gVersion & vbCrLf
    If App.CompanyName <> "" Then _
        DriverInfo = DriverInfo & vbCrLf & App.CompanyName
    If App.LegalCopyright <> "" Then _
        DriverInfo = DriverInfo & vbCrLf & App.LegalCopyright
    If App.Comments <> "" Then _
        DriverInfo = DriverInfo & vbCrLf & App.Comments
    
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("GET DriverInfo :" & DriverInfo)
        
End Property

Public Property Get FocalLength() As Double
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("GET FocalLength :NOT_SUPPORTED")
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Property FocalLength" & MSG_NOT_IMPLEMENTED
End Property

Public Property Get name() As String
    name = ASCOM_DESC                        ' 1-word name
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("GET Name :" & name)
End Property

Public Property Get RightAscension() As Double
    RightAscension = GetEmulRA_EQ()
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 2, ("GET RightAscension :" & RightAscension)
End Property

Public Property Get RightAscensionRate() As Double
    RightAscensionRate = gRightAscensionRate - SID_RATE
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 5, ("GET RightAscensionRate :" & RightAscensionRate)
End Property

Public Property Let RightAscensionRate(ByVal newval As Double)

    'newval is in arcseconds , convert to degrees
  If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 5, ("SET RightAscensionRate(" & newval & ")")
 
    ' don't action this if we're parked!
    m_RaRateAdjust = newval
    If gEQparkstatus = 0 Then
        newval = SID_RATE + newval      ' Treat newval as an offset
        ' if we're already tracking then apply the new rate.
        If gTrackingStatus <> 0 Then
            If (gRightAscensionRate * newval) <= 0 Then
                Call StartRA_by_Rate(newval)
            Else
                Call ChangeRA_by_Rate(newval)
            End If
            gTrackingStatus = 4
            ' Custom tracking!
            HC.TrackingFrame.Caption = oLangDll.GetLangString(121) & " " & oLangDll.GetLangString(189)
        End If
        gRightAscensionRate = newval
    End If
End Property

Public Property Get SiderealTime() As Double
    SiderealTime = EQnow_lst(gLongitude * DEG_RAD)
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("GET SiderealTime :" & SiderealTime)
End Property

Public Property Get SiteElevation() As Double
    SiteElevation = gElevation
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("GET SiteElevation :" & SiteElevation)
End Property

Public Property Let SiteElevation(ByVal newval As Double)
    If gAscomCompatibility.AllowSiteWrites Then
        If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("SET SiteElevation :" & CStr(newval))
        gElevation = newval
        Call UpdateSiteControls
    Else
        If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("SET SiteElevation :NOT_SUPPORTED")
        RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Property SiteElevation" & MSG_NOT_IMPLEMENTED
    End If
End Property

Public Property Get SiteLatitude() As Double
    SiteLatitude = gLatitude
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("GET SiteLatitude :" & SiteLatitude)
End Property

Public Property Let SiteLatitude(ByVal newval As Double)
    If gAscomCompatibility.AllowSiteWrites Then
        If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("SET SiteLatitude :" & CStr(newval))
        gLatitude = newval
        Call UpdateSiteControls
    Else
        If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("SET SiteLatitude :NOT_SUPPORTED")
        RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Property SiteLatitude" & MSG_NOT_IMPLEMENTED
    End If
End Property

Public Property Get SiteLongitude() As Double
    SiteLongitude = gLongitude
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("GET SiteLongitude :" & SiteLongitude)
End Property

Public Property Let SiteLongitude(ByVal newval As Double)
    If gAscomCompatibility.AllowSiteWrites Then
        If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("SET SiteLongitude :" & CStr(newval))
        gLongitude = newval
        Call UpdateSiteControls
    Else
        If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("SET SiteLongitude :NOT_SUPPORTED")
        RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Property SiteLongitude" & MSG_NOT_IMPLEMENTED
    End If
End Property

Public Property Get Slewing() As Boolean
    Select Case gEQparkstatus
        Case 0
            ' unparked
            Slewing = gSlewStatus
        Case 1
            ' parked
            Slewing = False
        Case 2
            ' parking
            Slewing = True
    End Select
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 6, ("GET Slewing :" & Slewing)
End Property

Public Property Get SlewSettleTime() As Integer
    SlewSettleTime = m_iSettleTime
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 6, ("GET SlewSettleTime :" & SlewSettleTime)
End Property

Public Property Let SlewSettleTime(ByVal newval As Integer)
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 6, ("LET SlewSettleTime(" & newval & ")")
    If (newval < 0) Or (newval > 100) Then
        RaiseError SCODE_VAL_OUTOFRANGE, ERR_SOURCE, MSG_VAL_OUTOFRANGE
    End If
    m_iSettleTime = newval
End Property

Public Property Get TargetDeclination() As Double
    If gTargetDec = EQ_INVALIDCOORDINATE Then
        RaiseError SCODE_PROP_NOT_SET, ERR_SOURCE, "Property TargetDeclination " & MSG_PROP_NOT_SET
    End If
    TargetDeclination = gTargetDec
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 6, ("GET TargetDeclination :" & TargetDeclination)
End Property

Public Property Let TargetDeclination(ByVal newval As Double)
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 6, ("Let TargetDeclination(" & newval & ")")
    If newval > 90 Or newval < -90 Then
        RaiseError SCODE_VAL_OUTOFRANGE, ERR_SOURCE, "Property TargetDeclination " & MSG_VAL_OUTOFRANGE
    Else
        gTargetDec = newval
    End If
End Property

Public Property Get TargetRightAscension() As Double
    If gTargetRA = EQ_INVALIDCOORDINATE Then
        RaiseError SCODE_PROP_NOT_SET, ERR_SOURCE, "Property TargetRightAscension " & MSG_PROP_NOT_SET
    End If
    TargetRightAscension = gTargetRA
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 6, ("GET TargetRightAscension :" & TargetRightAscension)
End Property

Public Property Let TargetRightAscension(ByVal newval As Double)
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 6, ("Let TargetRightAscension(" & newval & ")")
    If newval > 24 Or newval < 0 Then
        RaiseError SCODE_VAL_OUTOFRANGE, ERR_SOURCE, "Property TargetRightAscension " & MSG_VAL_OUTOFRANGE
    Else
        gTargetRA = newval
    End If
End Property

Public Property Get Tracking() As Boolean
    If gTrackingStatus <> 0 Then
        Tracking = True
    Else
        Tracking = False
    End If
   If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 5, ("GET Tracking :" & Tracking)
End Property

Public Property Let Tracking(ByVal newval As Boolean)
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 5, ("LET Tracking(" & newval & ")")
    If gEQparkstatus = 0 Then
        If newval Then
        
            If m_RaRateAdjust = 0 And m_DecRateAdjust = 0 Then
               ' track at sidereal
                Call EQStartSidereal2
                gEmulOneShot = True                 ' Get One shot cap
            Else
                ' track at custom rate
                gRA_LastRate = 0
                gDeclinationRate = m_DecRateAdjust
                gRightAscensionRate = SID_RATE + m_RaRateAdjust
                If gPEC_Enabled Then PEC_StopTracking
                Call CustomMoveAxis(0, gRightAscensionRate, True, oLangDll.GetLangString(189))
                Call CustomMoveAxis(1, gDeclinationRate, True, oLangDll.GetLangString(189))
            End If
        
        Else
        
            eqres = EQ_MotorStop(0)
            eqres = EQ_MotorStop(1)
            EQ_Beep (7)
            gTrackingStatus = 0
            ' not sure that we should be clearing the rate offests ASCOM Spec is no help
            gDeclinationRate = 0
            gRightAscensionRate = 0
            
            HC.TrackingFrame.Caption = oLangDll.GetLangString(121) & " " & oLangDll.GetLangString(178)
            
        End If
    Else
    
        HC.Add_Message (oLangDll.GetLangString(5013))
    
    End If

End Property

Public Property Get UTCDate() As Date
    UTCDate = CDate(CDbl(Now()) + (CDbl(utc_offs()) / 86400#))
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("GET Date :" & UTCDate)
End Property

Public Property Let UTCDate(ByVal newval As Date)
     'Impossible to set own PC time
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("SET Date :NOT_SUPPORTED")
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Property UTCDate" & MSG_NOT_IMPLEMENTED
End Property

'----------------------------------------------
'-------Some Extra COM interfaces -------------
'----------------------------------------------

Public Sub PecMoveAxis(axis As TelescopeAxes, rate As Double)
    ' special sub for PEC control - normal moveaxis results in background polling of motor positions
    If Not gSlewStatus Then
        If axis = axisPrimary Then
            Call PEC_MoveAxis(0, rate * 3600)
        End If
        If axis = axisSecondary Then
            Call PEC_MoveAxis(1, rate * 3600)
        End If
    End If
End Sub

Public Function MoveMotor(motor_id As Long, hemisphere As Long, direction As Long, Steps As Long, stepslowdown As Long) As Long
    MoveMotor = EQStartMoveMotor(motor_id, hemisphere, direction, Steps, stepslowdown)
End Function
Public Property Get ramotor() As Long
    'RAMotor = EQGetMotorValues(0)
     ramotor = GetEmulRA()
End Property

Public Property Get DECMotor() As Long
    'DECMotor = EQGetMotorValues(1)
    DECMotor = gEmulDEC
End Property

Public Property Get SyncRAMotor() As Long
     gEmulRA = EQGetMotorValues(0)
     SyncRAMotor = gEmulRA
     gLast_time = EQnow_lst_norange()
     gCurrent_time = gLast_time
     gEmulRA_Init = gEmulRA
End Property

Public Property Get SyncDECMotor() As Long
     gEmulDEC = EQGetMotorValues(1)
     SyncDECMotor = gEmulDEC
End Property
Public Property Get RAWormPeriod() As Long
     RAWormPeriod = gRAWormPeriod
End Property
Public Property Get RAWormPeriodFloat() As Double
    RAWormPeriodFloat = SPSD * gRAWormSteps / gTot_RA
End Property
Public Property Get PecGain() As Double
    PecGain = gPEC_Gain
End Property
Public Property Let PecGain(ByVal newval As Double)
    PECConfigFrm.GainScroll.Value = newval * 10
End Property



' ------------------------- Methods Portion of the Code -------------


Public Sub AbortSlew()

    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 6, ("COMMAND AbortSlew")
    If gEQparkstatus <> 0 Then
        ' no move axis if parked or parking!
        RaiseError SCODE_SCOPE_PARKED, ERR_SOURCE, "AbortSlew() " & MSG_SCOPE_PARKED
        Exit Sub
    End If

    If gSlewStatus Then
        gSlewStatus = False
        ' stop the slew if already slewing
        eqres = EQ_MotorStop(0)
        eqres = EQ_MotorStop(1)
        gRAStatus_slew = False
    
        ' restart tracking
        RestartTracking

    End If

End Sub

Public Sub CommandBlind(ByVal Command As String, Optional ByVal Raw As Boolean = False)
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("COMMAND CommandBlind(" & Command & " :NOT_SUPPORTED")
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Method xxx()" & MSG_NOT_IMPLEMENTED

End Sub

Public Function CommandBool(ByVal Command As String, Optional ByVal Raw As Boolean = False) As Boolean
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("COMMAND CommandBool(" & Command & " :NOT_SUPPORTED")
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Method CommandBlind()" & MSG_NOT_IMPLEMENTED

End Function

Public Function CommandString(ByVal Command As String, Optional ByVal Raw As Boolean = False) As String

Dim pos As Long
Dim tmpstr As String
Dim tmp2 As String
Dim arg1 As Double
Dim arg2 As Double
Dim arg3 As Double
Dim tmp1 As Integer

Dim comStr() As String
Dim ParseErr As Boolean

    On Error GoTo HandleError

    ParseErr = False
    If Right$(Command, 1) <> "#" Then
        ParseErr = True
    Else
        Command = Left(Command, Len(Command) - 1)
        comStr = Split(Command, ",")
    
        Select Case comStr(0)
    
            Case ":PECGAIN"
                Select Case UBound(comStr)
                    Case 1
                        'Update the PEC Gain value
                        If PEC_SetGain(comStr(1)) Then
                            CommandString = "1#"
                        Else
                            CommandString = "0#"
                        End If
                    Case 0
                        CommandString = FormatNumber(PECConfigFrm.GainScroll.Value, 0, , , 0) & "#"
                    Case Else
                        ParseErr = True
                End Select
                
            Case ":PECPHASE"
                Select Case UBound(comStr)
                    Case 1
                        If PEC_SetPhase(comStr(1)) Then
                            CommandString = "1#"
                        Else
                            CommandString = "0#"
                        End If
                    Case 0
                        CommandString = FormatNumber(PECConfigFrm.PhaseScroll.Value, 0, , , 0) & "#"
                    Case Else
                        ParseErr = True
                End Select
        
            Case ":PECFILE"
                CommandString = PECDef1.FileName & "#"
        
            Case ":PECLOAD"
                If UBound(comStr) = 1 Then
                    ' Load a PEC file
                    If PEC_LoadFile(comStr(1)) Then
                        CommandString = "1#"
                    Else
                        CommandString = "0#"
                    End If
                Else
                    ParseErr = True
                End If
            
            Case ":PECSAVE"
                If UBound(comStr) = 1 Then
                    ' Save a PEC file
                    If PEC_SaveFile(comStr(1), PECDef1) Then
                        CommandString = "1#"
                    Else
                        CommandString = "0#"
                    End If
                Else
                    ParseErr = True
                End If
            
            Case ":PECENA"
                If UBound(comStr) = 1 Then
                    Select Case comStr(1)
                        Case "1"
                            'enable PEC
                            HC.CheckPEC.Value = 1
                            PEC_StartTracking
                        Case "0"
                            'disable PEC
                            HC.CheckPEC.Value = 0
                            PEC_StopTracking
                        Case Else
                            ParseErr = True
                    End Select
                Else
                    ParseErr = True
                End If
            
            Case ":PECSTA":
                ' read pec state
                If gPEC_Enabled = True Then
                    CommandString = "1#"
                Else
                    CommandString = "0#"
                End If
            
            Case ":PECWTC":
                ' read worm tooth count
                If EQ_GetMountStatus() <> 1 Then
                    CommandString = FormatNumber(-1, 0, , , 0) & "#"
                Else
                    pos = gTot_RA / gRAWormSteps
                    CommandString = FormatNumber(pos, 0, , , 0) & "#"
                End If
            
            Case ":PECIDX"
                ' read worm position
                If EQ_GetMountStatus() <> 1 Then
                    CommandString = FormatNumber(-1, 0, , , 0) & "#"
                Else
                    pos = NormalisePosition(SyncRAMotor, gRAWormSteps)
                    CommandString = FormatNumber(pos, 0, , , 0) & "#"
                End If
        
            Case ":PECINFO"
                ' read worm
                If EQ_GetMountStatus() <> 1 Then
                    CommandString = FormatNumber(-1, 0, , , 0) & "," & FormatNumber(-1, 0, , , 0) & "#"
                Else
                    pos = Int(gRAWormSteps - 1)
                    CommandString = FormatNumber(Int(gRAWormPeriod), 0, , , 0) & "," & FormatNumber(pos, 0, , , 0) & "#"
                End If
            
            Case ":PECSET"
                ' Write a PEC table entry
                If UBound(comStr) = 3 Then
                    arg1 = CInt(comStr(1))
                    arg2 = CLng(comStr(2))
                    arg3 = CDbl(comStr(3)) / 1000
                    If PEC_Write_Table(arg1, arg2, arg3) Then
                        CommandString = "1#"
                    Else
                        CommandString = "0#"
                    End If
                Else
                    ParseErr = True
                End If
                
            Case ":PECGET"
                If UBound(comStr) = 1 Then
                    ' Read a PEC table entry
                    arg1 = CInt(comStr(1))
                    If arg1 >= 0 And arg1 <= UBound(PECDef1.PECCurve) Then
                        With PECDef1.PECCurve(arg1)
                            pos = .signal * 1000
                            tmpstr = FormatNumber(pos, 0, , , 0)
                            pos = .PEPosition
                            CommandString = "1," & FormatNumber(pos, 0, , , 0) & "," & tmpstr & "#"
                        End With
                    Else
                        CommandString = "#0"
                    End If
                Else
                    ParseErr = True
                End If
                
            Case ":PARK"
                If UBound(comStr) = 1 Then
                    tmp1 = val(comStr(1))
                    If tmp1 >= 0 And tmp1 <= 7 Then
                        If tmp1 = 0 Then
                            Call HC.ApplyParkMode
                        Else
                            Call ApplyParkMode2(tmp1 - 1)
                        End If
                    End If
                    If gEQparkstatus <> 0 Then
                        ' mount is parked or is parking
                        CommandString = "1#"
                    Else
                        ' mount is still unparked
                        CommandString = "0#"
                    End If
                Else
                    ParseErr = True
                End If
            
            Case ":UNPARK"
                If UBound(comStr) = 1 Then
                    tmp1 = val(comStr(1))
                    If tmp1 >= 0 And tmp1 <= 6 Then
                        If tmp1 = 0 Then
                            Call HC.ApplyUnParkMode
                        Else
                            Call ApplyUnParkMode2(tmp1 - 1)
                        End If
                    End If
                    If gEQparkstatus = 0 Then
                        CommandString = "1#"
                    Else
                        CommandString = "0#"
                    End If
                Else
                    ParseErr = True
                End If
                          
            Case ":RA_ENC"
                If EQ_GetMountStatus() <> 1 Then
                    CommandString = FormatNumber(-1, 0, , , 0) & "#"
                Else
                    pos = SyncRAMotor()
                    CommandString = FormatNumber(pos, 0, , , 0) & "#"
                End If
        
            Case ":DEC_ENC"
                If EQ_GetMountStatus() <> 1 Then
                    CommandString = FormatNumber(-1, 0, , , 0) & "#"
                Else
                    pos = SyncDECMotor()
                    CommandString = FormatNumber(pos, 0, , , 0) & "#"
                End If
    
            Case ":ST4_RARATE"
                Select Case UBound(comStr)
                    Case 0
                        ' read current RA_Rate
                        CommandString = Right$(HC.RAGuideRateList.Text, 4)
                    Case 1
                        ' write new RA rate
                        Select Case comStr(1)
                            Case "1.00", "0.75", "0.50", "0.25"
                                Call writeportrateRa("x" & comStr(1))
                                Call readportrate
                                ' success
                                CommandString = "1#"
                            Case Else
                                ' failure
                                CommandString = "0#"
                        End Select
                    Case Else
                        ParseErr = True
                End Select
            
            Case ":ST4_DECRATE"
                Select Case UBound(comStr)
                    Case 0
                        ' read current DEC_Rate
                        CommandString = Right$(HC.DECGuideRateList.Text, 4)
                    Case 1
                        ' write new dec rate
                        Select Case comStr(1)
                            Case "1.00", "0.75", "0.50", "0.25"
                                Call writeportrateDec("x" & comStr(1))
                                Call readportrate
                                ' success
                                CommandString = "1#"
                            Case Else
                                ' failure
                                CommandString = "0#"
                        End Select
                    Case Else
                        ParseErr = True
                End Select
            
            Case ":PG_RARATE"
                Select Case UBound(comStr)
                    Case 0
                        ' read current RA_Rate
                        CommandString = Right$(HC.Label14.Caption, 3)
                    Case 1
                        ' write new RA rate
                        ' assume success
                        CommandString = "1#"
                        Select Case comStr(1)
                            Case "0.9"
                                HC.HScrollRARate.Value = 9
                            Case "0.8"
                                HC.HScrollRARate.Value = 8
                            Case "0.7"
                                HC.HScrollRARate.Value = 7
                            Case "0.6"
                                HC.HScrollRARate.Value = 6
                            Case "0.5"
                                HC.HScrollRARate.Value = 5
                            Case "0.4"
                                HC.HScrollRARate.Value = 4
                            Case "0.3"
                                HC.HScrollRARate.Value = 3
                            Case "0.2"
                                HC.HScrollRARate.Value = 2
                            Case "0.1"
                                HC.HScrollRARate.Value = 1
                            Case Else
                                ' failure
                                CommandString = "0#"
                        End Select
                    Case Else
                        ParseErr = True
                End Select
            
            Case ":PG_DECRATE"
                Select Case UBound(comStr)
                    Case 0
                        ' read current DEC_Rate
                        CommandString = Right$(HC.Label15.Caption, 3)
                    Case 1
                        ' write new RA rate
                        ' assume success
                        CommandString = "1#"
                        Select Case comStr(1)
                            Case "0.9"
                                HC.HScrollDecRate.Value = 9
                            Case "0.8"
                                HC.HScrollDecRate.Value = 8
                            Case "0.7"
                                HC.HScrollDecRate.Value = 7
                            Case "0.6"
                                HC.HScrollDecRate.Value = 6
                            Case "0.5"
                                HC.HScrollDecRate.Value = 5
                            Case "0.4"
                                HC.HScrollDecRate.Value = 4
                            Case "0.3"
                                HC.HScrollDecRate.Value = 3
                            Case "0.2"
                                HC.HScrollDecRate.Value = 2
                            Case "0.1"
                                HC.HScrollDecRate.Value = 1
                            Case Else
                                ' failure
                                CommandString = "0#"
                        End Select
                    Case Else
                        ParseErr = True
                End Select
                
            Case ":FORM_RESET"
                Call ResetFormPosition
                
            Case ":PGENA_RA"
                If UBound(comStr) = 1 Then
                    Select Case comStr(1)
                        Case "1"
                            'enable RA Pulseguideing
                            HC.rapulse_enchk.Value = 1
                        Case "0"
                            'disable RA Pulseguideing
                            HC.rapulse_enchk.Value = 0
                        Case Else
                            ParseErr = True
                    End Select
                Else
                    If HC.rapulse_enchk.Value = 0 Then
                        CommandString = "0#"
                    Else
                        CommandString = "1#"
                    End If
                End If
            
            Case ":PGENA_DEC"
                If UBound(comStr) = 1 Then
                    Select Case comStr(1)
                        Case "1"
                            'enable DEC Pulseguideing
                            HC.decpulse_enchk.Value = 1
                            CommandString = "1#"
                        Case "0"
                            'disable DEC Pulseguideing
                            HC.decpulse_enchk.Value = 0
                            CommandString = "1#"
                        Case Else
                            ParseErr = True
                    End Select
                Else
                    If HC.decpulse_enchk.Value = 0 Then
                        CommandString = "0#"
                    Else
                        CommandString = "1#"
                    End If
                End If
            
                
                
            Case ":DRIFTCOMP"
                Select Case UBound(comStr)
                    Case 1
                        gDriftComp = val(comStr(1))
                        HC.DriftScroll.Value = gDriftComp
                        CommandString = "1#"
                    Case 0
                        CommandString = FormatNumber(gDriftComp, 0, , , 0) & "#"
                    Case Else
                        ParseErr = True
                End Select
            
            Case ":DLLVER"
                pos = EQ_DriverVersion
                tmpstr = Hex$((pos And &HF00000) / 1048576 And &HF) + Hex$((pos And &HF0000) / 65536 And &HF)
                tmpstr = tmpstr & "." & Hex$((pos And &HF000) / 4096 And &HF) + Hex$((pos And &HF00) / 256 And &HF)
                tmpstr = tmpstr & "." & Hex$((pos And &HF0) / 16 And &HF) + Hex$(pos And &HF)
                CommandString = tmpstr
        
            Case ":MOUNTVER"
                tmpstr = Hex$((gMount_Ver And &HF0) / 16 And &HF) + Hex$(gMount_Ver And &HF)
                tmpstr = tmpstr & "." & Hex$((gMount_Ver And &HF000) / 4096 And &HF) + Hex$((gMount_Ver And &HF00) / 256 And &HF)
                tmpstr = tmpstr & "." & Hex$((gMount_Ver And &HF00000) / 1048576 And &HF) + Hex$((gMount_Ver And &HF0000) / 65536 And &HF)
                CommandString = tmpstr
    
            Case ":DRIVERVER"
                CommandString = gVersion
                
            Case Else
                ' unknown command!
                ParseErr = True
                
        End Select
    End If

    If ParseErr Then
HandleError:
        RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Method CommandString():" & Command & " " & MSG_NOT_IMPLEMENTED
    End If

    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("COMMAND CommandString()" & Command & " :" & CommandString)
 
End Function

Public Sub FindHome()
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 7, ("COMMAND FindHome :NOT_SUPPORTED")
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Method FindHome()" & MSG_NOT_IMPLEMENTED
End Sub

Public Sub Park()
    Dim endtime As Double
    Dim nowtime As Double
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 7, ("COMMAND Park")
    If gEQparkstatus = 0 Then
        ' initiate park
        Call ParktoUserDefine
        
        ' ASCOM, in their wisdom (or lack of it), require that park blocks the client until completion.
        ' This is rather poor and we have chosen to ignor that part of the spec believing that
        ' non blocking asynchronous methods are a much better solution. However some clients may
        ' require a blocking function so we've provided an option to allow this.
        If gAscomCompatibility.BlockPark Then
            
            '----------------------------
            ' Blocking
            '----------------------------
            ' 3 minute timeout should be more than long enough to park from any position
            nowtime = EQnow_lst_norange
            endtime = nowtime + 180
            While gEQparkstatus <> 1
                DoEvents
                nowtime = EQnow_lst_norange
                If nowtime >= endtime Then
                    ' timeout - lets ingnore the crazy ASCOM syncronos park and do it right i.e. asynchronously
                    Exit Sub
                End If
            Wend
        
        End If
    End If
End Sub

Public Property Get IsPulseGuiding() As Boolean
    If gAscomCompatibility.AllowPulseGuide Then
        If (gEQRAPulseDuration + gEQDECPulseDuration) <> 0 Then
            IsPulseGuiding = True
        Else
            IsPulseGuiding = False
        End If
        If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 4, ("Get IsPulseGuiding :" & IsPulseGuiding)
    Else
        RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Read Property IsPulseGuidng :NOT SUPPORTED"
    End If
End Property

Public Sub PulseGuide(ByVal direction As GuideDirections, ByVal Duration As Long)
    
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 4, ("COMMAND PulseGuide(" & direction & "," & Duration & ")")
    
    If Not gAscomCompatibility.AllowPulseGuide Then
        RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "PulseGuide() :NOT SUPPORTED"
        Exit Sub
    End If
    
    If gSlewStatus Then
        ' no guiding whilst slewing - makes no sense and means the slew will terminate!
        Exit Sub
    End If
    
    Select Case gEQparkstatus
        Case 0
            ' unparked
        Case 1
            ' parked
            ' no guidng if parked
            If gAscomCompatibility.AllowPulseGuideExceptions Then
                RaiseError SCODE_SCOPE_PARKED, ERR_SOURCE, "Method PulseGuide() " & MSG_SCOPE_PARKED
            End If
            Exit Sub
        Case 2
            ' parking
            ' no guidng if parking
            Exit Sub
    End Select
    
    'Pulse guide implemtation for EQMOD
    'This uses the duration parameter and an asynchronous timer
    'that will decrement the duration count for every time tick
    'It then disables the guiderate upon expiration of the counter
    
    HC.Pulseguide_Timer.Enabled = False
    
    Select Case (direction)
        Case guideNorth, guideSouth  ''DEC+,DEC-
            If HC.decpulse_enchk.Value = 1 Then
                ' apply gain to duration
                If HC.decfixed_enchk.Value = 0 Then
                    ' apply gain
                    gEQDECPulseDuration = Duration * (HC.HScrollDECWidth.Value / 100)
                Else
                    ' use fixed period
                    gEQDECPulseDuration = HC.HScrollDecOride.Value * 100
                End If
                
                If direction <> LastGuideNS Then
                    gEQDECPulseDuration = gEQDECPulseDuration + gBacklashDec
                End If
                
                If gEQDECPulseDuration < (gpl_interval / 2) Then
                    ' pulse duration is too small to action - better to ignor it than over correct
                    Duration = 0
                End If
            Else
                gEQDECPulseDuration = 0
                Duration = 0
                ' nothing to do!
                GoTo endsub
            End If
        
        Case guideEast, guideWest  ''RA+,RA-
            If HC.rapulse_enchk.Value = 1 Then
                If HC.rafixed_enchk.Value = 0 Then
                    ' apply gain to duration
                    gEQRAPulseDuration = Duration * (HC.HScrollRAWidth.Value / 100)
                Else
                    'use fixed period
                    gEQRAPulseDuration = HC.HScrollRAOride.Value * 100
                End If
                
                If gEQRAPulseDuration < (gpl_interval / 2) Then
                    ' pulse duration is too small - better to ignor it than over correct
                    Duration = 0
                End If
            Else
                gEQRAPulseDuration = 0
                Duration = 0
                GoTo endsub
            End If
            
        Case Else
            ' invalid direction
            If gAscomCompatibility.AllowPulseGuideExceptions Then
                RaiseError SCODE_VAL_OUTOFRANGE, ERR_SOURCE, _
                    "Pulse Guide: GuideDirections" & MSG_VAL_OUTOFRANGE
            End If
     End Select
    
    If (gTrackingStatus = 4) Then
        If Duration = 0 Then
            Select Case (direction)
                Case guideNorth 'DEC+
                    gEQDECPulseDuration = 0
                    Call ChangeDEC_by_Rate(gDeclinationRate)
                    Call Plot_PG(1, 0, 0)
                
                Case guideSouth 'DEC-
                    gEQDECPulseDuration = 0
                    Call ChangeDEC_by_Rate(gDeclinationRate)
                    Call Plot_PG(1, 1, 0)
            
                Case guideEast  'RA+
                    If gRA_LastRate = 0 Then
                        eqres = EQ_SendGuideRate(0, gTrackingStatus - 1, 0, 0, gHemisphere, gHemisphere)
                    Else
                        ChangeRA_by_Rate (gRightAscensionRate)
                    End If
                    gEQRAPulseDuration = 0
                    Call Plot_PG(0, 0, 0)
            
                Case guideWest  'RA-
                    If gRA_LastRate = 0 Then
                        eqres = EQ_SendGuideRate(0, gTrackingStatus - 1, 0, 0, gHemisphere, gHemisphere)
                    Else
                        ChangeRA_by_Rate (gRightAscensionRate)
                    End If
                    Call Plot_PG(0, 1, 0)
                    gEQRAPulseDuration = 0
            End Select
        Else
          Select Case (direction)
                Case guideNorth 'DEC+
                    Call ChangeDEC_by_Rate(gDeclinationRate + (gDeclinationRate * HC.HScrollDecRate.Value * 0.1))
                    Call Plot_PG(1, 0, gEQDECPulseDuration)
                    
                Case guideSouth 'DEC-
                    Call ChangeDEC_by_Rate(gDeclinationRate - (gDeclinationRate * HC.HScrollDecRate.Value * 0.1))
                    eqres = EQ_SendGuideRate(1, gTrackingStatus - 1, val(HC.HScrollDecRate.Value), 0, 0, 0)
                    Call Plot_PG(1, 1, gEQDECPulseDuration)
                
                Case guideEast  'RA+
                    If gRA_LastRate = 0 Then
                        eqres = EQ_SendGuideRate(0, gTrackingStatus - 1, val(HC.HScrollRARate.Value), 1, gHemisphere, gHemisphere)
                    Else
                        Call ChangeRA_by_Rate(gRightAscensionRate - (gRightAscensionRate * HC.HScrollRARate.Value * 0.1))
                    End If
                    Call Plot_PG(0, 0, gEQRAPulseDuration)
                
                Case guideWest  'RA-
                    If gRA_LastRate = 0 Then
                       eqres = EQ_SendGuideRate(0, gTrackingStatus - 1, val(HC.HScrollRARate.Value), 0, gHemisphere, gHemisphere)
                    Else
                       Call ChangeRA_by_Rate(gRightAscensionRate + (gRightAscensionRate * HC.HScrollRARate.Value * 0.1))
                    End If
                    Call Plot_PG(0, 1, gEQRAPulseDuration)
              End Select
         End If
    Else
        ' Process if Equatorial Tracking
        If Duration = 0 Then
            Select Case (direction)
                
                Case guideNorth 'DEC+
                    eqres = EQ_MotorStop(1)
                    If eqres = 1 Then GoTo PError
'                    Do
'                      eqres = EQ_GetMotorStatus(1)
'                      If eqres = 1 Then GoTo PError
'                    Loop While (eqres And &H10) <> 0
                    gEQDECPulseDuration = 0
                    gPulseguideRateDec = 0
                    Call Plot_PG(1, 0, 0)
                
                Case guideSouth 'DEC-
                    eqres = EQ_MotorStop(1)
                    If eqres = 1 Then GoTo PError
 '                   Do
 '                     eqres = EQ_GetMotorStatus(1)
 '                     If eqres = 1 Then GoTo PError
 '                   Loop While (eqres And &H10) <> 0
                    gEQDECPulseDuration = 0
                    gPulseguideRateDec = 0
                    Call Plot_PG(1, 1, 0)
                
                Case guideEast  'RA+
                    If gTrackingStatus = 0 Then
                        eqres = EQ_SendGuideRate(0, 0, 0, 0, gHemisphere, gHemisphere)
                    Else
                        eqres = EQ_SendGuideRate(0, gTrackingStatus - 1, 0, 0, gHemisphere, gHemisphere)
                    End If
                    gEQRAPulseDuration = 0
                    gPulseguideRateRa = 0
                    Call Plot_PG(0, 0, 0)
                
                Case guideWest  'RA-
                    If gTrackingStatus = 0 Then
                        eqres = EQ_SendGuideRate(0, gTrackingStatus - 1, 0, 0, gHemisphere, gHemisphere)
                    Else
                        eqres = EQ_SendGuideRate(0, 0, 0, 0, gHemisphere, gHemisphere)
                    End If
                    gEQRAPulseDuration = 0
                    gPulseguideRateRa = 0
                    Call Plot_PG(0, 1, 0)
            
            End Select
       
       Else
            Select Case (direction)
                
                Case guideNorth 'DEC+
'                    eqres = EQ_MotorStop(1)
'                    If eqres = 1 Then GoTo PError
'                    Do
'                      eqres = EQ_GetMotorStatus(1)
'                      If eqres = 1 Then GoTo PError
'                    Loop While (eqres And &H10) <> 0
                    If gTrackingStatus = 0 Then
                        eqres = EQ_SendGuideRate(1, 0, val(HC.HScrollDecRate.Value), 1, 0, 0)
                    Else
                        eqres = EQ_SendGuideRate(1, gTrackingStatus - 1, val(HC.HScrollDecRate.Value), 1, 0, 0)
                    End If
                    
                    gPulseguideRateDec = HC.HScrollDecRate.Value * 0.1
                    
                    Call Plot_PG(1, 0, gEQDECPulseDuration)
                    LastGuideNS = guideNorth
                    
                Case guideSouth 'DEC-
'                    eqres = EQ_MotorStop(1)
'                    If eqres = 1 Then GoTo PError
'                    Do
'                      eqres = EQ_GetMotorStatus(1)
'                      If eqres = 1 Then GoTo PError
'                    Loop While (eqres And &H10) <> 0
                    If gTrackingStatus = 0 Then
                        eqres = EQ_SendGuideRate(1, 0, val(HC.HScrollDecRate.Value), 0, 0, 0)
                    Else
                        eqres = EQ_SendGuideRate(1, gTrackingStatus - 1, val(HC.HScrollDecRate.Value), 0, 0, 0)
                    End If
                    
                    gPulseguideRateDec = HC.HScrollDecRate.Value * -0.1
                    
                    Call Plot_PG(1, 1, gEQDECPulseDuration)
                    LastGuideNS = guideSouth
                
                Case guideEast  'RA+
                    If gTrackingStatus = 0 Then
                        eqres = EQ_SendGuideRate(0, 0, val(HC.HScrollRARate.Value), 1, gHemisphere, gHemisphere)
                    Else
                        eqres = EQ_SendGuideRate(0, gTrackingStatus - 1, val(HC.HScrollRARate.Value), 1, gHemisphere, gHemisphere)
                    End If
                    
                    gPulseguideRateRa = HC.HScrollRARate.Value * -0.1
                    
                    Call Plot_PG(0, 0, gEQRAPulseDuration)
                
                Case guideWest  'RA-
                    If gTrackingStatus = 0 Then
                        eqres = EQ_SendGuideRate(0, 0, val(HC.HScrollRARate.Value), 0, gHemisphere, gHemisphere)
                    Else
                        eqres = EQ_SendGuideRate(0, gTrackingStatus - 1, val(HC.HScrollRARate.Value), 0, gHemisphere, gHemisphere)
                    End If
                    gPulseguideRateRa = HC.HScrollRARate.Value * 0.1
                    Call Plot_PG(0, 1, gEQRAPulseDuration)
            End Select
        End If
    End If
    
endsub:
    HC.Pulseguide_Timer.Enabled = True
    
PError:

End Sub

Public Sub SetPark()
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 7, ("COMMAND SetPark")
    HC.Add_Message (oLangDll.GetLangString(5004))
    DefinePark (True)
End Sub

Public Sub SetupDialog()
    Unload Setupfrm
    Setupfrm.Show 1
End Sub

Public Sub SlewToCoordinates(ByVal RightAscension As Double, ByVal Declination As Double)
    Dim nowtime As Double
    Dim endtime As Double
    
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 6, ("COMMAND SlewToCoordinates(" & RightAscension & "," & Declination & ")")
    If gEQparkstatus = 0 Then
        If gAscomCompatibility.SlewWithTrackingOff = False And gTrackingStatus = 0 Then
            RaiseError SCODE_RADEC_SLEW_ERROR, ERR_SOURCE, "SlewToCoordinates() " & MSG_RADEC_SLEW_ERROR
        Else
            If Align.Visible Then
                Call Align.FillAlignmentStar(RightAscension, Declination)
            End If
            gTargetRA = RightAscension
            gTargetDec = Declination
            HC.Add_Message ("SyncCSlew: " & oLangDll.GetLangString(105) & "[ " & FmtSexa(gTargetRA, False) & " ] " & oLangDll.GetLangString(106) & "[ " & FmtSexa(gTargetDec, True) & " ]")
            gSlewCount = gMaxSlewCount   'NUM_SLEW_RETRIES               'Set initial iterative slew count
            Call radecAsyncSlew(gGotoRate)
            EQ_Beep (20)
            
            '----------------------------
            ' Blocking
            '----------------------------
            ' 3 minute timeout should be more than long enough to slew to any position
            nowtime = EQnow_lst_norange
            endtime = nowtime + 180
            While gSlewStatus <> 0
                DoEvents
                nowtime = EQnow_lst_norange
                If nowtime >= endtime Then
                    Exit Sub
                End If
            Wend
        
        End If
    Else
            HC.Add_Message (oLangDll.GetLangString(5000))
            RaiseError SCODE_SCOPE_PARKED, ERR_SOURCE, "SlewToCoordinates() " & MSG_SCOPE_PARKED
    End If
End Sub

Public Sub SlewToCoordinatesAsync(ByVal RightAscension As Double, ByVal Declination As Double)
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 6, ("COMMAND SlewToCoordinatesAsync(" & RightAscension & "," & Declination & ")")
    If gEQparkstatus = 0 Then
    
        If gAscomCompatibility.SlewWithTrackingOff = False And gTrackingStatus = 0 Then
            RaiseError SCODE_RADEC_SLEW_ERROR, ERR_SOURCE, "SlewToCoordinatesAsync() " & MSG_RADEC_SLEW_ERROR
        Else
            If Align.Visible Then
                Call Align.FillAlignmentStar(RightAscension, Declination)
            End If
            gTargetRA = RightAscension
            gTargetDec = Declination
            HC.Add_Message ("CoordSlew: " & oLangDll.GetLangString(105) & "[ " & FmtSexa(gTargetRA, False) & " ] " & oLangDll.GetLangString(106) & "[ " & FmtSexa(gTargetDec, True) & " ]")
            gSlewCount = gMaxSlewCount   'NUM_SLEW_RETRIES               'Set initial iterative slew count
            Call radecAsyncSlew(gGotoRate)
            EQ_Beep (20)
        End If
    Else
        HC.Add_Message (oLangDll.GetLangString(5000))
        RaiseError SCODE_SCOPE_PARKED, ERR_SOURCE, "SlewToCoordinatesAsync() " & MSG_SCOPE_PARKED
    End If
End Sub

Public Sub SlewToTarget()
    Dim nowtime As Double
    Dim endtime As Double

    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 6, ("COMMAND SlewToTarget")
    If gEQparkstatus = 0 Then
        If gAscomCompatibility.SlewWithTrackingOff = False And gTrackingStatus = 0 Then
            RaiseError SCODE_RADEC_SLEW_ERROR, ERR_SOURCE, "SlewToTarget() " & MSG_RADEC_SLEW_ERROR
        Else
            If Align.Visible Then
                Call Align.FillAlignmentStar(gTargetRA, gTargetDec)
            End If
            HC.Add_Message ("SyncSlew: " & oLangDll.GetLangString(105) & "[ " & FmtSexa(gTargetRA, False) & "] " & oLangDll.GetLangString(106) & "[ " & FmtSexa(gTargetDec, True) & " ]")
            gSlewCount = gMaxSlewCount   'NUM_SLEW_RETRIES               'Set initial iterative slew count
            Call radecAsyncSlew(gGotoRate)
            EQ_Beep (20)
            
            '----------------------------
            ' Blocking
            '----------------------------
            ' 3 minute timeout should be more than long enough to slew to any position
            nowtime = EQnow_lst_norange
            endtime = nowtime + 180
            While gSlewStatus <> 0
                DoEvents
                nowtime = EQnow_lst_norange
                If nowtime >= endtime Then
                    Exit Sub
                End If
            Wend
        
        End If
    Else
        HC.Add_Message (oLangDll.GetLangString(5000))
        RaiseError SCODE_SCOPE_PARKED, ERR_SOURCE, "SlewToTarget() " & MSG_SCOPE_PARKED
    End If
    
End Sub

Public Sub SlewToTargetAsync()

    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 6, ("COMMAND SlewToTargetAsync")
    If gEQparkstatus = 0 Then
        If gAscomCompatibility.SlewWithTrackingOff = False And gTrackingStatus = 0 Then
            RaiseError SCODE_RADEC_SLEW_ERROR, ERR_SOURCE, "SlewToTargetAsync() " & MSG_RADEC_SLEW_ERROR
        Else
            If Align.Visible Then
                Call Align.FillAlignmentStar(gTargetRA, gTargetDec)
            End If
            HC.Add_Message ("AsyncSlew: " & oLangDll.GetLangString(105) & "[ " & FmtSexa(gTargetRA, False) & "] " & oLangDll.GetLangString(106) & "[ " & FmtSexa(gTargetDec, True) & " ]")
            gSlewCount = gMaxSlewCount   'NUM_SLEW_RETRIES               'Set initial iterative slew count
            Call radecAsyncSlew(gGotoRate)
            EQ_Beep (20)
        End If
    Else
        HC.Add_Message (oLangDll.GetLangString(5000))
        RaiseError SCODE_SCOPE_PARKED, ERR_SOURCE, "SlewToTargetAsync() " & MSG_SCOPE_PARKED
    End If
    
End Sub

Public Sub SyncToCoordinates(ByVal RightAscension As Double, ByVal Declination As Double)
                                    
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 7, ("COMMAND SyncToCoordinates(" & RightAscension & "," & Declination & ")")
    If gEQparkstatus = 0 Then
        If gTrackingStatus = 0 Then
            RaiseError SCODE_RADEC_SLEW_ERROR, ERR_SOURCE, "SyncToCoordinates() " & MSG_RADEC_SLEW_ERROR
        Else
            HC.Add_Message ("SynCoor: " & oLangDll.GetLangString(105) & "[ " & FmtSexa(RightAscension, False) & "] " & oLangDll.GetLangString(106) & "[ " & FmtSexa(Declination, True) & " ]")
            SyncToRADEC RightAscension, Declination, gLongitude, gHemisphere
            EQ_Beep (4)
        End If
    Else
        RaiseError SCODE_SCOPE_PARKED, ERR_SOURCE, "SyncToCoordinates() " & MSG_SCOPE_PARKED
    End If

End Sub

Public Sub SyncToTarget()

    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 7, ("COMMAND SyncToTarget")
    If gEQparkstatus = 0 Then
        If gTrackingStatus = 0 Then
            RaiseError SCODE_RADEC_SLEW_ERROR, ERR_SOURCE, "SyncToTarget() " & MSG_RADEC_SLEW_ERROR
        Else
            HC.Add_Message ("SyncTaget: " & oLangDll.GetLangString(105) & "[ " & FmtSexa(gTargetRA, False) & "] " & oLangDll.GetLangString(106) & "[ " & FmtSexa(gTargetDec, True) & " ]")
            SyncToRADEC gTargetRA, gTargetDec, gLongitude, gHemisphere
            EQ_Beep (4)
        End If
    Else
        RaiseError SCODE_SCOPE_PARKED, ERR_SOURCE, "SyncToTarget() " & MSG_SCOPE_PARKED
    End If


End Sub

Public Sub UnPark()
Dim nowtime As Double
Dim endtime As Double

    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 7, ("COMMAND Unpark")
    Unparkscope
   
    If gAscomCompatibility.BlockPark Then
        '----------------------------
        ' Blocking
        '----------------------------
        ' 3 minute timeout should be more than long enough to park from any position
        nowtime = EQnow_lst_norange
        endtime = nowtime + 180
        While gEQparkstatus <> 0
            DoEvents
            nowtime = EQnow_lst_norange
            If nowtime >= endtime Then
                ' timeout - lets ingnore the crazy ASCOM syncronos park and do it right i.e. asynchronously
                Exit Sub
            End If
        Wend
    End If
    
End Sub

Public Sub StopClientCount()
    ' A client can use this to ensure when it closes EQASCOM will close too irrespective of other attached clients.
    ClientCount = 2
End Sub

Public Sub SetClientCount(ByVal Count As Double)
    ClientCount = Count
End Sub
Public Sub IncClientCount()
    ClientCount = ClientCount + 1
End Sub

Public Property Get DecNoAdjust() As Double
    DecNoAdjust = gDec_DegNoAdjust
End Property

Public Property Get PulseGuideRateRa()
    PulseGuideRateRa = gPulseguideRateRa
End Property

Public Property Get PulseGuideRateDec()
    PulseGuideRateDec = gPulseguideRateDec
End Property

Public Sub ReadComPortSettings()

   
     Dim tmptxt As String
   
     tmptxt = HC.oPersist.ReadIniValue("Port")
     If tmptxt <> "" Then
        gPort = tmptxt
     Else
        gPort = "COM1"
     End If
     
     tmptxt = HC.oPersist.ReadIniValue("Baud")
     If tmptxt <> "" Then
        gBaud = val(tmptxt)
     Else
        gBaud = 9600
     End If
     
     tmptxt = HC.oPersist.ReadIniValue("Timeout")
     If tmptxt <> "" Then
        gTimeout = val(tmptxt)
     Else
        gTimeout = 1000
     End If
     

     tmptxt = HC.oPersist.ReadIniValue("Retry")
     If tmptxt <> "" Then
        gRetry = val(tmptxt)
     Else
        gRetry = 1
     End If
     
End Sub

Public Sub ReadSyncMap()

     Dim tmptxt As String

     tmptxt = HC.oPersist.ReadIniValue("RSYNC01")
     If tmptxt <> "" Then
        gRASync01 = val(tmptxt)
     Else
        gRASync01 = 0
     End If
     
     tmptxt = HC.oPersist.ReadIniValue("DSYNC01")
     If tmptxt <> "" Then
        gDECSync01 = val(tmptxt)
     Else
        gDECSync01 = 0
     End If

    HC.DxSalbl.Caption = Format$(str(gRASync01), "000000000")
    HC.DxSblbl.Caption = Format$(str(gDECSync01), "000000000")

End Sub
Public Sub ReadAlignMap()

     Dim tmptxt As String

     tmptxt = HC.oPersist.ReadIniValue("RALIGN01")
     If tmptxt <> "" Then
        gRA1Star = val(tmptxt)
     Else
        gRA1Star = 0
     End If
     
     tmptxt = HC.oPersist.ReadIniValue("DALIGN01")
     If tmptxt <> "" Then
        gDEC1Star = val(tmptxt)
     Else
        gDEC1Star = 0
     End If
     
End Sub


Private Sub RaiseError(ErrNumber As Long, ErrSource As String, ErrDescription As String)
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("ERROR RAISED: errno=" & Hex(ErrNumber) & " " & ErrDescription)
    If gAscomCompatibility.AllowExceptions Then
        Err.Raise ErrNumber, ErrSource, ErrDescription
    End If

End Sub


' ============================
' Implementation of ITelescope
' ============================

Private Sub ITelescope_AbortSlew()
    AbortSlew
End Sub

Private Property Get ITelescope_AlignmentMode() As AlignmentModes
    ITelescope_AlignmentMode = AlignmentMode
End Property

Private Property Get ITelescope_Altitude() As Double
    ITelescope_Altitude = altitude
End Property

Private Property Get ITelescope_ApertureArea() As Double
    ITelescope_ApertureArea = ApertureArea
End Property

Private Property Get ITelescope_ApertureDiameter() As Double
    ITelescope_ApertureDiameter = ApertureDiameter
End Property

Private Property Get ITelescope_AtHome() As Boolean
    ITelescope_AtHome = AtHome
End Property

Private Property Get ITelescope_AtPark() As Boolean
    ITelescope_AtPark = AtPark
End Property

Private Function ITelescope_AxisRates(ByVal axis As TelescopeAxes) As IAxisRates
    '
    ' Note that this more or less "casts" our internal AxisRates
    ' object's interface to AxisRates.
    '
    Set ITelescope_AxisRates = AxisRates(axis)
End Function

Private Property Get ITelescope_Azimuth() As Double
    ITelescope_Azimuth = Azimuth
End Property

Private Property Get ITelescope_CanFindHome() As Boolean
    ITelescope_CanFindHome = CanFindHome
End Property

Private Function ITelescope_CanMoveAxis(ByVal axis As TelescopeAxes) As Boolean
    ITelescope_CanMoveAxis = CanMoveAxis(axis)
End Function

Private Property Get ITelescope_CanPark() As Boolean
    ITelescope_CanPark = CanPark
End Property

Private Property Get ITelescope_CanPulseGuide() As Boolean
    ITelescope_CanPulseGuide = CanPulseGuide
End Property

Private Property Get ITelescope_CanSetDeclinationRate() As Boolean
    ITelescope_CanSetDeclinationRate = CanSetDeclinationRate
End Property

Private Property Get ITelescope_CanSetGuideRates() As Boolean
    ITelescope_CanSetGuideRates = CanSetGuideRates
End Property

Private Property Get ITelescope_CanSetPark() As Boolean
    ITelescope_CanSetPark = CanSetPark
End Property

Private Property Get ITelescope_CanSetPierSide() As Boolean
    ITelescope_CanSetPierSide = CanSetPierSide
End Property

Private Property Get ITelescope_CanSetRightAscensionRate() As Boolean
    ITelescope_CanSetRightAscensionRate = CanSetRightAscensionRate
End Property

Private Property Get ITelescope_CanSetTracking() As Boolean
    ITelescope_CanSetTracking = CanSetTracking
End Property

Private Property Get ITelescope_CanSlew() As Boolean
    ITelescope_CanSlew = CanSlew
End Property

Private Property Get ITelescope_CanSlewAltAz() As Boolean
    ITelescope_CanSlewAltAz = CanSlewAltAz
End Property

Private Property Get ITelescope_CanSlewAltAzAsync() As Boolean
    ITelescope_CanSlewAltAzAsync = CanSlewAltAzAsync
End Property

Private Property Get ITelescope_CanSlewAsync() As Boolean
    ITelescope_CanSlewAsync = CanSlewAsync
End Property

Private Property Get ITelescope_CanSync() As Boolean
    ITelescope_CanSync = CanSync
End Property

Private Property Get ITelescope_CanSyncAltAz() As Boolean
    ITelescope_CanSyncAltAz = CanSyncAltAz
End Property

Private Property Get ITelescope_CanUnpark() As Boolean
    ITelescope_CanUnpark = CanUnpark
End Property

Private Sub ITelescope_CommandBlind(ByVal Command As String, Optional ByVal Raw As Boolean = False)
'    CommandBlind Command, Raw
    CommandBlind Command
End Sub

Private Function ITelescope_CommandBool(ByVal Command As String, Optional ByVal Raw As Boolean = False) As Boolean
'    ITelescope_CommandBool = CommandBool(Command, Raw)
    ITelescope_CommandBool = CommandBool(Command)
End Function

Private Function ITelescope_CommandString(ByVal Command As String, Optional ByVal Raw As Boolean = False) As String
    ITelescope_CommandString = CommandString(Command, Raw)
End Function

Private Property Let ITelescope_Connected(ByVal RHS As Boolean)
    Connected = RHS
End Property

Private Property Get ITelescope_Connected() As Boolean
    ITelescope_Connected = Connected
End Property

Private Property Get ITelescope_Declination() As Double
    ITelescope_Declination = Declination
End Property

Private Property Let ITelescope_DeclinationRate(ByVal RHS As Double)
    DeclinationRate = RHS
End Property

Private Property Get ITelescope_DeclinationRate() As Double
    ITelescope_DeclinationRate = DeclinationRate
End Property

Private Property Get ITelescope_Description() As String
    ITelescope_Description = Description
End Property

Private Function ITelescope_DestinationSideOfPier(ByVal RightAscension As Double, ByVal Declination As Double) As PierSide
    ITelescope_DestinationSideOfPier = DestinationSideOfPier(RightAscension, Declination)
End Function

Private Property Let ITelescope_DoesRefraction(ByVal RHS As Boolean)
    DoesRefraction = RHS
End Property

Private Property Get ITelescope_DoesRefraction() As Boolean
    ITelescope_DoesRefraction = DoesRefraction
End Property

Private Property Get ITelescope_DriverInfo() As String
    ITelescope_DriverInfo = DriverInfo
End Property

Private Property Get ITelescope_DriverVersion() As String
    ITelescope_DriverVersion = DriverVersion
End Property

Private Property Get ITelescope_EquatorialSystem() As EquatorialCoordinateType
    ITelescope_EquatorialSystem = EquatorialSystem
End Property

Private Sub ITelescope_FindHome()
    FindHome
End Sub

Private Property Get ITelescope_FocalLength() As Double
    ITelescope_FocalLength = FocalLength
End Property

Private Property Let ITelescope_GuideRateDeclination(ByVal RHS As Double)
    GuideRateDeclination = RHS
End Property

Private Property Get ITelescope_GuideRateDeclination() As Double
    ITelescope_GuideRateDeclination = GuideRateDeclination
End Property

Private Property Let ITelescope_GuideRateRightAscension(ByVal RHS As Double)
    GuideRateRightAscension = RHS
End Property

Private Property Get ITelescope_GuideRateRightAscension() As Double
    ITelescope_GuideRateRightAscension = GuideRateRightAscension
End Property

Private Property Get ITelescope_InterfaceVersion() As Integer
    ITelescope_InterfaceVersion = InterfaceVersion
End Property

Private Property Get ITelescope_IsPulseGuiding() As Boolean
    ITelescope_IsPulseGuiding = IsPulseGuiding
End Property

Private Sub ITelescope_MoveAxis(ByVal axis As TelescopeAxes, ByVal rate As Double)
    MoveAxis axis, rate
End Sub

Private Property Get ITelescope_Name() As String
    ITelescope_Name = "I_" & name
End Property

Private Sub ITelescope_Park()
    Park
End Sub

Private Sub ITelescope_PulseGuide(ByVal direction As GuideDirections, ByVal Duration As Long)
    PulseGuide direction, Duration
End Sub

Private Property Get ITelescope_RightAscension() As Double
    ITelescope_RightAscension = RightAscension
End Property

Private Property Let ITelescope_RightAscensionRate(ByVal RHS As Double)
    RightAscensionRate = RHS
End Property

Private Property Get ITelescope_RightAscensionRate() As Double
    ITelescope_RightAscensionRate = RightAscensionRate
End Property

Private Sub ITelescope_SetPark()
    SetPark
End Sub

Private Sub ITelescope_SetupDialog()
    SetupDialog
End Sub

Private Property Let ITelescope_SideOfPier(ByVal RHS As PierSide)
    SideOfPier = RHS
End Property

Private Property Get ITelescope_SideOfPier() As PierSide
    ITelescope_SideOfPier = SideOfPier
End Property

Private Property Get ITelescope_SiderealTime() As Double
    ITelescope_SiderealTime = SiderealTime
End Property

Private Property Let ITelescope_SiteElevation(ByVal RHS As Double)
    SiteElevation = RHS
End Property

Private Property Get ITelescope_SiteElevation() As Double
    ITelescope_SiteElevation = SiteElevation
End Property

Private Property Let ITelescope_SiteLatitude(ByVal RHS As Double)
    SiteLatitude = RHS
End Property

Private Property Get ITelescope_SiteLatitude() As Double
    ITelescope_SiteLatitude = SiteLatitude
End Property

Private Property Let ITelescope_SiteLongitude(ByVal RHS As Double)
    SiteLongitude = RHS
End Property

Private Property Get ITelescope_SiteLongitude() As Double
    ITelescope_SiteLongitude = SiteLongitude
End Property

Private Property Get ITelescope_Slewing() As Boolean
    ITelescope_Slewing = Slewing
End Property

Private Property Let ITelescope_SlewSettleTime(ByVal RHS As Integer)
    SlewSettleTime = RHS
End Property

Private Property Get ITelescope_SlewSettleTime() As Integer
    ITelescope_SlewSettleTime = SlewSettleTime
End Property

Private Sub ITelescope_SlewToAltAz(ByVal Azimuth As Double, ByVal altitude As Double)
    SlewToAltAz Azimuth, altitude
End Sub

Private Sub ITelescope_SlewToAltAzAsync(ByVal Azimuth As Double, ByVal altitude As Double)
    SlewToAltAzAsync Azimuth, altitude
End Sub

Private Sub ITelescope_SlewToCoordinates(ByVal RightAscension As Double, ByVal Declination As Double)
    SlewToCoordinates RightAscension, Declination
End Sub

Private Sub ITelescope_SlewToCoordinatesAsync(ByVal RightAscension As Double, ByVal Declination As Double)
    SlewToCoordinatesAsync RightAscension, Declination
End Sub

Private Sub ITelescope_SlewToTarget()
    SlewToTarget
End Sub

Private Sub ITelescope_SlewToTargetAsync()
    SlewToTargetAsync
End Sub

Private Sub ITelescope_SyncToAltAz(ByVal Azimuth As Double, ByVal altitude As Double)
    SyncToAltAz Azimuth, altitude
End Sub

Private Sub ITelescope_SyncToCoordinates(ByVal RightAscension As Double, ByVal Declination As Double)
    SyncToCoordinates RightAscension, Declination
End Sub

Private Sub ITelescope_SyncToTarget()
    SyncToTarget
End Sub

Private Property Let ITelescope_TargetDeclination(ByVal RHS As Double)
    TargetDeclination = RHS
End Property

Private Property Get ITelescope_TargetDeclination() As Double
    ITelescope_TargetDeclination = TargetDeclination
End Property

Private Property Let ITelescope_TargetRightAscension(ByVal RHS As Double)
    TargetRightAscension = RHS
End Property

Private Property Get ITelescope_TargetRightAscension() As Double
    ITelescope_TargetRightAscension = TargetRightAscension
End Property

Private Property Let ITelescope_Tracking(ByVal RHS As Boolean)
    Tracking = RHS
End Property

Private Property Get ITelescope_Tracking() As Boolean
    ITelescope_Tracking = Tracking
End Property

Private Property Let ITelescope_TrackingRate(ByVal RHS As DriveRates)
    TrackingRate = RHS
End Property

Private Property Get ITelescope_TrackingRate() As DriveRates
    ITelescope_TrackingRate = TrackingRate
End Property

Private Property Get ITelescope_TrackingRates() As ITrackingRates
    '
    ' Note that this more or less "casts" our internal TrackingRates
    ' object's interface to ITrackingRates.
    '
    Set ITelescope_TrackingRates = TrackingRates
End Property

Private Sub ITelescope_Unpark()
    UnPark
End Sub

Private Property Let ITelescope_UTCDate(ByVal RHS As Date)
    UTCDate = RHS
End Property

Private Property Get ITelescope_UTCDate() As Date
    ITelescope_UTCDate = UTCDate
End Property


