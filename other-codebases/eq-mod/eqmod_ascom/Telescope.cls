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
' When      Who     What
' --------- ---     --------------------------------------------------
' 27-Jun-00 rbd     Initial edit, for template project
' 21-Aug-00 rbd     Changed AlignmentModes enum to use 'alg' prefix
'                   and added algUnknown. Remove serial form and
'                   control. Add reference to DriverHelper for
'                   serial and utility stuff.
' 14-Oct-00 rbd     SlewSettleTime property
' 12-Nov-00 rbd     ByVal for all parameters and Lets - for JScript
' 13-Nov-00 rbd     Name and Description return String not Double,
'                   Park() does not take a command.
' 19-Nov-00 rbd     DriverInfo is a String.
' 21-Nov-00 rbd     Standardize DriverInfo to use VERS resource data.
' 13-Jan-01 rbd     Fill in missing ERR_SOURCE argument and specifics
'                   for calls to Err.Raise for "not implemented".
' 21-Jan-00 rbd     New Helper.Profile object, self-registration.
' 25-Jan-01 rbd     New properties CanSlew, CanSlewAsync, CanSync,
'                   CanPark, CanFindHome, CanSetPark, CanUnpark
'                   Interface V1.1
' 06-Feb-01 rbd     Add Tracking, CanSetTracking
' 08-Jun-01 rbd     New Profile.DeviceType to generalize Chooser and
'                   Profile.
' 08-Jun-01 rbd     Telescope 1.4: Add PulseGuide()
' 05-Oct-01 rbd     Add missing ByVals
' 27-Jul-02 rbd     ASCOm Platform 2.0
' 03-Sep-03 rbd     ASCOM Platform 3.0
' 24-Nov-04 rbd     ASCOM Platform 4.0 (Still V1 interface :-(( )
' 04-Nov-06 rcs     ASCOM EQMOD Driver modification
' 20-Nov-06 rcs     ASCOM Fix Elevation Registry Bug
' 01-Dec-06 rcs     Add GPS Function and Autoguider port save
' 01-Dec-06 rcs     Fix Autoguider port restore bug
' 11-Jan-07 rcs     Handle 0 duration errors
' 30-Jan-07 rcs     Activate SlewSettletime
' 01-Feb-07 rcs     Added Custom rate tracking
' 19-Mar-07 rcs     Initial Edit for Three star alignment
' 15-Aug-07 cs      Client Counting
' 30-Aug-07 cs      SyncRAMotor, SyncDECMotor added.
' 05-Oct-07 cs      PecMoveAxis added.
'---------------------------------------------------------------------

Option Explicit

Public Enum AlignmentModes
    algUnknown = -1
    algAltAz = 0
    algPolar = 1
    algGermanPolar = 2
End Enum

Public Enum GuideDirections
    guideNorth = 0
    guideSouth = 1
    guideEast = 2
    guideWest = 3
End Enum


Public Enum EquatorialCoordinateType
    equother = 0
    equLocalTopocentric = 1
    equJ2000 = 2
    equJ2050 = 3
    equB1950 = 4
End Enum

Public Enum DriveRates
    driveSidereal = 0
    driveLunar = 1
    driveSolar = 2
    driveKing = 3
End Enum

Public Enum AxisID
    axisPrimary = 0
    axisSecondary = 1
    axisTertiary = 2
End Enum

Public Enum PierSide
    pierUnknown = -1
    PierEast = 0
    PierWest = 1
End Enum


'Private Const id As String = "EQMOD.Telescope"
'Private Const DESC As String = "EQMOD ASCOM Scope Driver"

Private m_Serial As DriverHelper.Serial
Private m_Util As DriverHelper.Util
Private m_Profile As DriverHelper.Profile
Private m_iSettleTime As Integer
Private m_RaRateAdjust As Double
Private m_DecRateAdjust As Double

Public g_TrackingRatesFull As TrackingRates ' Collection of supported drive rates


Private Sub Class_Initialize()

    If ClientCount = 0 Then
    
        eqres = EQ_End()
    
        Set m_Serial = New DriverHelper.Serial
        Set m_Util = New DriverHelper.Util
        Set m_Profile = New DriverHelper.Profile
        m_Profile.DeviceType = "Telescope"          ' We're a Telescope driver
        m_iSettleTime = 0                           ' Default 0 slew settle time
        m_Profile.Register ASCOM_id, ASCOM_DESC     ' Self-register
        
        Set g_TrackingRatesFull = New TrackingRates
        g_TrackingRatesFull.Add driveSidereal
        g_TrackingRatesFull.Add driveLunar
        g_TrackingRatesFull.Add driveSolar
        
        ' initialise a rates colection for each axis but don't assign any rates yet
        ' this will be done on connection
        Set g_RAAxisRates = New Rates
        Set g_DECAxisRates = New Rates
        
        gHemisphere = 1
        gSidFlag = 1
        gTargetRA = EQ_INVALIDCOORDINATE
        gTargetDec = EQ_INVALIDCOORDINATE
        gEQTimeDelta = 0
        
        gTot_step = gDefault_step
        gEQ_MAXSYNC = EQ_MAXSYNC_Const
            
        
        'Initialize position values - They are actually done also on
        'Driver connect
        Call readAscomCompatibiity
        Call readSiderealRate
        Call ReadSiteValues
        Call ReadSyncMap
        Call ReadAlignMap
        Call readUserParkPos
    
        
        gEQparkstatus = 0
        gTrackingStatus = 0                        ' Initially not tracking
        gDeclinationRate = 0
        m_DecRateAdjust = 0
        gRightAscensionRate = 0
        m_RaRateAdjust = 0
        
        Call readRALimit
        If gRA_Limit_East <> 0 Or gRA_Limit_West <> 0 Then
            ' Limiting is crrently enabled - use default limit until unpark
            ' on unpark custom limits are applied.
            gRA_Limit_East = &H5D5500                  ' Set RA Limit West Side
            gRA_Limit_West = &HA2AB00                  ' Set RA Limit East Side
        End If
        
        Call Limits_Init
        
        ' At least with an initial value
        
        gDECEncoder_Home_pos = DECEncoder_Home_pos_const
        
        
        gRAStatus = EQ_MOTORBUSY                   ' RA Motor Busy Status
        gDECStatus = EQ_MOTORBUSY                  ' DEC Motor Busy Status
        gSlewStatus = False                        ' Not Slewing status
        gRAStatus_slew = False                     ' Slew to track condidition
        gSlewCount = 0                             ' Goto Iterative Counter
        
        
        ' Initialize these values for polling emulation
        
        gEmulRA = 0
        gEmulDEC = 0
        gEmulOneShot = False
        gEmulNudge = False
        
        
        ReadComPortSettings
        
      
        gTot_RA = gTot_step                        ' Set RA Total Encoder Step Count
        gTot_DEC = gTot_step                       ' Set DEC total Encoder Step Count
        gRAWormSteps = 50133
        gDECWormSteps = 50133
        gRAWormPeriod = 480
        gSOP = 0                                    'Set at unknown pier setting
    
    
        HC.HCMessage.Text = ""                           'Set Message Center to BLANK
        HC.HCTextAlign.Text = ""
     
        HC.DisplayTimer.Enabled = True
    End If
    
    ClientCount = ClientCount + 1
    HC.Add_Message (HC.oLangDll.GetLangString(5134) & CStr(ClientCount - 1))
    
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
        writeDriftVal
        writeRASyncCheckVal
        writePulseguidepwidth
        
        
        Unload HC
        Unload Align
        Unload Slewpad
        Unload GPSSetup
        Unload NStar_debug
        Unload StarEditform
        Unload JStickConfigForm
        Unload AscomTrace
        Unload LimitEditForm
        Unload polarfrm
        Unload DefineParkForm
        ClientCount = 1
    Else
        ClientCount = ClientCount - 1
        HC.Add_Message (HC.oLangDll.GetLangString(5134) & CStr(ClientCount - 1))
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

    If gEQparkstatus = 1 Then
        ' if parkered
        AtPark = True
    Else
        ' if unparked or parking
        AtPark = False
    End If
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, "GET AtPark :" & AtPark
    
End Property

'Later this will be changed to support the property

Public Property Get CanSetDeclinationRate() As Boolean
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 3, "GET CanSetDeclination :T"
    CanSetDeclinationRate = True
End Property

'Later this will be changed to support the property
Public Property Get CanSetGuideRates() As Boolean
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 3, "GET CanSetGuideRates :F"
    CanSetGuideRates = False
End Property


Public Property Get CanSetPierSide() As Boolean
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 3, "GET CanSetGuideRates :F"
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
    EquatorialSystem = equother
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, "GET EquatorialCoordinateType :" & equother
End Property

Public Property Get GuideRateDeclination() As Double
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 4, "GET GuideRateDEC :" & "NOT_SUPPORTED"
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
              "Property Get GuideRateDeclination" & MSG_NOT_IMPLEMENTED
End Property

Public Property Let GuideRateDeclination(ByVal newval As Double)
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 4, "LET GuideRateDEC(" & newval & ") :NOT_SUPPORTED"
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
              "Property Let GuideRateDeclination" & MSG_NOT_IMPLEMENTED
End Property

Public Property Get GuideRateRightAscension() As Double
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 4, "GET GuideRateRA :NOT_SUPPORTED"
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
              "Property Get GuideRateRightAscension" & MSG_NOT_IMPLEMENTED

End Property


Public Property Let GuideRateRightAscension(ByVal newval As Double)
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 4, "LET GuideRateRA(" & newval & ") :NOT_SUPPORTED"
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
              "Property Let GuideRateRightAscension" & MSG_NOT_IMPLEMENTED
End Property


Public Property Get InterfaceVersion() As Integer
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, "GET InterfaceVersion=2"
    InterfaceVersion = 2
End Property


Public Property Let SideOfPier(ByVal newval As PierSide)
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, "LET SideOfPier(" & newval & ") :NOT_SUPPORTED"
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
              "Property Let SideOfPier" & MSG_NOT_IMPLEMENTED
End Property
Public Property Get SideOfPier() As PierSide
    SideOfPier = gSideofPier
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, "GET SideOfPier :" & gSideofPier
'     RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Property Get SideOfPier" & MSG_NOT_IMPLEMENTED
End Property

Public Property Let TrackingRate(ByVal newval As DriveRates)
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 5, "LET TrackingRate :" & newval & ") :NOT_SUPPORTED"
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
              "Property Let TrackingRate" & MSG_NOT_IMPLEMENTED
End Property

Public Property Get TrackingRate() As DriveRates
    TrackingRate = driveSidereal
'    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
              "Property Get TrackingRate" & MSG_NOT_IMPLEMENTED
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 5, ("GET TrackingRate :" & TrackingRate)
End Property

Public Property Let TrackingRates(ByVal newval As Object)
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 5, "LET TrackingRates :NOT_SUPPORTED"
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
              "Property Let TrackingRates" & MSG_NOT_IMPLEMENTED
End Property

Public Property Get TrackingRates() As Object
    Set TrackingRates = g_TrackingRatesFull
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 5, ("GET Tracking Rates")
End Property

Public Function AxisRates(axis As AxisID) As Object
    
    Select Case axis
        Case axisPrimary
            Set AxisRates = g_RAAxisRates
        Case axisSecondary
            Set AxisRates = g_DECAxisRates
        Case Else
            Set AxisRates = New Rates
    End Select
    
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 6, ("AxisRates(" & axis & ")")
'    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE,"Method AxisRates()" & MSG_NOT_IMPLEMENTED
End Function

Public Function CanMoveAxis(axis As AxisID) As Boolean
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
        DestinationSideOfPier = PierEast 'looking West
    Else
        DestinationSideOfPier = PierWest 'looking East
    End If
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("DestinationSideOfPier(" & destRa & "," & destDec & ")")
'    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Method DestinationSideOfPier()" & MSG_NOT_IMPLEMENTED
End Function

Public Sub MoveAxis(axis As AxisID, rate As Double)
    
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

' special sub for PEC control - normal moveaxis results in background polling of motor positions
Public Sub PecMoveAxis(axis As AxisID, rate As Double)

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
    MoveMotor = EQ_StartMoveMotor(motor_id, hemisphere, direction, Steps, stepslowdown)
End Function

Public Sub SlewToAltAz(ByVal Azimuth As Double, _
                        ByVal altitude As Double)
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 6, ("SlewToAltAz(" & Azimuth & "," & altitude & ")")
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
              "Method SlewToAltAz()" & MSG_NOT_IMPLEMENTED
End Sub


Public Sub SlewToAltAzAsync(ByVal Azimuth As Double, _
                            ByVal altitude As Double)
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 6, ("SlewToAltAzAsync(" & Azimuth & "," & altitude & ")")
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
              "Method SlewToAltAzAsync()" & MSG_NOT_IMPLEMENTED
End Sub
Public Sub SyncToAltAz(ByVal Azimuth As Double, _
                        ByVal altitude As Double)
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 7, ("SyncToAltAz(" & Azimuth & "," & altitude & ")")
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
              "Method SyncToAltAz()" & MSG_NOT_IMPLEMENTED
End Sub


'-------------------------- V1 Properties Starts here -----------------

Public Property Get DoesRefraction() As Boolean
    DoesRefraction = False
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("GET DoesRfraction :F")
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
              "Property Get DoesRefraction" & MSG_NOT_IMPLEMENTED
End Property

Public Property Let DoesRefraction(ByVal newval As Boolean)
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("SET DoesRfraction :NOT_SUPPORTED")
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
              "Property Get DoesRefraction" & MSG_NOT_IMPLEMENTED

End Property

Public Property Get altitude() As Double

    altitude = gAlt
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 2, ("GET altitude :" & gAlt)
    
End Property

Public Property Get ApertureDiameter() As Double

    '** insert your implementation here and remove the following line
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("GET ApertureDiameter :NOT_SUPPORTED")
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property ApertureDiameter" & MSG_NOT_IMPLEMENTED

End Property

Public Property Get ApertureArea() As Double

    '** insert your implementation here and remove the following line
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("GET ApertureArea :NOT_SUPPORTED")
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property ApertureDiameter" & MSG_NOT_IMPLEMENTED

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
      Connected = True
  End If
  If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 1, ("GET Connected :" & Connected)
    
    
End Property

Public Property Let Connected(ByVal newval As Boolean)

    
Dim icast As Long
Dim MaxRate As Double
Dim strtmp As String

    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 1, ("LET Connected(" & newval & ")")
    If newval Then
    
    
'        Call ReadSiteValues
'        Call ReadSyncMap
'        Call ReadAlignMap
    
        If EQ_GetMountStatus() = 0 Then
        
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
            gSidFlag = 1
            gpl_interval = 50
        
            HC.Show
            
            ' prefix the com port to ensure com10+ works
            strtmp = "\\.\" & gPort
            
            gInitResult = EQ_Init(strtmp, gBaud, gTimeout, gRetry)
            
            ' Make sure we have the right board
            Call readCustomMount
            eqres = CheckMount(gInitResult)
        
            If eqres = EQ_OK And EQ_GetMountStatus() = 1 Then
                Call readCustomMount
                eqres = EQGetTotal360microstep(0)
                If eqres < &H1000000 Then
                    gTot_RA = eqres
                    gRAWormSteps = EQGP(0, 10006)
                    HC.Add_Message CStr(gRAWormSteps) & " RAWormSteps read"
                    If gCustomMount = 0 Then
                        Select Case gRAWormSteps
                            Case 50133
                            Case 66844
                            Case 35200
                            Case 31288
                            Case Else
                                ' mount isn't returning known values - read from ini instead
                                Call readWormSteps
                        End Select
                    End If
                    gRAWormPeriod = Int((SPSD * gRAWormSteps / gTot_RA) + 0.5)
                End If
                
                eqres = EQGetTotal360microstep(1)
                If eqres < &H1000000 Then
                    gTot_DEC = eqres
                    gDECWormSteps = EQGP(1, 10006)
                    HC.Add_Message CStr(gDECWormSteps) & " DECWormSteps read"
                    If gCustomMount = 0 Then
                        Select Case gDECWormSteps
                            Case 50133
                            Case 66844
                            Case 35200
                            Case 31288
                            Case Else
                                ' mount isn't returning known values - read from ini instead
                                Call readWormSteps
                        End Select
                    End If
                End If
                
                'Make sure motors are not running
                eqres = EQ_MotorStop(0)     ' Stop RA Motor
                eqres = EQ_MotorStop(1)     ' Stop DEC Motor
                
                HC.TrackingFrame.Caption = HC.oLangDll.GetLangString(121) & " " & HC.oLangDll.GetLangString(178)
                
                'Get state of at least one of the motors
                
                eqres = EQ_GetMotorStatus(0)
                
                ' If its an error then Initialize it
                
                If eqres = EQ_NOTINITIALIZED Then
                    icast = gDECEncoder_Home_pos ' Typecast
                    eqres = EQ_InitMotors(RAEncoder_Home_pos, icast)
                End If
                
                gMount_Ver = EQ_GetMountVersion()
                
                ' set up rates collection
                MaxRate = SID_RATE * 800 / 3600
                g_RAAxisRates.Add MaxRate, 0#
                g_DECAxisRates.Add MaxRate, 0#

                            
                'Make sure we get the latest data from the registry
               
               
                HC.Add_Message (HC.oLangDll.GetLangString(5132) & " " & gPort & ":" & str(gBaud))
                HC.Add_Message (HC.oLangDll.GetLangString(5133) & " " & printhex(EQ_GetMountVersion()) & " DLL Version:" & printhex(EQ_DriverVersion()))
                HC.Add_Message "Using " & CStr(gRAWormSteps) & "RAWormSteps"
                HC.EncoderTimer.Enabled = True
                HC.EncoderTimerFlag = True
                gEQPulsetimerflag = True
                HC.Pulseguide_Timer.Enabled = False     'Enabled only during pulseguide session
                
                Call readParkModes

                gEQparkstatus = readparkStatus()
    
                If gEQparkstatus = 1 Then
                    HC.Frame15.Caption = HC.oLangDll.GetLangString(146) & " " & HC.oLangDll.GetLangString(177)
                Else
                    HC.Frame15.Caption = HC.oLangDll.GetLangString(146) & " " & HC.oLangDll.GetLangString(179)
                End If
                Call SetParkCaption

    
                Call readportrate ' Read Autoguider port settings from registry and send to mount
                Call readRASyncCheckVal ' RA Sync Auto
                Call readDriftVal ' Read the Drift offset value
                Call readPulseguidepwidth ' Read Pulseguide interval
                Call PEC_Initialise   ' only initialise PEc when we've defaults for worm

            Else
                HC.EncoderTimer.Enabled = False
                HC.EncoderTimerFlag = False
                eqres = EQ_End()
                HC.Add_Message (HC.oLangDll.GetLangString(5135) & " " & gPort & ":" & str(gBaud))
            
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
           
           ' Added to save settings
           
            writeDriftVal
            writeRASyncCheckVal
            writePulseguidepwidth
            
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
            HC.TrackingFrame.Caption = HC.oLangDll.GetLangString(121) & " " & HC.oLangDll.GetLangString(189)
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

    '** insert your implementation here and remove the following line
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("GET FocalLength :NOT_SUPPORTED")
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property FocalLength" & MSG_NOT_IMPLEMENTED

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

    RightAscensionRate = gRightAscensionRate
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
            HC.TrackingFrame.Caption = HC.oLangDll.GetLangString(121) & " " & HC.oLangDll.GetLangString(189)
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

    '** insert your implementation here and remove the following line
  If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("SET SiteElevation :NOT_SUPPORTED")
  RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property SiteElevation" & MSG_NOT_IMPLEMENTED

End Property

Public Property Get SiteLatitude() As Double

    SiteLatitude = gLatitude
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("GET SiteLatitude :" & SiteLatitude)

End Property

Public Property Let SiteLatitude(ByVal newval As Double)

    '** insert your implementation here and remove the following line
  If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("SET SiteLatitude :NOT_SUPPORTED")
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property SiteLatitude" & MSG_NOT_IMPLEMENTED

End Property

Public Property Get SiteLongitude() As Double

    SiteLongitude = gLongitude
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("GET SiteLongitude :" & SiteLongitude)

End Property

Public Property Let SiteLongitude(ByVal newval As Double)

    '** insert your implementation here and remove the following line
  If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("SET SiteLongitude :NOT_SUPPORTED")
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property SiteLongitude" & MSG_NOT_IMPLEMENTED

End Property

Public Property Get Slewing() As Boolean

    Slewing = gSlewStatus
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 6, ("GET Slewing :" & Slewing)

End Property

Public Property Get SlewSettleTime() As Integer

    SlewSettleTime = m_iSettleTime
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 6, ("GET SlewSettleTime :" & SlewSettleTime)
    
End Property

Public Property Let SlewSettleTime(ByVal newval As Integer)

    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 6, ("LET SlewSettleTime(" & newval & ")")
    If (newval < 0) Or (newval > 100) Then _
        RaiseError SCODE_VAL_OUTOFRANGE, ERR_SOURCE, MSG_VAL_OUTOFRANGE
        
    m_iSettleTime = newval
    
End Property

Public Property Get TargetDeclination() As Double

    If gTargetDec = EQ_INVALIDCOORDINATE Then _
        RaiseError SCODE_PROP_NOT_SET, ERR_SOURCE, _
            "Property TargetDeclination " & MSG_PROP_NOT_SET

    TargetDeclination = gTargetDec
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 6, ("GET TargetDeclination :" & TargetDeclination)

End Property

Public Property Let TargetDeclination(ByVal newval As Double)

    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 6, ("Let TargetDeclination(" & newval & ")")
    If newval > 90 Or newval < -90 Then
        RaiseError SCODE_VAL_OUTOFRANGE, ERR_SOURCE, _
            "Property TargetDeclination " & MSG_VAL_OUTOFRANGE
    Else
        gTargetDec = newval
    End If
    
End Property

Public Property Get TargetRightAscension() As Double

    If gTargetRA = EQ_INVALIDCOORDINATE Then _
        RaiseError SCODE_PROP_NOT_SET, ERR_SOURCE, _
            "Property TargetRightAscension " & MSG_PROP_NOT_SET


    TargetRightAscension = gTargetRA
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 6, ("GET TargetRightAscension :" & TargetRightAscension)

End Property

Public Property Let TargetRightAscension(ByVal newval As Double)

    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 6, ("Let TargetRightAscension(" & newval & ")")
    If newval > 24 Or newval < 0 Then
        RaiseError SCODE_VAL_OUTOFRANGE, ERR_SOURCE, _
            "Property TargetRightAscension " & MSG_VAL_OUTOFRANGE
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
               ' track at sidreal
                Call EQStartSidereal
                gEmulOneShot = True                 ' Get One shot cap
            Else
                ' track at custom rate
                gRA_LastRate = 0
                gDeclinationRate = m_DecRateAdjust
                gRightAscensionRate = SID_RATE + m_RaRateAdjust
                If PEC_Enabled Then PEC_StopTracking
                Call CustomMoveAxis(0, gRightAscensionRate)
                Call CustomMoveAxis(1, gDeclinationRate)
            End If
        
        Else
        
            eqres = EQ_MotorStop(0)
            eqres = EQ_MotorStop(1)
            gTrackingStatus = 0
            ' not sure that we should be clearing the rate offests ASCOM Spec is no help
            gDeclinationRate = 0
            gRightAscensionRate = 0
            
            HC.TrackingFrame.Caption = HC.oLangDll.GetLangString(121) & " " & HC.oLangDll.GetLangString(178)
            
        End If
    Else
    
        HC.Add_Message (HC.oLangDll.GetLangString(5013))
    
    End If

End Property

Public Property Get UTCDate() As Date

    UTCDate = CDate(CDbl(Now()) + (CDbl(utc_offs()) / 86400#))
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("GET Date :" & UTCDate)
        
End Property

Public Property Let UTCDate(ByVal newval As Date)

     'Impossible to set own PC time
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("SET Date :NOT_SUPPORTED")

     RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property UTCDate" & MSG_NOT_IMPLEMENTED

End Property

Public Property Get ramotor() As Long

    'RAMotor = EQ_GetMotorValues(0)
     ramotor = GetEmulRA()
        
End Property

Public Property Get DECMotor() As Long

    'DECMotor = EQ_GetMotorValues(1)
    DECMotor = gEmulDEC
        
End Property

Public Property Get SyncRAMotor() As Long

     gEmulRA = EQ_GetMotorValues(0)
     SyncRAMotor = gEmulRA
    
     gLast_time = EQnow_lst_norange()
     gCurrent_time = gLast_time
     gEmulRA_Init = gEmulRA
        
End Property

Public Property Get SyncDECMotor() As Long

     gEmulDEC = EQ_GetMotorValues(1)
     SyncDECMotor = gEmulDEC
        
End Property
Public Property Get RAWormPeriod() As Long
     RAWormPeriod = gRAWormPeriod
End Property
Public Property Get RAWormPeriodFloat() As Double
    RAWormPeriodFloat = SPSD * gRAWormSteps / gTot_RA
End Property

Public Property Get PecGain() As Double
    PecGain = gPecGain
End Property

Public Property Let PecGain(ByVal newval As Double)
    HC.GainScroll.Value = newval * 10
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
        ' stop the slew if already slewing
        HC.GotoTimer.Enabled = False
        eqres = EQ_MotorStop(0)
        eqres = EQ_MotorStop(1)
        gSlewStatus = False
        gRAStatus_slew = False
    
        ' restart tracking
        RestartTracking

    End If

End Sub

Public Sub CommandBlind(ByVal Command As String, Optional ByVal Raw As Boolean = False)

    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("COMMAND CommandBlind(" & Command & " :NOT_SUPPORTED")
    '** insert your implementation here and remove the following line
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Method xxx()" & MSG_NOT_IMPLEMENTED

End Sub

Public Function CommandBool(ByVal Command As String, Optional ByVal Raw As Boolean = False) As Boolean

    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("COMMAND CommandBool(" & Command & " :NOT_SUPPORTED")
    '** insert your implementation here and remove the following line
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Method CommandBlind()" & MSG_NOT_IMPLEMENTED

End Function

Public Function CommandString(ByVal Command As String, Optional ByVal Raw As Boolean = False) As String

Dim pos As Long
Dim tmpstr As String
Dim tmp2 As String
Dim arg1 As Double
Dim arg2 As Double
Dim arg3 As Double
Dim tmp1 As Integer

    Select Case Command
    
        Case ":RA_ENC#"
            If EQ_GetMountStatus() <> 1 Then
                CommandString = FormatNumber(-1, 0, , , 0) & "#"
            Else
                pos = SyncRAMotor()
                CommandString = FormatNumber(pos, 0, , , 0) & "#"
            End If
        
        Case ":DEC_ENC#"
            If EQ_GetMountStatus() <> 1 Then
                CommandString = FormatNumber(-1, 0, , , 0) & "#"
            Else
                pos = SyncDECMotor()
                CommandString = FormatNumber(pos, 0, , , 0) & "#"
            End If
    
        Case ":DRIVERVER#"
            CommandString = gVersion
            
        Case ":DLLVER#"
            pos = EQ_DriverVersion
            tmpstr = Hex$((pos And &HF0) / 16 And &HF) + Hex$(pos And &HF)
            tmpstr = tmpstr & "." & Hex$((pos And &HF000) / 4096 And &HF) + Hex$((pos And &HF00) / 256 And &HF)
            tmpstr = tmpstr & "." & Hex$((pos And &HF00000) / 1048576 And &HF) + Hex$((pos And &HF0000) / 65536 And &HF)
            CommandString = tmpstr
    
        Case ":MOUNTVER#"
            tmpstr = Hex$((gMount_Ver And &HF0) / 16 And &HF) + Hex$(gMount_Ver And &HF)
            tmpstr = tmpstr & "." & Hex$((gMount_Ver And &HF000) / 4096 And &HF) + Hex$((gMount_Ver And &HF00) / 256 And &HF)
            tmpstr = tmpstr & "." & Hex$((gMount_Ver And &HF00000) / 1048576 And &HF) + Hex$((gMount_Ver And &HF0000) / 65536 And &HF)
            CommandString = tmpstr
    
        Case ":PECGAIN#"
            CommandString = FormatNumber(HC.GainScroll.Value, 0, , , 0) & "#"
            
        Case ":PECPHASE#"
            CommandString = FormatNumber(HC.PhaseScroll.Value, 0, , , 0) & "#"
    
        Case ":PECFILE#"
            CommandString = PECDef1.filename & "#"
    
        Case ":PECENA,1#"
            'enable PEC
            HC.CheckPEC.Value = 1
            PEC_StartTracking
        
        Case ":PECENA,0#"
            'disable PEC
            HC.CheckPEC.Value = 0
            PEC_StopTracking
        
        Case ":PECSTA#":
            ' read pec state
            If PEC_Enabled = True Then
                CommandString = "1#"
            Else
                CommandString = "0#"
            End If
        
        Case ":PECWTC#":
            ' read worm tooth count
            If EQ_GetMountStatus() <> 1 Then
                CommandString = FormatNumber(-1, 0, , , 0) & "#"
            Else
                pos = gTot_RA / gRAWormSteps
                CommandString = FormatNumber(pos, 0, , , 0) & "#"
            End If
        
        Case ":PECIDX#"
            ' read worm position
            If EQ_GetMountStatus() <> 1 Then
                CommandString = FormatNumber(-1, 0, , , 0) & "#"
            Else
                pos = NormalisePosition(SyncRAMotor, gRAWormSteps)
                CommandString = FormatNumber(pos, 0, , , 0) & "#"
            End If
    
        Case ":PECINFO#"
            ' read worm
            If EQ_GetMountStatus() <> 1 Then
                CommandString = FormatNumber(-1, 0, , , 0) & "," & FormatNumber(-1, 0, , , 0) & "#"
            Else
                pos = Int(gRAWormSteps - 1)
                CommandString = FormatNumber(Int(gRAWormPeriod), 0, , , 0) & "," & FormatNumber(pos, 0, , , 0) & "#"
            End If
        
        Case Else
            tmpstr = Left(Command, 6)
            Select Case tmpstr
                Case ":PARK,"
                    tmpstr = mid(Command, 7, Len(Command) - 7)
                    tmp1 = val(tmpstr)
                    If tmp1 >= 0 And tmp1 <= 7 Then
                        If tmp1 = 0 Then
                            Call HC.ApplyParkMode
                        Else
                            Call HC.ApplyParkMode2(tmp1 - 1)
                        End If
                    End If
                    If gEQparkstatus <> 0 Then
                        ' mount is parked or is parking
                        CommandString = "1#"
                    Else
                        ' mount is still unparked
                        CommandString = "0#"
                    End If
                    GoTo alldone
            End Select
            
            tmpstr = Left(Command, 8)
            Select Case tmpstr
                Case ":PECSET,"
                    ' Write a PEC table entry
                    tmpstr = mid(Command, 9, Len(Command) - 9)
                    pos = InStr(tmpstr, ",")
                    If pos <> 0 Then
                        tmp2 = Left$(tmpstr, pos - 1)
                        arg1 = CInt(tmp2)
                        tmpstr = Right$(tmpstr, Len(tmpstr) - pos)
                        pos = InStr(tmpstr, ",")
                        If pos <> 0 Then
                            arg2 = CLng(Left$(tmpstr, pos - 1))
                            tmp2 = Right$(tmpstr, Len(tmpstr) - pos)
                            arg3 = CDbl(tmp2) / 1000
                            If PEC_Write_Table(arg1, arg2, arg3) Then
                                CommandString = "1#"
                            Else
                                CommandString = "0#"
                            End If
                        Else
                            RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Method CommandString()" & MSG_NOT_IMPLEMENTED
                        End If
                    Else
                        RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Method CommandString()" & MSG_NOT_IMPLEMENTED
                    End If
                    GoTo alldone
                
                Case ":PECGET,"
                    ' Read a PEC table entry
                    tmpstr = mid(Command, 9, Len(Command) - 9)
                    arg1 = CInt(tmpstr)
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
                    GoTo alldone
                
                Case ":UNPARK,"
                    tmpstr = mid(Command, 9, Len(Command) - 9)
                    tmp1 = val(tmpstr)
                    If tmp1 >= 0 And tmp1 <= 6 Then
                        If tmp1 = 0 Then
                            Call HC.ApplyUnParkMode
                        Else
                            Call HC.ApplyUnParkMode2(tmp1 - 1)
                        End If
                    End If
                    If gEQparkstatus = 0 Then
                        CommandString = "1#"
                    Else
                        CommandString = "0#"
                    End If
                    GoTo alldone
                    
            End Select
            
            tmpstr = Left(Command, 9)
            Select Case tmpstr
                Case ":PECGAIN,"
                    'Update the PEC Gain value
                    tmpstr = mid(Command, 10, Len(Command) - 10)
                    If PEC_SetGain(tmpstr) Then
                        CommandString = "1#"
                    Else
                        CommandString = "0#"
                    End If
                    GoTo alldone
                
                Case ":PECLOAD,"
                    ' Load a PEC file
                    tmpstr = mid(Command, 10, Len(Command) - 10)
                    If PEC_LoadFile(tmpstr) Then
                        CommandString = "1#"
                    Else
                        CommandString = "0#"
                    End If
                    GoTo alldone
                
                Case ":PECSAVE,"
                    ' Save a PEC file
                    tmpstr = mid(Command, 10, Len(Command) - 10)
                    If PEC_SaveFile(tmpstr, PECDef1) Then
                        CommandString = "1#"
                    Else
                        CommandString = "0#"
                    End If
                    GoTo alldone
                    
            End Select
                            
            tmpstr = Left(Command, 10)
            Select Case tmpstr
                Case ":PECPHASE,"
                    tmpstr = mid(Command, 11, Len(Command) - 11)
                    If PEC_SetPhase(tmpstr) Then
                        CommandString = "1#"
                    Else
                        CommandString = "0#"
                    End If
                    GoTo alldone
            End Select
            
            ' unknown command!
            RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "Method CommandString()" & MSG_NOT_IMPLEMENTED
    
    End Select

alldone:
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("COMMAND CommandString()" & Command & " :" & CommandString)
 
End Function

Public Sub FindHome()

    '** insert your implementation here and remove the following line
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 7, ("COMMAND FindHome :NOT_SUPPORTED")
    RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Method FindHome()" & MSG_NOT_IMPLEMENTED

End Sub

Public Sub Park()

    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 7, ("COMMAND Park")
    If gEQparkstatus = 0 Then
        ParktoUserDefine (True)
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

Public Sub PulseGuide(ByVal direction As GuideDirections, _
                                            ByVal Duration As Long)
    
    If Not gAscomCompatibility.AllowPulseGuide Then
        RaiseError SCODE_NOT_IMPLEMENTED, ERR_SOURCE, "PulseGuide() :NOT SUPPORTED"
        Exit Sub
    End If
    
    If gSlewStatus Then
        ' no guiding whilst slewing - makes no sense and means the slew will terminate!
        Exit Sub
    End If
    
    
    'Pulse guide implentation for EQMOD
    'This uses the duration parameter and an asynchronous timer
    'that will decrement the duration count for every time tick
    'It then disables the guiderate upon expiration of the counter
    
    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 4, ("COMMAND PuleGuide(" & direction & "," & Duration & ")")
    HC.Pulseguide_Timer.Enabled = False
    
    If gEQparkstatus <> 0 Then
        ' no move axis if parked or parking!
        If gAscomCompatibility.AllowPulseGuideExceptions Then
            RaiseError SCODE_SCOPE_PARKED, ERR_SOURCE, "Method PulseGuide() " & MSG_SCOPE_PARKED
        End If
        Exit Sub
    End If
    
    
    If (gTrackingStatus = 4) Then
    If Duration = 0 Then
      Select Case (direction)
        Case guideNorth 'DEC+
            If HC.decpulse_enchk.Value = 1 Then
                gEQDECPulseDuration = 0
                Call ChangeDEC_by_Rate(gDeclinationRate)
                If HC.Frame16.Visible = True Then Call Plot_points(1, 0, 0)
             End If
 
        Case guideSouth 'DEC-
            If HC.decpulse_enchk.Value = 1 Then
                gEQDECPulseDuration = 0
                Call ChangeDEC_by_Rate(gDeclinationRate)
                If HC.Frame16.Visible = True Then Call Plot_points(1, 1, 0)
            End If
        Case guideEast  'RA+
            If HC.rapulse_enchk.Value = 1 Then
                If gRA_LastRate = 0 Then
                    eqres = EQ_SendGuideRate(0, gTrackingStatus - 1, 0, 0, gHemisphere, gHemisphere)
                Else
                    ChangeRA_by_Rate (gRightAscensionRate)
                End If
                gEQRAPulseDuration = 0
                If HC.Frame16.Visible = True Then Call Plot_points(0, 0, 0)
            End If
        Case guideWest  'RA-
            If HC.rapulse_enchk.Value = 1 Then
                If gRA_LastRate = 0 Then
                    eqres = EQ_SendGuideRate(0, gTrackingStatus - 1, 0, 0, gHemisphere, gHemisphere)
                Else
                    ChangeRA_by_Rate (gRightAscensionRate)
                End If
                If HC.Frame16.Visible = True Then Call Plot_points(0, 1, 0)
                gEQRAPulseDuration = 0
           End If
        Case Else
            If gAscomCompatibility.AllowPulseGuideExceptions Then
                RaiseError SCODE_VAL_OUTOFRANGE, ERR_SOURCE, _
                    "Pulse Guide: GuideDirections" & MSG_VAL_OUTOFRANGE
            End If
       End Select
    
    Else
    
      Select Case (direction)
        Case guideNorth 'DEC+
            If HC.decpulse_enchk.Value = 1 Then
            
                Call ChangeDEC_by_Rate(gDeclinationRate + (gDeclinationRate * HC.VScroll4.Value * 0.1))
                
                If HC.decfixed_enchk.Value = 0 Then
                    gEQDECPulseDuration = Duration * (HC.HScrollDECWidth.Value / 100)
                Else
                    gEQDECPulseDuration = HC.VScroll6.Value * 100
                End If
                
                If HC.Frame16.Visible = True Then Call Plot_points(1, 0, gEQDECPulseDuration)
            End If
            
        Case guideSouth 'DEC-
            If HC.decpulse_enchk.Value = 1 Then
            
                Call ChangeDEC_by_Rate(gDeclinationRate - (gDeclinationRate * HC.VScroll4.Value * 0.1))
                
                eqres = EQ_SendGuideRate(1, gTrackingStatus - 1, val(HC.VScroll4.Value), 0, 0, 0)
                If HC.decfixed_enchk.Value = 0 Then
                    gEQDECPulseDuration = Duration * (HC.HScrollDECWidth.Value / 100)
                Else
                    gEQDECPulseDuration = HC.VScroll6.Value * 100
                End If
                If HC.Frame16.Visible = True Then Call Plot_points(1, 1, gEQDECPulseDuration)
            End If
        Case guideEast  'RA+
            If HC.rapulse_enchk.Value = 1 Then
            
                If gRA_LastRate = 0 Then
                    eqres = EQ_SendGuideRate(0, gTrackingStatus - 1, val(HC.VScroll3.Value), 1, gHemisphere, gHemisphere)
                Else
                    Call ChangeRA_by_Rate(gRightAscensionRate - (gRightAscensionRate * HC.VScroll3.Value * 0.1))
                End If
              
                If HC.rafixed_enchk.Value = 0 Then
                    gEQRAPulseDuration = Duration * (HC.HScrollRAWidth.Value / 100)
                Else
                    gEQRAPulseDuration = HC.VScroll5.Value * 100
                End If
                
                If HC.Frame16.Visible = True Then Call Plot_points(0, 0, gEQRAPulseDuration)
            End If
        Case guideWest  'RA-
            If HC.rapulse_enchk.Value = 1 Then
                
                 If gRA_LastRate = 0 Then
                    eqres = EQ_SendGuideRate(0, gTrackingStatus - 1, val(HC.VScroll3.Value), 0, gHemisphere, gHemisphere)
                Else
                    Call ChangeRA_by_Rate(gRightAscensionRate + (gRightAscensionRate * HC.VScroll3.Value * 0.1))
                End If

                              
                If HC.rafixed_enchk.Value = 0 Then
                    gEQRAPulseDuration = Duration * (HC.HScrollRAWidth.Value / 100)
                Else
                    gEQRAPulseDuration = HC.VScroll5.Value * 100
                End If
                If HC.Frame16.Visible = True Then Call Plot_points(0, 1, gEQRAPulseDuration)
            End If
        Case Else
            If gAscomCompatibility.AllowPulseGuideExceptions Then
                RaiseError SCODE_VAL_OUTOFRANGE, ERR_SOURCE, _
                    "Pulse Guide: GuideDirections" & MSG_VAL_OUTOFRANGE
            End If
      End Select
     End If
    Else
     
    ' Process if Equatorial Tracking
     
     If Duration = 0 Then

      Select Case (direction)
        Case guideNorth 'DEC+
            If HC.decpulse_enchk.Value = 1 Then
                eqres = EQ_MotorStop(1)
                If eqres = 1 Then GoTo PError
                Do
                  eqres = EQ_GetMotorStatus(1)
                  If eqres = 1 Then GoTo PError
                Loop While (eqres And &H10) <> 0
                gEQDECPulseDuration = 0
                If HC.Frame16.Visible = True Then Call Plot_points(1, 0, 0)
            End If
 
        Case guideSouth 'DEC-
            If HC.decpulse_enchk.Value = 1 Then
                eqres = EQ_MotorStop(1)
                If eqres = 1 Then GoTo PError
                Do
                  eqres = EQ_GetMotorStatus(1)
                  If eqres = 1 Then GoTo PError
                Loop While (eqres And &H10) <> 0
                gEQDECPulseDuration = 0
                If HC.Frame16.Visible = True Then Call Plot_points(1, 1, 0)
            End If
        Case guideEast  'RA+
            If HC.rapulse_enchk.Value = 1 Then
                eqres = EQ_SendGuideRate(0, gTrackingStatus - 1, 0, 0, gHemisphere, gHemisphere)
                gEQRAPulseDuration = 0
                If HC.Frame16.Visible = True Then Call Plot_points(0, 0, 0)
            End If
        Case guideWest  'RA-
            If HC.rapulse_enchk.Value = 1 Then
                eqres = EQ_SendGuideRate(0, gTrackingStatus - 1, 0, 0, gHemisphere, gHemisphere)
                gEQRAPulseDuration = 0
                If HC.Frame16.Visible = True Then Call Plot_points(0, 1, 0)
            End If
        Case Else
            If gAscomCompatibility.AllowPulseGuideExceptions Then
                RaiseError SCODE_VAL_OUTOFRANGE, ERR_SOURCE, _
                    "Pulse Guide: GuideDirections" & MSG_VAL_OUTOFRANGE
            End If
       End Select
    
    Else
    
      Select Case (direction)
        Case guideNorth 'DEC+
            If HC.decpulse_enchk.Value = 1 Then
            
                eqres = EQ_MotorStop(1)
                If eqres = 1 Then GoTo PError
                Do
                  eqres = EQ_GetMotorStatus(1)
                  If eqres = 1 Then GoTo PError
                Loop While (eqres And &H10) <> 0
            
                eqres = EQ_SendGuideRate(1, gTrackingStatus - 1, val(HC.VScroll4.Value), 1, 0, 0)
                If HC.decfixed_enchk.Value = 0 Then
                    gEQDECPulseDuration = Duration * (HC.HScrollDECWidth.Value / 100)
                Else
                    gEQDECPulseDuration = HC.VScroll6.Value * 100
                End If
                If HC.Frame16.Visible = True Then Call Plot_points(1, 0, gEQDECPulseDuration)
            End If
            
        Case guideSouth 'DEC-
            If HC.decpulse_enchk.Value = 1 Then
            
                eqres = EQ_MotorStop(1)
                If eqres = 1 Then GoTo PError
                Do
                  eqres = EQ_GetMotorStatus(1)
                  If eqres = 1 Then GoTo PError
                Loop While (eqres And &H10) <> 0
                
                eqres = EQ_SendGuideRate(1, gTrackingStatus - 1, val(HC.VScroll4.Value), 0, 0, 0)
                If HC.decfixed_enchk.Value = 0 Then
                    gEQDECPulseDuration = Duration * (HC.HScrollDECWidth.Value / 100)
                Else
                    gEQDECPulseDuration = HC.VScroll6.Value * 100
                End If
                If HC.Frame16.Visible = True Then Call Plot_points(1, 1, gEQDECPulseDuration)
            End If
        Case guideEast  'RA+
            If HC.rapulse_enchk.Value = 1 Then
                eqres = EQ_SendGuideRate(0, gTrackingStatus - 1, val(HC.VScroll3.Value), 1, gHemisphere, gHemisphere)
                If HC.rafixed_enchk.Value = 0 Then
                    gEQRAPulseDuration = Duration * (HC.HScrollRAWidth.Value / 100)
                Else
                    gEQRAPulseDuration = HC.VScroll5.Value * 100
                End If
                If HC.Frame16.Visible = True Then Call Plot_points(0, 0, gEQRAPulseDuration)
            End If
        Case guideWest  'RA-
            If HC.rapulse_enchk.Value = 1 Then
                eqres = EQ_SendGuideRate(0, gTrackingStatus - 1, val(HC.VScroll3.Value), 0, gHemisphere, gHemisphere)
                If HC.rafixed_enchk.Value = 0 Then
                    gEQRAPulseDuration = Duration * (HC.HScrollRAWidth.Value / 100)
                Else
                    gEQRAPulseDuration = HC.VScroll5.Value * 100
                End If
                If HC.Frame16.Visible = True Then Call Plot_points(0, 1, gEQRAPulseDuration)
            End If
        Case Else
            If gAscomCompatibility.AllowPulseGuideExceptions Then
                RaiseError SCODE_VAL_OUTOFRANGE, ERR_SOURCE, _
                    "Pulse Guide: GuideDirections" & MSG_VAL_OUTOFRANGE
            End If
      End Select
     End If
    End If
    HC.Pulseguide_Timer.Enabled = True
    
PError:

End Sub

Public Sub SetPark()

    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 7, ("COMMAND SetPark")
    HC.Add_Message (HC.oLangDll.GetLangString(5004))
    DefinePark (True)

End Sub

Public Sub SetupDialog()

Unload Setupfrm
Setupfrm.Show 1

End Sub

Public Sub SlewToCoordinates(ByVal RightAscension As Double, _
                                    ByVal Declination As Double)
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
            HC.Add_Message ("SyncCSlew: " & HC.oLangDll.GetLangString(105) & "[ " & FmtSexa(gTargetRA, False) & " ] " & HC.oLangDll.GetLangString(106) & "[ " & FmtSexa(gTargetDec, True) & " ]")
            gSlewCount = NUM_SLEW_RETRIES               'Set initial iterative slew count
            Call radecAsyncSlew(False, gGotoRate)
        End If
    Else
            HC.Add_Message (HC.oLangDll.GetLangString(5000))
            RaiseError SCODE_SCOPE_PARKED, ERR_SOURCE, "SlewToCoordinates() " & MSG_SCOPE_PARKED
    End If
End Sub

Public Sub SlewToCoordinatesAsync(ByVal RightAscension As Double, _
                                    ByVal Declination As Double)
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
            HC.Add_Message ("CoordSlew: " & HC.oLangDll.GetLangString(105) & "[ " & FmtSexa(gTargetRA, False) & " ] " & HC.oLangDll.GetLangString(106) & "[ " & FmtSexa(gTargetDec, True) & " ]")
            gSlewCount = NUM_SLEW_RETRIES               'Set initial iterative slew count
            Call radecAsyncSlew(False, gGotoRate)
        End If
    Else
        HC.Add_Message (HC.oLangDll.GetLangString(5000))
        RaiseError SCODE_SCOPE_PARKED, ERR_SOURCE, "SlewToCoordinatesAsync() " & MSG_SCOPE_PARKED
    End If
End Sub

Public Sub SlewToTarget()

    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 6, ("COMMAND SlewToTarget")
    If gEQparkstatus = 0 Then
        If gAscomCompatibility.SlewWithTrackingOff = False And gTrackingStatus = 0 Then
            RaiseError SCODE_RADEC_SLEW_ERROR, ERR_SOURCE, "SlewToTarget() " & MSG_RADEC_SLEW_ERROR
        Else
            If Align.Visible Then
                Call Align.FillAlignmentStar(gTargetRA, gTargetDec)
            End If
            HC.Add_Message ("SyncSlew: " & HC.oLangDll.GetLangString(105) & "[ " & FmtSexa(gTargetRA, False) & "] " & HC.oLangDll.GetLangString(106) & "[ " & FmtSexa(gTargetDec, True) & " ]")
            gSlewCount = NUM_SLEW_RETRIES               'Set initial iterative slew count
            Call radecAsyncSlew(False, gGotoRate)
        End If
    Else
        HC.Add_Message (HC.oLangDll.GetLangString(5000))
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
            HC.Add_Message ("AsyncSlew: " & HC.oLangDll.GetLangString(105) & "[ " & FmtSexa(gTargetRA, False) & "] " & HC.oLangDll.GetLangString(106) & "[ " & FmtSexa(gTargetDec, True) & " ]")
            gSlewCount = NUM_SLEW_RETRIES               'Set initial iterative slew count
            Call radecAsyncSlew(False, gGotoRate)
        End If
    Else
        HC.Add_Message (HC.oLangDll.GetLangString(5000))
        RaiseError SCODE_SCOPE_PARKED, ERR_SOURCE, "SlewToTargetAsync() " & MSG_SCOPE_PARKED
    End If
    
End Sub

Public Sub SyncToCoordinates(ByVal RightAscension As Double, _
                                    ByVal Declination As Double)
                                    

    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 7, ("COMMAND SyncToCoordinates(" & RightAscension & "," & Declination & ")")
    If gEQparkstatus = 0 Then
        If gTrackingStatus = 0 Then
            RaiseError SCODE_RADEC_SLEW_ERROR, ERR_SOURCE, "SyncToCoordinates() " & MSG_RADEC_SLEW_ERROR
        Else
            HC.Add_Message ("SynCoor: " & HC.oLangDll.GetLangString(105) & "[ " & FmtSexa(RightAscension, False) & "] " & HC.oLangDll.GetLangString(106) & "[ " & FmtSexa(Declination, True) & " ]")
            SyncToRADEC RightAscension, Declination, gLongitude, gHemisphere
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
            HC.Add_Message ("SyncTaget: " & HC.oLangDll.GetLangString(105) & "[ " & FmtSexa(gTargetRA, False) & "] " & HC.oLangDll.GetLangString(106) & "[ " & FmtSexa(gTargetDec, True) & " ]")
            SyncToRADEC gTargetRA, gTargetDec, gLongitude, gHemisphere
        End If
    Else
        RaiseError SCODE_SCOPE_PARKED, ERR_SOURCE, "SyncToTarget() " & MSG_SCOPE_PARKED
    End If


End Sub

Public Sub UnPark()

    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 7, ("COMMAND Unpark")
    Unparkscope
    
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
     
     HC.Add_Message_Align ("---------------------------")
     HC.Add_Message_Align (HC.oLangDll.GetLangString(6000))
     HC.Add_Message_Align (HC.oLangDll.GetLangString(105) & ":" & Format$(str(gRA1Star), "000000000") & " " & HC.oLangDll.GetLangString(106) & ":" & Format$(str(gDEC1Star), "000000000"))
     HC.Add_Message_Align (HC.oLangDll.GetLangString(6002))

  
End Sub


Private Sub RaiseError(ErrNumber As Long, ErrSource As String, ErrDescription As String)

    If AscomTrace.AscomTraceEnabled Then AscomTrace.Add_log 0, ("ERROR RAISED: errno=" & Hex(ErrNumber) & " " & ErrDescription)
    If gAscomCompatibility.AllowExceptions Then
        ERR.Raise ErrNumber, ErrSource, ErrDescription
    End If

End Sub


