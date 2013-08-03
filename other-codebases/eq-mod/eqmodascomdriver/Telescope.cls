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
    equOther = 0
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

Private Const ID As String = "EQMOD.Telescope"
Private Const DESC As String = "EQMOD ASCOM Scope Driver"

Private m_Serial As DriverHelper.Serial
Private m_Util As DriverHelper.Util
Private m_Profile As DriverHelper.Profile
Private m_iSettleTime As Integer

Private m_DriveRates As Collection



Private Sub Class_Initialize()

    eqres = EQ_End()

    Set m_Serial = New DriverHelper.Serial
    Set m_Util = New DriverHelper.Util
    Set m_Profile = New DriverHelper.Profile
    m_Profile.DeviceType = "Telescope"          ' We're a Telescope driver
    m_iSettleTime = 0                           ' Default 0 slew settle time
    m_Profile.Register ID, DESC                 ' Self-register
    
    Set m_DriveRates = New Collection
    m_DriveRates.Add driveSidereal
    m_DriveRates.Add driveLunar
    m_DriveRates.Add driveSolar
    
    gHemisphere = 1
    gTargetRA = EQ_INVALIDCOORDINATE
    gTargetDec = EQ_INVALIDCOORDINATE
    
    
    'Initialize position values - They are actually done also on
    'Driver connect
    
    Call ReadSiteValues
    Call ReadSyncMap
    Call ReadAlignMap
    
    gEQparkstatus = 0
    gTrackingStatus = 0                        ' Initially not tracking
    gRAStatus = EQ_MOTORBUSY                   ' RA Motor Busy Status
    gDECStatus = EQ_MOTORBUSY                  ' DEC Motor Busy Status
    gSlewStatus = True                         ' Not Slewing status
    gRAStatus_slew = False                     ' Slew to track condidition
    gSlewCount = 0                             ' Goto Iterative Counter
    
    ReadComPortSettings
    
  
    gTot_RA = gTot_step                        ' Set RA Total Encoder Step Count
    gTot_DEC = gTot_step                       ' Set DEC total Encoder Step Count
    
    gSOP = 0                                    'Set at unknown pier setting

    HC.HCMessage = ""                           'Set Message Center to BLANK
 
    HC.DisplayTimer.Enabled = True
 
End Sub

Private Sub Class_Terminate()
    
    On Error Resume Next
    
    HC.DisplayTimer.Enabled = False

    eqres = EQ_End()
    
    writeratebarstateHC
    
    Unload HC
    Unload Align
    Unload Slewpad
    
End Sub

'------------------ V2 Compliance Properties

Public Property Get AlignmentMode() As AlignmentModes

    AlignmentMode = algGermanPolar
    
End Property

Public Property Get AtHome() As Boolean
    AtHome = False
End Property

Public Property Get AtPark() As Boolean

    If gEQparkstatus = 1 Then
        AtPark = True
    Else
        AtPark = False
    End If
    
End Property

'Later this will be changed to support the property

Public Property Get CanSetDeclinationRate() As Boolean
    CanSetDeclinationRate = False
End Property

'Later this will be changed to support the property
Public Property Get CanSetGuideRates() As Boolean
    CanSetGuideRates = False
End Property


Public Property Get CanSetPierSide() As Boolean
    CanSetPierSide = False
End Property

Public Property Get CanSetRightAscensionRate() As Boolean
    CanSetRightAscensionRate = False
End Property

Public Property Get CanSlewAltAz() As Boolean
    CanSlewAltAz = False
End Property

Public Property Get CanSlewAltAzAsync() As Boolean
    CanSlewAltAzAsync = False
End Property

Public Property Get CanSyncAltAz() As Boolean
    CanSyncAltAz = False
End Property

Public Property Get DriverVersion() As String
    DriverVersion = App.Major & "." & App.Minor
End Property

Public Property Get EquatorialSystem() As EquatorialCoordinateType
    EquatorialSystem = equOther
End Property

Public Property Get GuideRateDeclination() As Double
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
              "Property Get GuideRateDeclination" & MSG_NOT_IMPLEMENTED
End Property

Public Property Let GuideRateDeclination(ByVal newVal As Double)
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
              "Property Let GuideRateDeclination" & MSG_NOT_IMPLEMENTED
End Property

Public Property Get GuideRateRightAscension() As Double
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
              "Property Get GuideRateRightAscension" & MSG_NOT_IMPLEMENTED
End Property


Public Property Let GuideRateRightAscension(ByVal newVal As Double)
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
              "Property Let GuideRateRightAscension" & MSG_NOT_IMPLEMENTED
End Property


Public Property Get InterfaceVersion() As Integer
    InterfaceVersion = 2
End Property


Public Property Let SideOfPier(ByVal newVal As PierSide)
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
              "Property Let SideOfPier" & MSG_NOT_IMPLEMENTED
End Property
Public Property Get SideOfPier() As PierSide
    SideOfPier = pierUnknown
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
              "Property Get SideOfPier" & MSG_NOT_IMPLEMENTED
End Property


Public Property Get TrackingRate() As DriveRates
    TrackingRate = driveSidereal
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
              "Property Get TrackingRate" & MSG_NOT_IMPLEMENTED
End Property

Public Property Get TrackingRates() As Object
    Set TrackingRates = m_DriveRates
End Property

Public Function AxisRates(axis As AxisID) As Object
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
              "Method AxisRates()" & MSG_NOT_IMPLEMENTED
End Function

Public Function CanMoveAxis(axis As AxisID) As Boolean
    CanMoveAxis = False
End Function

Public Function DestinationSideOfPier(ByVal destRa As Double, destDec As Double) As Long
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
              "Method DestinationSideOfPier()" & MSG_NOT_IMPLEMENTED
End Function

Public Sub MoveAxis(axis As AxisID, rate As Double)
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
              "Method MoveAxis()" & MSG_NOT_IMPLEMENTED
End Sub
Public Sub SlewToAltAz(ByVal Azimuth As Double, _
                        ByVal Altitude As Double)
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
              "Method SlewToAltAz()" & MSG_NOT_IMPLEMENTED
End Sub


Public Sub SlewToAltAzAsync(ByVal Azimuth As Double, _
                            ByVal Altitude As Double)
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
              "Method SlewToAltAzAsync()" & MSG_NOT_IMPLEMENTED
End Sub
Public Sub SyncToAltAz(ByVal Azimuth As Double, _
                        ByVal Altitude As Double)
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
              "Method SyncToAltAz()" & MSG_NOT_IMPLEMENTED
End Sub


'-------------------------- V1 Properties Starts here -----------------

Public Property Get DoesRefraction() As Boolean
    DoesRefraction = False
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
              "Property Get DoesRefraction" & MSG_NOT_IMPLEMENTED
End Property

Public Property Get Altitude() As Double

    Altitude = gAlt
    
End Property

Public Property Get ApertureDiameter() As Double

    '** insert your implementation here and remove the following line
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property ApertureDiameter" & MSG_NOT_IMPLEMENTED

End Property

Public Property Get ApertureArea() As Double

    '** insert your implementation here and remove the following line
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property ApertureDiameter" & MSG_NOT_IMPLEMENTED

End Property

Public Property Get Azimuth() As Double

    Azimuth = gAz

End Property

Public Property Get CanFindHome() As Boolean

    CanFindHome = False

End Property

Public Property Get CanPark() As Boolean

    CanPark = True

End Property

Public Property Get CanSetPark() As Boolean

    CanSetPark = True

End Property

Public Property Get CanSetTracking() As Boolean

    CanSetTracking = True

End Property

Public Property Get CanSlew() As Boolean

    CanSlew = True

End Property

Public Property Get CanSlewAsync() As Boolean

    CanSlewAsync = True

End Property

Public Property Get CanSync() As Boolean

    CanSync = True

End Property

Public Property Get CanUnpark() As Boolean

    CanUnpark = True

End Property
Public Property Get CanPulseGuide() As Boolean

    CanPulseGuide = True

End Property

Public Property Get Connected() As Boolean

   
  If EQ_GetMountStatus() <> 1 Then
      Connected = False
  Else
      Connected = True
  End If
    
    
End Property

Public Property Let Connected(ByVal newVal As Boolean)


    If newVal Then
    
        gEQRAPulseDuration = 0
        gEQDECPulseDuration = 0
        gEQPulsetimerflag = True
        
    
        HC.Show
        
        eqres = EQ_Init(gPort, gBaud, gTimeout, gRetry)
    
        If eqres = EQ_OK And EQ_GetMountStatus() = 1 Then
        
            eqres = EQ_GetTotal360microstep(0)
            If eqres < &H1000000 Then gTot_RA = eqres
            eqres = EQ_GetTotal360microstep(1)
            If eqres < &H1000000 Then gTot_DEC = eqres
            
            'Make sure motors are not running
            
            eqres = EQ_MotorStop(0)     ' Stop RA Motor
            eqres = EQ_MotorStop(1)     ' Stop DEC Motor
            
            'Get state of at least one of the motors
            
            eqres = EQ_GetMotorStatus(0)
            
            ' If its an error then Initialize it
            
            If eqres = EQ_NOTINITIALIZED Then
                eqres = EQ_InitMotors(RAEncoder_Home_pos, DECEncoder_Home_pos)
            End If
            
            'Make sure we get the latest data from the registry
            
           
            Call ReadSiteValues
            Call ReadSyncMap
            Call ReadAlignMap
            
           
            HC.Add_Message ("EQ Modded Mount found at " & gPort & ":" & Str(gBaud))
            HC.Add_Message ("Mount Version : " & printhex(EQ_GetMountVersion()))
            HC.EncoderTimer.Enabled = True
            HC.EncoderTimerFlag = True
            gEQPulsetimerflag = True
            HC.Pulseguide_Timer.Enabled = False     'Enabled only during pulseguide session
            
            
            gEQparkstatus = readparkStatus()

            If gEQparkstatus = 1 Then
                HC.parklbl.Caption = "Mount Park Status: PARKED"
            Else
                HC.parklbl.Caption = "Mount Park Status: NOT PARKED"
            End If

        Else
            HC.EncoderTimer.Enabled = False
            HC.EncoderTimerFlag = False
            eqres = EQ_End()
            HC.Add_Message ("EQ Modded Mount NOT found at " & gPort & ":" & Str(gBaud))
            
        End If

    Else
    
        If EQ_GetMountStatus() = 1 Then     ' We update only if the mount is online
            writeParkStatus gEQparkstatus
        End If
    
        HC.EncoderTimer.Enabled = False
        HC.EncoderTimerFlag = False
        
        'Save alignment and Sync data if scope is parked
        'Otherwise an re-alignment / re-Sync process has to be made on restart
        
        If gEQparkstatus = 1 Then
            Call WriteSyncMap
            Call WriteAlignMap
        End If
    
        eqres = EQ_End()
        
        'Save all information to Windows Registry
        
        WriteSiteValues
        
    End If

End Property

Public Property Get Declination() As Double

    Declination = gDec

End Property

Public Property Get DeclinationRate() As Double

    '** insert your implementation here and remove the following line
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property DeclinationRate" & MSG_NOT_IMPLEMENTED

End Property

Public Property Let DeclinationRate(ByVal newVal As Double)

    '** insert your implementation here and remove the following line
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property DeclinationRate" & MSG_NOT_IMPLEMENTED

End Property

Public Property Get Description() As String

    Description = "EQMOD ASCOM Driver"
    
End Property

Public Property Get DriverInfo() As String
    '
    ' Use the Project/Properties sheet, Make tab, to set these
    ' items. That way they will show in the Version tab of the
    ' Explorer property sheet, and the exact same data will
    ' show in Telescope.DriverInfo.
    '
    DriverInfo = App.FileDescription & " " & _
                App.Major & "." & App.Minor & "." & App.Revision
    If App.CompanyName <> "" Then _
        DriverInfo = DriverInfo & vbCrLf & App.CompanyName
    If App.LegalCopyright <> "" Then _
        DriverInfo = DriverInfo & vbCrLf & App.LegalCopyright
    If App.Comments <> "" Then _
        DriverInfo = DriverInfo & vbCrLf & App.Comments
        
End Property

Public Property Get FocalLength() As Double

    '** insert your implementation here and remove the following line
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property FocalLength" & MSG_NOT_IMPLEMENTED

End Property

Public Property Get Name() As String

    Name = "EQMOD ASCOM Driver"                        ' 1-word name
    
End Property

Public Property Get RightAscension() As Double


    RightAscension = gRA

End Property

Public Property Get RightAscensionRate() As Double

    '** insert your implementation here and remove the following line
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property RightAscensionRate" & MSG_NOT_IMPLEMENTED

End Property

Public Property Let RightAscensionRate(ByVal newVal As Double)

    '** insert your implementation here and remove the following line
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property RightAscensionRate" & MSG_NOT_IMPLEMENTED

End Property

Public Property Get SiderealTime() As Double

    SiderealTime = EQnow_lst(gLongitude * DEG_RAD)

End Property

Public Property Get SiteElevation() As Double

   SiteElevation = gElevation

End Property

Public Property Let SiteElevation(ByVal newVal As Double)

    '** insert your implementation here and remove the following line
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property SiteElevation" & MSG_NOT_IMPLEMENTED

End Property

Public Property Get SiteLatitude() As Double

    SiteLatitude = gLatitude

End Property

Public Property Let SiteLatitude(ByVal newVal As Double)

    '** insert your implementation here and remove the following line
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property SiteLatitude" & MSG_NOT_IMPLEMENTED

End Property

Public Property Get SiteLongitude() As Double

    SiteLongitude = gLongitude

End Property

Public Property Let SiteLongitude(ByVal newVal As Double)

    '** insert your implementation here and remove the following line
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property SiteLongitude" & MSG_NOT_IMPLEMENTED

End Property

Public Property Get Slewing() As Boolean

    Slewing = gSlewStatus

End Property

Public Property Get SlewSettleTime() As Integer

    '** If you implement this feature, remove the following line
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property SlewSettleTime" & MSG_NOT_IMPLEMENTED
    SlewSettleTime = m_iSettleTime
    
End Property

Public Property Let SlewSettleTime(ByVal newVal As Integer)

    '** If you implement this feature, remove the following line
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property SlewSettleTime" & MSG_NOT_IMPLEMENTED
    If (newVal < 0) Or (newVal > 100) Then _
        Err.Raise SCODE_PROP_RANGE_ERROR, ERR_SOURCE, MSG_PROP_RANGE_ERROR
        
    m_iSettleTime = newVal
    
End Property

Public Property Get TargetDeclination() As Double

    If gTargetDec = EQ_INVALIDCOORDINATE Then _
        Err.Raise SCODE_PROP_NOT_SET, ERR_SOURCE, _
            "Property Declination " & MSG_PROP_NOT_SET

    TargetDeclination = gTargetDec

End Property

Public Property Let TargetDeclination(ByVal newVal As Double)

    gTargetDec = newVal
    
End Property

Public Property Get TargetRightAscension() As Double

    If gTargetRA = EQ_INVALIDCOORDINATE Then _
        Err.Raise SCODE_PROP_NOT_SET, ERR_SOURCE, _
            "Property Ascension " & MSG_PROP_NOT_SET


    TargetRightAscension = gTargetRA

End Property

Public Property Let TargetRightAscension(ByVal newVal As Double)

    gTargetRA = newVal
    

End Property

Public Property Get Tracking() As Boolean

    If gTrackingStatus <> 0 Then
    
        Tracking = True
    Else
        Tracking = False
        
    End If
    
End Property

Public Property Let Tracking(ByVal newVal As Boolean)


   If gEQparkstatus = 0 Then
    If newVal Then
    
        eqres = EQ_StartRATrack(0, gHemisphere, gHemisphere)
        gTrackingStatus = 1                 'Sidreal rate tracking'
        HC.TrackingFrame.Caption = "Track Rate : Tracking at Sidreal"
    
    Else
    
        eqres = EQ_MotorStop(0)
        gTrackingStatus = 0
        HC.TrackingFrame.Caption = "Track Rate : Not Tracking "
    
    End If
   Else
   
    HC.Add_Message ("Cannot start tracking. Mount still parked. Please unpark")
   
   End If

End Property

Public Property Get UTCDate() As Date

    UTCDate = CDate(CDbl(Now()) + (CDbl(utc_offs()) / 86400#))
        
End Property

Public Property Let UTCDate(ByVal newVal As Date)

     'Impossible to set own PC time

     Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property UTCDate" & MSG_NOT_IMPLEMENTED

End Property


' ------------------------- Methods Portion of the Code -------------


Public Sub AbortSlew()

    'Stop all motors

    eqres = EQ_MotorStop(0)
    eqres = EQ_MotorStop(1)
    
    'Disable HC Slew timers
    
    gSlewStatus = False
    gRAStatus_slew = False
    HC.GotoTimer.Enabled = False

End Sub

Public Sub CommandBlind(ByVal Command As String)

    '** insert your implementation here and remove the following line
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Method xxx()" & MSG_NOT_IMPLEMENTED

End Sub

Public Function CommandBool(ByVal Command As String) As Boolean

    '** insert your implementation here and remove the following line
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Method CommandBlind()" & MSG_NOT_IMPLEMENTED

End Function

Public Function CommandString(ByVal Command As String) As String

    '** insert your implementation here and remove the following line
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Method CommandString()" & MSG_NOT_IMPLEMENTED

End Function

Public Sub FindHome()

    '** insert your implementation here and remove the following line
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Method FindHome()" & MSG_NOT_IMPLEMENTED

End Sub

Public Sub Park()

    If gEQparkstatus = 0 Then
        ParktoUserDefine (True)
        gEQparkstatus = 1
        HC.TrackingFrame.Caption = "Track Rate : Not Tracking"
        HC.parklbl.Caption = "Mount Park Status: PARKED"
        HC.Add_Message ("Scope parked. You may turn it off after slewing")
    End If

End Sub

Public Property Get IsPulseGuiding() As Boolean

    If (gEQRAPulseDuration + gEQDECPulseDuration) <> 0 Then
        IsPulseGuiding = True
    Else
        IsPulseGuiding = False
    End If
    
End Property

Public Sub PulseGuide(ByVal direction As GuideDirections, _
                                            ByVal Duration As Long)
    
    'Pulse guide implentation for EQMOD
    'This uses the duration parameter and an asynchronous timer
    'that will decrement the duration count for every time tick
    'It then disables the guiderate upon expiration of the counter
    
    HC.Pulseguide_Timer.Enabled = False
    Select Case (direction)
        Case guideNorth 'DEC+
            If HC.decpulse_enchk.Value = 1 Then
                eqres = EQ_SendGuideRate(1, gTrackingStatus - 1, val(HC.VScroll4.Value), 1, 0, 0)
                If HC.decfixed_enchk.Value = 0 Then
                    gEQDECPulseDuration = Duration
                Else
                    gEQDECPulseDuration = HC.VScroll6.Value * 100
                End If
            End If
        Case guideSouth 'DEC-
            If HC.decpulse_enchk.Value = 1 Then
                eqres = EQ_SendGuideRate(1, gTrackingStatus - 1, val(HC.VScroll4.Value), 0, 0, 0)
                If HC.decfixed_enchk.Value = 0 Then
                    gEQDECPulseDuration = Duration
                Else
                    gEQDECPulseDuration = HC.VScroll6.Value * 100
                End If
            End If
        Case guideEast  'RA+
            If HC.rapulse_enchk.Value = 1 Then
                eqres = EQ_SendGuideRate(0, gTrackingStatus - 1, val(HC.VScroll3.Value), 1, gHemisphere, gHemisphere)
                If HC.rafixed_enchk.Value = 0 Then
                    gEQRAPulseDuration = Duration
                Else
                    gEQRAPulseDuration = HC.VScroll5.Value * 100
                End If
            End If
        Case guideWest  'RA-
            If HC.rapulse_enchk.Value = 1 Then
                eqres = EQ_SendGuideRate(0, gTrackingStatus - 1, val(HC.VScroll3.Value), 0, gHemisphere, gHemisphere)
                If HC.rafixed_enchk.Value = 0 Then
                    gEQRAPulseDuration = Duration
                Else
                    gEQRAPulseDuration = HC.VScroll5.Value * 100
                End If
            End If
        Case Else
            Err.Raise SCODE_PROP_RANGE_ERROR, ERR_SOURCE, _
                "Pulse Guide: GuideDirections" & MSG_PROP_RANGE_ERROR
    End Select
    HC.Pulseguide_Timer.Enabled = True
    

End Sub

Public Sub SetPark()

    HC.Add_Message ("Current location set as new park position.")
    DefinePark (True)

End Sub

Public Sub SetupDialog()

Unload Setupfrm
Setupfrm.Show 1

End Sub

Public Sub SlewToCoordinates(ByVal RightAscension As Double, _
                                    ByVal Declination As Double)
    
    If gEQparkstatus = 0 Then
        gTargetRA = RightAscension
        gTargetDec = Declination
 
        HC.Add_Message ("SyncCSlew: RA[ " & FmtSexa(gTargetRA, False) & " ] DEC[ " & FmtSexa(gTargetDec, True) & " ]")
   
        gSlewCount = NUM_SLEW_RETRIES               'Set initial iterative slew count
        Call radecAsyncSlew(False)
    Else
    
        HC.Add_Message ("Slew Error: Scope still parked. Please unpark")
    
    End If

End Sub

Public Sub SlewToCoordinatesAsync(ByVal RightAscension As Double, _
                                    ByVal Declination As Double)
    
    If gEQparkstatus = 0 Then
        gTargetRA = RightAscension
        gTargetDec = Declination
 
        HC.Add_Message ("CoordSlew: RA[ " & FmtSexa(gTargetRA, False) & " ] DEC[ " & FmtSexa(gTargetDec, True) & " ]")
   
        gSlewCount = NUM_SLEW_RETRIES               'Set initial iterative slew count
        Call radecAsyncSlew(False)
    Else
        HC.Add_Message ("Slew Error: Scope still parked. Please unpark")
    End If
 
End Sub

Public Sub SlewToTarget()

    If gEQparkstatus = 0 Then
        HC.Add_Message ("SyncSlew: RA[ " & FmtSexa(gTargetRA, False) & "] DEC[ " & FmtSexa(gTargetDec, True) & " ]")
        gSlewCount = NUM_SLEW_RETRIES               'Set initial iterative slew count
        Call radecAsyncSlew(False)
    Else
        HC.Add_Message ("Slew Error: Scope still parked. Please unpark")
    End If
    
End Sub

Public Sub SlewToTargetAsync()

    If gEQparkstatus = 0 Then
        HC.Add_Message ("AsyncSlew: RA[ " & FmtSexa(gTargetRA, False) & "] DEC[ " & FmtSexa(gTargetDec, True) & " ]")
        gSlewCount = NUM_SLEW_RETRIES               'Set initial iterative slew count
        Call radecAsyncSlew(False)
    Else
        HC.Add_Message ("Slew Error: Scope still parked. Please unpark")
    End If
    
End Sub

Public Sub SyncToCoordinates(ByVal RightAscension As Double, _
                                    ByVal Declination As Double)
                                    

    HC.Add_Message ("SynCoor: RA[ " & FmtSexa(RightAscension, False) & "] DEC[ " & FmtSexa(Declination, True) & " ]")
    SyncToRADEC RightAscension, Declination, gLongitude, gHemisphere

End Sub

Public Sub SyncToTarget()

    HC.Add_Message ("SyncTaget: RA[ " & FmtSexa(gTargetRA, False) & "] DEC[ " & FmtSexa(gTargetDec, True) & " ]")
    SyncToRADEC gTargetRA, gTargetDec, gLongitude, gHemisphere


End Sub

Public Sub Unpark()


    If EQ_GetMountStatus() = 1 Then     ' Make sure that we unpark only if the mount is online
        If gEQparkstatus = 1 Then
            HC.TrackingFrame.Caption = "Track Rate : Not Tracking"
            Unparkscope
            gEQparkstatus = 0
            HC.parklbl.Caption = "Mount Park Status: NOT PARKED"
        Else
            HC.Add_Message ("Unpark error: Mount is already unparked")
        End If
    End If
    
End Sub



Public Sub ReadSiteValues()
    
     Dim tmptxt As String
  
     HC.cbNS.ListIndex = 0
     HC.cbEW.ListIndex = 0
     HC.cbhem.ListIndex = 0
   
     tmptxt = m_Profile.GetValue(ID, "LongitudeDeg")
     If tmptxt <> "" Then HC.txtLongDeg.Text = tmptxt
     
     tmptxt = m_Profile.GetValue(ID, "LongitudeMin")
     If tmptxt <> "" Then HC.txtLongMin.Text = tmptxt
     
     tmptxt = m_Profile.GetValue(ID, "LongitudeEW")
     If tmptxt <> "" Then HC.cbEW.ListIndex = val(tmptxt)
     
     tmptxt = m_Profile.GetValue(ID, "LatitudeDeg")
     If tmptxt <> "" Then HC.txtLatDeg.Text = tmptxt
   
     tmptxt = m_Profile.GetValue(ID, "LatitudeMin")
     If tmptxt <> "" Then HC.txtLatMin.Text = tmptxt
     
     tmptxt = m_Profile.GetValue(ID, "LatitudeNS")
     If tmptxt <> "" Then HC.cbNS.ListIndex = val(tmptxt)
     
     tmptxt = m_Profile.GetValue(ID, "Elevation")
     If tmptxt <> "" Then HC.txtElevation = tmptxt
     
      tmptxt = m_Profile.GetValue(ID, "HemisphereNS")
     If tmptxt <> "" Then HC.cbhem.ListIndex = val(tmptxt)
     
     gLongitude = CDbl(HC.txtLongDeg) + (CDbl(HC.txtLongMin) / 60#)
     If HC.cbEW.Text = "W" Then gLongitude = -gLongitude  ' W is neg
    
     gLatitude = CDbl(HC.txtLatDeg) + (CDbl(HC.txtLatMin) / 60#)
     If HC.cbNS.Text = "S" Then gLatitude = -gLatitude
     gElevation = CDbl(HC.txtElevation)
     
     If HC.cbhem.Text = "North" Then
        gHemisphere = 0
     Else
        gHemisphere = 1
     End If
     
End Sub

Public Sub WriteSiteValues()

        m_Profile.WriteValue ID, "LatitudeDeg", CStr(HC.txtLatDeg.Text)
        m_Profile.WriteValue ID, "LatitudeMin", CStr(HC.txtLatMin.Text)
        m_Profile.WriteValue ID, "LatitudeNS", CStr(HC.cbNS.ListIndex)
        m_Profile.WriteValue ID, "LongitudeDeg", CStr(HC.txtLongDeg.Text)
        m_Profile.WriteValue ID, "LongitudeMin", CStr(HC.txtLongMin.Text)
        m_Profile.WriteValue ID, "LongitudeEW", CStr(HC.cbEW.ListIndex)
        m_Profile.WriteValue ID, "HemisphereNS", CStr(HC.cbhem.ListIndex)
        m_Profile.WriteValue ID, "Elevation", CStr(HC.txtElevation.Text)
      

End Sub

Public Sub ReadComPortSettings()

   
     Dim tmptxt As String
   
     tmptxt = m_Profile.GetValue(ID, "Port")
     If tmptxt <> "" Then
        gPort = tmptxt
     Else
        gPort = "COM1"
     End If
     
     tmptxt = m_Profile.GetValue(ID, "Baud")
     If tmptxt <> "" Then
        gBaud = val(tmptxt)
     Else
        gBaud = 9600
     End If
     
     tmptxt = m_Profile.GetValue(ID, "Timeout")
     If tmptxt <> "" Then
        gTimeout = val(tmptxt)
     Else
        gTimeout = 1000
     End If
     

     tmptxt = m_Profile.GetValue(ID, "Retry")
     If tmptxt <> "" Then
        gRetry = val(tmptxt)
     Else
        gRetry = 1
     End If
     
End Sub

Public Sub ReadSyncMap()

     Dim tmptxt As String

     tmptxt = m_Profile.GetValue(ID, "RSYNC01")
     If tmptxt <> "" Then
        gRASync01 = val(tmptxt)
     Else
        gRASync01 = 0
     End If
     
     tmptxt = m_Profile.GetValue(ID, "DSYNC01")
     If tmptxt <> "" Then
        gDECSync01 = val(tmptxt)
     Else
        gDECSync01 = 0
     End If

    HC.DxSalbl.Caption = Format$(Str(gRASync01), "000000000")
    HC.DxSblbl.Caption = Format$(Str(gDECSync01), "000000000")

End Sub
Public Sub ReadAlignMap()

     Dim tmptxt As String

     tmptxt = m_Profile.GetValue(ID, "RALIGN01")
     If tmptxt <> "" Then
        gRA1Star = val(tmptxt)
     Else
        gRA1Star = 0
     End If
     
     tmptxt = m_Profile.GetValue(ID, "DALIGN01")
     If tmptxt <> "" Then
        gDEC1Star = val(tmptxt)
     Else
        gDEC1Star = 0
     End If

     HC.DxAlbl.Caption = Format$(Str(gRA1Star), "000000000")
     HC.DxBlbl.Caption = Format$(Str(gDEC1Star), "000000000")

End Sub


