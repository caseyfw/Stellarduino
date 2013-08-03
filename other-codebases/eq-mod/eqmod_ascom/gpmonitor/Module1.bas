Attribute VB_Name = "Module1"

Option Explicit

Type JOYINFOEX
   dwSize As Long                      ' size of structure
   dwFlags As Long                     ' flags to indicate what to return
   dwXpos As Long                      ' x position
   dwYpos As Long                      ' y position
   dwZpos As Long                      ' z position
   dwRpos As Long                      ' rudder/4th axis position
   dwUpos As Long                      ' 5th axis position
   dwVpos As Long                      ' 6th axis position
   dwButtons As Long                   ' button states
   dwButtonNumber As Long              ' current button number pressed
   dwPOV As Long                       ' point of view state
   dwReserved1 As Long                 ' reserved for communication between winmm driver
   dwReserved2 As Long                 ' reserved for future expansion
End Type

Public Const MAX_JOYSTICKOEMVXDNAME = 260
Public Const MAXPNAMELEN = 32

' The JOYCAPS user-defined type contains information about the joystick capabilities
Type JOYCAPS
    wMid As Integer ' Manufacturer identifier of the device driver for the MIDI output device
    ' For a list of identifiers, see the Manufacturer Indentifier topic in the
    ' Multimedia Reference of the Platform SDK.
    
    wPid As Integer ' Product Identifier Product of the MIDI output device. For a list of
    ' product identifiers, see the Product Identifiers topic in the Multimedia
    ' Reference of the Platform SDK.
    szPname As String * MAXPNAMELEN ' Null-terminated string containing the joystick product name
    wXmin As Long ' Minimum X-coordinate.
    wXmax As Long ' Maximum X-coordinate.
    wYmin As Long ' Minimum Y-coordinate
    wYmax As Long ' Maximum Y-coordinate
    wZmin As Long ' Minimum Z-coordinate
    wZmax As Long ' Maximum Z-coordinate
    wNumButtons As Long ' Number of joystick buttons
    wPeriodMin As Long ' Smallest polling frequency supported when captured by the joySetCapture function.
    wPeriodMax As Long ' Largest polling frequency supported when captured by the joySetCapture function.
    wRmin As Long ' Minimum rudder value. The rudder is a fourth axis of movement.
    wRmax As Long ' Maximum rudder value. The rudder is a fourth axis of movement.
    wUmin As Long ' Minimum u-coordinate (fifth axis) values.
    wUmax As Long ' Maximum u-coordinate (fifth axis) values.
    wVmin As Long ' Minimum v-coordinate (sixth axis) values.
    wVmax As Long ' Maximum v-coordinate (sixth axis) values.
    wCaps As Long ' Joystick capabilities as defined by the following flags
    ' JOYCAPS_HASZ- Joystick has z-coordinate information.
    ' JOYCAPS_HASR- Joystick has rudder (fourth axis) information.
    ' JOYCAPS_HASU- Joystick has u-coordinate (fifth axis) information.
    ' JOYCAPS_HASV- Joystick has v-coordinate (sixth axis) information.
    ' JOYCAPS_HASPOV- Joystick has point-of-view information.
    ' JOYCAPS_POV4DIR- Joystick point-of-view supports discrete values (centered, forward, backward, left, and right).
    ' JOYCAPS_POVCTS Joystick point-of-view supports continuous degree bearings.
    wMaxAxes As Long ' Maximum number of axes supported by the joystick.
    wNumAxes As Long ' Number of axes currently in use by the joystick.
    wMaxButtons As Long ' Maximum number of buttons supported by the joystick.
    szRegKey As String * MAXPNAMELEN ' String containing the registry key for the joystick.
    szOEMVxD As String * MAX_JOYSTICKOEMVXDNAME ' OEM VxD in use
End Type

Type JOYCAPS1
  wMid As Integer
  wPid As Integer
  szPname As String * 32
  wXmin As Long
  wXmax As Long
  wYmin As Long
  wYmax As Long
  wZmin As Long
  wZmax As Long
  wNumButtons As Long
  wPeriodMin As Long
 wPeriodMax As Long
' wRmin As Long
' wRmax As Long
' wUmin As Long
' wUmax As Long
' wVmin As Long
' wVmax As Long
' wMaxAxes As Long
' wNumAxes As Long
' wMaxButtons As Long
' szRegKey As String * 32
' szOEMVxD As String * 240
End Type

Declare Function joyGetPosEx Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFOEX) As Long
Declare Function joyGetDevCaps Lib "winmm.dll" Alias "joyGetDevCapsA" (ByVal id As Long, lpCaps As JOYCAPS1, ByVal uSize As Long) As Long
Declare Function joyGetDevCaps2 Lib "winmm.dll" Alias "joyGetDevCapsA" (ByVal id As Long, lpCaps As JOYCAPS, ByVal uSize As Long) As Long

 Public Const JOYSTICKID1 = 0
   Public Const JOYSTICKID2 = 1
   Public Const JOY_RETURNBUTTONS = &H80&
   Public Const JOY_RETURNCENTERED = &H400&
   Public Const JOY_RETURNPOV = &H40&
   Public Const JOY_RETURNR = &H8&
   Public Const JOY_RETURNU = &H10
   Public Const JOY_RETURNV = &H20
   Public Const JOY_RETURNX = &H1&
   Public Const JOY_RETURNY = &H2&
   Public Const JOY_RETURNZ = &H4&
   Public Const JOY_RETURNALL = (JOY_RETURNX Or JOY_RETURNY Or JOY_RETURNZ Or JOY_RETURNR Or JOY_RETURNU Or JOY_RETURNV Or JOY_RETURNPOV Or JOY_RETURNBUTTONS)
   Public Const JOYCAPS_HASZ = &H1&
   Public Const JOYCAPS_HASR = &H2&
   Public Const JOYCAPS_HASU = &H4&
   Public Const JOYCAPS_HASV = &H8&
   Public Const JOYCAPS_HASPOV = &H10&
   Public Const JOYCAPS_POV4DIR = &H20&
   Public Const JOYCAPS_POVCTS = &H40&
   Public Const JOYERR_BASE = 160
   Public Const JOYERR_UNPLUGGED = (JOYERR_BASE + 7)

   Public Const JOYERR_NOCANDO = (JOYERR_BASE + 6)   ' Request Not Completed
   Public Const JOYERR_NOERROR = (0)                 ' No Error
   Public Const JOYERR_PARMS = (JOYERR_BASE + 5)     ' Bad Parameters

Public joyinfo As JOYCAPS1  ' receives joystick information
Public joyinfo2 As JOYCAPS  ' receives joystick information
Public JoystickDat As JOYINFOEX  ' receives joystick information
 
Public gpID As Long
