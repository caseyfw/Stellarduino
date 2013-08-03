Attribute VB_Name = "JStickKbrd"
'---------------------------------------------------------------------
' Copyright © 2006 Raymund Sarmiento
'
' Permission is hereby granted to use this Software for any purpose
' including combining with commercial products, creating derivative
' works, and redistribution of source or binary code, without
' limitation or consideration. Any redistributed copies of this
' Software must include the above Copyright Notice.
'
' THIS SOFTWARE IS PROVIDED "AS IS". THE AUTHOR OF THIS CODE MAKES NO
' WARRANTIES REGARDING THIS SOFTWARE, EXPRESS OR IMPLIED, AS TO ITS
' SUITABILITY OR FITNESS FOR A PARTICULAR PURPOSE.
'---------------------------------------------------------------------
'
' JStickKbrd.bas - Keyboard/Joystick functions for EQMOD ASCOM Driver
'
'
' Written:  07-Oct-06   Raymund Sarmiento
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' 17-Dec-06 rcs     Initial edit for EQ Mount Driver Function Prototype
' 09-Jun-07 rcs     Fixed bug DPAD E-W Reversal
' 25-Jul-07 cs      User defined joystick button assignments
' 29-Jul-07 cs      Joystick Caibration save/load; Slew presets added.
' 30-Jul-07 cs      Alignment end button handling
' 31-Jul-07 cs      Fixed rate buttons handling
'---------------------------------------------------------------------
'
'
'  SYNOPSIS:
'
'  This is a demonstration of a EQ6/ATLAS/EQG direct stepper motor control access
'  using the EQCONTRL.DLL driver code.
'
'  File EQCONTROL.bas contains all the function prototypes of all subroutines
'  encoded in the EQCONTRL.dll
'
'  The EQ6CONTRL.DLL simplifies execution of the Mount controller board stepper
'  commands.
'
'  The mount circuitry needs to be modified for this test program to work.
'  Circuit details can be found at http://www.freewebs.com/eq6mod/
'

'  DISCLAIMER:

'  You can use the information on this site COMPLETELY AT YOUR OWN RISK.
'  The modification steps and other information on this site is provided
'  to you "AS IS" and WITHOUT WARRANTY OF ANY KIND, express, statutory,
'  implied or otherwise, including without limitation any warranty of
'  merchantability or fitness for any particular or intended purpose.
'  In no event the author will  be liable for any direct, indirect,
'  punitive, special, incidental or consequential damages or loss of any
'  kind whether or not the author  has been advised of the possibility
'  of such loss.

'  WARNING:

'  Circuit modifications implemented on your setup could invalidate
'  any warranty that you may have with your product. Use this
'  information at your own risk. The modifications involve direct
'  access to the stepper motor controls of your mount. Any "mis-control"
'  or "mis-command"  / "invalid parameter" or "garbage" data sent to the
'  mount could accidentally activate the stepper motors and allow it to
'  rotate "freely" damaging any equipment connected to your mount.
'  It is also possible that any garbage or invalid data sent to the mount
'  could cause its firmware to generate mis-steps pulse sequences to the
'  motors causing it to overheat. Make sure that you perform the
'  modifications and testing while there is no physical "load" or
'  dangling wires on your mount. Be sure to disconnect the power once
'  this event happens or if you notice any unusual sound coming from
'  the motor assembly.
'
'  CREDITS:
'
'  Portions of the information on this code should be attributed
'  to Mr. John Archbold from his initial observations and analysis
'  of the interface circuits and of the ASCII data stream between
'  the Hand Controller (HC) and the Go To Controller.
'

Option Explicit


Declare Function joyGetPosEx Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFOEX) As Long
Declare Function joyGetDevCaps Lib "winmm.dll" Alias "joyGetDevCapsA" (ByVal id As Long, lpCaps As JOYCAPS1, ByVal uSize As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias _
   "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long


Public gJoyTimerFlag As Boolean
Public gJoyTimerFlag2 As Boolean

Public gSpiralTimerFlag As Boolean
Public gParkTimerFlag As Boolean

Public gdwXpos As Long
Public gdwYpos As Long
Public gdwZpos As Long
Public gdwRpos As Long
Public gdwButtons As Long
Public gdwPov As Long
Public gZoneX As Integer
Public gZoneY As Integer

Public gEQjbuttons As Long

Public gPrevCode As Long
Public SyncPressCount As Integer
Public PollCount As Integer

Public BTN_STARTSIDREAL As Long
Public BTN_EMERGENCYSTOP As Long
Public BTN_SPIRAL As Long
Public BTN_RARATEINC As Long
Public BTN_DECRATEINC As Long
Public BTN_RARATEDEC As Long
Public BTN_DECRATEDEC As Long
Public BTN_HOMEPARK As Long
Public BTN_USERPARK As Long
Public BTN_ALIGNACCEPT As Long
Public BTN_ALIGNCANCEL As Long
Public BTN_ALIGNEND As Long
Public BTN_UNPARK As Long
Public BTN_EAST As Long
Public BTN_WEST As Long
Public BTN_NORTH As Long
Public BTN_SOUTH As Long
Public BTN_RAREVERSE As Long
Public BTN_DECREVERSE As Long
Public BTN_CUSTOMTRACKSTART As Long
Public BTN_CURRENTPARK As Long
Public BTN_STARTSOLAR As Long
Public BTN_STARTLUNAR As Long
Public BTN_INCRATEPRESET As Long
Public BTN_DECRATEPRESET As Long
Public BTN_RATE1 As Long
Public BTN_RATE2 As Long
Public BTN_RATE3 As Long
Public BTN_RATE4 As Long
Public BTN_PEC As Long
Public BTN_SYNC As Long
Public BTN_NORTHEAST As Long
Public BTN_NORTHWEST As Long
Public BTN_SOUTHEAST As Long
Public BTN_SOUTHWEST As Long
Public BTN_POLARSCOPEALIGN As Long
Public BTN_DEADMANSHANDLE As Long
Public BTN_TOGGLELOCK As Long
Public BTN_TOGGLESCREENSAVER As Long

Public POV_Enabled As Integer


' default joystick button assignments
Public Const BTN_UNDEFINED = 0
Public Const BTN_JOY1 = 1
Public Const BTN_JOY2 = 2
Public Const BTN_JOY3 = 4
Public Const BTN_JOY4 = 8
Public Const BTN_JOY5 = 16
Public Const BTN_JOY6 = 32
Public Const BTN_JOY7 = 64
Public Const BTN_JOY8 = 128
Public Const BTN_JOY9 = 256
Public Const BTN_JOY10 = 512
Public Const BTN_JOY11 = 1024
Public Const BTN_JOY12 = 2048
Public Const BTN_POVN = 0 + 65536
Public Const BTN_POVS = 18000 + 65536
Public Const BTN_POVE = 9000 + 65536
Public Const BTN_POVW = 27000 + 65536
Public Const BTN_POVNE = 4500 + 65536
Public Const BTN_POVNW = 31500 + 65536
Public Const BTN_POVSW = 22500 + 65536
Public Const BTN_POVSE = 13500 + 65536


Type JOYCALIB
    dwMinXpos As Long
    dwMaxXpos As Long
    dwX25Left As Long
    dwX25Right As Long
    dwX75left As Long
    dwX75Right As Long
    dwMinYpos As Long
    dwMaxYpos As Long
    dwY25Left As Long
    dwY25Right As Long
    dwY75left As Long
    dwY75Right As Long
    dwMinZpos As Long
    dwMaxZpos As Long
    dwMinRpos As Long
    dwMaxRpos As Long
    Enabled As Integer
    DualSpeed As Integer
    id As Long
End Type


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
    wMid As Integer                 ' Manufacturer identifier of the device driver for the MIDI output device
                                    ' For a list of identifiers, see the Manufacturer Indentifier topic in the
                                    ' Multimedia Reference of the Platform SDK.
    
    wPid As Integer                 ' Product Identifier Product of the MIDI output device. For a list of
                                    ' product identifiers, see the Product Identifiers topic in the Multimedia
                                    ' Reference of the Platform SDK.
    szPname As String * MAXPNAMELEN ' Null-terminated string containing the joystick product name
    wXmin As Long                   ' Minimum X-coordinate.
    wXmax As Long                   ' Maximum X-coordinate.
    wYmin As Long                   ' Minimum Y-coordinate
    wYmax As Long                   ' Maximum Y-coordinate
    wZmin As Long                   ' Minimum Z-coordinate
    wZmax As Long                   ' Maximum Z-coordinate
    wNumButtons As Long             ' Number of joystick buttons
    wPeriodMin As Long              ' Smallest polling frequency supported when captured by the joySetCapture function.
    wPeriodMax As Long              ' Largest polling frequency supported when captured by the joySetCapture function.
    wRmin As Long                   ' Minimum rudder value. The rudder is a fourth axis of movement.
    wRmax As Long                   ' Maximum rudder value. The rudder is a fourth axis of movement.
    wUmin As Long                   ' Minimum u-coordinate (fifth axis) values.
    wUmax As Long                   ' Maximum u-coordinate (fifth axis) values.
    wVmin As Long                   ' Minimum v-coordinate (sixth axis) values.
    wVmax As Long                   ' Maximum v-coordinate (sixth axis) values.
    wCaps As Long                   ' Joystick capabilities as defined by the following flags
                                        ' JOYCAPS_HASZ- Joystick has z-coordinate information.
                                        ' JOYCAPS_HASR- Joystick has rudder (fourth axis) information.
                                        ' JOYCAPS_HASU- Joystick has u-coordinate (fifth axis) information.
                                        ' JOYCAPS_HASV- Joystick has v-coordinate (sixth axis) information.
                                        ' JOYCAPS_HASPOV- Joystick has point-of-view information.
                                        ' JOYCAPS_POV4DIR- Joystick point-of-view supports discrete values (centered, forward, backward, left, and right).
                                        ' JOYCAPS_POVCTS Joystick point-of-view supports continuous degree bearings.
    wMaxAxes As Long                ' Maximum number of axes supported by the joystick.
    wNumAxes As Long                ' Number of axes currently in use by the joystick.
    wMaxButtons As Long             ' Maximum number of buttons supported by the joystick.
    szRegKey As String * MAXPNAMELEN ' String containing the registry key for the joystick.
    szOEMVxD As String * MAX_JOYSTICKOEMVXDNAME ' OEM VxD in use
End Type

Type JOYCAPS1
   wMid As Integer                     ' Manufacturer identifier of the device driver for the MIDI output device
                                       ' For a list of identifiers, see the Manufacturer Indentifier topic in the
                                       ' Multimedia Reference of the Platform SDK.
   
   wPid As Integer                     ' Product Identifier Product of the MIDI output device. For a list of
                                       ' product identifiers, see the Product Identifiers topic in the Multimedia
                                       ' Reference of the Platform SDK.
   szPname As String * 32     ' Null-terminated string containing the joystick product name
   wXmin As Long                       ' Minimum X-coordinate.
   wXmax As Long                       ' Maximum X-coordinate.
   wYmin As Long                       ' Minimum Y-coordinate
   wYmax As Long                       ' Maximum Y-coordinate
   wZmin As Long                       ' Minimum Z-coordinate
   wZmax As Long                       ' Maximum Z-coordinate
   wNumButtons As Long                 ' Number of joystick buttons
   wPeriodMin As Long                  ' Smallest polling frequency supported when captured by the joySetCapture function.
   wPeriodMax As Long                  ' Largest polling frequency supported when captured by the joySetCapture function.
End Type


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

   Public JoystickDat As JOYINFOEX
   Public JoystickInfo As JOYCAPS1
   Public JoystickCal As JOYCALIB
   
   Dim RAGuidingNudge As Boolean
   Dim DECGuidingNudge As Boolean
   Public SlewActive As Integer


Public Function EQ_JoystickPoller(RRATE As Long, DRATE As Long) As Boolean
Dim i As Long
Dim dwXpos As Long
Dim dwYpos As Long
Dim dwZpos As Long
Dim dwRpos As Long
Dim dwButtons As Long
Dim dwPOV As Long
Dim ZoneX As Integer
Dim ZoneY As Integer
Dim rate As Double

        PollCount = PollCount + 1
        If PollCount > 10 Then
            If SyncPressCount <> 0 Then
                SyncPressCount = 0
            End If
            PollCount = 0
        End If

        JoystickDat.dwSize = Len(JoystickDat)
        JoystickDat.dwFlags = JOY_RETURNALL
        
        If JoystickCal.id = -1 Then
            'Auto search first two ids
            i = joyGetPosEx(JOYSTICKID1, JoystickDat)
            If i <> JOYERR_NOERROR Then i = joyGetPosEx(JOYSTICKID2, JoystickDat)
        Else
            ' use specific id
            i = joyGetPosEx(JoystickCal.id, JoystickDat)
        End If
        
        If i <> JOYERR_NOERROR Then
            
            ' Joystick not found disable joystick scan
            
            EQ_JoystickPoller = False
            Exit Function
            
        End If
        
        ' Start Polling for JoyStick routines here
        
        If i = JOYERR_NOERROR Then
        
            dwXpos = JoystickDat.dwXpos
            dwYpos = JoystickDat.dwYpos
            dwZpos = JoystickDat.dwZpos
            dwRpos = JoystickDat.dwRpos
            dwButtons = JoystickDat.dwButtons
            dwPOV = JoystickDat.dwPOV
            
            ZoneX = 0
            If dwXpos >= JoystickCal.dwMaxXpos Then
                ZoneX = 4
            Else
                If dwXpos > JoystickCal.dwX25Right And dwXpos < JoystickCal.dwX75Right Then
                    If JoystickCal.DualSpeed = 1 Then
                        ZoneX = 3
                    End If
                Else
                    If dwXpos > JoystickCal.dwX75left And dwXpos < JoystickCal.dwX25Left Then
                        If JoystickCal.DualSpeed = 1 Then
                            ZoneX = 1
                        End If
                    Else
                        If dwXpos <= JoystickCal.dwMinXpos Then
                            ZoneX = 2
                        Else
                            ZoneX = 0
                        End If
                   End If
                End If
            End If
            
         
            ZoneY = 0
            If dwYpos >= JoystickCal.dwMaxYpos Then
                ZoneY = 4
            Else
                If dwYpos > JoystickCal.dwY25Right And dwYpos < JoystickCal.dwY75Right Then
                    If JoystickCal.DualSpeed = 1 Then
                        ZoneY = 3
                    End If
                Else
                    If dwYpos > JoystickCal.dwY75left And dwYpos < JoystickCal.dwY25Left Then
                        If JoystickCal.DualSpeed = 1 Then
                            ZoneY = 1
                        End If
                    Else
                        If dwYpos <= JoystickCal.dwMinYpos Then
                            ZoneY = 2
                        End If
                   End If
                End If
            End If
         
         
            ' Debouncing on Both Axis
           
            If (dwXpos <> gdwXpos) And (dwYpos <> gdwYpos) Then
            
                If (gdwXpos <= JoystickCal.dwMinXpos) And (gdwYpos <= JoystickCal.dwMinYpos) Then
                    Call NorthWest_Up
                    gdwXpos = dwXpos
                    gdwYpos = dwYpos
                    EQ_JoystickPoller = True
                    Exit Function
                End If
                
                If (gdwXpos <= JoystickCal.dwMinXpos) And (gdwYpos >= JoystickCal.dwMaxYpos) Then
                    Call SouthWest_Up
                    gdwXpos = dwXpos
                    gdwYpos = dwYpos
                    EQ_JoystickPoller = True
                    Exit Function
                End If
                
                If (gdwXpos >= JoystickCal.dwMaxXpos) And (gdwYpos <= JoystickCal.dwMinYpos) Then
                    Call NorthEast_Up
                    gdwXpos = dwXpos
                    gdwYpos = dwYpos
                    EQ_JoystickPoller = True
                    Exit Function
                End If
                
                If (gdwXpos >= JoystickCal.dwMaxXpos) And (gdwYpos >= JoystickCal.dwMaxYpos) Then
                    Call SouthEast_Up
                    gdwXpos = dwXpos
                    gdwYpos = dwYpos
                    EQ_JoystickPoller = True
                    Exit Function
                End If
            
                If (dwXpos <= JoystickCal.dwMinXpos) And (dwYpos <= JoystickCal.dwMinYpos) Then
                    Call HC.Add_Message(oLangDll.GetLangString(5126))
                    Call NorthWest_Down(RRATE, DRATE)
                    gdwXpos = dwXpos
                    gdwYpos = dwYpos
                    EQ_JoystickPoller = True
                    Exit Function
                End If
                
                If (dwXpos <= JoystickCal.dwMinXpos) And (dwYpos >= JoystickCal.dwMaxYpos) Then
                    Call HC.Add_Message(oLangDll.GetLangString(5127))
                    Call SouthWest_Down(RRATE, DRATE)
                    gdwXpos = dwXpos
                    gdwYpos = dwYpos
                    EQ_JoystickPoller = True
                    Exit Function
                End If
                
                If (dwXpos >= JoystickCal.dwMaxXpos) And (dwYpos <= JoystickCal.dwMinYpos) Then
                    Call HC.Add_Message(oLangDll.GetLangString(5128))
                    Call NorthEast_Down(RRATE, DRATE)
                    gdwXpos = dwXpos
                    gdwYpos = dwYpos
                    EQ_JoystickPoller = True
                    Exit Function
                End If
                
                If (dwXpos >= JoystickCal.dwMaxXpos) And (dwYpos >= JoystickCal.dwMaxYpos) Then
                    Call HC.Add_Message(oLangDll.GetLangString(5129))
                    Call SouthEast_Down(RRATE, DRATE)
                    gdwXpos = dwXpos
                    gdwYpos = dwYpos
                    EQ_JoystickPoller = True
                    Exit Function
                End If
                

            End If
        
                
            ' Debouncing on X Axis
            
            rate = gPresetSlewRates(1)
            If rate > 0 Then
                If rate >= 1 Then
                    rate = rate + 9
                Else
                    rate = rate * 10
                End If
            End If

            
            If gZoneX <> ZoneX Then
                ' change in zone so stop current slew
                Call Slew_Release_RA
                SlewActive = 0
                ' Decide what to do now
                Select Case ZoneX
                    Case 0
                        ' released
                    Case 1
                        Call Slew_Release_RA
                        SlewActive = 0
                        Call HC.Add_Message(oLangDll.GetLangString(5112))
                        If RRATE > rate Then
                            Call West_Down(CInt(rate))
                        Else
                            Call West_Down(RRATE)
                        End If
                    Case 2
                        Call HC.Add_Message(oLangDll.GetLangString(5112))
                        Call West_Down(RRATE)
                    Case 3
                        Call HC.Add_Message(oLangDll.GetLangString(5111))
                        If RRATE > rate Then
                            Call East_Down(CInt(rate))
                        Else
                            Call East_Down(CInt(RRATE))
                        End If
                    Case 4
                        Call HC.Add_Message(oLangDll.GetLangString(5111))
                        Call East_Down(RRATE)
                End Select
                gZoneX = ZoneX
            End If
            
'            If dwXpos <> gdwXpos Then
'
'                'Scan for Joystick Release here
'
'                If (gdwXpos <= JoystickCal.dwMinXpos) And (dwXpos > JoystickCal.dwMinXpos) Then Call West_Up
'                If (gdwXpos >= JoystickCal.dwMaxXpos) And (dwXpos < JoystickCal.dwMaxXpos) Then Call East_Up
'
'                ' Scan for Joystick Activate Here
'
'                If (dwXpos <= JoystickCal.dwMinXpos) And (gdwXpos > JoystickCal.dwMinXpos) Then
'                    Call HC.Add_Message(oLangDll.GetLangString(5112))
'                    Call West_Down(RRATE)
'                End If
'
'                If (dwXpos >= JoystickCal.dwMaxXpos) And (gdwXpos < JoystickCal.dwMaxXpos) Then
'                    Call HC.Add_Message(oLangDll.GetLangString(5111))
'                    Call East_Down(RRATE)
'                End If
'
                gdwXpos = dwXpos
        
'            End If
        
            ' Debouncing on Y Axis
        
            If gZoneY <> ZoneY Then
                ' change in zone so stop current slew
                Call Slew_Release_DEC
                SlewActive = 0
                ' Decide what to do now
                Select Case ZoneY
                    Case 0
                        ' released
                    Case 1
                        Call HC.Add_Message(oLangDll.GetLangString(5109))
                        If DRATE > rate Then
                            Call North_Down(CInt(rate))
                        Else
                            Call North_Down(DRATE)
                        End If
                    Case 2
                        Call HC.Add_Message(oLangDll.GetLangString(5109))
                        Call North_Down(DRATE)
                    Case 3
                        Call HC.Add_Message(oLangDll.GetLangString(5110))
                        If DRATE > rate Then
                            Call South_Down(CInt(rate))
                        Else
                            Call South_Down(DRATE)
                        End If
                    Case 4
                        Call HC.Add_Message(oLangDll.GetLangString(5110))
                        Call South_Down(DRATE)
                End Select
                gZoneY = ZoneY
            End If
            
'            If dwYpos <> gdwYpos Then
'
'                'Scan for Joystick Release here
'                If (gdwYpos <= JoystickCal.dwMinYpos) And (dwYpos > JoystickCal.dwMinYpos) Then Call North_Up
'                If (gdwYpos >= JoystickCal.dwMaxYpos) And (dwYpos < JoystickCal.dwMaxYpos) Then Call South_Up
'
'                ' Scan for Joystick Activate Here
'                If (dwYpos <= JoystickCal.dwMinYpos) And (gdwYpos > JoystickCal.dwMinYpos) Then
'                    Call HC.Add_Message(oLangDll.GetLangString(5109))
'                    Call North_Down(DRATE)
'                End If
'
'                If (dwYpos >= JoystickCal.dwMaxYpos) And (gdwYpos < JoystickCal.dwMaxYpos) Then
'                    Call HC.Add_Message(oLangDll.GetLangString(5110))
'                    Call South_Down(DRATE)
'                End If
'
                gdwYpos = dwYpos
        
'            End If
            
            ' Debouncing on R Axis
'            If dwRpos <> gdwRpos Then

                'Scan for Joystick Release here
'                If gdwRpos = 0 Then
'                End If
'                If gdwRpos = 65535 Then
'                End If

                ' Scan for Joystick Activate Here
'                If dwRpos = 0 Then
'                End If
'                If dwRpos = 65535 Then
'                End If

'                gdwRpos = dwRpos

'            End If

           ' Debouncing on Z Axis
'            If dwZpos <> gdwZpos Then
'                ' Scan for Joystick Activate Here
'                If gdwZpos = 0 Then
'                End If
'                If gdwZpos = 65535 Then
'                End If
'
'                ' Scan for Joystick Activate Here
'                If dwZpos = 0 Then
'                End If
'                If dwZpos = 65535 Then
'                End If
'
'                gdwZpos = dwZpos
'            End If
            
            ' check for button preses
            Call ButtonHandler(dwButtons, gdwButtons, RRATE, DRATE)
            
            If POV_Enabled Then
                ' Debouncing on the POV Pads
                Select Case dwPOV
                    Case 9000, 27000, 0, 18000, 31500, 4500, 22500, 13500
                        dwPOV = dwPOV + 65536
                    Case Else
                        dwPOV = 0
                End Select
                Call POVHandler(dwPOV, gdwPov, RRATE, DRATE)
            End If
            
        End If
        EQ_JoystickPoller = True

End Function

Public Function EQ_JoystickPoller2() As Boolean
Dim i As Long
Dim dwButtons As Long
Dim dwPOV As Long

        JoystickDat.dwSize = Len(JoystickDat)
        JoystickDat.dwFlags = JOY_RETURNALL
        
        If JoystickCal.id = -1 Then
            'Auto search first two ids
            i = joyGetPosEx(JOYSTICKID1, JoystickDat)
            If i <> JOYERR_NOERROR Then i = joyGetPosEx(JOYSTICKID2, JoystickDat)
        Else
            ' use specific id
            i = joyGetPosEx(JoystickCal.id, JoystickDat)
        End If
        
        If i <> JOYERR_NOERROR Then
            
            ' Joystick not found disable joystick scan
            
            EQ_JoystickPoller2 = False
            Exit Function
            
        End If
        
        ' Start Polling for Joytick routines here
        
        If i = JOYERR_NOERROR Then
        
            dwButtons = JoystickDat.dwButtons
            dwPOV = JoystickDat.dwPOV
        
            ' check for button preses
            If dwButtons <> gdwButtons Then
            
                If BTN_EMERGENCYSTOP <> BTN_UNDEFINED Then
                    If (dwButtons And BTN_EMERGENCYSTOP) = BTN_EMERGENCYSTOP Then
                        Call EmergencyStopPark
                        GoTo skiplock1:
                    End If
                Else
                End If
                
                If BTN_TOGGLELOCK <> BTN_UNDEFINED Then
                    If (dwButtons And BTN_TOGGLELOCK) = BTN_TOGGLELOCK Then
                        If JoystickCal.Enabled Then
                            JStickConfigForm.Check1.Value = 0
                            Call EQ_Beep(33)
                        Else
                            JStickConfigForm.Check1.Value = 1
                            Call EQ_Beep(34)
                        End If
                    Else
                        If dwButtons <> 0 Then
                            Call EQ_Beep(33)
                        End If
                    End If
                End If
skiplock1:
                gdwButtons = dwButtons
            End If
            
            
            If POV_Enabled Then
                ' Debouncing on the POV Pads
                Select Case dwPOV
                    Case 9000, 27000, 0, 18000, 31500, 4500, 22500, 13500
                        dwPOV = dwPOV + 65536
                    Case Else
                        dwPOV = 0
                End Select
                If dwPOV <> gdwPov Then
                    If BTN_EMERGENCYSTOP <> BTN_UNDEFINED Then
                        If dwPOV = BTN_EMERGENCYSTOP Then
                            Call EmergencyStopPark 'Call emergency_stop
                            GoTo skiplock2
                        End If
                    End If
                    
                    If BTN_TOGGLELOCK <> BTN_UNDEFINED Then
                        If (dwPOV = BTN_TOGGLELOCK) Then
                            If JoystickCal.Enabled Then
                                JStickConfigForm.Check1.Value = 0
                                Call EQ_Beep(33)
                            Else
                                JStickConfigForm.Check1.Value = 1
                                Call EQ_Beep(34)
                            End If
                        Else
                            If dwPOV <> 0 Then
                                Call EQ_Beep(33)
                            End If
                        End If
                    End If
skiplock2:
                    gdwPov = dwPOV
                End If
            End If
            
        End If
        EQ_JoystickPoller2 = True

End Function




Public Sub ButtonHandler(ByRef CURRENT As Long, ByRef last As Long, RRATE As Long, DRATE As Long)
    ' Debouncing on Buttons
    
    If CURRENT <> last Then
    
        If BTN_SPIRAL <> BTN_UNDEFINED Then
            If (last And BTN_SPIRAL) Then Call Spiral_Slew_Stop
            If (CURRENT And BTN_SPIRAL) Then
                Call HC.Add_Message(oLangDll.GetLangString(5113))
                Call Spiral_Slew
            End If
        End If
        
        ' alignment buttons are handled on the alignment form
        ' apply a mask
        If CURRENT = BTN_ALIGNACCEPT Or CURRENT = BTN_ALIGNCANCEL Or CURRENT = BTN_ALIGNEND Then
            gEQjbuttons = CURRENT
        End If
        If BTN_STARTSIDREAL <> BTN_UNDEFINED Then
            If (CURRENT And BTN_STARTSIDREAL) = BTN_STARTSIDREAL Then Call Start_sidereal
        End If
        If BTN_STARTLUNAR <> BTN_UNDEFINED Then
            If (CURRENT And BTN_STARTLUNAR) = BTN_STARTLUNAR Then Call Start_Lunar
        End If
        If BTN_STARTSOLAR <> BTN_UNDEFINED Then
            If (CURRENT And BTN_STARTSOLAR) = BTN_STARTSOLAR Then Call Start_Solar
        End If
        If BTN_EMERGENCYSTOP <> BTN_UNDEFINED Then
            If (CURRENT And BTN_EMERGENCYSTOP) = BTN_EMERGENCYSTOP Then Call EmergencyStopPark ' Call emergency_stop
        End If
        If BTN_HOMEPARK <> BTN_UNDEFINED Then
            If (CURRENT And BTN_HOMEPARK) = BTN_HOMEPARK Then Call ParkToHome
        End If
        If BTN_USERPARK <> BTN_UNDEFINED Then
            If (CURRENT And BTN_USERPARK) = BTN_USERPARK Then Call ParkToUser
        End If
        If BTN_CURRENTPARK <> BTN_UNDEFINED Then
            If (CURRENT And BTN_CURRENTPARK) = BTN_CURRENTPARK Then Call ParkToCurrent
        End If
        If BTN_UNPARK <> BTN_UNDEFINED Then
            If (CURRENT And BTN_UNPARK) = BTN_UNPARK Then Call UnPark
        End If
        If BTN_RAREVERSE <> BTN_UNDEFINED Then
            If (CURRENT And BTN_RAREVERSE) = BTN_RAREVERSE Then Call RAReverse
        End If
        If BTN_DECREVERSE <> BTN_UNDEFINED Then
            If (CURRENT And BTN_DECREVERSE) = BTN_DECREVERSE Then Call DecReverse
        End If
        If BTN_CUSTOMTRACKSTART <> BTN_UNDEFINED Then
            If (CURRENT And BTN_CUSTOMTRACKSTART) = BTN_CUSTOMTRACKSTART Then Call Start_CustomTracking2
        End If
        If BTN_INCRATEPRESET <> BTN_UNDEFINED Then
            If (CURRENT And BTN_INCRATEPRESET) = BTN_INCRATEPRESET Then Call ChangeRatePreset(1)
        End If
        If BTN_DECRATEPRESET <> BTN_UNDEFINED Then
            If (CURRENT And BTN_DECRATEPRESET) = BTN_DECRATEPRESET Then Call ChangeRatePreset(-1)
        End If
        If BTN_RATE1 <> BTN_UNDEFINED Then
            If (CURRENT And BTN_RATE1) = BTN_RATE1 Then Call SetRate(1)
        End If
        If BTN_RATE2 <> BTN_UNDEFINED Then
            If (CURRENT And BTN_RATE2) = BTN_RATE2 Then Call SetRate(2)
        End If
        If BTN_RATE3 <> BTN_UNDEFINED Then
            If (CURRENT And BTN_RATE3) = BTN_RATE3 Then Call SetRate(3)
        End If
        If BTN_RATE4 <> BTN_UNDEFINED Then
            If (CURRENT And BTN_RATE4) = BTN_RATE4 Then Call SetRate(4)
        End If
        If BTN_POLARSCOPEALIGN <> BTN_UNDEFINED Then
            If (CURRENT And BTN_POLARSCOPEALIGN) = BTN_POLARSCOPEALIGN Then
                If polarfrm.Visible = True Then
                    gEQjbuttons = BTN_POLARSCOPEALIGN
                End If
            End If
        End If
        
        If (BTN_DEADMANSHANDLE <> BTN_UNDEFINED) Then
            If (last = BTN_DEADMANSHANDLE) Then
                If gSlewStatus Then
                    Call ParkToCurrent
                Else
                    EQ_Beep (32)
                End If
            End If
            If (CURRENT And BTN_DEADMANSHANDLE) = BTN_DEADMANSHANDLE Then Call EQ_Beep(31)
        End If
        
        If BTN_SYNC <> BTN_UNDEFINED Then
            If (CURRENT And BTN_SYNC) = BTN_SYNC Then
                SyncPressCount = SyncPressCount + 1
                If SyncPressCount >= 2 Then
                    Call DoSync
                End If
            End If
        End If
        
        If BTN_TOGGLELOCK <> BTN_UNDEFINED Then
            If (CURRENT And BTN_TOGGLELOCK) = BTN_TOGGLELOCK Then
                If JoystickCal.Enabled Then
                    JStickConfigForm.Check1.Value = 0
                    Call EQ_Beep(33)
               Else
                    JStickConfigForm.Check1.Value = 1
                    Call EQ_Beep(34)
                End If
             End If
        End If
        
        If BTN_TOGGLESCREENSAVER <> BTN_UNDEFINED Then
            If (CURRENT And BTN_TOGGLESCREENSAVER) = BTN_TOGGLESCREENSAVER Then
                 Call ToggleMonitorPower
            End If
        End If
        
        
        If (BTN_SOUTH <> BTN_UNDEFINED) Then
            If (last = BTN_SOUTH) Then Call South_Up
            If CURRENT = BTN_SOUTH Then
                Call HC.Add_Message(oLangDll.GetLangString(5114))
                Call South_Down(DRATE)
            End If
        End If
        If (BTN_EAST <> BTN_UNDEFINED) Then
            If (last = BTN_EAST) Then Call East_Up
            If (CURRENT = BTN_EAST) Then
                Call HC.Add_Message(oLangDll.GetLangString(5115))
                Call East_Down(RRATE)
            End If
        End If
        If (BTN_WEST <> BTN_UNDEFINED) Then
            If last = BTN_WEST Then Call West_Up
            If CURRENT = BTN_WEST Then
                Call HC.Add_Message(oLangDll.GetLangString(5116))
                Call West_Down(RRATE)
            End If
        End If
        If (BTN_NORTH <> BTN_UNDEFINED) Then
            If last = BTN_NORTH Then Call North_Up
            If CURRENT = BTN_NORTH Then
                Call HC.Add_Message(oLangDll.GetLangString(5117))
                Call North_Down(DRATE)
            End If
        End If
        If (BTN_NORTHWEST <> BTN_UNDEFINED) Then
            If last = BTN_NORTHWEST Then Call NorthWest_Up
            If CURRENT = BTN_NORTHWEST Then
'                Call HC.Add_Message(oLangDll.GetLangString(5117))
                Call NorthWest_Down(DRATE, RRATE)
            End If
        End If
        If (BTN_NORTHEAST <> BTN_UNDEFINED) Then
            If last = BTN_NORTHEAST Then Call NorthEast_Up
            If CURRENT = BTN_NORTHEAST Then
'                Call HC.Add_Message(oLangDll.GetLangString(5117))
                Call NorthEast_Down(DRATE, RRATE)
            End If
        End If
        If (BTN_SOUTHWEST <> BTN_UNDEFINED) Then
            If last = BTN_SOUTHWEST Then Call SouthWest_Up
            If CURRENT = BTN_SOUTHWEST Then
'                Call HC.Add_Message(oLangDll.GetLangString(5117))
                Call SouthWest_Down(DRATE, RRATE)
            End If
        End If
        If (BTN_SOUTHEAST <> BTN_UNDEFINED) Then
            If last = BTN_SOUTHEAST Then Call SouthEast_Up
            If CURRENT = BTN_SOUTHEAST Then
'                Call HC.Add_Message(oLangDll.GetLangString(5117))
                Call SouthEast_Down(DRATE, RRATE)
            End If
        End If
        
    
    End If
    
    ' auto repeating buttons here
    ' Slew Rate Adjustment Buttons
        
     If BTN_RARATEINC <> BTN_UNDEFINED Then
         If (CURRENT And BTN_RARATEINC) = BTN_RARATEINC Then Call Adjust_rate(0, 1)
     End If
     If BTN_RARATEDEC <> BTN_UNDEFINED Then
         If (CURRENT And BTN_RARATEDEC) = BTN_RARATEDEC Then Call Adjust_rate(0, -1)
     End If
     If BTN_DECRATEINC <> BTN_UNDEFINED Then
         If (CURRENT And BTN_DECRATEINC) = BTN_DECRATEINC Then Call Adjust_rate(1, 1)
     End If
     If BTN_DECRATEDEC <> BTN_UNDEFINED Then
         If (CURRENT And BTN_DECRATEDEC) = BTN_DECRATEDEC Then Call Adjust_rate(1, -1)
     End If
        
     last = CURRENT
End Sub
Public Sub POVHandler(ByRef CURRENT As Long, ByRef last As Long, RRATE As Long, DRATE As Long)
    ' Debouncing on Buttons
    
    If CURRENT <> last Then
    
        If BTN_SPIRAL <> BTN_UNDEFINED Then
            If (last = BTN_SPIRAL) Then Call Spiral_Slew_Stop
            If (CURRENT = BTN_SPIRAL) Then
                Call HC.Add_Message(oLangDll.GetLangString(5113))
                Call Spiral_Slew
            End If
        End If
        
        ' alignment buttons are handled on the alignment form
        ' apply a mask
        If CURRENT = BTN_ALIGNACCEPT Or CURRENT = BTN_ALIGNCANCEL Or CURRENT = BTN_ALIGNEND Then
                gEQjbuttons = CURRENT
        End If
        
        If BTN_STARTSIDREAL <> BTN_UNDEFINED Then
            If CURRENT = BTN_STARTSIDREAL Then Call Start_sidereal
        End If
        If BTN_STARTLUNAR <> BTN_UNDEFINED Then
            If CURRENT = BTN_STARTLUNAR Then Call Start_Lunar
        End If
        If BTN_STARTSOLAR <> BTN_UNDEFINED Then
            If CURRENT = BTN_STARTSOLAR Then Call Start_Solar
        End If
        If BTN_EMERGENCYSTOP <> BTN_UNDEFINED Then
            If CURRENT = BTN_EMERGENCYSTOP Then Call EmergencyStopPark 'Call emergency_stop
        End If
        If BTN_HOMEPARK <> BTN_UNDEFINED Then
            If CURRENT = BTN_HOMEPARK Then Call ParkToHome
        End If
        If BTN_USERPARK <> BTN_UNDEFINED Then
            If CURRENT = BTN_USERPARK Then Call ParkToUser
        End If
        If BTN_CURRENTPARK <> BTN_UNDEFINED Then
            If CURRENT = BTN_CURRENTPARK Then Call ParkToCurrent
        End If
        If BTN_UNPARK <> BTN_UNDEFINED Then
            If CURRENT = BTN_UNPARK Then Call UnPark
        End If
        If BTN_RAREVERSE <> BTN_UNDEFINED Then
            If CURRENT = BTN_RAREVERSE Then Call RAReverse
        End If
        If BTN_DECREVERSE <> BTN_UNDEFINED Then
            If CURRENT = BTN_DECREVERSE Then Call DecReverse
        End If
        If BTN_CUSTOMTRACKSTART <> BTN_UNDEFINED Then
            If CURRENT = BTN_CUSTOMTRACKSTART Then Call Start_CustomTracking2
        End If
        If BTN_INCRATEPRESET <> BTN_UNDEFINED Then
            If CURRENT = BTN_INCRATEPRESET Then Call ChangeRatePreset(1)
        End If
        If BTN_DECRATEPRESET <> BTN_UNDEFINED Then
            If CURRENT = BTN_DECRATEPRESET Then Call ChangeRatePreset(-1)
        End If
        If BTN_RATE1 <> BTN_UNDEFINED Then
            If CURRENT = BTN_RATE1 Then Call SetRate(1)
        End If
        If BTN_RATE2 <> BTN_UNDEFINED Then
            If CURRENT = BTN_RATE2 Then Call SetRate(2)
        End If
        If BTN_RATE3 <> BTN_UNDEFINED Then
            If CURRENT = BTN_RATE3 Then Call SetRate(3)
        End If
        If BTN_RATE4 <> BTN_UNDEFINED Then
            If CURRENT = BTN_RATE4 Then Call SetRate(4)
        End If
        If BTN_POLARSCOPEALIGN <> BTN_UNDEFINED Then
            If CURRENT = BTN_POLARSCOPEALIGN Then
                If polarfrm.Visible = True Then
                    gEQjbuttons = BTN_POLARSCOPEALIGN
                End If
            End If
        End If
        
        If (BTN_DEADMANSHANDLE <> BTN_UNDEFINED) Then
            If (last = BTN_DEADMANSHANDLE) Then
                If gSlewStatus Then
                    Call ParkToCurrent
                Else
                    EQ_Beep (32)
                End If
            End If
            If CURRENT = BTN_DEADMANSHANDLE Then Call EQ_Beep(31)
        End If
        
        If BTN_SYNC <> BTN_UNDEFINED Then
            If CURRENT = BTN_SYNC Then
                SyncPressCount = SyncPressCount + 1
                If SyncPressCount >= 2 Then
                    Call DoSync
                End If
            End If
        End If
        
        If (BTN_SOUTH <> BTN_UNDEFINED) Then
            If (last = BTN_SOUTH) Then Call South_Up
            If (CURRENT = BTN_SOUTH) Then
                Call HC.Add_Message(oLangDll.GetLangString(5114))
                Call South_Down(DRATE)
            End If
        End If
        If (BTN_EAST <> BTN_UNDEFINED) Then
            If (last = BTN_EAST) Then Call East_Up
            If (CURRENT = BTN_EAST) Then
                Call HC.Add_Message(oLangDll.GetLangString(5115))
                Call East_Down(RRATE)
            End If
        End If
        If (BTN_WEST <> BTN_UNDEFINED) Then
            If (last = BTN_WEST) Then Call West_Up
            If (CURRENT = BTN_WEST) Then
                Call HC.Add_Message(oLangDll.GetLangString(5116))
                Call West_Down(RRATE)
            End If
        End If
        If (BTN_NORTH <> BTN_UNDEFINED) Then
            If (last = BTN_NORTH) Then Call North_Up
            If (CURRENT = BTN_NORTH) Then
                Call HC.Add_Message(oLangDll.GetLangString(5117))
                Call North_Down(DRATE)
            End If
        End If
        If (BTN_NORTHWEST <> BTN_UNDEFINED) Then
            If (last = BTN_NORTHWEST) Then Call NorthWest_Up
            If (CURRENT = BTN_NORTHWEST) Then
'                Call HC.Add_Message(oLangDll.GetLangString(5117))
                Call NorthWest_Down(DRATE, RRATE)
            End If
        End If
        If (BTN_NORTHEAST <> BTN_UNDEFINED) Then
            If (last = BTN_NORTHEAST) Then Call NorthEast_Up
            If (CURRENT = BTN_NORTHEAST) Then
'                Call HC.Add_Message(oLangDll.GetLangString(5117))
                Call NorthEast_Down(DRATE, RRATE)
            End If
        End If
        If (BTN_SOUTHWEST <> BTN_UNDEFINED) Then
            If (last = BTN_SOUTHWEST) Then Call SouthWest_Up
            If (CURRENT = BTN_SOUTHWEST) Then
'                Call HC.Add_Message(oLangDll.GetLangString(5117))
                Call SouthWest_Down(DRATE, RRATE)
            End If
        End If
        If (BTN_SOUTHEAST <> BTN_UNDEFINED) Then
            If (last = BTN_SOUTHEAST) Then Call SouthEast_Up
            If (CURRENT = BTN_SOUTHEAST) Then
'                Call HC.Add_Message(oLangDll.GetLangString(5117))
                Call SouthEast_Down(DRATE, RRATE)
            End If
        End If
        
        If BTN_TOGGLELOCK <> BTN_UNDEFINED Then
            If (CURRENT = BTN_TOGGLELOCK) Then
                If JoystickCal.Enabled Then
                    JStickConfigForm.Check1.Value = 0
                    Call EQ_Beep(33)
                Else
                    JStickConfigForm.Check1.Value = 1
                    Call EQ_Beep(34)
                End If
            End If
        End If
        
        If BTN_TOGGLESCREENSAVER <> BTN_UNDEFINED Then
            If CURRENT = BTN_TOGGLESCREENSAVER Then
                Call ToggleMonitorPower
            End If
        End If
        
    
    End If
    
    ' auto repeating buttons here
    ' Slew Rate Adjustment Buttons
        
     If BTN_RARATEINC <> BTN_UNDEFINED Then
         If CURRENT = BTN_RARATEINC Then Call Adjust_rate(0, 1)
     End If
     If BTN_RARATEDEC <> BTN_UNDEFINED Then
         If CURRENT = BTN_RARATEDEC Then Call Adjust_rate(0, -1)
     End If
     If BTN_DECRATEINC <> BTN_UNDEFINED Then
         If CURRENT = BTN_DECRATEINC Then Call Adjust_rate(1, 1)
     End If
     If BTN_DECRATEDEC <> BTN_UNDEFINED Then
         If CURRENT = BTN_DECRATEDEC Then Call Adjust_rate(1, -1)
     End If
        
     last = CURRENT
End Sub


Public Sub West_Down(rate As Long)
    If HC.RA_inv.Value Then
        Slew_East (rate)
    Else
        Slew_West (rate)
    End If
End Sub

Private Sub Slew_West(rate As Long)

    ' no slewing possible if parked!
    If gEQparkstatus <> 0 Then
        HC.Add_Message (oLangDll.GetLangString(5000))
        Exit Sub
    End If
    
    ' stop pec from sending updates
    StopTrackingUpdates
    
    If rate > 9 Then
        
        ' Stop RA Motor
        eqres = EQ_MotorStop(0)

'        If eqres <> EQ_OK Then GoTo SWEND01
'
        'Wait until RA motor is stable
        Do
            eqres = EQ_GetMotorStatus(0)
            If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo SWEND01
        Loop While (eqres And EQ_MOTORBUSY) <> 0
        
        rate = rate - 9
        If gTrackingStatus Then
            If rate < 800 Then
                rate = rate + 1
            End If
        End If
        If Not gHemisphere = 1 Then
            eqres = EQ_Slew(0, 0, 0, rate)
        Else
            eqres = EQ_Slew(0, 0, 1, rate)
        End If
        RAGuidingNudge = False
    Else
        RAGuidingNudge = True
        eqres = EQ_SendGuideRate(0, 0, rate, 0, gHemisphere, gHemisphere)
    End If
    
   ' Stop Emulation
    gEmulNudge = True
    SlewActive = 7

SWEND01:

End Sub

Public Sub East_Down(rate As Long)
    If HC.RA_inv.Value Then
        Slew_West (rate)
    Else
        Slew_East (rate)
    End If
End Sub


Private Sub Slew_East(rate As Long)

    ' no slewing possible if parked!
    If gEQparkstatus <> 0 Then
        HC.Add_Message (oLangDll.GetLangString(5000))
        Exit Sub
    End If
    
    ' stop pec from sending updates
    StopTrackingUpdates
    
    If rate > 9 Then
        ' Stop RA Motor
        eqres = EQ_MotorStop(0)
'        If eqres <> EQ_OK Then GoTo EDEND01

        'Wait until RA motor is stable
        Do
            eqres = EQ_GetMotorStatus(0)
            If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo EDEND01
        Loop While (eqres And EQ_MOTORBUSY) <> 0
        
        rate = rate - 9
        If gTrackingStatus Then
            ' allow for the fact that sidereal drift gives us a boost in this direction.
            rate = rate - 1
        End If
        
        If rate <> 0 Then
            If Not gHemisphere = 1 Then
                eqres = EQ_Slew(0, 0, 1, rate)
            Else
                eqres = EQ_Slew(0, 0, 0, rate)
            End If
        End If
        RAGuidingNudge = False
    Else
        RAGuidingNudge = True
        eqres = EQ_SendGuideRate(0, 0, rate, 1, gHemisphere, gHemisphere)
    End If
    
    ' Stop Emulation
    gEmulNudge = True
    SlewActive = 3

            
EDEND01:

End Sub

Public Sub North_Down(rate As Long)
    If HC.DEC_Inv.Value = 1 Then
        Slew_South (rate)
    Else
        Slew_North (rate)
    End If
End Sub

Private Sub Slew_North(rate As Long)

    ' no slewing possible if parked!
    If gEQparkstatus <> 0 Then
        HC.Add_Message (oLangDll.GetLangString(5000))
        Exit Sub
    End If
            
    ' stop PEC sending update
    StopTrackingUpdates
    
    If rate > 9 Then
        eqres = EQ_MotorStop(1)          ' Stop DEC Motor
'        If eqres <> EQ_OK Then GoTo NDEND01

        ' Wait for motor stop
        Do
            eqres = EQ_GetMotorStatus(1)
            If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo NDEND01
        Loop While (eqres And EQ_MOTORBUSY) <> 0

        rate = rate - 9
        eqres = EQ_Slew(1, 0, 0, rate)
        DECGuidingNudge = False
    Else
        DECGuidingNudge = True
        eqres = EQ_SendGuideRate(1, 0, rate, 0, 0, 0)
    End If
    gEmulNudge = True               ' Stop Emulation
    SlewActive = 1
NDEND01:

End Sub


Public Sub South_Down(rate As Long)
    If HC.DEC_Inv.Value = 1 Then
        Slew_North (rate)
    Else
        Slew_South (rate)
    End If
End Sub

Private Sub Slew_South(rate As Long)
    ' no slewing possible if parked!
    If gEQparkstatus <> 0 Then
        HC.Add_Message (oLangDll.GetLangString(5000))
        Exit Sub
    End If
    
    ' stop PEC sending update
    StopTrackingUpdates
    
    If rate > 9 Then
        eqres = EQ_MotorStop(1)          ' Stop DEC Motor
'        If eqres <> EQ_OK Then GoTo SDEND01
'
        ' Wait for motor stop
        Do
            eqres = EQ_GetMotorStatus(1)
            If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo SDEND01
        Loop While (eqres And EQ_MOTORBUSY) <> 0
        
        rate = rate - 9
        eqres = EQ_Slew(1, 0, 1, rate)
        DECGuidingNudge = False
    Else
        DECGuidingNudge = True
        eqres = EQ_SendGuideRate(1, 0, rate, 1, 0, 0)
    End If
    
    gEmulNudge = True               ' Stop Emulation
    SlewActive = 5
SDEND01:

End Sub


Public Sub keydown(KeyCode As Integer, RRATE As Long, DRATE As Long)

    If KeyCode = 16 Then Exit Sub

    If gPrevCode = KeyCode Then Exit Sub

    gPrevCode = KeyCode

    Select Case (KeyCode)
    
            Case 38
                Call North_Down(DRATE)
            Case 104
                Call North_Down(DRATE)
            Case 56
                Call North_Down(DRATE)
            Case 27
                Call North_Down(DRATE)
            Case 116
                Call North_Down(DRATE)
    
    
            Case 40
                Call South_Down(DRATE)
            Case 98
                Call South_Down(DRATE)
            Case 75
                Call South_Down(DRATE)
            Case 66
                Call South_Down(DRATE)
    
    
            Case 37
                Call West_Down(RRATE)
            Case 100
                Call West_Down(RRATE)
            Case 85
                Call West_Down(RRATE)
    
    
            Case 39
                Call East_Down(RRATE)
            Case 102
                Call East_Down(RRATE)
            Case 79
                Call East_Down(RRATE)
                
            Case 36
                Call NorthWest_Down(RRATE, DRATE)
            Case 103
                Call NorthWest_Down(RRATE, DRATE)
            Case 55
                Call NorthWest_Down(RRATE, DRATE)
                
                
            Case 33
                Call NorthEast_Down(RRATE, DRATE)
'                 Call West_Down(RRATE)
            Case 105
                Call NorthEast_Down(RRATE, DRATE)
            Case 57
                Call NorthEast_Down(RRATE, DRATE)
                
                
            Case 35
                Call SouthWest_Down(RRATE, DRATE)
            Case 97
                Call SouthWest_Down(RRATE, DRATE)
            Case 74
                Call SouthWest_Down(RRATE, DRATE)
                
                
            Case 34
                Call SouthEast_Down(RRATE, DRATE)
'                 Call East_Down(RRATE)
            Case 99
                Call SouthEast_Down(RRATE, DRATE)
            Case 76
                Call SouthEast_Down(RRATE, DRATE)
                
                
            Case 12
                Call emergency_stop
            Case 101
                Call emergency_stop
            Case 73
                Call emergency_stop
                
            Case 45
                Call Start_sidereal
            Case 96
                Call Start_sidereal
            Case 77
                Call Start_sidereal
            
            Case 177, 109 ' Presenter Button 2
                Call ChangeRatePreset(-1)
'               Call Adjust_rate2(-1)
                If Slewpad.Visible Then
                    Slewpad.SetFocus
                End If
            
            Case 176, 107 ' Presenter Button 3
                Call ChangeRatePreset(1)
'               Call Adjust_rate2(1)
                If Slewpad.Visible Then
                    Slewpad.SetFocus
                End If

            Case 106
                Call Spiral_Slew
                
            Case Else
                eqres = 0
    End Select

End Sub

Public Sub keyup(KeyCode As Integer)

    If KeyCode = 16 Then Exit Sub

    Select Case (KeyCode)
    
            Case 38
                Call North_Up
            Case 104
                Call North_Up
            Case 56
                Call North_Up
            Case 27
                Call North_Up
            Case 116
                Call North_Up
   
    
            Case 40
                Call South_Up
            Case 98
                Call South_Up
            Case 75
                Call South_Up
            Case 66
                Call South_Up
    
    
            Case 37
                Call West_Up
            Case 100
                Call West_Up
            Case 85
                Call West_Up
    
            Case 39
                Call East_Up
            Case 102
                Call East_Up
            Case 79
                Call East_Up
                
                
            Case 36
                Call NorthWest_Up
            Case 103
                Call NorthWest_Up
            Case 55
                Call NorthWest_Up
                
                
            Case 33
'                Call NorthEast_Up
                 Call West_Up
            Case 105
                Call NorthEast_Up
            Case 57
                Call NorthEast_Up
                
            Case 35
                Call SouthWest_Up
            Case 97
                Call SouthWest_Up
            Case 77
                Call SouthWest_Up
                
            Case 34
'                Call SouthEast_Up
                 Call East_Up
                 
            Case 99
                Call SouthEast_Up
            Case 76
                Call SouthEast_Up
                
            Case 106
                Call Spiral_Slew_Stop

            Case Else
                eqres = 0
    End Select
    
    gPrevCode = 0

End Sub

Public Sub NorthEast_Down(RRATE As Long, DRATE As Long)
    If HC.DEC_Inv.Value = 1 Then
        If HC.RA_inv.Value = 1 Then
            Call Slew_SouthWest(RRATE, DRATE)
        Else
            Call Slew_SouthEast(RRATE, DRATE)
        End If
    Else
        If HC.RA_inv.Value = 1 Then
            Call Slew_NorthWest(RRATE, DRATE)
        Else
            Call Slew_NorthEast(RRATE, DRATE)
        End If
    End If
End Sub

Private Sub Slew_NorthEast(RRATE As Long, DRATE As Long)

    If gEQparkstatus <> 0 Then
        HC.Add_Message (oLangDll.GetLangString(5000))
        Exit Sub
    End If

    ' stop PEC sending update
    StopTrackingUpdates
    
    If RRATE > 9 Then
         ' Stop RA Motor
         eqres = EQ_MotorStop(0)
'         If eqres <> EQ_OK Then GoTo NEEND01
'
         ' Wait for RA motor stop
         Do
             eqres = EQ_GetMotorStatus(0)
             If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo NEEND01
         Loop While (eqres And EQ_MOTORBUSY) <> 0
    
        RRATE = RRATE - 9
        If gTrackingStatus Then
            ' allow for the fact that sidereal drift gives us a boost in this direction.
            RRATE = RRATE - 1
        End If
        If RRATE > 0 Then
            eqres = EQ_Slew(0, 0, 1, RRATE)
        End If
        RAGuidingNudge = False
    Else
        RAGuidingNudge = True
        eqres = EQ_SendGuideRate(0, 0, RRATE, 1, gHemisphere, gHemisphere)
    End If
    
    If DRATE > 9 Then
        ' Stop DEC Motor
        eqres = EQ_MotorStop(1)
'        If eqres <> EQ_OK Then GoTo NEEND01
'
        ' Wait for DEC motor stop
        Do
            eqres = EQ_GetMotorStatus(1)
            If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo NEEND01
        Loop While (eqres And EQ_MOTORBUSY) <> 0
       
        eqres = EQ_Slew(1, 0, 0, DRATE - 9)
        DECGuidingNudge = False
    Else
        DECGuidingNudge = True
        eqres = EQ_SendGuideRate(1, 0, DRATE, 0, 0, 0)
    End If
    
    gEmulNudge = True               ' Stop Emulation
    SlewActive = 2
    
NEEND01:

End Sub

Public Sub NorthWest_Down(RRATE As Long, DRATE As Long)
    If HC.DEC_Inv.Value = 1 Then
        If HC.RA_inv.Value = 1 Then
            Call Slew_SouthEast(RRATE, DRATE)
        Else
            Call Slew_SouthWest(RRATE, DRATE)
        End If
    Else
        If HC.RA_inv.Value = 1 Then
            Call Slew_NorthEast(RRATE, DRATE)
        Else
            Call Slew_NorthWest(RRATE, DRATE)
        End If
    End If
End Sub


Private Sub Slew_NorthWest(RRATE As Long, DRATE As Long)

    ' no slewing possible if parked!
    If gEQparkstatus <> 0 Then
        HC.Add_Message (oLangDll.GetLangString(5000))
        Exit Sub
    End If

    ' stop PEC sending update
    StopTrackingUpdates
    
    If RRATE > 9 Then
        ' Stop RA Motor
        eqres = EQ_MotorStop(0)
'        If eqres <> EQ_OK Then GoTo NWEND01

        ' Wait for motor stop
        Do
            eqres = EQ_GetMotorStatus(0)
            If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo NWEND01
        Loop While (eqres And EQ_MOTORBUSY) <> 0
    
        RRATE = RRATE - 9
        If gTrackingStatus Then
            If RRATE < 800 Then
                RRATE = RRATE + 1
            End If
        End If
        eqres = EQ_Slew(0, 0, 0, RRATE)
        RAGuidingNudge = False
    Else
        RAGuidingNudge = True
        eqres = EQ_SendGuideRate(0, 0, RRATE, 0, gHemisphere, gHemisphere)
    End If
    
    If DRATE > 9 Then
        ' Stop DEC Motor
        eqres = EQ_MotorStop(1)
'        If eqres <> EQ_OK Then GoTo NWEND01

        ' Wait for motor stop
        Do
           eqres = EQ_GetMotorStatus(1)
            If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo NWEND01
        Loop While (eqres And EQ_MOTORBUSY) <> 0

        eqres = EQ_Slew(1, 0, 0, DRATE - 9)
        DECGuidingNudge = False
    Else
        DECGuidingNudge = True
        eqres = EQ_SendGuideRate(1, 0, DRATE, 0, 0, 0)
    End If
    
    gEmulNudge = True               ' Stop Emulation
    SlewActive = 8
            
NWEND01:

End Sub

Public Sub SouthEast_Down(RRATE As Long, DRATE As Long)
    If HC.DEC_Inv.Value = 1 Then
        If HC.RA_inv.Value = 1 Then
            Call Slew_NorthWest(RRATE, DRATE)
        Else
            Call Slew_NorthEast(RRATE, DRATE)
        End If
    Else
        If HC.RA_inv.Value = 1 Then
            Call Slew_SouthWest(RRATE, DRATE)
        Else
            Call Slew_SouthEast(RRATE, DRATE)
        End If
    End If
End Sub

Private Sub Slew_SouthEast(RRATE As Long, DRATE As Long)

    ' no sleing possible if parked!
    If gEQparkstatus <> 0 Then
        HC.Add_Message (oLangDll.GetLangString(5000))
        Exit Sub
    End If
    
    ' stop PEC sending update
    StopTrackingUpdates
    
    If RRATE > 9 Then
        ' Stop RA Motor
        eqres = EQ_MotorStop(0)
'        If eqres <> EQ_OK Then GoTo SEEND01
'
        ' Wait for motor stop
        Do
             eqres = EQ_GetMotorStatus(0)
            If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo SEEND01
        Loop While (eqres And EQ_MOTORBUSY) <> 0
    
        RRATE = RRATE - 9
        If gTrackingStatus Then
            ' allow for the fact that sidereal drift gives us a boost in this direction.
            RRATE = RRATE - 1
        End If
        If RRATE > 0 Then
            eqres = EQ_Slew(0, 0, 1, RRATE)
        End If
        DECGuidingNudge = True
    Else
        RAGuidingNudge = True
        eqres = EQ_SendGuideRate(0, 0, RRATE, 1, gHemisphere, gHemisphere)
    End If
    
    If DRATE > 9 Then
        ' Stop DEC Motor
        eqres = EQ_MotorStop(1)
'        If eqres <> EQ_OK Then GoTo SEEND01
'
        ' Wait for motor stop
        Do
            eqres = EQ_GetMotorStatus(1)
            If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo SEEND01
        Loop While (eqres And EQ_MOTORBUSY) <> 0

        eqres = EQ_Slew(1, 0, 1, DRATE - 9)
        DECGuidingNudge = False
    Else
        DECGuidingNudge = True
        eqres = EQ_SendGuideRate(1, 0, DRATE, 1, 0, 0)
    End If
    
    gEmulNudge = True               ' Stop Emulation
    SlewActive = 4

SEEND01:

End Sub

Public Sub SouthWest_Down(RRATE As Long, DRATE As Long)
    If HC.DEC_Inv.Value = 1 Then
        If HC.RA_inv.Value = 1 Then
            Call Slew_NorthEast(RRATE, DRATE)
        Else
            Call Slew_NorthWest(RRATE, DRATE)
        End If
    Else
        If HC.RA_inv.Value = 1 Then
            Call Slew_SouthEast(RRATE, DRATE)
        Else
            Call Slew_SouthWest(RRATE, DRATE)
        End If
    End If
End Sub

Private Sub Slew_SouthWest(RRATE As Long, DRATE As Long)

    ' no sleing possible if parked!
    If gEQparkstatus <> 0 Then
        HC.Add_Message (oLangDll.GetLangString(5000))
        Exit Sub
    End If
    
    ' stop PEC sending update
    StopTrackingUpdates
    
    If RRATE > 9 Then
       ' Stop RA Motor
        eqres = EQ_MotorStop(0)
'        If eqres <> EQ_OK Then GoTo SWEND01
'
        ' Wait for ra motor stop
        Do
            eqres = EQ_GetMotorStatus(0)
            If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo SWEND01
        Loop While (eqres And EQ_MOTORBUSY) <> 0
        
        RRATE = RRATE - 9
        If gTrackingStatus Then
            If RRATE < 800 Then
                RRATE = RRATE + 1
            End If
        End If
        eqres = EQ_Slew(0, 0, 0, RRATE)
        RAGuidingNudge = False
    Else
        RAGuidingNudge = True
        eqres = EQ_SendGuideRate(0, 0, RRATE, 0, gHemisphere, gHemisphere)
    End If
    
    If DRATE > 9 Then
        ' Stop DEC Motor
        eqres = EQ_MotorStop(1)
'        If eqres <> EQ_OK Then GoTo SWEND01
        ' Wait for dec motor stop
        Do
            eqres = EQ_GetMotorStatus(1)
            If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo SWEND01
        Loop While (eqres And EQ_MOTORBUSY) <> 0
    
        eqres = EQ_Slew(1, 0, 1, DRATE - 9)
        DECGuidingNudge = False
    Else
        DECGuidingNudge = True
        eqres = EQ_SendGuideRate(1, 0, DRATE, 1, 0, 0)
    End If
    
    gEmulNudge = True               ' Stop Emulation
    SlewActive = 6

SWEND01:

End Sub

Public Sub North_Up()
    Call Slew_Release_DEC
    SlewActive = 0
End Sub

Public Sub South_Up()
    Call Slew_Release_DEC
    SlewActive = 0
End Sub

Public Sub East_Up()
    Call Slew_Release_RA
    SlewActive = 0
End Sub

Public Sub West_Up()
    Call Slew_Release_RA
    SlewActive = 0
End Sub

Public Sub NorthEast_Up()
    Call Slew_Release
    SlewActive = 0
End Sub

Public Sub NorthWest_Up()
    Call Slew_Release
    SlewActive = 0
End Sub

Public Sub SouthEast_Up()
    Call Slew_Release
    SlewActive = 0
End Sub

Public Sub SouthWest_Up()
    Call Slew_Release
    SlewActive = 0
End Sub

Private Sub Slew_Release_RA()

    ' no slewing possible if parked!
    If gEQparkstatus <> 0 Then
        HC.Add_Message (oLangDll.GetLangString(5000))
        Exit Sub
    End If
    
    If RAGuidingNudge Then
        Call EQ_SendGuideRate(0, gTrackingStatus - 1, 0, 0, gHemisphere, gHemisphere)
        If HC.CheckPEC.Value = 1 Then
            PEC_StartTracking
        End If
    Else
        ' Stop Motors
        eqres = EQ_MotorStop(1)
        eqres = EQ_MotorStop(0)
    
        'Wait until RA motor is stable
        Do
            eqres = EQ_GetMotorStatus(0)
            If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo SRRA1
        Loop While (eqres And EQ_MOTORBUSY) <> 0

SRRA1:
        Do
            eqres = EQ_GetMotorStatus(1)
            If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo SRRA2
        Loop While (eqres And EQ_MOTORBUSY) <> 0
        
SRRA2:
        Call RestartTracking
    End If
    RAGuidingNudge = False
    
    gEmulNudge = False               ' Enable Emulation
    gEmulOneShot = True              ' Get One shot cap
    

End Sub

Private Sub Slew_Release_DEC()

    ' no slewing possible if parked!
    If gEQparkstatus <> 0 Then
        HC.Add_Message (oLangDll.GetLangString(5000))
        Exit Sub
    End If
    
    If DECGuidingNudge Then
        Call EQ_MotorStop(1)
    Else
        ' Stop Motors
        eqres = EQ_MotorStop(1)
        eqres = EQ_MotorStop(0)
    
        'Wait until RA motor is stable
        Do
            eqres = EQ_GetMotorStatus(0)
            If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo SRDEC1
        Loop While (eqres And EQ_MOTORBUSY) <> 0

SRDEC1:
        Do
            eqres = EQ_GetMotorStatus(1)
            If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo SRDEC2
        Loop While (eqres And EQ_MOTORBUSY) <> 0
        
SRDEC2:
        Call RestartTracking
    End If
    DECGuidingNudge = False
    
    gEmulNudge = False               ' Enable Emulation
    gEmulOneShot = True              ' Get One shot cap

End Sub


Private Sub Slew_Release()

    ' no slewing possible if parked!
    If gEQparkstatus <> 0 Then
        HC.Add_Message (oLangDll.GetLangString(5000))
        Exit Sub
    End If
    
    eqres = EQ_MotorStop(1)
    eqres = EQ_MotorStop(0)          ' Stop RA Motor

    'Wait until RA motor is stable

    Do
        eqres = EQ_GetMotorStatus(0)
        If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo SR1
    Loop While (eqres And EQ_MOTORBUSY) <> 0

SR1:
    Do
        eqres = EQ_GetMotorStatus(1)
        If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo SR2
    Loop While (eqres And EQ_MOTORBUSY) <> 0
        
SR2:
    Call RestartTracking
    
    RAGuidingNudge = False
    DECGuidingNudge = False
    gEmulNudge = False               ' Enable Emulation
    gEmulOneShot = True              ' Get One shot cap

End Sub


Public Sub emergency_stop()
    
    gSlewStatus = False
    
    If gEQparkstatus = 2 Then
        ' we were slewing to park position
        ' well its not happening now!
        gEQparkstatus = 0
        HC.ParkTimer.Enabled = False
        HC.Frame15.Caption = oLangDll.GetLangString(146) & " " & oLangDll.GetLangString(179)
        Call SetParkCaption
    End If
    
    If gPEC_Enabled Then
       PEC_StopTracking
    End If
    
    eqres = EQ_MotorStop(0)
    eqres = EQ_MotorStop(1)
    
    gRA_LastRate = 0
    Do
        eqres = EQ_GetMotorStatus(0)
        If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo STOPEND1
    Loop While (eqres And EQ_MOTORBUSY) <> 0

STOPEND1:
    Do
        eqres = EQ_GetMotorStatus(1)
        If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo STOPEND2
    Loop While (eqres And EQ_MOTORBUSY) <> 0
    
STOPEND2:
    ' clear an active flips
    HC.ChkForceFlip.Value = 0
    gCWUP = False
    gGotoParams.SuperSafeMode = 0
    
    gRAStatus_slew = False
    gTrackingStatus = 0
    gDeclinationRate = 0
    gRightAscensionRate = 0
    HC.TrackingFrame.Caption = oLangDll.GetLangString(121) & " " & oLangDll.GetLangString(178)
    HC.Add_Message (oLangDll.GetLangString(5130))
    
    gEmulNudge = False               ' Enable Emulation
    gEmulOneShot = True              ' Get One shot cap
    
    EQ_Beep (7)

End Sub

Public Sub Start_sidereal()
    EQStartSidereal2
    gEmulNudge = False               ' Enable Emulation
End Sub

Public Sub Start_Lunar()
        
    gRA_LastRate = 0
    If gEQparkstatus <> 0 Then
        HC.Add_Message (oLangDll.GetLangString(5013))
        Exit Sub
    End If
    
    If gPEC_Enabled Then
       PEC_StopTracking
    End If
    
    eqres = EQ_StartRATrack(1, gHemisphere, gHemisphere)
    eqres = EQ_MotorStop(1)
    gTrackingStatus = 2                 'Lunar rate tracking'
    gDeclinationRate = 0
    gRightAscensionRate = LUN_RATE
    
    HC.TrackingFrame.Caption = oLangDll.GetLangString(121) & " " & oLangDll.GetLangString(123)
    HC.Add_Message (oLangDll.GetLangString(5015))
    gEmulNudge = False               ' Enable Emulation
    EQ_Beep (11)

End Sub

Public Sub Start_Solar()

    gRA_LastRate = 0
    If gEQparkstatus <> 0 Then
        HC.Add_Message (oLangDll.GetLangString(5013))
        Exit Sub
    End If
    
    If gPEC_Enabled Then
       PEC_StopTracking
    End If
    
    eqres = EQ_StartRATrack(2, gHemisphere, gHemisphere)
    eqres = EQ_MotorStop(1)
    
    gTrackingStatus = 3                 'Solar rate tracking'
    gDeclinationRate = 0
    gRightAscensionRate = SOL_RATE
    
    HC.TrackingFrame.Caption = oLangDll.GetLangString(121) & " " & oLangDll.GetLangString(124)
    HC.Add_Message (oLangDll.GetLangString(5016))

    gEmulNudge = False               ' Enable Emulation
    EQ_Beep (12)

End Sub

Public Sub Adjust_rate(axis As Integer, direction As Integer)

Dim i As Integer
Dim j As Integer

    If axis = 0 Then
        i = HC.VScrollRASlewRate.Value
        
        If i > 50 Then
            j = 20
        Else
            j = 1
        End If
        
        If direction > 0 Then
            i = i + j
            If i >= 800 Then i = 800
        Else
            i = i - j
            If i <= 0 Then i = 2
        End If
        HC.VScrollRASlewRate.Value = i
    Else
        i = HC.VScrollDecSlewRate.Value
        
        If i > 50 Then
            j = 20
        Else
            j = 1
        End If
        
        If direction > 0 Then
            i = i + j
            If i >= 800 Then i = 800
        Else
            i = i - j
            If i <= 0 Then i = 2
        End If
        HC.VScrollDecSlewRate.Value = i
    End If
    
    Call ReApplySlew
End Sub

Private Sub ReApplySlew()

    Select Case SlewActive
        Case 0
            'none
        Case 1
            ' north
            Call Slew_North(HC.VScrollDecSlewRate.Value)
        Case 2
            ' northeast
            Call Slew_NorthEast(HC.VScrollRASlewRate.Value, HC.VScrollDecSlewRate.Value)
        Case 3
            ' east
            Call Slew_East(HC.VScrollRASlewRate.Value)
        Case 4
            ' southeast
            Call Slew_SouthEast(HC.VScrollRASlewRate.Value, HC.VScrollDecSlewRate.Value)
        Case 5
            'south
            Call Slew_South(HC.VScrollDecSlewRate.Value)
        Case 6
            'southwest
            Call Slew_SouthWest(HC.VScrollRASlewRate.Value, HC.VScrollDecSlewRate.Value)
        Case 7
            ' west
            Call Slew_West(HC.VScrollRASlewRate.Value)
        Case 8
            ' northwest
            Call Slew_NorthWest(HC.VScrollRASlewRate.Value, HC.VScrollDecSlewRate.Value)
    End Select
    
End Sub

Public Sub Adjust_rate2(direction As Integer)

Dim i As Integer

      i = Slewpad.VScroll1.Value
      If direction > 0 Then
                i = i + 20
                If i >= 800 Then i = 800
      Else
                i = i - 20
                If i <= 0 Then i = 2
      End If
      Slewpad.VScroll1.Value = i
      Slewpad.VScroll2.Value = i
        
End Sub

Public Sub Spiral_Slew()

    ' no slewing possible if parked!
    If gEQparkstatus <> 0 Then
        HC.Add_Message (oLangDll.GetLangString(5000))
        Exit Sub
    End If
    
    ' stop pec from senidng rate updates
    StopTrackingUpdates
  
    eqres = EQ_MotorStop(0)
    eqres = EQ_MotorStop(1)
    
    Do
        eqres = EQ_GetMotorStatus(0)
        If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo SSLEW01
    Loop While (eqres And EQ_MOTORBUSY) <> 0

SSLEW01:

    Do
        eqres = EQ_GetMotorStatus(1)
        If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo SSLEW02
    Loop While (eqres And EQ_MOTORBUSY) <> 0

SSLEW02:

' Initialize Slew Parameters

    gSPIRAL_JUMP = HC.SpiralHScroll1.Value

    gRightAscension_Start = EQGetMotorValues(0)
    gDeclination_Start = EQGetMotorValues(1)
    gDeclination_Dir = 0
    gRightAscension_Dir = 0
    gDeclination_Len = gSPIRAL_JUMP
    gRightAscension_Len = gSPIRAL_JUMP
    gSpiral_AxisFlag = 0
   
    If gRightAscension_Dir = 0 Then
        eqres = EQStartMoveMotor(0, 0, 0, gRightAscension_Len, GetSlowdown(gRightAscension_Len))
        gRightAscension_Dir = 1
    Else
        eqres = EQStartMoveMotor(0, 0, 1, gRightAscension_Len, GetSlowdown(gRightAscension_Len))
        gRightAscension_Dir = 0
    End If
    
    gSpiralTimerFlag = True
    HC.Spiral_Timer.Enabled = True

    gEmulNudge = True               ' Stop Emulation

End Sub

Public Sub Spiral_Slew_Stop()

    ' no slewing possible if parked!
    If gEQparkstatus <> 0 Then
        HC.Add_Message (oLangDll.GetLangString(5000))
        Exit Sub
    End If
    
    gSpiralTimerFlag = False
    HC.Spiral_Timer.Enabled = False

    eqres = EQ_MotorStop(0)
    eqres = EQ_MotorStop(1)
    
    Do
        eqres = EQ_GetMotorStatus(0)
        If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo SSLEWSTOP01
    Loop While (eqres And EQ_MOTORBUSY) <> 0

SSLEWSTOP01:
    Do
        eqres = EQ_GetMotorStatus(1)
        If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo SSLEWSTOP02
    Loop While (eqres And EQ_MOTORBUSY) <> 0

SSLEWSTOP02:
    Call RestartTracking
    gEmulNudge = False               ' Enable Emulation
    gEmulOneShot = True              ' Get One shot cap
        
    HC.Add_Message (oLangDll.GetLangString(5131))
        
End Sub
Public Sub ParkToCurrent()
    ' no slewing possible if parked!
    If gEQparkstatus <> 0 Then
        HC.Add_Message (oLangDll.GetLangString(5000))
        Exit Sub
    End If
    ' do the park
    Call Park2Current
    
End Sub

Public Sub ParkToHome()
    ' no slewing possible if parked!
    If gEQparkstatus <> 0 Then
        HC.Add_Message (oLangDll.GetLangString(5000))
        Exit Sub
    End If
    Call ParkHome
End Sub

Public Sub ParkToUser()
    ' no slewing possible if parked!
    If gEQparkstatus <> 0 Then
        HC.Add_Message (oLangDll.GetLangString(5000))
        Exit Sub
    End If
    Call ParktoUserDefine2(UserParks(1))
End Sub

Public Sub UnPark()
    Call Unparkscope
End Sub

Public Sub RAReverse()
    If HC.RA_inv.Value = 0 Then
        HC.RA_inv.Value = 1
    Else
        HC.RA_inv.Value = 0
    End If
End Sub
Public Sub DecReverse()
    If HC.DEC_Inv.Value = 0 Then
        HC.DEC_Inv.Value = 1
    Else
        HC.DEC_Inv.Value = 0
    End If
End Sub

Public Function ChangeRatePreset(Shift As Integer)
Dim newrate As Double

    On Error GoTo errhandle

    gCurrentRatePreset = gCurrentRatePreset + Shift
    If gCurrentRatePreset > gPresetSlewRatesCount Then
        ' no more preset in this direction
        gCurrentRatePreset = gPresetSlewRatesCount
        Call EQ_Beep(30)
    Else
        If gCurrentRatePreset < 1 Then
            ' no more preset in this direction
            gCurrentRatePreset = 1
            Call EQ_Beep(30)
        Else
            ' make a click so the user knows the press has been actioned
        End If
    End If
    Call EQ_Beep(100 + gCurrentRatePreset)
    
    newrate = gPresetSlewRates(gCurrentRatePreset)
    ' set the new rates
    If newrate > 0 And newrate <= 800 Then
        If newrate < 1 Then
            newrate = newrate * 10
        Else
            newrate = newrate + 9
        End If
        
        HC.VScrollRASlewRate.Value = newrate
        HC.VScrollDecSlewRate.Value = newrate
    End If
    HC.PresetRateCombo.ListIndex = gCurrentRatePreset - 1
    HC.PresetRate2Combo.ListIndex = gCurrentRatePreset - 1

    Call ReApplySlew

errhandle:
End Function

Public Function SetRate(rate As Integer)
Dim newrate As Integer

    If rate <= HC.PresetRateCombo.ListCount Then
        gCurrentRatePreset = gRateButtons(rate)
        newrate = gPresetSlewRates(gCurrentRatePreset)
        EQ_Beep (100 + gCurrentRatePreset)
     
        ' set the new rates
        If newrate > 0 And newrate <= 800 Then
            If newrate < 1 Then
                newrate = newrate * 10
            Else
                newrate = newrate + 9
            End If
            
            HC.VScrollRASlewRate.Value = newrate
            HC.VScrollDecSlewRate.Value = newrate
        End If
        HC.PresetRateCombo.ListIndex = gCurrentRatePreset - 1
        HC.PresetRate2Combo.ListIndex = gCurrentRatePreset - 1
    
        Call ReApplySlew
    End If
End Function

Public Sub DoSync()
    If gTargetRA <> EQ_INVALIDCOORDINATE And gTargetDec <> EQ_INVALIDCOORDINATE Then
        HC.Add_Message ("SyncTaget: " & oLangDll.GetLangString(105) & "[ " & FmtSexa(gTargetRA, False) & "] " & oLangDll.GetLangString(106) & "[ " & FmtSexa(gTargetDec, True) & " ]")
        SyncToRADEC gTargetRA, gTargetDec, gLongitude, gHemisphere
        ' force a beep - sounds even if user has selected sounds to be off
        EQ_Beep (4)
    End If
End Sub

Public Sub LoadJoystickBtns()

Dim tmptxt As String
Dim VarStr As String
Dim key As String
Dim Ini As String

    ' set up a file path for the align.ini file
    Ini = HC.oPersist.GetIniPath & "\JOYSTICK.ini"

    key = "[buttondefs]"

    tmptxt = HC.oPersist.ReadIniValueEx("StartSidreal", key, Ini)
    If tmptxt <> "" Then
        BTN_STARTSIDREAL = val(tmptxt)
    Else
        tmptxt = CStr(BTN_JOY10)
        Call HC.oPersist.WriteIniValueEx("StartSidreal", tmptxt, key, Ini)
        BTN_STARTSIDREAL = BTN_JOY10
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("StartPEC", key, Ini)
    If tmptxt <> "" Then
        BTN_PEC = val(tmptxt)
    Else
        tmptxt = CStr(BTN_UNDEFINED)
        Call HC.oPersist.WriteIniValueEx("StartPEC", tmptxt, key, Ini)
        BTN_STARTSIDREAL = BTN_UNDEFINED
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("StartCustom", key, Ini)
    If tmptxt <> "" Then
        BTN_CUSTOMTRACKSTART = val(tmptxt)
    Else
        tmptxt = CStr(BTN_UNDEFINED)
        Call HC.oPersist.WriteIniValueEx("StartCustom", tmptxt, key, Ini)
        BTN_CUSTOMTRACKSTART = BTN_UNDEFINED
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("StartLunar", key, Ini)
    If tmptxt <> "" Then
        BTN_STARTLUNAR = val(tmptxt)
    Else
        tmptxt = CStr(BTN_UNDEFINED)
        Call HC.oPersist.WriteIniValueEx("StartLunar", tmptxt, key, Ini)
        BTN_STARTLUNAR = BTN_UNDEFINED
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("StartSolar", key, Ini)
    If tmptxt <> "" Then
        BTN_STARTSOLAR = val(tmptxt)
    Else
        tmptxt = CStr(BTN_UNDEFINED)
        Call HC.oPersist.WriteIniValueEx("StartSolar", tmptxt, key, Ini)
        BTN_STARTSOLAR = BTN_UNDEFINED
    End If
    
    tmptxt = HC.oPersist.ReadIniValueEx("SpiralSearch", key, Ini)
    If tmptxt <> "" Then
        BTN_SPIRAL = val(tmptxt)
    Else
        tmptxt = CStr(BTN_JOY1)
        Call HC.oPersist.WriteIniValueEx("SpiralSearch", tmptxt, key, Ini)
        BTN_SPIRAL = BTN_JOY1
    End If
    
    tmptxt = HC.oPersist.ReadIniValueEx("EmergencyStop", key, Ini)
    If tmptxt <> "" Then
        BTN_EMERGENCYSTOP = val(tmptxt)
    Else
        tmptxt = CStr(BTN_JOY11)
        Call HC.oPersist.WriteIniValueEx("EmergencyStop", tmptxt, key, Ini)
        BTN_EMERGENCYSTOP = BTN_JOY11
    End If
    
    tmptxt = HC.oPersist.ReadIniValueEx("RARateInc", key, Ini)
    If tmptxt <> "" Then
        BTN_RARATEINC = val(tmptxt)
    Else
        tmptxt = CStr(BTN_JOY5)
        Call HC.oPersist.WriteIniValueEx("RARateInc", tmptxt, key, Ini)
        BTN_RARATEINC = BTN_JOY5
    End If
    
    tmptxt = HC.oPersist.ReadIniValueEx("RARateDec", key, Ini)
    If tmptxt <> "" Then
        BTN_RARATEDEC = val(tmptxt)
    Else
        tmptxt = CStr(BTN_JOY7)
        Call HC.oPersist.WriteIniValueEx("RARateDec", tmptxt, key, Ini)
        BTN_RARATEDEC = BTN_JOY7
    End If
    
    tmptxt = HC.oPersist.ReadIniValueEx("DecRateInc", key, Ini)
    If tmptxt <> "" Then
        BTN_DECRATEINC = val(tmptxt)
    Else
        tmptxt = CStr(BTN_JOY6)
        Call HC.oPersist.WriteIniValueEx("DecRateInc", tmptxt, key, Ini)
        BTN_DECRATEINC = BTN_JOY6
    End If
    
    tmptxt = HC.oPersist.ReadIniValueEx("DecRateDec", key, Ini)
    If tmptxt <> "" Then
        BTN_DECRATEDEC = val(tmptxt)
    Else
        tmptxt = CStr(BTN_JOY8)
        Call HC.oPersist.WriteIniValueEx("DecRateDec", tmptxt, key, Ini)
        BTN_DECRATEDEC = BTN_JOY8
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("ParkHome", key, Ini)
    If tmptxt <> "" Then
        BTN_HOMEPARK = val(tmptxt)
    Else
        tmptxt = CStr(BTN_UNDEFINED)
        Call HC.oPersist.WriteIniValueEx("ParkHome", tmptxt, key, Ini)
        BTN_HOMEPARK = BTN_UNDEFINED
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("ParkUser", key, Ini)
    If tmptxt <> "" Then
        BTN_USERPARK = val(tmptxt)
    Else
        tmptxt = CStr(BTN_UNDEFINED)
        Call HC.oPersist.WriteIniValueEx("ParkUser", tmptxt, key, Ini)
        BTN_USERPARK = BTN_UNDEFINED
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("UnPark", key, Ini)
    If tmptxt <> "" Then
        BTN_UNPARK = val(tmptxt)
    Else
        tmptxt = CStr(BTN_UNDEFINED)
        Call HC.oPersist.WriteIniValueEx("UnPark", tmptxt, key, Ini)
        BTN_UNPARK = BTN_UNDEFINED
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("AlignAccept", key, Ini)
    If tmptxt <> "" Then
        BTN_ALIGNACCEPT = val(tmptxt)
    Else
        tmptxt = CStr(BTN_JOY3)
        Call HC.oPersist.WriteIniValueEx("AcceptAlign", tmptxt, key, Ini)
        BTN_ALIGNACCEPT = BTN_JOY3
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("AlignCancel", key, Ini)
    If tmptxt <> "" Then
        BTN_ALIGNCANCEL = val(tmptxt)
    Else
        tmptxt = CStr(BTN_JOY2)
        Call HC.oPersist.WriteIniValueEx("AlignCancel", tmptxt, key, Ini)
        BTN_ALIGNCANCEL = BTN_JOY2
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("AlignEnd", key, Ini)
    If tmptxt <> "" Then
        BTN_ALIGNEND = val(tmptxt)
    Else
        tmptxt = CStr(BTN_UNDEFINED)
        Call HC.oPersist.WriteIniValueEx("AlignEnd", tmptxt, key, Ini)
        BTN_ALIGNEND = BTN_UNDEFINED
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("PolarScopeAlign", key, Ini)
    If tmptxt <> "" Then
        BTN_POLARSCOPEALIGN = val(tmptxt)
    Else
        tmptxt = CStr(BTN_UNDEFINED)
        Call HC.oPersist.WriteIniValueEx("PolarScopeAlign", tmptxt, key, Ini)
        BTN_POLARSCOPEALIGN = BTN_UNDEFINED
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("East", key, Ini)
    If tmptxt <> "" Then
        BTN_EAST = val(tmptxt)
    Else
        tmptxt = CStr(BTN_POVE)
        Call HC.oPersist.WriteIniValueEx("East", tmptxt, key, Ini)
        BTN_EAST = BTN_POVE
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("West", key, Ini)
    If tmptxt <> "" Then
        BTN_WEST = val(tmptxt)
    Else
        tmptxt = CStr(BTN_POVW)
        Call HC.oPersist.WriteIniValueEx("West", tmptxt, key, Ini)
        BTN_WEST = BTN_POVW
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("North", key, Ini)
    If tmptxt <> "" Then
        BTN_NORTH = val(tmptxt)
    Else
        tmptxt = CStr(BTN_POVN)
        Call HC.oPersist.WriteIniValueEx("North", tmptxt, key, Ini)
        BTN_NORTH = BTN_POVN
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("South", key, Ini)
    If tmptxt <> "" Then
        BTN_SOUTH = val(tmptxt)
    Else
        tmptxt = CStr(BTN_POVS)
        Call HC.oPersist.WriteIniValueEx("South", tmptxt, key, Ini)
        BTN_SOUTH = BTN_POVS
    End If


    tmptxt = HC.oPersist.ReadIniValueEx("NorthEast", key, Ini)
    If tmptxt <> "" Then
        BTN_NORTHEAST = val(tmptxt)
    Else
        tmptxt = CStr(BTN_POVNE)
        Call HC.oPersist.WriteIniValueEx("NorthEast", tmptxt, key, Ini)
        BTN_NORTHEAST = BTN_POVNE
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("NorthWest", key, Ini)
    If tmptxt <> "" Then
        BTN_NORTHWEST = val(tmptxt)
    Else
        tmptxt = CStr(BTN_POVNW)
        Call HC.oPersist.WriteIniValueEx("NorthWest", tmptxt, key, Ini)
        BTN_NORTHWEST = BTN_POVNW
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("SouthEast", key, Ini)
    If tmptxt <> "" Then
        BTN_SOUTHEAST = val(tmptxt)
    Else
        tmptxt = CStr(BTN_POVSE)
        Call HC.oPersist.WriteIniValueEx("SouthEast", tmptxt, key, Ini)
        BTN_SOUTHEAST = BTN_POVSE
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("SouthWest", key, Ini)
    If tmptxt <> "" Then
        BTN_SOUTHWEST = val(tmptxt)
    Else
        tmptxt = CStr(BTN_POVSW)
        Call HC.oPersist.WriteIniValueEx("SouthWest", tmptxt, key, Ini)
        BTN_SOUTHWEST = BTN_POVSW
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("ReverseRA", key, Ini)
    If tmptxt <> "" Then
        BTN_RAREVERSE = val(tmptxt)
    Else
        tmptxt = CStr(BTN_UNDEFINED)
        Call HC.oPersist.WriteIniValueEx("ReverseRA", tmptxt, key, Ini)
        BTN_RAREVERSE = BTN_UNDEFINED
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("ReverseDec", key, Ini)
    If tmptxt <> "" Then
        BTN_DECREVERSE = val(tmptxt)
    Else
        tmptxt = CStr(BTN_UNDEFINED)
        Call HC.oPersist.WriteIniValueEx("ReverseDec", tmptxt, key, Ini)
        BTN_DECREVERSE = BTN_UNDEFINED
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("IncRatePreset", key, Ini)
    If tmptxt <> "" Then
        BTN_INCRATEPRESET = val(tmptxt)
    Else
        tmptxt = CStr(BTN_UNDEFINED)
        Call HC.oPersist.WriteIniValueEx("IncRatePreset", tmptxt, key, Ini)
        BTN_INCRATEPRESET = BTN_UNDEFINED
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("DecRatePreset", key, Ini)
    If tmptxt <> "" Then
        BTN_DECRATEPRESET = val(tmptxt)
    Else
        tmptxt = CStr(BTN_UNDEFINED)
        Call HC.oPersist.WriteIniValueEx("DecRatePreset", tmptxt, key, Ini)
        BTN_DECRATEPRESET = BTN_UNDEFINED
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("Rate1", key, Ini)
    If tmptxt <> "" Then
        BTN_RATE1 = val(tmptxt)
    Else
        Call HC.oPersist.WriteIniValueEx("Rate1", CStr(BTN_UNDEFINED), key, Ini)
        BTN_RATE1 = BTN_UNDEFINED
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("Rate2", key, Ini)
    If tmptxt <> "" Then
        BTN_RATE2 = val(tmptxt)
    Else
        Call HC.oPersist.WriteIniValueEx("Rate2", CStr(BTN_UNDEFINED), key, Ini)
        BTN_RATE2 = BTN_UNDEFINED
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("Rate3", key, Ini)
    If tmptxt <> "" Then
        BTN_RATE3 = val(tmptxt)
    Else
        Call HC.oPersist.WriteIniValueEx("Rate3", CStr(BTN_UNDEFINED), key, Ini)
        BTN_RATE3 = BTN_UNDEFINED
    End If
    
    tmptxt = HC.oPersist.ReadIniValueEx("Rate4", key, Ini)
    If tmptxt <> "" Then
        BTN_RATE4 = val(tmptxt)
    Else
        Call HC.oPersist.WriteIniValueEx("Rate4", CStr(BTN_UNDEFINED), key, Ini)
        BTN_RATE4 = BTN_UNDEFINED
    End If
    
    tmptxt = HC.oPersist.ReadIniValueEx("GPSync", key, Ini)
    If tmptxt <> "" Then
        BTN_SYNC = val(tmptxt)
    Else
        Call HC.oPersist.WriteIniValueEx("GPsync", CStr(BTN_UNDEFINED), key, Ini)
        BTN_SYNC = BTN_UNDEFINED
    End If
    
    tmptxt = HC.oPersist.ReadIniValueEx("DeadMansHandle", key, Ini)
    If tmptxt <> "" Then
        BTN_DEADMANSHANDLE = val(tmptxt)
    Else
        Call HC.oPersist.WriteIniValueEx("DeadMansHandle", CStr(BTN_DEADMANSHANDLE), key, Ini)
        BTN_DEADMANSHANDLE = BTN_UNDEFINED
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("ToggleLock", key, Ini)
    If tmptxt <> "" Then
        BTN_TOGGLELOCK = val(tmptxt)
    Else
        Call HC.oPersist.WriteIniValueEx("ToggleLock", CStr(BTN_TOGGLELOCK), key, Ini)
        BTN_TOGGLELOCK = BTN_UNDEFINED
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("ToggleScreenSaver", key, Ini)
    If tmptxt <> "" Then
        BTN_TOGGLESCREENSAVER = val(tmptxt)
    Else
        Call HC.oPersist.WriteIniValueEx("ToggleScreenSaver", CStr(BTN_TOGGLESCREENSAVER), key, Ini)
        BTN_TOGGLESCREENSAVER = BTN_UNDEFINED
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("POV_Enabled", key, Ini)
    If tmptxt <> "" Then
        POV_Enabled = val(tmptxt)
    Else
        tmptxt = "1"
        Call HC.oPersist.WriteIniValueEx("POV_Enabled", tmptxt, key, Ini)
        POV_Enabled = 1
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("MonitorMode", key, Ini)
    If tmptxt <> "" Then
        gMonitorMode = val(tmptxt)
    Else
        Call HC.oPersist.WriteIniValueEx("MonitorMode", "0", key, Ini)
        gMonitorMode = 0
    End If


End Sub

Public Sub SaveJoystickBtns()
Dim key As String
Dim Ini As String

    ' set up a file path for the align.ini file
    Ini = HC.oPersist.GetIniPath & "\JOYSTICK.ini"

    key = "[buttondefs]"

    Call HC.oPersist.WriteIniValueEx("StartSidreal", CStr(BTN_STARTSIDREAL), key, Ini)
    Call HC.oPersist.WriteIniValueEx("StartPEC", CStr(BTN_PEC), key, Ini)
    Call HC.oPersist.WriteIniValueEx("SpiralSearch", CStr(BTN_SPIRAL), key, Ini)
    Call HC.oPersist.WriteIniValueEx("EmergencyStop", CStr(BTN_EMERGENCYSTOP), key, Ini)
    Call HC.oPersist.WriteIniValueEx("RARateInc", CStr(BTN_RARATEINC), key, Ini)
    Call HC.oPersist.WriteIniValueEx("RARateDec", CStr(BTN_RARATEDEC), key, Ini)
    Call HC.oPersist.WriteIniValueEx("DecRateInc", CStr(BTN_DECRATEINC), key, Ini)
    Call HC.oPersist.WriteIniValueEx("DecRateDec", CStr(BTN_DECRATEDEC), key, Ini)
    Call HC.oPersist.WriteIniValueEx("ParkHome", CStr(BTN_HOMEPARK), key, Ini)
    Call HC.oPersist.WriteIniValueEx("ParkUser", CStr(BTN_USERPARK), key, Ini)
    Call HC.oPersist.WriteIniValueEx("ParkCurrent", CStr(BTN_CURRENTPARK), key, Ini)
    Call HC.oPersist.WriteIniValueEx("UnPark", CStr(BTN_UNPARK), key, Ini)
    Call HC.oPersist.WriteIniValueEx("AlignAccept", CStr(BTN_ALIGNACCEPT), key, Ini)
    Call HC.oPersist.WriteIniValueEx("AlignCancel", CStr(BTN_ALIGNCANCEL), key, Ini)
    Call HC.oPersist.WriteIniValueEx("AlignEnd", CStr(BTN_ALIGNEND), key, Ini)
    Call HC.oPersist.WriteIniValueEx("PolarScopeAlign", CStr(BTN_POLARSCOPEALIGN), key, Ini)
    Call HC.oPersist.WriteIniValueEx("East", CStr(BTN_EAST), key, Ini)
    Call HC.oPersist.WriteIniValueEx("West", CStr(BTN_WEST), key, Ini)
    Call HC.oPersist.WriteIniValueEx("North", CStr(BTN_NORTH), key, Ini)
    Call HC.oPersist.WriteIniValueEx("South", CStr(BTN_SOUTH), key, Ini)
    Call HC.oPersist.WriteIniValueEx("NorthEast", CStr(BTN_NORTHEAST), key, Ini)
    Call HC.oPersist.WriteIniValueEx("NorthWest", CStr(BTN_NORTHWEST), key, Ini)
    Call HC.oPersist.WriteIniValueEx("SouthEast", CStr(BTN_SOUTHEAST), key, Ini)
    Call HC.oPersist.WriteIniValueEx("SouthWest", CStr(BTN_SOUTHWEST), key, Ini)
    Call HC.oPersist.WriteIniValueEx("ReverseRA", CStr(BTN_RAREVERSE), key, Ini)
    Call HC.oPersist.WriteIniValueEx("ReverseDec", CStr(BTN_DECREVERSE), key, Ini)
    Call HC.oPersist.WriteIniValueEx("StartCustom", CStr(BTN_CUSTOMTRACKSTART), key, Ini)
    Call HC.oPersist.WriteIniValueEx("StartLunar", CStr(BTN_STARTLUNAR), key, Ini)
    Call HC.oPersist.WriteIniValueEx("StartSolar", CStr(BTN_STARTSOLAR), key, Ini)
    Call HC.oPersist.WriteIniValueEx("IncRatePreset", CStr(BTN_INCRATEPRESET), key, Ini)
    Call HC.oPersist.WriteIniValueEx("DecRatePreset", CStr(BTN_DECRATEPRESET), key, Ini)
    Call HC.oPersist.WriteIniValueEx("Rate1", CStr(BTN_RATE1), key, Ini)
    Call HC.oPersist.WriteIniValueEx("Rate2", CStr(BTN_RATE2), key, Ini)
    Call HC.oPersist.WriteIniValueEx("Rate3", CStr(BTN_RATE3), key, Ini)
    Call HC.oPersist.WriteIniValueEx("Rate4", CStr(BTN_RATE4), key, Ini)
    Call HC.oPersist.WriteIniValueEx("GPSync", CStr(BTN_SYNC), key, Ini)
    Call HC.oPersist.WriteIniValueEx("DeadMansHandle", CStr(BTN_DEADMANSHANDLE), key, Ini)
    Call HC.oPersist.WriteIniValueEx("ToggleLock", CStr(BTN_TOGGLELOCK), key, Ini)
    Call HC.oPersist.WriteIniValueEx("ToggleScreenSaver", CStr(BTN_TOGGLESCREENSAVER), key, Ini)
    Call HC.oPersist.WriteIniValueEx("POV_Enabled", CStr(POV_Enabled), key, Ini)
    Call HC.oPersist.WriteIniValueEx("MonitorMode", CStr(gMonitorMode), key, Ini)

End Sub

Public Sub LoadJoystickCalib()
Dim tmptxt As String
Dim VarStr As String
Dim key As String
Dim Ini As String

    ' set up a file path for the align.ini file
    Ini = HC.oPersist.GetIniPath & "\JOYSTICK.ini"

    key = "[calibration]"

    tmptxt = HC.oPersist.ReadIniValueEx("MinX", key, Ini)
    If tmptxt <> "" Then
        JoystickCal.dwMinXpos = val(tmptxt)
    Else
        tmptxt = "0"
        Call HC.oPersist.WriteIniValueEx("MinX", tmptxt, key, Ini)
        JoystickCal.dwMinXpos = 0
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("75LX", key, Ini)
    If tmptxt <> "" Then
        JoystickCal.dwX75left = val(tmptxt)
    Else
        tmptxt = "8192"
        Call HC.oPersist.WriteIniValueEx("75LX", tmptxt, key, Ini)
        JoystickCal.dwX75left = 8192
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("25LX", key, Ini)
    If tmptxt <> "" Then
        JoystickCal.dwX25Left = val(tmptxt)
    Else
        tmptxt = "24576"
        Call HC.oPersist.WriteIniValueEx("25LX", tmptxt, key, Ini)
        JoystickCal.dwX25Left = 24576
    End If
    
    tmptxt = HC.oPersist.ReadIniValueEx("25RX", key, Ini)
    If tmptxt <> "" Then
        JoystickCal.dwX25Right = val(tmptxt)
    Else
        tmptxt = "40960"
        Call HC.oPersist.WriteIniValueEx("25RX", tmptxt, key, Ini)
        JoystickCal.dwX25Right = 40960
    End If
    
    tmptxt = HC.oPersist.ReadIniValueEx("75RX", key, Ini)
    If tmptxt <> "" Then
        JoystickCal.dwX75Right = val(tmptxt)
    Else
        tmptxt = "57344"
        Call HC.oPersist.WriteIniValueEx("75RX", tmptxt, key, Ini)
        JoystickCal.dwX75Right = 57344
    End If
    
    tmptxt = HC.oPersist.ReadIniValueEx("MaxX", key, Ini)
    If tmptxt <> "" Then
        JoystickCal.dwMaxXpos = val(tmptxt)
    Else
        tmptxt = "65535"
        Call HC.oPersist.WriteIniValueEx("MaxX", tmptxt, key, Ini)
        JoystickCal.dwMaxXpos = 65535
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("MinY", key, Ini)
    If tmptxt <> "" Then
        JoystickCal.dwMinYpos = val(tmptxt)
    Else
        tmptxt = "0"
        Call HC.oPersist.WriteIniValueEx("MinY", tmptxt, key, Ini)
        JoystickCal.dwMinYpos = 0
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("75LY", key, Ini)
    If tmptxt <> "" Then
        JoystickCal.dwY75left = val(tmptxt)
    Else
        tmptxt = "8192"
        Call HC.oPersist.WriteIniValueEx("75LY", tmptxt, key, Ini)
        JoystickCal.dwY75left = 8192
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("25LY", key, Ini)
    If tmptxt <> "" Then
        JoystickCal.dwY25Left = val(tmptxt)
    Else
        tmptxt = "24576"
        Call HC.oPersist.WriteIniValueEx("25LY", tmptxt, key, Ini)
        JoystickCal.dwY25Left = 24576
    End If
    
    tmptxt = HC.oPersist.ReadIniValueEx("25RY", key, Ini)
    If tmptxt <> "" Then
        JoystickCal.dwY25Right = val(tmptxt)
    Else
        tmptxt = "40960"
        Call HC.oPersist.WriteIniValueEx("25RY", tmptxt, key, Ini)
        JoystickCal.dwY25Right = 40960
    End If
    
    tmptxt = HC.oPersist.ReadIniValueEx("75RY", key, Ini)
    If tmptxt <> "" Then
        JoystickCal.dwY75Right = val(tmptxt)
    Else
        tmptxt = "57344"
        Call HC.oPersist.WriteIniValueEx("75RY", tmptxt, key, Ini)
        JoystickCal.dwY75Right = 57344
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("MaxY", key, Ini)
    If tmptxt <> "" Then
        JoystickCal.dwMaxYpos = val(tmptxt)
    Else
        tmptxt = "65535"
        Call HC.oPersist.WriteIniValueEx("MaxY", tmptxt, key, Ini)
        JoystickCal.dwMaxYpos = 65535
    End If
    
    tmptxt = HC.oPersist.ReadIniValueEx("MinZ", key, Ini)
    If tmptxt <> "" Then
        JoystickCal.dwMinZpos = val(tmptxt)
    Else
        tmptxt = "0"
        Call HC.oPersist.WriteIniValueEx("MinZ", tmptxt, key, Ini)
        JoystickCal.dwMinZpos = 0
    End If
    
    tmptxt = HC.oPersist.ReadIniValueEx("MaxZ", key, Ini)
    If tmptxt <> "" Then
        JoystickCal.dwMaxZpos = val(tmptxt)
    Else
        tmptxt = "65535"
        Call HC.oPersist.WriteIniValueEx("MaxZ", tmptxt, key, Ini)
        JoystickCal.dwMaxZpos = 65535
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("MinR", key, Ini)
    If tmptxt <> "" Then
        JoystickCal.dwMinRpos = val(tmptxt)
    Else
        tmptxt = "0"
        Call HC.oPersist.WriteIniValueEx("MinR", tmptxt, key, Ini)
        JoystickCal.dwMinRpos = 0
    End If
    
    tmptxt = HC.oPersist.ReadIniValueEx("MaxR", key, Ini)
    If tmptxt <> "" Then
        JoystickCal.dwMaxRpos = val(tmptxt)
    Else
        tmptxt = "65535"
        Call HC.oPersist.WriteIniValueEx("MaxR", tmptxt, key, Ini)
        JoystickCal.dwMaxRpos = 65535
    End If
    
    tmptxt = HC.oPersist.ReadIniValueEx("Enabled", key, Ini)
    If tmptxt <> "" Then
        JoystickCal.Enabled = val(tmptxt)
    Else
        tmptxt = "1"
        Call HC.oPersist.WriteIniValueEx("Enabled", tmptxt, key, Ini)
        JoystickCal.Enabled = 1
    End If

    tmptxt = HC.oPersist.ReadIniValueEx("DualSpeed", key, Ini)
    If tmptxt <> "" Then
        JoystickCal.DualSpeed = val(tmptxt)
    Else
        tmptxt = "0"
        Call HC.oPersist.WriteIniValueEx("DualSpeed", tmptxt, key, Ini)
        JoystickCal.DualSpeed = 0
    End If


    tmptxt = HC.oPersist.ReadIniValueEx("Id", key, Ini)
    If tmptxt <> "" Then
        JoystickCal.id = val(tmptxt)
    Else
        Call HC.oPersist.WriteIniValueEx("id", "-1", key, Ini)
        JoystickCal.id = -1
    End If


End Sub

Public Sub SaveJoystickCalib()

Dim tmptxt As String
Dim VarStr As String
Dim key As String
Dim Ini As String

    ' set up a file path for the align.ini file
    Ini = HC.oPersist.GetIniPath & "\JOYSTICK.ini"

    key = "[calibration]"

    tmptxt = CStr(JoystickCal.dwMinXpos)
    Call HC.oPersist.WriteIniValueEx("MinX", tmptxt, key, Ini)

    tmptxt = CStr(JoystickCal.dwX25Left)
    Call HC.oPersist.WriteIniValueEx("25LX", tmptxt, key, Ini)

    tmptxt = CStr(JoystickCal.dwX75left)
    Call HC.oPersist.WriteIniValueEx("75LX", tmptxt, key, Ini)

    tmptxt = CStr(JoystickCal.dwX25Right)
    Call HC.oPersist.WriteIniValueEx("25RX", tmptxt, key, Ini)
    
    tmptxt = CStr(JoystickCal.dwX75Right)
    Call HC.oPersist.WriteIniValueEx("75RX", tmptxt, key, Ini)
    
    tmptxt = CStr(JoystickCal.dwMaxXpos)
    Call HC.oPersist.WriteIniValueEx("MaxX", tmptxt, key, Ini)

    tmptxt = CStr(JoystickCal.dwMinYpos)
    Call HC.oPersist.WriteIniValueEx("MinY", tmptxt, key, Ini)

    tmptxt = CStr(JoystickCal.dwY25Left)
    Call HC.oPersist.WriteIniValueEx("25LY", tmptxt, key, Ini)

    tmptxt = CStr(JoystickCal.dwY75left)
    Call HC.oPersist.WriteIniValueEx("75LY", tmptxt, key, Ini)

    tmptxt = CStr(JoystickCal.dwY25Right)
    Call HC.oPersist.WriteIniValueEx("25RY", tmptxt, key, Ini)
    
    tmptxt = CStr(JoystickCal.dwY75Right)
    Call HC.oPersist.WriteIniValueEx("75RY", tmptxt, key, Ini)
        
    tmptxt = CStr(JoystickCal.dwMaxYpos)
    Call HC.oPersist.WriteIniValueEx("MaxY", tmptxt, key, Ini)

    tmptxt = CStr(JoystickCal.dwMinZpos)
    Call HC.oPersist.WriteIniValueEx("MinZ", tmptxt, key, Ini)

    tmptxt = CStr(JoystickCal.dwMaxZpos)
    Call HC.oPersist.WriteIniValueEx("MaxZ", tmptxt, key, Ini)

    tmptxt = CStr(JoystickCal.dwMinRpos)
    Call HC.oPersist.WriteIniValueEx("MinR", tmptxt, key, Ini)

    tmptxt = CStr(JoystickCal.dwMaxRpos)
    Call HC.oPersist.WriteIniValueEx("MaxR", tmptxt, key, Ini)

    tmptxt = CStr(JoystickCal.Enabled)
    Call HC.oPersist.WriteIniValueEx("Enabled", tmptxt, key, Ini)
    
    tmptxt = CStr(JoystickCal.DualSpeed)
    Call HC.oPersist.WriteIniValueEx("DualSpeed", tmptxt, key, Ini)

    tmptxt = CStr(JoystickCal.id)
    Call HC.oPersist.WriteIniValueEx("id", tmptxt, key, Ini)

End Sub

