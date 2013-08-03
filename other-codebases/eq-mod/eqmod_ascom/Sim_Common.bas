Attribute VB_Name = "Sim_Common"
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
'
' Written:  07-Oct-06   Raymund Sarmiento
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' 28-Oct-06 rcs     Initial edit for EQ Mount Driver Function Prototype
'---------------------------------------------------------------------
'
'
'  SYNOPSIS:
'
'  This is a demonstration of a EQ6/ATLAS/EQG direct stepper motor control access
'  using a SIMULATED EQCONTRL.DLL driver code.
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
'  the Hand Controller(HC) and the Go To Controller.
'

Option Explicit

'Public Const GMS = 10.5         'Geared Microstep per 100 millisecond
Public Const GMS = 10.4730403903004         'Geared Microstep per 100 millisecond
Public Const PIEDISP = 2        'number of 100ms ticks before a pie chart is updated
'Public Const GMS As Double = 104.730403903004         ' (9024000/86164.0905)
                                                            
                                                            ' 104.73040390300411747513310083625


Public emulCurrent_time As Double
Public emulLast_time As Double
Public emulEmulRA_Init As Double

Public emulRAEncoder_Zero_pos As Double
Public emulDECEncoder_Zero_pos As Double

Public emulRA_shift As Double
Public emulDEC_Shift As Double

Public emulRA_target As Double
Public emulDEC_target As Double

Public emulRA_gotorate As Double
Public emulDEC_gotorate As Double

Public emulSimConnected As Long

Public emulTot_RA As Double
Public emulTot_DEC As Double

Public emulRA_track As Double
Public emulDEC_track As Double

Public emulRA_Hours As Double
Public emulDEC_Degrees As Double
Public emulDec_DegNoAdjust As Double

Public emulLatitude As Double
Public emulLongitude As Double
Public emulElevation As Double
Public emulHemisphere As Long

Public emulRA As Double
Public emulDEC As Double
Public emulAlt As Double
Public emulAz As Double

Public emulRA_Encoder As Double
Public emulDEC_Encoder As Double
Public emulpieCounter As Double

'///// Conection-Initalization Functions /////

'
' Function name    : EQ_Init()
' Description      : Connect to the EQ Controller via Serial and initialize the stepper board
' Return type      : DOUBLE
'                      000 - Success
'                      001 - COM Port Not available
'                      002 - COM Port already Open
'                      003 - COM Timeout Error
'                      005 - Mount Initialized on using non-standard parameters
'                      010 - Cannot execute command at the current stepper controller state
'                      999 - Invalid parameter
' Argument         : STRING COMPORT Name
' Argument         : DOUBLE baud - Baud Rate
' Argument         : DOUBLE timeout - COMPORT Timeout(1 - 50000)
' Argument         : DOUBLE retry - COMPORT Retry(0 - 100)
'
Public Function EQ_Init(COMPORT As String, baud As Long, timeout As Long, retry As Long) As Long

    If emulSimConnected = 1 Then
        EQ_Init = 2
        Exit Function
    End If
    

    emulSimConnected = 1
    EQ_Init = 0
    
    EQSIM.Show
End Function

'
' Function name    : EQ_End()
' Description      : Close the COM Port and end EQ Connection
' Return type      : DOUBLE
'          00 - Success
'          01 - COM Port Not Openavailable
'
Public Function EQ_End() As Long

    emulSimConnected = 0
    EQ_End = 0
    
 
    Unload EQSIM

End Function

'
' Function name    : EQ_InitMotors()
' Description      : Initialize RA/DEC Motors and activate Motor Driver Coils
' Return type      : DOUBLE
'                     000 - Success
'                     001 - COM PORT Not available
'                     003 - COM Timeout Error
'                     006 - RA Motor still running
'                     007 - DEC Motor still running
'                     008 - Error Initializing RA Motor
'                     009 - Error Initilizing DEC Motor
'                     010 - Cannot execute command at the current stepper controller state
' Argument         : DOUBLE RA_val       Initial ra microstep counter value
' Argument         : DOUBLE DEC_val     Initial dec microstep counter value
'
Public Function EQ_InitMotors(pRA As Long, pDEC As Long) As Long

    If emulSimConnected <> 1 Then
        
        EQ_InitMotors = 1
        Exit Function
    
    End If
    
    If (emulRA_shift + emulRA_track) <> 0 Then
    
        EQ_InitMotors = 6
        Exit Function
    End If

    If (emulDEC_Shift + emulDEC_track) <> 0 Then
    
        EQ_InitMotors = 7
        Exit Function
    End If


    emulRA_Encoder = pRA
    emulDEC_Encoder = pDEC
    
    EQ_InitMotors = 0

End Function


'///// Motor Status Functions /////


'
' Function name    : EQ_GetMotorValues()
' Description      : Get RA/DEC Motor microstep counts
' Return type      : Double - Stepper Counter Values
'                     0 - 16777215  Valid Count Values
'                     0x1000000 - Mount Not available
'                     0x1000005 - COM TIMEOUT
'                     0x10000FF - Illegal Mount reply
'                     0x3000000 - Invalid Parameter
' Argument         : DOUBLE motor_id
'                     00 - RA Motor
'                     01 - DEC Motor
'
Public Function EQ_GetMotorValues(motor_id As Long) As Long
                      
    If emulSimConnected <> 1 Then
        
        EQ_GetMotorValues = &H1000000
        Exit Function
    
    End If
    
    Select Case (motor_id)
    
        Case 0
            ' keep returned value in range 0-ffffff
            EQ_GetMotorValues = emulRA_Encoder Mod 16777216
        Case 1
            ' keep returned value in range 0-ffffff
            EQ_GetMotorValues = emulDEC_Encoder Mod 16777216
        Case Else
            EQ_GetMotorValues = &H3000000
    End Select
                      
                      
End Function

'
' Function name    : EQ_GetMotorStatus()
' Description      : Get RA/DEC Stepper Motor Status
' Return type      : DOUBLE
'                     128 - Motor not rotating, Teeth at front contact
'                     144 - Motor rotating, Teeth at front contact
'                     160 - Motor not rotating, Teeth at rear contact
'                     176 - Motor rotating, Teeth at rear contact
'                     200 - Motor not initialized
'                     001 - COM Port Not available
'                     003 - COM Timeout Error
'                     999 - Invalid Parameter
' Argument         : DOUBLE motor_id
'                     00 - RA Motor
'                     01 - DEC Motor
'
Public Function EQ_GetMotorStatus(motor_id As Long) As Long
 
    If emulSimConnected <> 1 Then
        
        EQ_GetMotorStatus = 1
        Exit Function
    
    End If
    
    Select Case (motor_id)
    
        Case 0
        
            If (emulRA_shift + emulRA_track + emulRA_target) = 0 Then
        
                EQ_GetMotorStatus = 0
            Else
                EQ_GetMotorStatus = &H10
            End If

        Case 1
            If (emulDEC_Shift + emulDEC_track + emulDEC_target) = 0 Then
                EQ_GetMotorStatus = 0
            Else
                EQ_GetMotorStatus = &H10
            End If
            
        Case Else
            EQ_GetMotorStatus = 999
    End Select
     
 
End Function

 
'
' Function name    : EQ_SeTMotorValues()
' Description      : Sets RA/DEC Motor microstep counters(pseudo encoder position)
' Return type      : DOUBLE - Stepper Counter Values
'                     000 - Success
'                     001 - Comport Not available
'                     003 - COM Timeout Error
'                     010 - Cannot execute command at the current stepper controller state
'                     011 - Motor not initialized
'                     999 - Invalid Parameter
' Argument         : DOUBLE motor_id
'                     00 - RA Motor
'                     01 - DEC Motor
' Argument         : DOUBLE motor_val
'                     0 - 16777215  Valid Count Values
'
 
Public Function EQ_SetMotorValues(motor_id As Long, motor_val As Long) As Long

    If emulSimConnected <> 1 Then
        
        EQ_SetMotorValues = 1
        Exit Function
    
    End If
    
    Select Case (motor_id)
    
        Case 0
            EQ_SetMotorValues = 0
            emulRA_Encoder = motor_val
        Case 1
            EQ_SetMotorValues = 0
            emulDEC_Encoder = motor_val
        Case Else
            EQ_SetMotorValues = 999
    End Select
               

End Function

'///// Motor Movement Functions /////

'
' Function name    : EQ_StartMoveMotor
' Description      : Slew RA/DEC Motor based on provided microstep counts
' Return type      : DOUBLE
'                     000 - Success
'                     001 - COM PORT Not available
'                     003 - COM Timeout Error
'                     004 - Motor still busy, aborted
'                     010 - Cannot execute command at the current stepper controller state
'                     011 - Motor not initialized
'                     999 - Invalid Parameter
' Argument         : DOUBLE motor_id
'                     00 - RA Motor
'                     01 - DEC Motor
' Argument         : DOUBLE hemisphere
'                     00 - North
'                     01 - South
' Argument         : DOUBLE direction
'                     00 - Forward(+)
'                     01 - Reverse(-)
' Argument         : DOUBLE steps count
' Argument         : DOUBLE motor de-acceleration  point(set between 50% t0 90% of total steps)
'


Public Function EQ_StartMoveMotor(motor_id As Long, hemisphere As Long, direction As Long, Steps As Long, stepslowdown As Long) As Long

    If emulSimConnected <> 1 Then
        
        EQ_StartMoveMotor = 1
        Exit Function
    
    End If
    
    If (direction < 0) Or (direction > 1) Then
    
        EQ_StartMoveMotor = 999
        Exit Function
    
    End If
    
    If (hemisphere < 0) Or (hemisphere > 1) Then
    
        EQ_StartMoveMotor = 999
        Exit Function
    
    End If
    
'    emulHemisphere = hemisphere
    
    Select Case (motor_id)
    
        Case 0
            EQ_StartMoveMotor = 0
            
            EQSIM.Timer.Enabled = False
            If direction = 0 Then
                emulRA_target = emulRA_Encoder + Steps
            Else
                emulRA_target = emulRA_Encoder - Steps
            End If
            EQSIM.Timer.Enabled = True
        Case 1
            EQ_StartMoveMotor = 0
            EQSIM.Timer.Enabled = False
            If direction = 0 Then
                emulDEC_target = emulDEC_Encoder + Steps
            Else
                emulDEC_target = emulDEC_Encoder - Steps
            End If
            EQSIM.Timer.Enabled = True
        Case Else
            EQ_StartMoveMotor = 999
    End Select

End Function

'
' Function name    : EQ_Slew()
' Description      : Slew RA/DEC Motor based on given rate
' Return type      : DOUBLE
'                     000 - Success
'                     001 - Comport Not available
'                     003 - COM Timeout Error
'                     004 - Motor still busy
'                     010 - Cannot execute command at the current stepper controller state
'                     011 - Motor not initialized
'                     999 - Invalid Parameter
' Argument         : DOUBLE motor_id
'                     00 - RA Motor
'                     01 - DEC Motor
' Argument         : INTEGER direction
'                    00 - Forward(+)
'                    01 - Reverse(-)
' Argument         : INTEGER rate
'                         1-800 of Sidreal Rate
'
Public Function EQ_Slew(motor_id As Long, hemisphere As Long, direction As Long, DRATE As Long) As Long

    If emulSimConnected <> 1 Then
        
        EQ_Slew = 1
        Exit Function
    
    End If

    If (direction < 0) Or (direction > 1) Then
    
        EQ_Slew = 999
        Exit Function
    
    End If

    If (DRATE < 0) Or (DRATE > 800) Then
    
        EQ_Slew = 999
        Exit Function
    
    End If

    If (hemisphere < 0) Or (hemisphere > 1) Then
    
        EQ_Slew = 999
        Exit Function
    
    End If
    
'    emulHemisphere = hemisphere

   Select Case (motor_id)
    
        Case 0
            EQ_Slew = 0
            If direction = 0 Then
                emulRA_shift = (GMS * DRATE)
            Else
                emulRA_shift = (-GMS * DRATE)
            End If

        Case 1
            EQ_Slew = 0
            If direction = 0 Then
                emulDEC_Shift = (GMS * DRATE)
            Else
                emulDEC_Shift = (-GMS * DRATE)
            End If

        Case Else
            EQ_Slew = 999
    End Select

End Function

'
' Function name    : EQ_StartRATrack()
' Description      : Track or rotate RA/DEC Stepper Motors at the specified rate
' Return type      : DOUBLE
'                     000 - Success
'                     001 - Comport Not available
'                     003 - COM Timeout Error
'                     010 - Cannot execute command at the current stepper controller state
'                     011 - Motor not initialized
'                     999 - Invalid Parameter
' Argument         : DOUBLE trackrate
'                     00 - Sidreal
'                     01 - Lunar
'                     02 - Solar
' Argument         : DOUBLE hemisphere
'                     00 - North
'                     01 - South
' Argument         : DOUBLE direction
'                     00 - Forward(+)
'                     01 - Reverse(-)
'
Public Function EQ_StartRATrack(trackrate As Long, hemisphere As Long, direction As Long) As Long

    If emulSimConnected <> 1 Then
        
        EQ_StartRATrack = 1
        Exit Function
    
    End If



    If (direction < 0) Or (direction > 1) Then
    
        EQ_StartRATrack = 999
        Exit Function
    
    End If
    
    If (hemisphere < 0) Or (hemisphere > 1) Then
    
        EQ_StartRATrack = 999
        Exit Function
    
    End If
    
    emulHemisphere = hemisphere
    
 '   eqres = EQ_MotorStop(1)
    eqres = EQ_MotorStop(0)


    Select Case (trackrate)
    
        Case 0
            EQ_StartRATrack = 0
            If direction = 0 Then
                emulRA_track = GMS
            Else
                emulRA_track = -GMS
            End If
            
        Case 1
            EQ_StartRATrack = 0
            If direction = 0 Then
                emulRA_track = GMS / 1.0371
            Else
                emulRA_track = -GMS / 1.0371
            End If
            
            
            
        Case 2
            EQ_StartRATrack = 0
            If direction = 0 Then
                emulRA_track = GMS / 1.001613
            Else
                emulRA_track = -GMS / 1.001613
            End If
            
            
        Case Else
            EQ_StartRATrack = 999
    End Select

End Function

'
' Function name    : EQ_SendGuideRate()
' Description      : Adjust the RA/DEC rotation trackrate based on a given speed adjustment rate
' Return type      : int
'                     000 - Success
'                     001 - Comport Not available
'                     003 - COM Timeout Error
'                     004 - Motor still busy
'                     010 - Cannot execute command at the current stepper controller state
'                     011 - Motor not initialized
'                     999 - Invalid Parameter
'
' Argument         : DOUBLE motor_id
'                     00 - RA Motor
'                     01 - DEC Motor
' Argument         : DOUBLE trackrate
'                     00 - Sidreal
'                     01 - Lunar
'                     02 - Solar
' Argument         : DOUBLE guiderate
'                     00 - No Change
'                     01 - 10%
'                     02 - 20%
'                     03 - 30%
'                     04 - 40%
'                     05 - 50%
'                     06 - 60%
'                     07 - 70%
'                     08 - 80%
'                     09 - 90%
' Argument         : DOUBLE guidedir
'                     00 - Positive
'                     01 - Negative
' Argument         : DOUBLE hemisphere(used for DEC Motor control)
'                     00 - North
'                     01 - South
' Argument         : DOUBLE direction(used for DEC Motor control)
'                     00 - Forward(+)
'                     01 - Reverse(-)
'

Public Function EQ_SendGuideRate(motor_id As Long, trackrate As Long, guiderate As Long, guidedir As Long, hemisphere As Long, direction As Long) As Long
    
    Dim i As Double
    
     If emulSimConnected <> 1 Then
        
        EQ_SendGuideRate = 1
        Exit Function
    
    End If

    If (motor_id < 0) Or (motor_id > 1) Then
    
        EQ_SendGuideRate = 999
        Exit Function
    
    End If

    If (guidedir < 0) Or (guidedir > 1) Then
    
        EQ_SendGuideRate = 999
        Exit Function
    
    End If

    If (hemisphere < 0) Or (hemisphere > 1) Then
    
        EQ_SendGuideRate = 999
        Exit Function
    
    End If
    
'    emulHemisphere = hemisphere
    
    i = 0.1 * guiderate


    If motor_id = 0 Then
      Select Case (trackrate)
    
        Case 0
            EQ_SendGuideRate = 0
            If guidedir = 0 Then
                emulRA_track = GMS + (GMS * i)
            Else
                emulRA_track = GMS - (GMS * i)
            End If
            
        Case 1
            EQ_SendGuideRate = 0
            If guidedir = 0 Then
                emulRA_track = (GMS / 1.0371) + (GMS / 1.0371 * i)
            Else
                emulRA_track = (GMS / 1.0371) - (GMS / 1.0371 * i)
            End If
            
            
            
        Case 2
            EQ_SendGuideRate = 0
            If guidedir = 0 Then
                emulRA_track = (GMS / 1.001613) + (GMS / 1.001613 * i)
            Else
                emulRA_track = (GMS / 1.001613) - (GMS / 1.001613 * i)
            End If
            
            
        Case Else
            EQ_SendGuideRate = 999
      End Select
    Else
         Select Case (trackrate)
    
        Case 0
            EQ_SendGuideRate = 0
            If guidedir = 0 Then
                emulDEC_track = GMS * i
            Else
                emulDEC_track = -GMS * i
            End If
        Case 1
            EQ_SendGuideRate = 0
            If guidedir = 0 Then
                emulDEC_track = GMS / 1.0371 * i
            Else
                emulDEC_track = -GMS / 1.0371 * i
            End If
            
        Case 2
            EQ_SendGuideRate = 0
            If guidedir = 0 Then
                emulDEC_track = GMS / 1.001613 * i
            Else
                emulDEC_track = -GMS / 1.001613 * i
            End If
            
        Case Else
            EQ_SendGuideRate = 999
      End Select
    
    End If
    
    
End Function

'
' Function name    : EQ_SendCustomTrackRate()
' Description      : Adjust the RA/DEC rotation trackrate based on a given speed adjustment offset
' Return type      : DOUBLE
'                     000 - Success
'                     001 - Comport Not available
'                     003 - COM Timeout Error
'                     004 - Motor still busy
'                     010 - Cannot Execute command at the current state
'                     011 - Motor not initialized
'                     999 - Invalid Parameter
'
' Argument         : DOUBLE motor_id
'                     00 - RA Motor
'                     01 - DEC Motor
' Argument         : DOUBLE trackrate
'                     00 - Sidreal
'                     01 - Lunar
'                     02 - Solar
' Argument         : DOUBLE trackoffset
'                     0 - 300
' Argument         : DOUBLE trackdir
'                     00 - Positive
'                     01 - Negative
' Argument         : DOUBLE hemisphere(used for DEC Motor)
'                     00 - North
'                     01 - South
' Argument         : DOUBLE direction(used for DEC Motor)
'                     00 - Forward(+)
'                     01 - Reverse(-)
'


Public Function EQ_SendCustomTrackRate(motor_id As Long, trackrate As Long, trackoffset As Long, trackdir As Long, hemisphere As Long, direction As Long) As Long

   Dim i As Double
   Dim j As Double
   
    
     If emulSimConnected <> 1 Then
        
        EQ_SendCustomTrackRate = 1
        Exit Function
    
    End If

    If (motor_id < 0) Or (motor_id > 1) Then
    
        EQ_SendCustomTrackRate = 999
        Exit Function
    
    End If

    If (trackdir < 0) Or (trackdir > 1) Then
    
        EQ_SendCustomTrackRate = 999
        Exit Function
    
    End If
    
    If (direction < 0) Or (direction > 1) Then
    
        EQ_SendCustomTrackRate = 999
        Exit Function
    
    End If
    
    If (trackoffset < 0) Or (trackoffset > 400) Then
    
        EQ_SendCustomTrackRate = 999
        Exit Function
    
    End If

    If (hemisphere < 0) Or (hemisphere > 1) Then
    
        EQ_SendCustomTrackRate = 999
        Exit Function
    
    End If
    
    emulHemisphere = hemisphere

    If direction = 1 Then
        j = -1
    Else
        j = 1
    End If


    If motor_id = 0 Then
      Select Case (trackrate)
    
        Case 0
            EQ_SendCustomTrackRate = 0
            If trackdir = 0 Then
                emulRA_track = (620 * GMS / (620 - trackoffset)) * j
            Else
                emulRA_track = (620 * GMS / (620 + trackoffset)) * j
            End If
            
        Case 1
            EQ_SendCustomTrackRate = 0
            If trackdir = 0 Then
                emulRA_track = (620 * GMS / (643 - trackoffset)) * j
            Else
                emulRA_track = (620 * GMS / (643 + trackoffset)) * j
            End If
            
            
            
        Case 2
            EQ_SendCustomTrackRate = 0
            If trackdir = 0 Then
                emulRA_track = (620 * GMS / (621 - trackoffset)) * j
            Else
                emulRA_track = (620 * GMS / (621 + trackoffset)) * j
            End If
            
            
        Case Else
            EQ_SendCustomTrackRate = 999
      End Select
    Else
         Select Case (trackrate)
    
        Case 0
            EQ_SendCustomTrackRate = 0
            If trackdir = 0 Then
                emulDEC_track = (620 * GMS / (620 - trackoffset)) * j
            Else
                emulDEC_track = (620 * GMS / (620 + trackoffset)) * j
            End If
        Case 1
            EQ_SendCustomTrackRate = 0
            If trackdir = 0 Then
                emulDEC_track = (620 * GMS / (643 - trackoffset)) * j
            Else
                emulDEC_track = (620 * GMS / (643 + trackoffset)) * j
            End If
            
        Case 2
            EQ_SendCustomTrackRate = 0
            If trackdir = 0 Then
                emulDEC_track = (620 * GMS / (621 - trackoffset)) * j
            Else
                emulDEC_track = (620 * GMS / (621 + trackoffset)) * j
            End If
            
        Case Else
            EQ_SendCustomTrackRate = 999
      End Select
    
    End If
    
    

End Function

'
' Function name    : EQ_MotorStop()
' Description      : Stop RA/DEC Motor
' Return type      : DOUBLE
'                     000 - Success
'                     001 - Comport Not available
'                     003 - COM Timeout Error
'                     010 - Cannot execute command at the current stepper controller state
'                     011 - Motor not initialized
'                     999 - Invalid Parameter
' Argument         : DOUBLE motor_id
'                     00 - RA Motor
'                     01 - DEC Motor
'
Public Function EQ_MotorStop(motor_id As Long) As Long

    If emulSimConnected <> 1 Then
        
        EQ_MotorStop = 1
        Exit Function
    
    End If
    
    Select Case (motor_id)
    
        Case 0
            EQ_MotorStop = 0
            emulRA_shift = 0
            emulRA_track = 0
            emulRA_target = 0
        Case 1
            EQ_MotorStop = 0
            emulDEC_Shift = 0
            emulDEC_track = 0
            emulDEC_target = 0
        Case Else
            EQ_MotorStop = 999
    End Select
    

End Function

'
' Function name    : EQ_SetAutoguiderPortRate()
' Description      : Sets RA/DEC Autoguideport rate
' Return type      : DOUBLE - Stepper Counter Values
'                     000 - Success
'                     001 - Comport Not available
'                     003 - COM Timeout Error
'                     999 - Invalid Parameter
' Argument         : motor_id
'                       00 - RA Motor
'                       01 - DEC Motor
' Argument         : DOUBLE guideportrate
'                       00 - 0.25x
'                       01 - 0.50x
'                       02 - 0.75x
'                       03 - 1.00x
'
 
Public Function EQ_SetAutoguiderPortRate(motor_id As Long, guideportrate As Long) As Long

    If emulSimConnected <> 1 Then
        
        EQ_SetAutoguiderPortRate = 1
        Exit Function
    
    End If

    If (motor_id < 0) Or (motor_id > 1) Then
    
        EQ_SetAutoguiderPortRate = 999
        Exit Function
    
    End If
    
    If (guideportrate < 0) And (guideportrate > 3) Then
    
        EQ_SetAutoguiderPortRate = 999
        Exit Function
    
    End If
    
    EQ_SetAutoguiderPortRate = 0

End Function

' Function name    : EQ_GetTotal360microstep()
' Description      : Get RA/DEC Motor Total 360 degree microstep counts
' Return type      : Double - Stepper Counter Values
'                     0 - 16777215  Valid Count Values
'                     0x1000000 - Mount Not available
'                     0x3000000 - Invalid Parameter
' Argument         : DOUBLE motor_id
'                     00 - RA Motor
'                     01 - DEC Motor
'
Public Function EQ_GetTotal360microstep(motor_id As Long) As Long


    If emulSimConnected <> 1 Then
        
        EQ_GetTotal360microstep = &H1000000
        Exit Function
    
    End If

    Select Case (motor_id)
        Case 0
            EQ_GetTotal360microstep = 9024000
        Case 1
            EQ_GetTotal360microstep = 9024000
        Case Else
            EQ_GetTotal360microstep = &H3000000
    End Select
    
End Function


' Function name    : EQ_GetMountVersion()
' Description      : Get Mount's Firmware version
' Return type      : Double - Mount's Firmware Version
'
'                     0x1000000 - Mount Not available
'
Public Function EQ_GetMountVersion() As Long

    EQ_GetMountVersion = &H501

End Function

' Function name    : EQ_GetMountStatus()
' Description      : Get Mount's Firmware version
' Return type      : Double - Mount Status
'
'                     000 - Not Connected
'                     001 - Connected
'
Public Function EQ_GetMountStatus() As Long

    EQ_GetMountStatus = emulSimConnected

End Function


Public Function EQ_DriverVersion() As Long

    EQ_DriverVersion = &H215

End Function

Public Function EQ_GP(motor_id As Long, p_id As Long) As Long

    Select Case p_id
        Case 10000
            EQ_GP = 1281
        Case 10001
            EQ_GP = 9024000
        Case 10002
            EQ_GP = 64935
        Case 10003
            EQ_GP = 16
        Case 10004
            If motor_id = 1 Then
                EQ_GP = 9920
            Else
                EQ_GP = 620
            End If
        Case 10005
            If motor_id = 1 Then
                EQ_GP = 9920
            Else
                EQ_GP = 620
            End If
        Case 10006
            ' get worm steps from ini file this way we can easilly simulate heq5
            readWormSteps
            EQ_GP = gRAWormSteps
'            EQ_GP = 66844
'            EQ_GP = 50133
        Case 10007
            EQ_GP = 0
        Case Else
            EQ_GP = 32
    End Select
End Function

Public Function EQ_SP(motor_id As Long, p_id As Long, val As Long) As Long
    EQ_SP = 0
End Function

Public Function EQ_SetOffset(motor_id As Long, doffset As Long) As Long
    EQ_SetOffset = 0
End Function

'
' Function name    : EQ_SetCustomTrackRate()
' Description      : Adjust the RA/DEC rotation trackrate based on a given speed adjustment offset
' Return type      : DOUBLE
'                     000 - Success
'                     001 - Comport Not available
'                     003 - COM Timeout Error
'                     004 - Motor still busy
'                     010 - Cannot Execute command at the current state
'                     011 - Motor not initialized
'                     999 - Invalid Parameter
'
' Argument         : DOUBLE motor_id
'                     00 - RA Motor
'                     01 - DEC Motor
' Argument         : DOUBLE trackmode
'                     01 - Initial
'                     00 - Update
' Argument         : DOUBLE trackoffset
' Argument         : DOUBLE trackbase
'                     00 - LowSpeed
' Argument         : DOUBLE hemisphere
'                     00 - North
'                     01 - South
' Argument         : DOUBLE direction
'                     00 - Forward(+)
'                     01 - Reverse(-)
'

Public Function EQ_SetCustomTrackRate(motor_id As Long, trackmode As Long, trackoffset As Double, trackbase As Double, hemisphere As Long, direction As Long) As Long

Dim j As Double
Dim k As Double

     If emulSimConnected <> 1 Then
        
        EQ_SetCustomTrackRate = 1
        Exit Function
    
    End If

    If (motor_id < 0) Or (motor_id > 1) Then
    
        EQ_SetCustomTrackRate = 999
        Exit Function
    
    End If

    If (trackmode < 0) Or (trackmode > 1) Then
    
        EQ_SetCustomTrackRate = 999
        Exit Function
    
    End If
    
    If (direction < 0) Or (direction > 1) Then
    
        EQ_SetCustomTrackRate = 999
        Exit Function
    
    End If
    
    If (trackoffset < 0) Then
    
        EQ_SetCustomTrackRate = 999
        Exit Function
    
    End If
    
    If (trackbase < 0) Or (trackbase > 1) Then
    
        EQ_SetCustomTrackRate = 999
        Exit Function
    
    End If

    If (hemisphere < 0) Or (hemisphere > 1) Then
    
        EQ_SetCustomTrackRate = 999
        Exit Function
    
    End If
    
    emulHemisphere = hemisphere

    If direction = 1 Then
        j = -1
    Else
        j = 1
    End If

    If trackbase = 1 Then
        k = 32
    Else
        k = 1
    End If

    If motor_id = 0 Then
    
        If trackoffset = 0 Then
            emulRA_track = 0
            EQ_SetCustomTrackRate = 0
            Exit Function
        End If
        
        emulRA_track = (620 * GMS / ((trackoffset - 30000) / k)) * j
        
        
    
    Else
    
        If trackoffset = 0 Then
            emulDEC_track = 0
            EQ_SetCustomTrackRate = 0
            Exit Function
        End If
    
        emulDEC_track = (620 * GMS / ((trackoffset - 30000) / k)) * j
        
    
    End If


    EQ_SetCustomTrackRate = 0

End Function



