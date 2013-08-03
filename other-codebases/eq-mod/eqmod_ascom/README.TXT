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
' 07-Oct-06 rcs     README.TXT for EQMOD EQCONTRL.DLL DRIVER
'---------------------------------------------------------------------
'
'
'  SYNOPSIS:
'
'  This is a demonstration of a EQ6/ATLAS/EQG direct stepper motor control access
'  using the EQCONTRL.DLL driver file.
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

' ////////////////// EQCONTRL.DLL FUNCTION PROTOTYPES //////////////////

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
' Argument         : DOUBLE timeout - COMPORT Timeout (1 - 50000)
' Argument         : DOUBLE retry - COMPORT Retry (0 - 100)
'
Public Declare Function EQ_Init Lib "EQContrl" (ByVal COMPORT As String, ByVal baud As Long, ByVal timeout As Long, ByVal retry As Long) As Long


'
' Function name    : EQ_End()
' Description      : Close the COM Port and end EQ Connection
' Return type      : DOUBLE
'          00 - Success
'          01 - COM Port Not Openavailable
'
Public Declare Function EQ_End Lib "EQContrl.dll" () As Long


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
'
' Argument         : DOUBLE RA_val	 Initial ra microstep counter value
' Argument	   : DOUBLE DEC_val     Initial dec microstep counter value

Public Declare Function EQ_InitMotors Lib "EQContrl" (ByVal RA as Long, ByVal DEC As Long) As Long



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
Public Declare Function EQ_GetMotorValues Lib "EQContrl" (ByVal value As Long) As Long
                      

'
' Function name    : EQ_GetMotorStatus()
' Description      : Get RA/DEC Stepper Motor Status
' Return type      : DOUBLE
'                     128 - Motor not rotating, Teeth at front contact
'                     144 - Motor rotating, Teeth at front contact
'                     160 - Motor not rotating, Teeth at rear contact
'                     176 - Motor rotating, Teeth at rear contact
'		      200 - Motor not initialized
'                     001 - COM Port Not available
'                     003 - COM Timeout Error
'                     999 - Invalid Parameter
' Argument         : DOUBLE motor_id
'                     00 - RA Motor
'                     01 - DEC Motor
'
Public Declare Function EQ_GetMotorStatus Lib "EQContrl" (ByVal value As Long) As Long
 
 
 
'
' Function name    : EQ_SeTMotorValues()
' Description      : Sets RA/DEC Motor microstep counters (pseudo encoder position)
' Return type      : DOUBLE - Stepper Counter Values
'                     000 - Success
'                     001 - Comport Not available
'                     003 - COM Timeout Error
'                     010 - Cannot execute command at the current stepper controller state
'		      011 - Motor not initialized
'                     999 - Invalid Parameter
' Argument         : DOUBLE motor_id
'                     00 - RA Motor
'                     01 - DEC Motor
' Argument         : DOUBLE motor_val
'                     0 - 16777215  Valid Count Values
'
 
Public Declare Function EQ_SetMotorValues Lib "EQContrl" (ByVal motor_id As Long, ByVal motor_val As Long) As Long


'
' Function name    : EQ_StartMoveMotor
' Description      : Slew RA/DEC Motor based on provided microstep counts
' Return type      : DOUBLE
'                     000 - Success
'                     001 - COM PORT Not available
'                     003 - COM Timeout Error
'                     004 - Motor still busy, aborted
'                     010 - Cannot execute command at the current stepper controller state
'		      011 - Motor not initialized
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
' Argument         : DOUBLE motor de-acceleration  point (set between 50% t0 90% of total steps)
'


Public Declare Function EQ_StartMoveMotor Lib "EQContrl" (ByVal motor_id As Long, ByVal hemisphere As Long, ByVal direction As Long, ByVal steps As Long, ByVal stepslowdown As Long) As Long

                 
'
' Function name    : EQ_Slew()
' Description      : Slew RA/DEC Motor based on given rate
' Return type      : DOUBLE
'                     000 - Success
'                     001 - Comport Not available
'                     003 - COM Timeout Error
'                     004 - Motor still busy
'                     010 - Cannot execute command at the current stepper controller state
'		      011 - Motor not initialized
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
Public Declare Function EQ_Slew Lib "EQContrl" (ByVal motor_id As Long, ByVal hemisphere As Long, ByVal direction As Long, ByVal rate As Long) As Long


'
' Function name    : EQ_StartRATrack()
' Description      : Track or rotate RA/DEC Stepper Motors at the specified rate
' Return type      : DOUBLE
'                     000 - Success
'                     001 - Comport Not available
'                     003 - COM Timeout Error
'                     010 - Cannot execute command at the current stepper controller state
'		      011 - Motor not initialized
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
Public Declare Function EQ_StartRATrack Lib "EQContrl" (ByVal trackrate As Long, ByVal hemisphere As Long, ByVal direction As Long) As Long

'
' Function name    : EQ_SendGuideRate()
' Description      : Adjust the RA/DEC rotation trackrate based on a given speed adjustment rate
' Return type      : int
'                     000 - Success
'                     001 - Comport Not available
'                     003 - COM Timeout Error
'                     004 - Motor still busy
'                     010 - Cannot execute command at the current stepper controller state
'		      011 - Motor not initialized
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
' Argument         : DOUBLE hemisphere (used for DEC Motor control)
'                     00 - North
'                     01 - South
' Argument         : DOUBLE direction (used for DEC Motor control)
'                     00 - Forward(+)
'                     01 - Reverse(-)
'

Public Declare Function EQ_SendGuideRate Lib "EQContrl" (ByVal motor_id As Long, ByVal trackrate As Long, ByVal guiderate As Long, ByVal guidedir As Long, ByVal hemisphere As Long, ByVal direction As Long) As Long


'
' Function name    : EQ_SendCustomTrackRate()
' Description      : Adjust the RA/DEC rotation trackrate based on a given speed adjustment offset
' Return type      : DOUBLE
'                     000 - Success
'                     001 - Comport Not available
'                     003 - COM Timeout Error
'                     004 - Motor still busy
'                     010 - Cannot Execute command at the current state
'		      011 - Motor not initialized
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
' Argument         : DOUBLE hemisphere (used for DEC Motor)
'                     00 - North
'                     01 - South
' Argument         : DOUBLE direction (used for DEC Motor)
'                     00 - Forward(+)
'                     01 - Reverse(-)
'
Public Declare Function EQ_SendCustomTrackRate Lib "EQContrl" (ByVal motor_id As Long, ByVal trackrate As Long, ByVal trackoffset As Long, ByVal trackdir As Long, ByVal hemisphere As Long, ByVal direction As Long) As Long

'
' Function name    : EQ_MotorStop()
' Description      : Stop RA/DEC Motor
' Return type      : DOUBLE
'                     000 - Success
'                     001 - Comport Not available
'                     003 - COM Timeout Error
'                     010 - Cannot execute command at the current stepper controller state
'		      011 - Motor not initialized
'                     999 - Invalid Parameter
' Argument         : DOUBLE motor_id
'                     00 - RA Motor
'                     01 - DEC Motor
'
Public Declare Function EQ_MotorStop Lib "EQContrl" (ByVal value As Long) As Long


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
Public Declare Function EQ_GetTotal360microstep Lib "EQContrl" (ByVal value As Long) As Long

' Function name    : EQ_GetMountVersion()
' Description      : Get Mount's Firmware version
' Return type      : Double - Mount's Firmware Version
'
'                     0x1000000 - Mount Not available
'
Public Declare Function EQ_GetMountVersion Lib "EQContrl" () As Long

' Function name    : EQ_GetMountStatus()
' Description      : Get Mount's Firmware version
' Return type      : Double - Mount Status
'			0 - Not Connected
'                       1 - Connected
'
Public Declare Function EQ_GetMountStatus Lib "EQContrl" () As Long


'
' Function name    : EQ_SetAutoguiderPortRate()
' Description      : Sets RA/DEC AUTOGUIDEPORT rate
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
 
Public Declare Function EQ_SetAutoguiderPortRate Lib "EQCONTRL" (ByVal motor_id As Long, ByVal guideportrate As Long) As Long


