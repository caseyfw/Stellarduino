//---------------------------------------------------------------------
//Copyright © 2006 Raymund Sarmiento

//Permission is hereby granted to use this Software for any purpose
//including combining with commercial products, creating derivative
//works, and redistribution of source or binary code, without
//limitation or consideration. Any redistributed copies of this
//Software must include the above Copyright Notice.

//THIS SOFTWARE IS PROVIDED "AS IS". THE AUTHOR OF THIS CODE MAKES NO
//WARRANTIES REGARDING THIS SOFTWARE, EXPRESS OR IMPLIED, AS TO ITS
//SUITABILITY OR FITNESS FOR A PARTICULAR PURPOSE.
//---------------------------------------------------------------------
//
//
//Written:  07-Oct-06   Raymund Sarmiento
//
//Edits:
//
//When      Who     What
//--------- ---     --------------------------------------------------
//24-Oct-03 rcs     Initial edit for EQ Mount Driver Function Prototype
//---------------------------------------------------------------------
//
//
// SYNOPSIS:
//
// This is a demonstration of a EQ6/ATLAS/EQG direct stepper motor control access
// using the EQCONTRL.DLL driver code.
//
// File EQCONTROL.bas contains all the function prototypes of all subroutines
// encoded in the EQCONTRL.dll
//
// The EQ6CONTRL.DLL simplifies execution of the Mount controller board stepper
// commands.
//
// The mount circuitry needs to be modified for this test program to work.
// Circuit details can be found at http://www.freewebs.com/eq6mod/

// DISCLAIMER:

// You can use the information on this site COMPLETELY AT YOUR OWN RISK.
// The modification steps and other information on this site is provided
// to you "AS IS" and WITHOUT WARRANTY OF ANY KIND, express, statutory,
// implied or otherwise, including without limitation any warranty of
// merchantability or fitness for any particular or intended purpose.
// In no event the author will  be liable for any direct, indirect,
// punitive, special, incidental or consequential damages or loss of any
// kind whether or not the author  has been advised of the possibility
// of such loss.

// WARNING:

// Circuit modifications implemented on your setup could invalidate
// any warranty that you may have with your product. Use this
// information at your own risk. The modifications involve direct
// access to the stepper motor controls of your mount. Any "mis-control"
// or "mis-command"  / "invalid parameter" or "garbage" data sent to the
// mount could accidentally activate the stepper motors and allow it to
// rotate "freely" damaging any equipment connected to your mount.
// It is also possible that any garbage or invalid data sent to the mount
// could cause its firmware to generate mis-steps pulse sequences to the
// motors causing it to overheat. Make sure that you perform the
// modifications and testing while there is no physical "load" or
// dangling wires on your mount. Be sure to disconnect the power once
// this event happens or if you notice any unusual sound coming from
// the motor assembly.






// ************* Mount initialization ***************

// Intialize Mount / Driver
DWORD __stdcall EQ_Init(char *comportname,DWORD baud,DWORD timeout, DWORD retry);

// Initialize RA/DEC Motors
DWORD __stdcall EQ_InitMotors(DWORD RA_val, DWORD DEC_val);

// Disconnect Driver
DWORD __stdcall EQ_End();

// ************* Motor Stop ************************
DWORD __stdcall EQ_MotorStop(DWORD motor_id);

// Move motor based on microstep values
DWORD __stdcall EQ_StartMoveMotor(DWORD motor_id, DWORD hemisphere, DWORD direction,  DWORD steps, DWORD stepslowdown);

// ************* Motor Parameters ******************

// Get Motor microstep position
DWORD __stdcall EQ_GetMotorValues(DWORD motor_id);

// Get Motor status
DWORD __stdcall EQ_GetMotorStatus(DWORD motor_id);

// Set Motor microstep position (counter value set only)
DWORD __stdcall EQ_SetMotorValues(DWORD motor_id, DWORD motor_val);

// ************* Slewing and GOTO *****************


// Slew Motor based on specified slew speed 
DWORD __stdcall EQ_Slew(DWORD motor_id, DWORD hemisphere, DWORD direction, DWORD rate);

// ************* Tracking ************************

// Move motor at specified tracking speed
DWORD __stdcall EQ_StartRATrack(DWORD trackrate, DWORD hemisphere, DWORD direction);

// Customized speed (for orbital data)
DWORD __stdcall EQ_SendCustomTrackRate(DWORD motor_id, DWORD trackrate, DWORD trackoffset, DWORD trackdir, DWORD hemisphere, DWORD direction);

// ************* Guiding / PEC ******************

// Initiate speed change (guiding)
DWORD __stdcall EQ_SendGuideRate(DWORD motor_id, DWORD trackrate, DWORD guiderate, DWORD guidedir, DWORD hemisphere, DWORD direction);


// Autoguider port speed
DWORD __stdcall EQ_SetAutoguiderPortRate(DWORD motor_id, DWORD portguiderate);


// ************* Mount and Driver Parameters *************

// Get mount stepping info
DWORD __stdcall EQ_GetTotal360microstep(DWORD motor_id);

// Mount version
DWORD __stdcall EQ_GetMountVersion();

// Mount/Driver state
DWORD __stdcall EQ_GetMountStatus();
