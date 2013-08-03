Attribute VB_Name = "Common"
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
' Common.bas - Common functions for EQMOD ASCOM driver
'
'
' Written:  07-Oct-06   Raymund Sarmiento
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' 04-Nov-06 rcs     Initial edit for EQ Mount Driver Function Prototype
' 15-Nov-06 rcs     Fix bug on OnestarAlign Message Center display (discovered by Sander)
' 21-Nov-06 rcs     Append RA GOTO Compenstation to minimize discrepancy
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

Private oProfile As DriverHelper.Profile
Private Const oID As String = "EQMOD.Telescope"
Private Const oDESC As String = "EQMOD ASCOM Scope Driver"


Private Const SC_CLOSE As Long = &HF060&
Private Const MIIM_STATE As Long = &H1&
Private Const MIIM_ID As Long = &H2&
Private Const MFS_GRAYED As Long = &H3&
Private Const WM_NCACTIVATE As Long = &H86

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Private Declare Function GetSystemMenu Lib "user32" ( _
    ByVal hWnd As Long, ByVal bRevert As Long) As Long

Private Declare Function GetMenuItemInfo Lib "user32" Alias _
    "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, _
    ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long

Private Declare Function SetMenuItemInfo Lib "user32" Alias _
    "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, _
    ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long

Private Declare Function SendMessage Lib "user32" Alias _
    "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Any) As Long

Private Declare Function IsWindow Lib "user32" _
    (ByVal hWnd As Long) As Long


Public Function EnableCloseButton(ByVal hWnd As Long, Enable As Boolean) _
                                                                As Integer
    Const xSC_CLOSE As Long = -10

    ' Check that the window handle passed is valid
    
    EnableCloseButton = -1
    If IsWindow(hWnd) = 0 Then Exit Function
    
    ' Retrieve a handle to the window's system menu
    
    Dim hMenu As Long
    hMenu = GetSystemMenu(hWnd, 0)
    
    ' Retrieve the menu item information for the close menu item/button
    
    Dim MII As MENUITEMINFO
    MII.cbSize = Len(MII)
    MII.dwTypeData = String(80, 0)
    MII.cch = Len(MII.dwTypeData)
    MII.fMask = MIIM_STATE
    
    If Enable Then
        MII.wID = xSC_CLOSE
    Else
        MII.wID = SC_CLOSE
    End If
    
    EnableCloseButton = -0
    If GetMenuItemInfo(hMenu, MII.wID, False, MII) = 0 Then Exit Function
    
    ' Switch the ID of the menu item so that VB can not undo the action itself
    
    Dim lngMenuID As Long
    lngMenuID = MII.wID
    
    If Enable Then
        MII.wID = SC_CLOSE
    Else
        MII.wID = xSC_CLOSE
    End If
    
    MII.fMask = MIIM_ID
    EnableCloseButton = -2
    If SetMenuItemInfo(hMenu, lngMenuID, False, MII) = 0 Then Exit Function
    
    ' Set the enabled / disabled state of the menu item
    
    If Enable Then
        MII.fState = (MII.fState Or MFS_GRAYED)
        MII.fState = MII.fState - MFS_GRAYED
    Else
        MII.fState = (MII.fState Or MFS_GRAYED)
    End If
    
    MII.fMask = MIIM_STATE
    EnableCloseButton = -3
    If SetMenuItemInfo(hMenu, MII.wID, False, MII) = 0 Then Exit Function
    
    ' Activate the non-client area of the window to update the titlebar, and
    ' draw the close button in its new state.
    
    SendMessage hWnd, WM_NCACTIVATE, True, 0
    
    EnableCloseButton = 0
    
End Function

Public Sub Main()

Set oProfile = New DriverHelper.Profile
Set m_telescope = New Telescope

oProfile.DeviceType = "Telescope"

If App.StartMode = vbSModeStandalone Then
        MsgBox "This is an ASCOM driver. It cannot be run stand-alone", _
                (vbOKOnly + vbCritical + vbMsgBoxSetForeground), App.FileDescription
        Exit Sub
   End If


End Sub

'Routine to Slew the mount to target location
'RAOnly flag when active only activates the RA Motor

Public Sub radecAsyncSlew(ByVal RAOnly As Boolean)

Dim targetRAEncoder As Double
Dim targetDECEncoder As Double
Dim currentRAEncoder As Double
Dim currentDECEncoder As Double

Dim DeltaRAStep As Long
Dim DeltaDECStep As Long

Dim RASlowdown As Long
Dim DECSlowdown As Long

Dim tRA As Double
Dim tha As Double
Dim tPier As Double

      'Make sure non of the motors are running, if yes then stop it
         
      eqres = EQ_MotorStop(0)
      eqres = EQ_MotorStop(1)
      
      'Wait for motor stop , Need to add timeout routines here
      
      Do
         DoEvents
      Loop While ((gRAStatus And EQ_MOTORBUSY) + (gDECStatus And EQ_MOTORBUSY)) <> 0

      tha = RangeHA(gTargetRA - EQnow_lst(gLongitude * DEG_RAD))
  
      If tha < 0 Then
    
        If gHemisphere = 0 Then
            tPier = 1
        Else
            tPier = 0
        End If
        tRA = Range24(gTargetRA - 12)
      
      Else
        
        If gHemisphere = 0 Then
            tPier = 0
        Else
            tPier = 1
        End If
        
        tRA = gTargetRA
      
      End If

    'Compute for Target RA/DEC Encoder

     targetRAEncoder = Get_RAEncoderfromRA(tRA, 0, gLongitude, gRAEncoder_Zero_pos, gTot_RA, gHemisphere)
     targetDECEncoder = Get_DECEncoderfromDEC(gTargetDec, tPier, gDECEncoder_Zero_pos, gTot_DEC, gHemisphere)
          
     
     currentRAEncoder = Delta_RA_Map(EQ_GetMotorValues(0))
     currentDECEncoder = Delta_DEC_Map(EQ_GetMotorValues(1))
     
     DeltaRAStep = Abs(targetRAEncoder - currentRAEncoder)
     DeltaDECStep = Abs(targetDECEncoder - currentDECEncoder)
               
               
    ' Compensate for the smallest discrepancy after the final slew
         
     If gTrackingStatus > 0 Then
        If targetRAEncoder > currentRAEncoder Then
            If gHemisphere = 0 Then
                DeltaRAStep = DeltaRAStep + gRA_Compensate
            Else
                DeltaRAStep = DeltaRAStep - gRA_Compensate
            End If
        Else
            If gHemisphere = 0 Then
                DeltaRAStep = DeltaRAStep - gRA_Compensate
            Else
                DeltaRAStep = DeltaRAStep + gRA_Compensate
            End If
        End If
               
        If DeltaRAStep < 0 Then DeltaRAStep = 0
        
     End If
               
     RASlowdown = GetSlowdown(DeltaRAStep)
     DECSlowdown = GetSlowdown(DeltaDECStep)
  
     'Execute the actual slew
         
     If DeltaRAStep <> 0 Then
      If targetRAEncoder > currentRAEncoder Then
        eqres = EQ_StartMoveMotor(0, 0, 0, DeltaRAStep, RASlowdown)
      Else
        eqres = EQ_StartMoveMotor(0, 0, 1, DeltaRAStep, RASlowdown)
      End If
     End If
     
      If (Not RAOnly) And (DeltaDECStep <> 0) Then
      
        If targetDECEncoder > currentDECEncoder Then
            eqres = EQ_StartMoveMotor(1, 0, 0, DeltaDECStep, DECSlowdown)
        Else
            eqres = EQ_StartMoveMotor(1, 0, 1, DeltaDECStep, DECSlowdown)
        End If
        
      End If

     ' Activate Asynchronous Slew Monitoring Routine

     gSlewStatus = True
     gRAStatus = EQ_MOTORBUSY
     gDECStatus = EQ_MOTORBUSY
     gRAStatus_slew = False
     HC.GotoTimer.Enabled = True
     
 End Sub

Public Sub ParkHome()

    Dim currentdecpos As Double
    Dim currentrapos As Double
    Dim i As Long
    Dim j As Long

    eqres = EQ_MotorStop(0)          ' Stop RA Motor
    If eqres <> EQ_OK Then
            GoTo Endhome
    End If
    eqres = EQ_MotorStop(1)          ' Stop DEC Motor
    If eqres <> EQ_OK Then
            GoTo Endhome
    End If

    'Wait until RA motor is stable

    Do
       eqres = EQ_GetMotorStatus(0)
       If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then
            GoTo Endhome
       End If
    Loop While (eqres And EQ_MOTORBUSY) <> 0
    
    'Wait until DEC motor is stable

    Do
       eqres = EQ_GetMotorStatus(1)
       If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then
            GoTo Endhome
       End If
    Loop While (eqres And EQ_MOTORBUSY) <> 0
    
    'Read Motor Values
    
    currentrapos = Delta_RA_Map(EQ_GetMotorValues(0))
    currentdecpos = Delta_DEC_Map(EQ_GetMotorValues(1))
    
    gRAEncoderlastpos = currentrapos
    gDECEncoderlastpos = currentdecpos
    
    writelastpos   ' Save current position
    
    gRAEncoderUNPark = RAEncoder_Home_pos
    gDECEncoderUNPark = DECEncoder_Home_pos
    
    writeUnpark
    
    i = Abs(currentrapos - RAEncoder_Home_pos)
    j = Abs(currentdecpos - DECEncoder_Home_pos)
    
    ' Slew Both motors to the home position
    
    If i <> 0 Then
        If currentrapos < RAEncoder_Home_pos Then
            eqres = EQ_StartMoveMotor(0, 0, 0, i, GetSlowdown(i))
        Else
            eqres = EQ_StartMoveMotor(0, 0, 1, i, GetSlowdown(i))
        End If
    End If
    
    If j <> 0 Then
        If currentdecpos < DECEncoder_Home_pos Then
            eqres = EQ_StartMoveMotor(1, 0, 0, j, GetSlowdown(j))
        Else
            eqres = EQ_StartMoveMotor(1, 0, 1, j, GetSlowdown(j))
        End If
    End If

    
    HC.Add_Message ("Slewing mount to home position")
    ' No need to wait at this point - return control to main routine

Endhome:

End Sub

Public Sub SyncToRADEC(ByVal RightAscension As Double, ByVal Declination As Double, ByVal pLongitude As Double, ByVal pHemisphere As Long)

                                    
Dim targetRAEncoder As Double
Dim targetDECEncoder As Double
Dim currentRAEncoder As Double
Dim currentDECEncoder As Double

Dim tRA As Double
Dim tha As Double
Dim tPier As Double


    HC.EncoderTimer.Enabled = False
    currentRAEncoder = EQ_GetMotorValues(0) + gRA1Star
    currentDECEncoder = EQ_GetMotorValues(1) + gDEC1Star
    HC.EncoderTimer.Enabled = True
    
    tha = RangeHA(RightAscension - EQnow_lst(pLongitude * DEG_RAD))
  
    If tha < 0 Then
        If pHemisphere = 0 Then
            tPier = 1
        Else
            tPier = 0
        End If
        tRA = Range24(RightAscension - 12)
    Else
        If pHemisphere = 0 Then
            tPier = 0
        Else
            tPier = 1
        End If
        
        tRA = RightAscension
    End If

    'Compute for Sync RA/DEC Encoder Values

     targetRAEncoder = Get_RAEncoderfromRA(tRA, 0, pLongitude, gRAEncoder_Zero_pos, gTot_RA, pHemisphere)
     targetDECEncoder = Get_DECEncoderfromDEC(Declination, tPier, gDECEncoder_Zero_pos, gTot_DEC, pHemisphere)
          
     gRASync01 = targetRAEncoder - currentRAEncoder
     gDECSync01 = targetDECEncoder - currentDECEncoder
     
     HC.DxSalbl.Caption = Format$(Str(gRASync01), "000000000")
     
     HC.DxSblbl.Caption = Format$(Str(gDECSync01), "000000000")
     

End Sub

Public Sub ParktoUserDefine(ByVal withslew As Boolean)

    Dim currentdecpos As Double
    Dim currentrapos As Double
    
    Dim i As Long
    Dim j As Long


    eqres = EQ_MotorStop(0)          ' Stop RA Motor
    If eqres <> 0 Then
            GoTo Endparkuser
    End If
    eqres = EQ_MotorStop(1)          ' Stop DEC Motor
    If eqres <> 0 Then
            GoTo Endparkuser
    End If

    'Wait until RA motor is stable

    Do
       eqres = EQ_GetMotorStatus(0)
       If eqres = 1 Then
            GoTo Endparkuser
       End If
    Loop While (eqres And EQ_MOTORBUSY) <> 0
    
    'Wait until DEC motor is stable

    Do
       eqres = EQ_GetMotorStatus(1)
       If eqres = 1 Then
            GoTo Endparkuser
       End If
    Loop While (eqres And EQ_MOTORBUSY) <> 0
    
    'Read Motor Values

    
    currentrapos = Delta_RA_Map(EQ_GetMotorValues(0))
    currentdecpos = Delta_DEC_Map(EQ_GetMotorValues(1))
    
    gRAEncoderlastpos = currentrapos
    gDECEncoderlastpos = currentdecpos
    
    writelastpos   ' Save current position
        
    readpark       ' Read Userdefined Park data
 
    gRAEncoderUNPark = gRAEncoderPark
    gDECEncoderUNPark = gDECEncoderPark
    
    writeUnpark    ' Save Unpark Data
    
    
    If withslew Then
    
        i = Abs(currentrapos - gRAEncoderPark)
        j = Abs(currentdecpos - gDECEncoderPark)
    

        ' Slew Both motors to the user-defined park position
    
    
        If i <> 0 Then
            If currentrapos < gRAEncoderPark Then
                eqres = EQ_StartMoveMotor(0, 0, 0, i, GetSlowdown(i))
            Else
                eqres = EQ_StartMoveMotor(0, 0, 1, i, GetSlowdown(i))
            End If
        End If
    
        If j <> 0 Then
            If currentdecpos < gDECEncoderPark Then
                eqres = EQ_StartMoveMotor(1, 0, 0, j, GetSlowdown(j))
            Else
                eqres = EQ_StartMoveMotor(1, 0, 1, j, GetSlowdown(j))
            End If
        End If
    End If

Endparkuser:

End Sub
Public Sub DefinePark(ByVal StopMotors As Boolean)

    If StopMotors Then
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
    
    gRAEncoderPark = Delta_RA_Map(EQ_GetMotorValues(0))
    gDECEncoderPark = Delta_DEC_Map(EQ_GetMotorValues(1))

    writepark

ENDDefinePark:

End Sub
Public Sub Unparkscope()

    If gEQparkstatus = 1 Then

        readUnpark
    
        'Just make sure motors are not moving
    
        eqres = EQ_MotorStop(0)
        eqres = EQ_MotorStop(1)
    
        ' Restore Encoder values
    
        eqres = EQ_SetMotorValues(0, Delta_RA_Map_encoder(gRAEncoderUNPark))
        eqres = EQ_SetMotorValues(1, Delta_DEC_Map_encoder(gDECEncoderUNPark))
        
        HC.Add_Message ("Scope unparked. Encoder values restored.")

    Else
    
        HC.Add_Message ("Unpark Error: Scope not even parked.")
    
    End If


End Sub

Public Sub UnparkscopeToLastPos()

    Dim i As Long
    Dim j As Long

    If gEQparkstatus = 1 Then

        'Unpark Scope first

        readUnpark
    
        'Just make sure motors are not moving
    
        eqres = EQ_MotorStop(0)
        eqres = EQ_MotorStop(1)
    
        ' Restore encoder values
    
        eqres = EQ_SetMotorValues(0, Delta_RA_Map_encoder(gRAEncoderUNPark))
        eqres = EQ_SetMotorValues(1, Delta_DEC_Map_encoder(gDECEncoderUNPark))

        'Then Slew Scope to last position prior to park command
    
        readlastpos
    
        i = Abs(gRAEncoderUNPark - gRAEncoderlastpos)
        j = Abs(gDECEncoderUNPark - gDECEncoderlastpos)
    
    
        If i <> 0 Then
            If gRAEncoderUNPark < gRAEncoderlastpos Then
                eqres = EQ_StartMoveMotor(0, 0, 0, i, GetSlowdown(i))
            Else
                eqres = EQ_StartMoveMotor(0, 0, 1, i, GetSlowdown(i))
            End If
        End If
    
        If j <> 0 Then
            If gDECEncoderUNPark < gDECEncoderlastpos Then
                eqres = EQ_StartMoveMotor(1, 0, 0, j, GetSlowdown(j))
            Else
                eqres = EQ_StartMoveMotor(1, 0, 1, j, GetSlowdown(j))
            End If
        End If
        
        HC.Add_Message ("Scope unparked. Slewing to last position.")
 
    Else
    
          HC.Add_Message ("Unpark Error: Scope not even parked.")
    
    End If
 
End Sub

Public Sub readUnpark()

     Dim tmptxt As String

     tmptxt = oProfile.GetValue(oID, "UNPARK_RA")
     If tmptxt <> "" Then
        gRAEncoderUNPark = val(tmptxt)
     Else
        gRAEncoderUNPark = RAEncoder_Home_pos
     End If
     
     tmptxt = oProfile.GetValue(oID, "UNPARK_DEC")
     If tmptxt <> "" Then
        gDECEncoderUNPark = val(tmptxt)
     Else
        gDECEncoderUNPark = DECEncoder_Home_pos
     End If
 
End Sub
Public Sub readpark()

     Dim tmptxt As String

     tmptxt = oProfile.GetValue(oID, "PARK_RA")
     If tmptxt <> "" Then
        gRAEncoderPark = val(tmptxt)
     Else
        gRAEncoderPark = RAEncoder_Home_pos
     End If
     
     tmptxt = oProfile.GetValue(oID, "PARK_DEC")
     If tmptxt <> "" Then
        gDECEncoderPark = val(tmptxt)
     Else
        gDECEncoderPark = DECEncoder_Home_pos
     End If
 
End Sub
Public Sub readlastpos()

     Dim tmptxt As String

     tmptxt = oProfile.GetValue(oID, "LASTPOS_RA")
     If tmptxt <> "" Then
        gRAEncoderlastpos = val(tmptxt)
     Else
        gRAEncoderlastpos = RAEncoder_Home_pos
     End If
     
     tmptxt = oProfile.GetValue(oID, "LASTPOS_DEC")
     If tmptxt <> "" Then
        gDECEncoderlastpos = val(tmptxt)
     Else
        gDECEncoderlastpos = DECEncoder_Home_pos
     End If
 
End Sub

Public Function readparkStatus() As Long

     tmptxt = oProfile.GetValue(oID, "EQPARKSTATUS")
     If tmptxt = "parked" Then
        readparkStatus = 1
     Else
        readparkStatus = 0
     End If

End Function

Public Sub writeParkStatus(ByVal pval As Long)

    If pval = 0 Then
        oProfile.WriteValue oID, "EQPARKSTATUS", CStr("unparked")
    Else
        oProfile.WriteValue oID, "EQPARKSTATUS", CStr("parked")
    End If
    
End Sub


Public Sub writeUnpark()

    oProfile.WriteValue oID, "UNPARK_RA", CStr(gRAEncoderUNPark)
    oProfile.WriteValue oID, "UNPARK_DEC", CStr(gDECEncoderUNPark)

End Sub

Public Sub writepark()

    oProfile.WriteValue oID, "PARK_RA", CStr(gRAEncoderPark)
    oProfile.WriteValue oID, "PARK_DEC", CStr(gDECEncoderPark)

End Sub

Public Sub writelastpos()

    oProfile.WriteValue oID, "LASTPOS_RA", CStr(gRAEncoderlastpos)
    oProfile.WriteValue oID, "LASTPOS_DEC", CStr(gDECEncoderlastpos)

End Sub
Public Sub WriteSyncMap()

    oProfile.WriteValue oID, "RSYNC01", CStr(gRASync01)
    oProfile.WriteValue oID, "DSYNC01", CStr(gDECSync01)

End Sub

Public Sub WriteAlignMap()

    oProfile.WriteValue oID, "RALIGN01", CStr(gRA1Star)
    oProfile.WriteValue oID, "DALIGN01", CStr(gDEC1Star)

End Sub
Public Sub OneStarAlign(ByVal vRA As Double, ByVal vDEC As Double)

Dim targetRAEncoder As Double
Dim targetDECEncoder As Double
Dim currentRAEncoder As Double
Dim currentDECEncoder As Double

Dim tRA As Double
Dim tha As Double
Dim tPier As Double

    HC.Add_Message ("AlignTaget: RA[ " & FmtSexa(vRA, False) & "] DEC[ " & FmtSexa(vDEC, True) & " ]")
    HC.EncoderTimer.Enabled = False
    

    currentRAEncoder = EQ_GetMotorValues(0)
    currentDECEncoder = EQ_GetMotorValues(1)
    HC.EncoderTimer.Enabled = True
    
    tha = RangeHA(vRA - EQnow_lst(gLongitude * DEG_RAD))
  
    If tha < 0 Then
        
        If gHemisphere = 0 Then
            tPier = 1
        Else
            tPier = 0
        End If
        
        tRA = Range24(vRA - 12)
    
    Else
        
        If gHemisphere = 0 Then
            tPier = 0
        Else
            tPier = 1
        End If
        
        tRA = vRA
    
    End If

    'Compute for Sync RA/DEC Encoder Values

     targetRAEncoder = Get_RAEncoderfromRA(tRA, 0, gLongitude, gRAEncoder_Zero_pos, gTot_RA, gHemisphere)
     targetDECEncoder = Get_DECEncoderfromDEC(vDEC, tPier, gDECEncoder_Zero_pos, gTot_DEC, gHemisphere)
          
     gRA1Star = targetRAEncoder - currentRAEncoder
     gDEC1Star = targetDECEncoder - currentDECEncoder
     Call resetsync  ' Sync data will now be invalid at this point
      
     HC.DxAlbl.Caption = Format$(Str(gRA1Star), "000000000")
     HC.DxBlbl.Caption = Format$(Str(gDEC1Star), "000000000")
     

End Sub

Public Sub resetsync()

    gRASync01 = 0
    gDECSync01 = 0
    
    WriteSyncMap
    
    HC.DxSalbl.Caption = Format$(Str(gRASyncy01), "000000000")
    HC.DxSblbl.Caption = Format$(Str(gDECSync01), "000000000")
    
End Sub

Public Sub writeratebarstateHC()

    oProfile.WriteValue oID, "BAR01_1", CStr(HC.VScroll1.value)
    oProfile.WriteValue oID, "BAR01_2", CStr(HC.VScroll2.value)
    oProfile.WriteValue oID, "BAR01_3", CStr(HC.VScroll3.value)
    oProfile.WriteValue oID, "BAR01_4", CStr(HC.VScroll4.value)
    oProfile.WriteValue oID, "BAR01_5", CStr(HC.VScroll5.value)
    oProfile.WriteValue oID, "BAR01_6", CStr(HC.VScroll6.value)

End Sub

Public Sub readratebarstateHC()

     Dim tmptxt As String

     tmptxt = oProfile.GetValue(oID, "BAR01_1")
     If tmptxt <> "" Then
        HC.VScroll1.value = val(tmptxt)
     End If
     
     tmptxt = oProfile.GetValue(oID, "BAR01_2")
     If tmptxt <> "" Then
        HC.VScroll2.value = val(tmptxt)
     End If
 
     tmptxt = oProfile.GetValue(oID, "BAR01_3")
     If tmptxt <> "" Then
        HC.VScroll3.value = val(tmptxt)
     End If
     
     tmptxt = oProfile.GetValue(oID, "BAR01_4")
     If tmptxt <> "" Then
        HC.VScroll4.value = val(tmptxt)
     End If

     tmptxt = oProfile.GetValue(oID, "BAR01_5")
     If tmptxt <> "" Then
        HC.VScroll5.value = val(tmptxt)
     End If
     
     tmptxt = oProfile.GetValue(oID, "BAR01_6")
     If tmptxt <> "" Then
        HC.VScroll6.value = val(tmptxt)
     End If
End Sub

Public Sub writeratebarstateAlign()

    oProfile.WriteValue oID, "BAR02_1", CStr(Align.HScroll1.value)
    oProfile.WriteValue oID, "BAR02_2", CStr(Align.HScroll2.value)


End Sub

Public Sub readratebarstateAlign()

     Dim tmptxt As String

     tmptxt = oProfile.GetValue(oID, "BAR02_1")
     If tmptxt <> "" Then
        Align.HScroll1.value = val(tmptxt)
     End If
     
     tmptxt = oProfile.GetValue(oID, "BAR02_2")
     If tmptxt <> "" Then
        Align.HScroll2.value = val(tmptxt)
     End If
 
  
End Sub

Public Sub writeratebarstatePad()

    oProfile.WriteValue oID, "BAR03_1", CStr(Slewpad.VScroll1.value)
    oProfile.WriteValue oID, "BAR03_2", CStr(Slewpad.VScroll2.value)


End Sub

Public Sub readratebarstatePad()

     Dim tmptxt As String

     tmptxt = oProfile.GetValue(oID, "BAR03_1")
     If tmptxt <> "" Then
        Slewpad.VScroll1.value = val(tmptxt)
     End If
     
     tmptxt = oProfile.GetValue(oID, "BAR03_2")
     If tmptxt <> "" Then
        Slewpad.VScroll2.value = val(tmptxt)
     End If
   
End Sub
