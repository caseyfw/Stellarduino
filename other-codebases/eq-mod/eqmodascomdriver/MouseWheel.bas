Attribute VB_Name = "Mousewheel"
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
' Mousewheel.bas - Mousewheel Module for slewpad
'
'
' Written:  07-Oct-06   Raymund Sarmiento
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' 20-Nov-06 rcs     Initial edit for EQ Mount Driver Function Prototype
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

' Store WndProcs
Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" ( _
                ByVal hWnd As Long, _
                ByVal lpString As String) As Long

Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" ( _
                ByVal hWnd As Long, _
                ByVal lpString As String, _
                ByVal hData As Long) As Long

Private Declare Function RemoveProp Lib "user32.dll" Alias "RemovePropA" ( _
                ByVal hWnd As Long, _
                ByVal lpString As String) As Long

' Hooking
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" ( _
                ByVal lpPrevWndFunc As Long, _
                ByVal hWnd As Long, _
                ByVal Msg As Long, _
                ByVal wParam As Long, _
                ByVal lParam As Long) As Long

Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" ( _
                ByVal hWnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
                ByVal hWnd As Long, _
                ByVal Msg As Long, _
                wParam As Any, _
                lParam As Any) As Long

' Position Checking
Private Declare Function GetWindowRect Lib "user32" ( _
                ByVal hWnd As Long, _
                lpRect As RECT) As Long
                
Private Declare Function GetParent Lib "user32" ( _
                ByVal hWnd As Long) As Long

Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL = &H20A
Private Const CB_GETDROPPEDSTATE = &H157

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

' Check Messages
' ================================================
Private Function WindowProc(ByVal Lwnd As Long, ByVal Lmsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim MouseKeys As Long
  Dim Rotation As Long
  Dim Xpos As Long
  Dim Ypos As Long
  Dim fFrm As Form

  Select Case Lmsg
  
    Case WM_MOUSEWHEEL
    
      MouseKeys = wParam And 65535
      Rotation = wParam / 65536
      Xpos = lParam And 65535
      Ypos = lParam / 65536
      
      Set fFrm = GetForm(Lwnd)
      If fFrm Is Nothing Then
        ' it's not a form
        If Not IsOver(Lwnd, Xpos, Ypos) And IsOver(GetParent(Lwnd), Xpos, Ypos) Then
          ' it's not over the control and is over the form,
          ' so fire mousewheel on form (if it's not a dropped down combo)
          If SendMessage(Lwnd, CB_GETDROPPEDSTATE, 0&, 0&) <> 1 Then
            GetForm(GetParent(Lwnd)).Mousewheel MouseKeys, Rotation, Xpos, Ypos
            Exit Function ' Discard scroll message to control
          End If
        End If
      Else
        ' it's a form so fire mousewheel
        If IsOver(fFrm.hWnd, Xpos, Ypos) Then fFrm.Mousewheel MouseKeys, Rotation, Xpos, Ypos
      End If
  End Select
  
  WindowProc = CallWindowProc(GetProp(Lwnd, "PrevWndProc"), Lwnd, Lmsg, wParam, lParam)
End Function

' Hook / UnHook
' ================================================
Public Sub WheelHook(ByVal hWnd As Long)
  On Error Resume Next
  SetProp hWnd, "PrevWndProc", SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub WheelUnHook(ByVal hWnd As Long)
  On Error Resume Next
  SetWindowLong hWnd, GWL_WNDPROC, GetProp(hWnd, "PrevWndProc")
  RemoveProp hWnd, "PrevWndProc"
End Sub

' Window Checks
' ================================================
Public Function IsOver(ByVal hWnd As Long, ByVal lX As Long, ByVal lY As Long) As Boolean
  Dim rectCtl As RECT
  GetWindowRect hWnd, rectCtl
  With rectCtl
    IsOver = (lX >= .Left And lX <= .Right And lY >= .Top And lY <= .Bottom)
  End With
End Function

Private Function GetForm(ByVal hWnd As Long) As Form
  For Each GetForm In Forms
    If GetForm.hWnd = hWnd Then Exit Function
  Next GetForm
  Set GetForm = Nothing
End Function





