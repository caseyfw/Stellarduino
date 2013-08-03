VERSION 5.00
Begin VB.Form Slewpad 
   BackColor       =   &H00000000&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "EQMOD Mouse Slew Pad"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   2055
      Left            =   6720
      Max             =   809
      Min             =   1
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1080
      Value           =   400
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0095C1CB&
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6960
      Width           =   1335
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   2055
      Left            =   6720
      Max             =   809
      Min             =   1
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3840
      Value           =   400
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Slew Region"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6495
   End
   Begin VB.VScrollBar VScroll3 
      Height          =   615
      Left            =   6840
      TabIndex        =   12
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "RA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6720
      TabIndex        =   10
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000040&
      Caption         =   "Put the mouse cursor on the Slew Region below and click the mouse buttons to slew the mount"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   6975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "EQMOD MOUSE BUTTON SLEW PAD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "DEC"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6720
      TabIndex        =   11
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000040&
      Caption         =   "Slew: NONE"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   6960
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000040&
      Caption         =   "BUTTONS: Left:West, Right: East, Middle: Switch Axis"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   6360
      Width           =   6975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000040&
      Caption         =   "400"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6720
      TabIndex        =   4
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
      Caption         =   "400"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6720
      TabIndex        =   3
      Top             =   3240
      Width           =   375
   End
End
Attribute VB_Name = "Slewpad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
' SlewPad.frm - Slew Window
'
'
' Written:  07-Oct-06   Raymund Sarmiento
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' 04-Nov-06 rcs     Initial edit for EQ Mount Driver Function Prototype
' 17-Dec-06 rcs     Added Numeric Keypad Access
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
Public sPadMode As Boolean
Public sPadLastState As Long


Private Sub Command1_Click()
    writeratebarstatePad
    Unload Slewpad
End Sub



Private Sub Form_Activate()
    Slewpad.VScroll1.Value = HC.VScrollRASlewRate.Value
    Slewpad.VScroll2.Value = HC.VScrollDecSlewRate.Value
    WheelHook (Me.hwnd)
End Sub


Private Sub Form_Deactivate()
    WheelUnHook (Me.hwnd)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Call keydown(KeyCode, val(Slewpad.VScroll1.Value), val(Slewpad.VScroll2.Value))
    Slewpad.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    Call keyup(KeyCode)
    Slewpad.SetFocus

End Sub

Private Sub Form_Load()
Dim tmptxt As String
      
    Call SetText
    
    EnableCloseButton Me.hwnd, False
    
    If HC.HCOnTop.Value = 1 Then Call PutWindowOnTop(Slewpad)
    
    tmptxt = HC.oPersist.ReadIniValue("SlewPadWidth")
    If tmptxt <> "" Then Slewpad.width = val(tmptxt)
    tmptxt = HC.oPersist.ReadIniValue("SlewPadHeight")
    If tmptxt <> "" Then Slewpad.Height = val(tmptxt)
    
    sPadMode = False
    sPadLastState = 0
    readratebarstatePad
    


End Sub



Private Sub Form_Resize()
    On Error Resume Next
    VScroll1.Left = Slewpad.width - 630
    VScroll2.Left = Slewpad.width - 630
    VScroll3.Left = Slewpad.width - 630
    Label1.Left = Slewpad.width - 630
    Label2.Left = Slewpad.width - 630
    Label8.Left = Slewpad.width - 630
    Label7.Left = Slewpad.width - 630
    Label5.width = Slewpad.width - 255
    Label3.width = Slewpad.width - 375

    Frame1.width = Slewpad.width - 855
    Frame1.Height = Slewpad.Height - 2340
    
    Label3.Top = Slewpad.Height - 1395
    Label4.Top = Slewpad.Height - 795
    Command1.Top = Slewpad.Height - 795
    Command1.Left = Slewpad.width - 1590
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call HC.oPersist.WriteIniValue("SlewPadWidth", CStr(Slewpad.width))
    Call HC.oPersist.WriteIniValue("SlewPadHeight", CStr(Slewpad.Height))
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)


    Select Case (Button)
    
    Case 1 ' Left Button
    
        If sPadMode = False Then
        
            Call West_Down(val(Slewpad.VScroll1.Value))
            Label4.Caption = oLangDll.GetLangString(807)
        
        Else
        
            Call North_Down(val(Slewpad.VScroll2.Value))
            Label4.Caption = oLangDll.GetLangString(808)
        
        End If

    
    Case 2 ' Right Button
    
        If sPadMode = False Then
        
            Call East_Down(val(Slewpad.VScroll1.Value))
            Label4.Caption = oLangDll.GetLangString(809)
            
        Else
            Call South_Down(val(Slewpad.VScroll2.Value))
            Label4.Caption = oLangDll.GetLangString(810)
            
        End If

    
    Case Else ' Assume Middle Button
    
        If sPadMode = False Then
            sPadMode = True
            Slewpad.Label3.Caption = oLangDll.GetLangString(811)
        Else
            sPadMode = False
            Slewpad.Label3.Caption = oLangDll.GetLangString(804)
        End If
            
    End Select

    sPadLastState = Button
    
spadEND01:
End Sub

Private Sub Frame1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    
    
    Select Case (sPadLastState)
        Case 1
            If sPadMode = False Then
            
                Call West_Up
            Else
            
                Call North_Up
            
            End If

        Case 2
            If sPadMode = False Then
            
                Call East_Up
                
            Else
                Call South_Up
                
            End If
        Case Else
        
                If (sPadLastState & 1) <> 0 Then
                        If sPadMode = False Then
                            Call West_Up
                        Else
                            Call North_Up
            
                        End If
                End If
                
                If (sPadLastState & 2) <> 0 Then
                        If sPadMode = False Then
                            Call East_Up
                        Else
                            Call South_Up
                        End If
                End If
 
    End Select
    
   Label4.Caption = oLangDll.GetLangString(812)
    
spadEND02:
End Sub



Private Sub Vscroll3_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Label4.Caption = oLangDll.GetLangString(813)
    Call keydown(KeyCode, val(Slewpad.VScroll1.Value), val(Slewpad.VScroll2.Value))
    Me.SetFocus
End Sub

Private Sub Vscroll3_KeyUp(KeyCode As Integer, Shift As Integer)
    Call keyup(KeyCode)
    Label4.Caption = oLangDll.GetLangString(812)
    Me.SetFocus
End Sub

Private Sub VScroll1_Change()
    If VScroll1.Value >= 10 Then
        Slewpad.Label1.Caption = VScroll1.Value - 9
    Else
        Slewpad.Label1.Caption = "0." & VScroll1.Value
    End If
End Sub

Private Sub VScroll1_GotFocus()
    VScroll3.SetFocus
End Sub

Private Sub VScroll1_Scroll()
    Call VScroll1_Change
End Sub

Private Sub VScroll2_Change()
    If VScroll2.Value >= 10 Then
        Slewpad.Label2.Caption = VScroll2.Value - 9
    Else
        Slewpad.Label2.Caption = "0." & VScroll2.Value
    End If
End Sub

Private Sub VScroll2_GotFocus()
    VScroll3.SetFocus
End Sub

Private Sub VScroll2_Scroll()
    Call VScroll2_Change
End Sub

Public Sub Mousewheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
   
   Dim i As Double
      
   If sPadMode = False Then
        i = Slewpad.VScroll1.Value
        If Rotation > 0 Then
            i = i + 20
            If i >= 809 Then i = 809
        Else
            i = i - 20
            If i <= 0 Then i = 1
        End If
        Slewpad.VScroll1.Value = i
   Else
        i = Slewpad.VScroll2.Value
        If Rotation > 0 Then
            i = i + 20
            If i >= 809 Then i = 809
        Else
            i = i - 20
            If i <= 0 Then i = 1
        End If
        Slewpad.VScroll2.Value = i
   End If
   
End Sub

Private Sub SetText()
    Slewpad.Caption = oLangDll.GetLangString(800)
    Label6.Caption = oLangDll.GetLangString(801)
    Label5.Caption = oLangDll.GetLangString(802)
    Frame1.Caption = oLangDll.GetLangString(803)
    Label3.Caption = oLangDll.GetLangString(804)
    Label4.Caption = oLangDll.GetLangString(805)
    Command1.Caption = oLangDll.GetLangString(806)
    Label7.Caption = oLangDll.GetLangString(105)
    Label8.Caption = oLangDll.GetLangString(106)
End Sub

