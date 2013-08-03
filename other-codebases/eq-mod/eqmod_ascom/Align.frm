VERSION 5.00
Begin VB.Form Align 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Alignment"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      BackColor       =   &H0095C1CB&
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0095C1CB&
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton End_Command 
      BackColor       =   &H0095C1CB&
      Caption         =   "END"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2760
      Top             =   2760
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H0095C1CB&
      Caption         =   "PAD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3360
      Width           =   975
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   135
      Left            =   2160
      Max             =   800
      Min             =   1
      TabIndex        =   8
      Top             =   4080
      Value           =   400
      Width           =   1935
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   135
      Left            =   120
      Max             =   800
      Min             =   1
      TabIndex        =   7
      Top             =   4080
      Value           =   400
      Width           =   1935
   End
   Begin VB.CommandButton Accept_Command 
      BackColor       =   &H0095C1CB&
      Caption         =   "Accept"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0095C1CB&
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Abort_Command 
      BackColor       =   &H0095C1CB&
      Caption         =   "ABORT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0095C1CB&
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
   Begin VB.Label RA_Target 
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label RA_Tgt_Label 
      BackColor       =   &H00000000&
      Caption         =   "RA Target"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Dec_Target 
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1560
      TabIndex        =   17
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Star_TextLabel 
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "DEC Slew Rate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "RA Slew Rate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000040&
      Caption         =   "400"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3720
      TabIndex        =   10
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000040&
      Caption         =   "400"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label SecondLabel 
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   3975
      WordWrap        =   -1  'True
   End
   Begin VB.Label Star_Label 
      BackColor       =   &H00000000&
      Caption         =   "Alignmnt. star"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   3000
      TabIndex        =   15
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label DEC_Tgt_Label 
      BackColor       =   &H00000000&
      Caption         =   "DEC Target"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   1560
      TabIndex        =   16
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Align"
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
' Align.frm - 1 Star alignment form
'
' Written:  07-Oct-06   Raymund Sarmiento
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' 04-Nov-06 rcs     Initial edit for EQ Mount Driver Function Prototype
' 05-Jan-07 rcs     Added Joystick Button Accept/Abort
' 06-May-07 rcs     Added Gamepad accept on N-Star Align
' 26-Jul-07 cs      Gamepad accept/cancel user 'variable' button definition
' 30-Jul-07 cs      Gamepad alignment end handling
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

Private aUtil As DriverHelper.Util
Public CurrentAlignmentStar As Integer



Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    If gEQparkstatus <> 0 Then
            HC.Add_Message (oLangDll.GetLangString(5000))
            Exit Sub
    End If
    
    Call West_Down(val(HScroll1.Value) + 9)
    HC.Add_Message (oLangDll.GetLangString(5010))
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
   
   Call West_Up
   
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If gEQparkstatus <> 0 Then
            HC.Add_Message (oLangDll.GetLangString(5000))
            Exit Sub
    End If
    
    Call South_Down(val(HScroll2.Value) + 9)
    HC.Add_Message (oLangDll.GetLangString(5011))
End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call South_Up
    
End Sub

Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    If gEQparkstatus <> 0 Then
            HC.Add_Message (oLangDll.GetLangString(5000))
            Exit Sub
    End If
        
    Call North_Down(val(HScroll2.Value) + 9)
    HC.Add_Message (oLangDll.GetLangString(5001))

End Sub

Private Sub Command3_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call North_Up
End Sub



Private Sub Command4_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  
   If gEQparkstatus <> 0 Then
            HC.Add_Message (oLangDll.GetLangString(5000))
            Exit Sub
   End If
  
   Call East_Down(val(HScroll1.Value) + 9)
   HC.Add_Message (oLangDll.GetLangString(5009))
   
End Sub

Private Sub Command4_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    Call East_Up
    
End Sub

Private Sub Accept_Command_Click()

    Call AcceptClick

End Sub

Private Sub End_Command_Click()

    Align.Timer1.Enabled = False
    HC.Add_Message (oLangDll.GetLangString(5022) & str(gAlignmentStars_count))
    writeratebarstateAlign
    Unload Align
        
End Sub

Private Sub Abort_Command_Click()
    Align.Timer1.Enabled = False
    HC.Add_Message (oLangDll.GetLangString(5023))
    writeratebarstateAlign
    Unload Align
End Sub

Private Sub Command7_Click()
    Load Slewpad
    Slewpad.Show
End Sub

Private Sub Form_Load()

    Call SetText

    If HC.HCOnTop.Value = 1 Then Call PutWindowOnTop(Align)

    End_Command.Enabled = False
    gEQjbuttons = 0
    Align.Timer1.Enabled = True
   
    EnableCloseButton Me.hwnd, False
    Set aUtil = New DriverHelper.Util
    readratebarstateAlign
    Accept_Command.Enabled = False
    SecondLabel.Caption = oLangDll.GetLangString(508)
    If gAlignmentStars_count >= 3 Then
        End_Command.Enabled = True
    Else
        End_Command.Enabled = False
    End If

    gJoyTimerFlag2 = True
    
    RA_Target = ""
    Dec_Target = ""
    Abort_Command.Visible = True
    End_Command.Visible = True
    End_Command.Enabled = False
    If gAlignmentStars_count <= 1 Then
        CurrentAlignmentStar = 1
        Star_TextLabel = "1 / N"
    Else
        CurrentAlignmentStar = gAlignmentStars_count + 1
        Star_TextLabel = CurrentAlignmentStar & " / N"
    End If
    Abort_Command.Visible = False
    End_Command.Enabled = True

End Sub

Private Sub HScroll1_Change()
    Label2.Caption = HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
    Label2.Caption = HScroll1.Value
End Sub

Private Sub HScroll2_Change()
    Label3.Caption = HScroll2.Value
End Sub

Private Sub HScroll2_Scroll()
    Label3.Caption = HScroll2.Value
End Sub

Private Sub Timer1_Timer()

  If gJoyTimerFlag2 = True Then         ' Avoid Overruns

    gJoyTimerFlag2 = False
    
    If BTN_ALIGNACCEPT <> BTN_UNDEFINED Then
        If (gEQjbuttons = BTN_ALIGNACCEPT) Then ' Accept
            If Accept_Command.Enabled = True Then
                gEQjbuttons = 0
                EQ_Beep (21)
                Call AcceptClick
            End If
        End If
    End If
    
    If BTN_ALIGNCANCEL <> BTN_UNDEFINED Then
        If (gEQjbuttons = BTN_ALIGNCANCEL) Then ' Abort
            gEQjbuttons = 0
            Align.Timer1.Enabled = False
            writeratebarstateAlign
            HC.Add_Message (oLangDll.GetLangString(5023))
            EQ_Beep (22)
            Unload Align
        End If
    End If

    If BTN_ALIGNEND <> BTN_UNDEFINED Then
        If (gEQjbuttons = BTN_ALIGNEND) Then ' End
            gEQjbuttons = 0
            EQ_Beep (23)
            Call End_Command_Click
        End If
    End If

    gJoyTimerFlag2 = True
    
  End If

End Sub

Public Sub FillAlignmentStar(RA As Double, DEC As Double)

    HC.Add_Message (oLangDll.GetLangString(5024) & " " & oLangDll.GetLangString(105) & "[ " & FmtSexa(RA, False) & " ] " & oLangDll.GetLangString(106) & "[ " & FmtSexa(DEC, True) & " ]")
    
    gRA_GOTO = RA
    gDEC_GOTO = DEC

    RA_Target = RA
    Dec_Target = DEC
    
    SecondLabel = oLangDll.GetLangString(5025)
    Accept_Command.Enabled = True
    
End Sub



Public Sub AcceptClick()
Dim tRA As Double
Dim tha As Double
Dim tPier As Double
Dim vRA As Double
Dim vDEC As Double

Dim deltaRA As Double
Dim deltadec As Double

    If gSlewStatus = True Then
        ' still slewing
        SecondLabel = oLangDll.GetLangString(5028)
        Exit Sub
    End If

    Align.Timer1.Enabled = False
    
    ' add new point
    Call EQ_NPointAppend(RA_Target, Dec_Target, gLongitude, gHemisphere)
    ' need to do more stars, pop to the next and disable the accept button
    SecondLabel = oLangDll.GetLangString(5030)
    CurrentAlignmentStar = CurrentAlignmentStar + 1
    Star_TextLabel = CurrentAlignmentStar & " / N"
    RA_Target = ""
    Dec_Target = ""
    Accept_Command.Enabled = False
    writeratebarstateAlign
    Align.Timer1.Enabled = True

End Sub


Private Sub SetText()
    Align.Caption = oLangDll.GetLangString(500)
    RA_Tgt_Label.Caption = oLangDll.GetLangString(501)
    DEC_Tgt_Label.Caption = oLangDll.GetLangString(502)
    Star_Label.Caption = oLangDll.GetLangString(503)
    Command1.Caption = oLangDll.GetLangString(115)
    Command2.Caption = oLangDll.GetLangString(116)
    Command3.Caption = oLangDll.GetLangString(113)
    Command4.Caption = oLangDll.GetLangString(114)
    Command7.Caption = oLangDll.GetLangString(112)
    Label4.Caption = oLangDll.GetLangString(117)
    Accept_Command.Caption = oLangDll.GetLangString(504)
    End_Command.Caption = oLangDll.GetLangString(505)
    Abort_Command.Caption = oLangDll.GetLangString(506)
    Label5.Caption = oLangDll.GetLangString(118)
End Sub
