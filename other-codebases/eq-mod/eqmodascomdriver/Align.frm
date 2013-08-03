VERSION 5.00
Begin VB.Form Align 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "1 Star Align"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "PAD"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   120
      TabIndex        =   15
      Top             =   5040
      Width           =   855
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   120
      Max             =   800
      Min             =   1
      TabIndex        =   9
      Top             =   6600
      Value           =   400
      Width           =   2775
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      Max             =   800
      Min             =   1
      TabIndex        =   8
      Top             =   6000
      Value           =   400
      Width           =   2775
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Accept"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1680
      TabIndex        =   5
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3240
      TabIndex        =   4
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1680
      TabIndex        =   3
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ABORT"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   7200
      Width           =   4695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1680
      TabIndex        =   0
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "EQMOD 1-STAR ALIGNMENT TOOL"
      BeginProperty Font 
         Name            =   "Arial Black"
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
      TabIndex        =   14
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "DEC Slew Rate"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "RA Slew Rate"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000080&
      Caption         =   "400"
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
      Left            =   3240
      TabIndex        =   11
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000080&
      Caption         =   "400"
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
      Left            =   3240
      TabIndex        =   10
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
      Caption         =   "text"
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
      TabIndex        =   7
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Label MainLabel 
      BackColor       =   &H00000080&
      Caption         =   "text"
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
      TabIndex        =   6
      Top             =   840
      Width           =   4575
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


Private aUtil As DriverHelper.Util


Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    eqres = EQ_MotorStop(0)          ' Stop RA Motor
    If eqres <> EQ_OK Then
            GoTo END04
    End If

    'Wait until RA motor is stable

   Do
       eqres = EQ_GetMotorStatus(0)
       If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo END04
       
   Loop While (eqres And EQ_MOTORBUSY) <> 0
   eqres = EQ_Slew(0, 0, 0, val(HScroll1.Value))
   
END04:
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   eqres = EQ_MotorStop(0)          ' Stop RA Motor
    If eqres <> EQ_OK Then
            GoTo END01
    End If

    'Wait until RA motor is stable

   Do
       eqres = EQ_GetMotorStatus(0)
       If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo END01
       
   Loop While (eqres And EQ_MOTORBUSY) <> 0

    If gTrackingStatus <> 0 Then eqres = EQ_StartRATrack(gTrackingStatus - 1, gHemisphere, gHemisphere)

END01:
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    eqres = EQ_MotorStop(1)          ' Stop DEC Motor
    If eqres <> EQ_OK Then GoTo END12


   ' Wait for motor stop

   Do
        
       eqres = EQ_GetMotorStatus(1)
       If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo END12
    
   Loop While (eqres And EQ_MOTORBUSY) <> 0

   eqres = EQ_Slew(1, 0, 1, val(HScroll2.Value))

END12:
End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    eqres = EQ_MotorStop(1)
End Sub

Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    eqres = EQ_MotorStop(1)          ' Stop DEC Motor
    If eqres <> EQ_OK Then
            GoTo END13
    End If

   ' Wait for motor stop

   Do
        
       eqres = EQ_GetMotorStatus(1)
       If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo END13
    
   Loop While (eqres And EQ_MOTORBUSY) <> 0
   
   eqres = EQ_Slew(1, 0, 0, val(HScroll2.Value))
    
END13:
End Sub

Private Sub Command3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    eqres = EQ_MotorStop(1)
End Sub



Private Sub Command4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
    eqres = EQ_MotorStop(0)          ' Stop RA Motor
    
    If eqres <> EQ_OK Then
            GoTo END05
    End If

    'Wait until RA motor is stable

   Do
       eqres = EQ_GetMotorStatus(0)
       If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo END05

   Loop While (eqres And EQ_MOTORBUSY) <> 0

   eqres = EQ_Slew(0, 0, 1, val(HScroll1.Value))

END05:
End Sub

Private Sub Command4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    eqres = EQ_MotorStop(0)          ' Stop RA Motor
    If eqres <> EQ_OK Then
            GoTo END02
    End If


   Do
        
       eqres = EQ_GetMotorStatus(0)
       If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo END02
    
   Loop While (eqres And EQ_MOTORBUSY) <> 0
   
   If gTrackingStatus > 0 Then eqres = EQ_StartRATrack(gTrackingStatus - 1, gHemisphere, gHemisphere)            ' Track RA Motor at Sidreal Rate

END02:
End Sub

Private Sub Command5_Click()
    Call OneStarAlign(aUtil.HMSToHours(Mid(HC.StarList.Text, 3, 10)), aUtil.DMSToDegrees(Mid(HC.StarList.Text, 14, 10)))
    writeratebarstateAlign
    Unload Align
End Sub

Private Sub Command6_Click()
    writeratebarstateAlign
    Unload Align
End Sub

Private Sub Command7_Click()
    Load Slewpad
    Slewpad.Show
End Sub

Private Sub Form_Load()
    EnableCloseButton Me.hWnd, False
    Set aUtil = New DriverHelper.Util
    readratebarstateAlign
    MainLabel.Caption = "Align your scope's FOV to " & Mid(HC.StarList.Text, 24) & " then click 'ACCEPT' "
    Label1.Caption = "Location: " & "RA[ " & Mid(HC.StarList.Text, 3, 10) & "] DEC[ " & Mid(HC.StarList.Text, 14, 10) & " ]"
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
