VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "EQ Variable Rate Guider v1.0"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8310
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox gHemisphere 
      Height          =   195
      Left            =   7080
      TabIndex        =   46
      Top             =   1800
      Width           =   255
   End
   Begin VB.Timer PulseGuide_timer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7320
      Top             =   600
   End
   Begin VB.CheckBox ScopeType 
      Height          =   195
      Left            =   6720
      TabIndex        =   44
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear Logs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   40
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Start Tracking"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7080
      TabIndex        =   38
      Top             =   960
      Width           =   975
   End
   Begin VB.Frame Frame4 
      Caption         =   " Guide Indicator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3120
      TabIndex        =   33
      Top             =   840
      Width           =   1695
      Begin VB.Label DirLabel4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   600
         TabIndex        =   37
         Top             =   1200
         Width           =   465
      End
      Begin VB.Label DirLabel2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "W"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Width           =   495
      End
      Begin VB.Label DirLabel1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   600
         TabIndex        =   35
         Top             =   240
         Width           =   495
      End
      Begin VB.Label DirLabel3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1080
         TabIndex        =   34
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Mount Centering Commands"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   25
      Top             =   840
      Width           =   2775
      Begin VB.CommandButton Command4 
         Caption         =   "W"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton Command7 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   29
         Top             =   1200
         Width           =   495
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1335
         Left            =   1680
         Max             =   800
         Min             =   1
         TabIndex        =   28
         Top             =   360
         Value           =   2
         Width           =   255
      End
      Begin VB.CommandButton Command5 
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   27
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton Command6 
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   26
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "SLEW"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   41
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "RATE"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   32
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080FF80&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   31
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.VScrollBar VScroll3 
      Height          =   1215
      Left            =   6000
      Max             =   9
      Min             =   1
      TabIndex        =   20
      Top             =   1320
      Value           =   5
      Width           =   255
   End
   Begin VB.Frame Frame2 
      Caption         =   "GuideRate ( x Sidreal)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4920
      TabIndex        =   18
      Top             =   840
      Width           =   2055
      Begin VB.VScrollBar VScroll2 
         Height          =   1215
         Left            =   120
         Max             =   9
         Min             =   1
         TabIndex        =   19
         Top             =   480
         Value           =   5
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "RATE"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   43
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "RATE"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   42
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label17 
         BackColor       =   &H0080FF80&
         Caption         =   ".5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   24
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label16 
         BackColor       =   &H0080FF80&
         Caption         =   " .5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   23
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label15 
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
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label14 
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
         Height          =   255
         Left            =   1080
         TabIndex        =   21
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CheckBox Check_NS_Enable 
      Caption         =   "Check1"
      Height          =   195
      Left            =   2040
      TabIndex        =   15
      Top             =   2880
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Check_WE_Enable 
      Caption         =   "Check1"
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   2880
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2520
      TabIndex        =   12
      Text            =   "COM1"
      Top             =   360
      Width           =   1095
   End
   Begin VB.CheckBox CheckN_S 
      Caption         =   "Check2"
      Height          =   195
      Left            =   5400
      TabIndex        =   9
      Top             =   2880
      Width           =   255
   End
   Begin VB.CheckBox CheckW_E 
      Caption         =   "Check1"
      Height          =   195
      Left            =   3840
      TabIndex        =   8
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   120
      Width           =   2295
   End
   Begin VB.CheckBox CheckLOG1 
      Caption         =   "Check1"
      Height          =   195
      Left            =   7320
      TabIndex        =   3
      Top             =   240
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.TextBox txtPort 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      TabIndex        =   2
      Text            =   "999"
      Top             =   360
      Width           =   975
   End
   Begin MSWinsockLib.Winsock wsClient 
      Left            =   0
      Top             =   6480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Command Logs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   8055
      Begin VB.TextBox List1 
         Height          =   3495
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   39
         Top             =   240
         Width           =   7815
      End
   End
   Begin VB.Label Label19 
      Caption         =   "Reverse"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   47
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label18 
      Caption         =   "LX200GPS Mode"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   45
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Enable N-S Guide"
      Height          =   255
      Left            =   2280
      TabIndex        =   17
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "Enable W-E Guide"
      Height          =   255
      Left            =   480
      TabIndex        =   16
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "EQInterface :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   13
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Swap N<--> S"
      Height          =   255
      Left            =   5640
      TabIndex        =   11
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Swap W<--> E"
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   8160
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label3 
      Caption         =   "LX200 PROTOCOL EMULATOR"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "EQ/ATLAS/EQG GUIDER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2295
   End
   Begin VB.Line Line4 
      X1              =   8160
      X2              =   8160
      Y1              =   2760
      Y2              =   720
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   120
      Y1              =   720
      Y2              =   2760
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8160
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Log01 
      Caption         =   "LOG"
      Height          =   255
      Left            =   7680
      TabIndex        =   4
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "TCP Port :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
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
'
' Written:  07-Oct-06   Raymund Sarmiento
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' 24-Oct-06 rcs     Initial edit for EQ Mount Driver GUIDER PROGRAM
' 26-Nov-06 rcs     Added LX200 PulseDuration functions
'---------------------------------------------------------------------
'
'
'  SYNOPSIS:
'
'  This is a demonstration of a EQ6/ATLAS/EQG direct stepper motor control access
'  using the EQCONTRL.DLL driver code. The code utilizes the LX200 protocol to
'  accept guiding commands from an external autoguiding software via Virtual COM
'  PORT -> TCP/IP. The program then accepts the stream and converts them to
'  EQContrl guide commands.
'
'  File EQCONTROL.bas contains all the function prototypes of all subroutines
'  encoded in the EQCONTRL.dll
'
'  The EQCONTRL.DLL simplifies execution of the Mount controller board stepper
'  commands.
'
'  The mount circuitry needs to be modified for this  program to work.
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


Option Explicit


Const L_PIPE As String = "#"


Dim Command_Format As Integer

Dim gEQDECPulseDuration As Double
Dim gEQRAPulseDuration As Double

Dim gEQPulsetimerflag As Boolean


Dim eqres As Long
Dim lxlastcommand As Integer            'Variable to block duplicate commands



Private Sub Command1_Click()



    DirLabel1.BackColor = &HFFFF00
    DirLabel2.BackColor = &HFFFF00
    DirLabel3.BackColor = &HFFFF00
    DirLabel4.BackColor = &HFFFF00
    
    Call AddLog("EVT: Ports Cleared")
    
    

End Sub
' Routine to clear the log window
Private Sub Command2_Click()
    List1.Text = ""
End Sub

' Routine to connect to the mount and listen for LX200 streams from the TCP Port

Private Sub Command3_Click()

   If Command3.Caption <> "Stop/Disconnect" Then
   
    Command3.Caption = "Connecting .."
    eqres = EQ_Init(Combo1.Text, "9600", 1000, 1)
    If eqres <> 0 Then
      Call AddLog0("Cannot Connect to the EQModded mount, Check COM port settings")
      Command3.Caption = "START"
    Else
      Call AddLog0("EQModded mount found at port " & Combo1.Text)
      If txtPort = "" Then
            Call AddLog0("Invalid Port. Enter value between 888 to 32000")
            Command3.Caption = "START"
       Else
        wsClient.LocalPort = txtPort
        wsClient.Listen
        Command3.Caption = "Stop/Disconnect"
        Call AddLog0("Server started and listening at TCP port " & txtPort)
        
        Combo1.Locked = True
        Combo1.Enabled = False
     End If
    End If
    
  Else 'Caption = "Stop/Disconnect"
  
    wsClient.Close
    Command3.Caption = "START"
    Call AddLog0("Server Closed")
    
    eqres = EQ_End()
    Combo1.Locked = False
    Combo1.Enabled = True
    
  
  End If
    
    
End Sub


' RA- Button Slew

Private Sub Command4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    eqres = EQ_MotorStop(0)          ' Stop RA Motor
    If eqres <> 0 Then
            AddLog ("Error[RA+ ]: Cannot stop RA motor, mount not connected")
            GoTo END05
    Else
            AddLog ("Stopping RA Motor")
    End If

    'Wait until RA motor is stable

   Do
       eqres = EQ_GetMotorStatus(0)
       If eqres = 1 Then
            AddLog ("Error Getting mount status")
            GoTo END05
       End If
   Loop While (eqres And &H10) <> 0


    eqres = EQ_Slew(0, 0, 1, Val(VScroll1.value))
    If eqres <> 0 Then
        AddLog ("Error[RA- ]: Mount not connected")
    Else
        AddLog ("Slewing at RA-")
    End If

END05:

End Sub

Private Sub Command4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    eqres = EQ_MotorStop(0)          ' Stop RA Motor
    If eqres <> 0 Then
            AddLog ("Error[RA- ]: Cannot stop RA motor, mount not connected")
            GoTo END02
    Else
            AddLog ("Stopping RA Motor")
    End If


   Do
        
       eqres = EQ_GetMotorStatus(0)
       If eqres = 1 Then
            AddLog ("Error Getting mount status")
            GoTo END02
       End If
    
    Loop While (eqres And &H10) <> 0
    
    If Command8.Caption = "Stop Tracking" Then
    
        eqres = EQ_StartRATrack(0, 0, 0)             ' Track RA Motor at Sidreal Rate
        If eqres <> 0 Then
            AddLog ("Error[RA- ]: Cannot resume tracking, mount not connected")
        Else
            AddLog ("Restoring RA tracking rate to Sidreal")
        End If
    End If

    
END02:

End Sub

'DEC + Button slew

Private Sub Command5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    eqres = EQ_Slew(1, 0, 0, Val(VScroll1.value))
    If eqres <> 0 Then
        AddLog ("Error[DEC+ ]: Mount not connected")
    Else
        AddLog ("Slewing at DEC+")
    End If

End Sub



Private Sub Command5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    eqres = EQ_MotorStop(1)
    If eqres <> 0 Then
        AddLog ("Error[DEC+ ]: Cannot Stop DEC Motor, Mount not connected")
    Else
        AddLog ("Stopping DEC Motor")
    End If
End Sub

'RA+ Button Slew

Private Sub Command6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    eqres = EQ_MotorStop(0)          ' Stop RA Motor
    If eqres <> 0 Then
            AddLog ("Error[RA+ ]: Cannot stop RA motor, mount not connected")
            GoTo END04
    Else
            AddLog ("Stopping RA Motor")
    End If

    'Wait until RA motor is stable

   Do
       eqres = EQ_GetMotorStatus(0)
       If eqres = 1 Then
            AddLog ("Error Getting mount status")
            GoTo END04
       End If
   Loop While (eqres And &H10) <> 0
   
    eqres = EQ_Slew(0, 0, 0, Val(VScroll1.value))
    If eqres <> 0 Then
       AddLog ("Error[RA+ ]: Mount not connected")
    Else
        AddLog ("Slewing at RA+")
    End If
    
END04:

End Sub



Private Sub Command6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


    eqres = EQ_MotorStop(0)          ' Stop RA Motor
    If eqres <> 0 Then
            AddLog ("Error[RA+ ]: Cannot stop RA motor, mount not connected")
            GoTo END01
    Else
            AddLog ("Stopping RA Motor")
    End If

    'Wait until RA motor is stable

   Do
       eqres = EQ_GetMotorStatus(0)
       If eqres = 1 Then
            AddLog ("Error Getting mount status")
            GoTo END01
       End If
   Loop While (eqres And &H10) <> 0

    If Command8.Caption = "Stop Tracking" Then
    
        eqres = EQ_StartRATrack(0, 0, 0)             ' Track RA Motor at Sidreal Rate
        If eqres <> 0 Then
            AddLog ("Error[RA+ ]: Cannot resume tracking, mount not connected")
        Else
            AddLog ("Restoring RA tracking rate to Sidreal")
        End If
    
    End If
  
END01:
Exit Sub
End Sub

'DEC- Slew button

Private Sub Command7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    eqres = EQ_Slew(1, 0, 1, Val(VScroll1.value))
    If eqres <> 0 Then
        AddLog ("Error[DEC- ]: Mount not connected")
    Else
        AddLog ("Slewing at DEC-")
    End If

End Sub

Private Sub Command7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    eqres = EQ_MotorStop(1)
    If eqres <> 0 Then
        AddLog ("Error[DEC- ]: Cannto Stop DEC motor, Mount not connected")
    Else
        AddLog ("Stopping DEC Motor")
    End If
    
End Sub

'Routine to START/STOP Sidreal tracking

Private Sub Command8_Click()

    If Command8.Caption = "Start Tracking" Then
    
        eqres = EQ_StartRATrack(0, gHemisphere.value, gHemisphere.value)
        
        If eqres <> 0 Then
            AddLog ("Cannot Start tracking, mount not connected")
        Else
           AddLog ("Mount now tracking at the sidreal rate")
           Command8.Caption = "Stop Tracking"
        End If
    Else
    
        PulseGuide_timer.Enabled = False
        
        eqres = EQ_MotorStop(0)                      ' Stop RA Motor
        eqres = EQ_MotorStop(1)
        Command8.Caption = "Start Tracking"
        AddLog ("Mount RA tracking disabled")
        
        PulseGuide_timer.Enabled = False
        
        gEQDECPulseDuration = 0
        gEQRAPulseDuration = 0
    
    End If
    
End Sub

Private Sub Form_Load()

    Combo1.AddItem "COM1"
    Combo1.AddItem "COM2"
    Combo1.AddItem "COM3"
    Combo1.AddItem "COM4"
    Combo1.AddItem "COM5"
    Combo1.AddItem "COM6"
    Combo1.AddItem "COM7"
    Combo1.AddItem "COM8"
    Combo1.AddItem "COM9"
    Combo1.AddItem "COM10"
    Combo1.AddItem "COM11"
    Combo1.AddItem "COM12"
    

    lxlastcommand = 0
    gEQDECPulseDuration = 0
    gEQRAPulseDuration = 0
    
    gEQPulsetimerflag = True
    PulseGuide_timer.Enabled = False
    
   
    
End Sub


Private Sub PulseGuide_timer_Timer()
  ' This is a 100 millisecond time ticker interval for Pulse guide

    If gEQPulsetimerflag Then   ' Guide only when the scope is tracking
        gEQPulsetimerflag = False                      ' Prevent Overruns
        
       
        If gEQDECPulseDuration > 0 Then
            gEQDECPulseDuration = gEQDECPulseDuration - 100
            If gEQDECPulseDuration <= 0 Then
                eqres = EQ_MotorStop(1)
                gEQDECPulseDuration = 0
                Call AddLog("EQMOD: DEC Stopped")
                DirLabel4.BackColor = &HFFFF00
                DirLabel1.BackColor = &HFFFF00
              
            End If
                
        End If
        If gEQRAPulseDuration > 0 Then
            gEQRAPulseDuration = gEQRAPulseDuration - 100
            If gEQRAPulseDuration <= 0 Then
                eqres = EQ_SendGuideRate(0, 0, 0, 0, gHemisphere.value, gHemisphere.value)
                gEQRAPulseDuration = 0
                Call AddLog("EQMOD: RA Normal Speed")
                DirLabel3.BackColor = &HFFFF00
                DirLabel2.BackColor = &HFFFF00
            End If
        End If
        gEQPulsetimerflag = True
    End If
End Sub

Private Sub VScroll1_Change()
    Label5.Caption = VScroll1.value
End Sub

Private Sub VScroll2_Change()
    Label16.Caption = Str(Val(VScroll2.value) * 0.1)
End Sub

Private Sub VScroll3_Change()
    Label17.Caption = Str(Val(VScroll3.value) * 0.1)
End Sub

Private Sub wsClient_Close()
    wsClient.Close
    wsClient.Listen
    Call AddLog("Client Disconnected")
End Sub

Private Sub wsClient_ConnectionRequest(ByVal requestID As Long)
    If wsClient.State <> sckClosed Then wsClient.Close
    wsClient.Accept requestID
    Call AddLog0("Server Connected")
    Command_Format = 0
End Sub

Private Sub SendData(Data As String)
    If wsClient.State = sckConnected Then
        wsClient.SendData Data
    End If
End Sub

'Routine to process LX200 Protocol commands

Private Sub wsClient_DataArrival(ByVal bytesTotal As Long)
    Dim Data As String
    Dim arr1() As String
    Dim arr2() As String
    Dim i As Integer
    Dim dblPrev As Double
    Dim intStatus As Integer
    Dim duration As Integer
    
    wsClient.GetData Data
    
    If Mid(Data, 1, 1) = Chr(6) Then
         Call SendData("P")
    End If
    
    
    
    arr1 = Split(Data, L_PIPE)
    For i = 0 To UBound(arr1) - 1
    
        Select Case arr1(i)
 
            Case ":GA", ":GD", ":Gd"
            
              If Command_Format = 0 Then
                Call SendData("+00" & Chr(223) & "00#")
              Else
                Call SendData("+00" & Chr(223) & "00:00#")
              End If
              
            Case ":Gg"
            
              Call SendData("000" & Chr(223) & "00#")
            
            Case ":GM", ":GN", ":GO", ":GP"
            
              Call SendData("???#")
              
            Case ":GR", ":Gr"
            
              If Command_Format = 0 Then
                Call SendData("00:00.0#")
              Else
                Call SendData("00:00:00#")
              End If
              
            Case ":GS"
              Call SendData("00:00:00#")
              
            Case ":Gt"
            
              Call SendData("+00" & Chr(223) & "00#")
              
            Case ":GT"
            
              Call SendData("00.0#")
              
            Case ":GVF"
              Call SendData("LX200 Emulation 1.0#")
              
            Case ":GVP"
            
            
              If ScopeType.value = 1 Then
              
                 Call SendData("LX2001#")
                 
              Else
                
                 Call SendData("LX200#")
                  
              End If
              
            Case ":GVN"
            
               If ScopeType.value = 1 Then
            
                 Call SendData("4.01#")

               Else
               
                 Call SendData("3.01#")
                 
               End If

            Case ":GZ"
            
              If Command_Format = 0 Then
                Call SendData("000" & Chr(223) & "00#")

              Else
                Call SendData("000" & Chr(223) & "00:00#")

              End If
              
            Case ":F+", ":F-", ":FQ", ":FF", ":FS"
              
            Case ":MA", ":MS"
            
              Call SendData("0")
              

            Case ":Mw"
               If Check_WE_Enable = 1 Then
                
                If CheckW_E.value = 1 Then
                
                    Call SetPort(4)
                    DirLabel2.BackColor = &HFF00&
                
                Else
                
                    Call SetPort(5)
                    DirLabel2.BackColor = &HFF00&

                End If
              
              End If

            Case ":Me"
            
               If Check_WE_Enable = 1 Then
            
            
                If CheckW_E.value = 1 Then
                    
                    Call SetPort(5)
                    DirLabel3.BackColor = &HFF00&
                
                Else
                    Call SetPort(4)
                    DirLabel3.BackColor = &HFF00&
                End If

               End If

            Case ":Mn"
            
               If Check_NS_Enable = 1 Then
           
                If CheckN_S.value = 1 Then
                    
                    Call SetPort(7)
                    DirLabel1.BackColor = &HFF00&
                
                Else
                    
                    Call SetPort(6)
                    DirLabel1.BackColor = &HFF00&
                
                End If
                
              End If

            Case ":Ms"
            
              If Check_NS_Enable = 1 Then

            
                If CheckN_S.value = 1 Then
                
                    Call SetPort(6)
                    DirLabel4.BackColor = &HFF00&
                
                Else
                
                    Call SetPort(7)
                    DirLabel4.BackColor = &HFF00&
                
                End If

              End If

            Case ":Qw"
            
               If Check_WE_Enable = 1 Then

                If CheckW_E.value = 1 Then
            
                    Call ResetPort(4)
                    DirLabel2.BackColor = &HFFFF00
                
                Else
                
                    Call ResetPort(5)
                    DirLabel2.BackColor = &HFFFF00
 
                End If
 
              End If
            
            Case ":Qe"
            
            
              If Check_WE_Enable = 1 Then

            
                If CheckW_E.value = 1 Then
                
                    Call ResetPort(5)
                    DirLabel3.BackColor = &HFFFF00
            
                Else
                
                    Call ResetPort(4)
                    DirLabel3.BackColor = &HFFFF00

                End If

             End If

            Case ":Qn"
            
               If Check_NS_Enable = 1 Then
          
                If CheckN_S.value = 1 Then
                
                    Call ResetPort(7)
                    DirLabel1.BackColor = &HFFFF00
           
                Else
                
                    Call ResetPort(6)
                    DirLabel1.BackColor = &HFFFF00

                End If
            
            
              End If
            
            Case ":Qs"
            
         
               If Check_NS_Enable = 1 Then
            
                If CheckN_S.value = 1 Then
                
                    Call ResetPort(6)
                    DirLabel4.BackColor = &HFFFF00
                
                Else
                
                    Call ResetPort(7)
                    DirLabel4.BackColor = &HFFFF00

                End If

              End If

            Case ":Q"
            
                Call ResetPort(5)
                Call ResetPort(4)
                Call ResetPort(6)
                Call ResetPort(7)
                DirLabel1.BackColor = &HFFFF00
                DirLabel2.BackColor = &HFFFF00
                DirLabel3.BackColor = &HFFFF00
                DirLabel4.BackColor = &HFFFF00

               
            Case ":RG"
            
                
            Case ":R1"
 
                               
              
            Case ":U"
              If Command_Format = 0 Then
                 Command_Format = 1
              Else
                 Command_Format = 0
              End If
              

            Case Else
            
              If Mid(arr1(i), 2, 1) = "S" Then
              
                 Call AddLog("SND: 1")
                 
              ElseIf Mid(arr1(i), 2, 1) = "L" Then
              
                 Call AddLog("SND: 0")

              Else
            
                
               If Mid(arr1(i), 1, 1) = Chr(6) Then
                 Call AddLog("SND: P")
               Else
               
                ' Attempt to detect Pulseguide commands here
               
                    Call AddLog("EQMOD Command: " & arr1(i))
                    If Mid(arr1(i), 2, 3) = "Mge" Then
                        gEQRAPulseDuration = Val(Mid(arr1(i), 5, 4))
                    
                        If Check_WE_Enable = 1 Then
            
                            If CheckW_E.value = 1 Then
                    
                                Call SetPort(5)
                                DirLabel3.BackColor = &HFF00&
                                Call AddLog("EQMOD: Duration: " & Str(gEQRAPulseDuration))
                             Else
                             
                                Call SetPort(4)
                                Call AddLog("EQMOD: Duration: " & Str(gEQRAPulseDuration))
                                DirLabel3.BackColor = &HFF00&
                                
                            End If

                        PulseGuide_timer.Enabled = True
                        End If
                        
                    End If

                    If Mid(arr1(i), 2, 3) = "Mgw" Then
                        gEQRAPulseDuration = Val(Mid(arr1(i), 5, 4))
                        
                        
                        If Check_WE_Enable = 1 Then
                
                            If CheckW_E.value = 1 Then
                
                                Call SetPort(4)
                                Call AddLog("EQMOD: Duration: " & Str(gEQRAPulseDuration))
                                DirLabel2.BackColor = &HFF00&
                
                            Else
                
                                Call SetPort(5)
                                Call AddLog("EQMOD: Duration: " & Str(gEQRAPulseDuration))
                                DirLabel2.BackColor = &HFF00&

                            End If
              
              
                        PulseGuide_timer.Enabled = True
                        End If
                    
                    End If

                    If Mid(arr1(i), 2, 3) = "Mgn" Then
                        gEQDECPulseDuration = Val(Mid(arr1(i), 5, 4))
                        
                        
                        If Check_NS_Enable = 1 Then
           
                            If CheckN_S.value = 1 Then
                    
                                Call SetPort(7)
                                Call AddLog("EQMOD: Duration: " & Str(gEQDECPulseDuration))
                                DirLabel1.BackColor = &HFF00&
                
                            Else
                    
                                Call SetPort(6)
                                Call AddLog("EQMOD: Duration: " & Str(gEQDECPulseDuration))
                                DirLabel1.BackColor = &HFF00&
                
                            End If
                            
                        PulseGuide_timer.Enabled = True
                        End If
                        
                    End If

                    If Mid(arr1(i), 2, 3) = "Mgs" Then
                        gEQDECPulseDuration = Val(Mid(arr1(i), 5, 4))
                        If Check_NS_Enable = 1 Then
            
                            If CheckN_S.value = 1 Then
                
                                Call SetPort(6)
                                Call AddLog("EQMOD: Duration: " & Str(gEQDECPulseDuration))
                                DirLabel4.BackColor = &HFF00&
                
                            Else
                
                                Call SetPort(7)
                                Call AddLog("EQMOD: Duration: " & Str(gEQDECPulseDuration))
                                DirLabel4.BackColor = &HFF00&
                
                            End If
                            
                        PulseGuide_timer.Enabled = True
                        End If
                    
                    End If
               
               End If
                                
              End If
                        
        End Select
    Next
End Sub


Private Sub Pulseguide(dir As Double, duration As Double)




End Sub


Private Sub wsClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    wsClient.Close
    wsClient.Listen
    Call AddLog("Connection Error")
End Sub


Private Sub AddLog0(dtaLog As String)
        List1.Text = Right(List1.Text & "[" & Time & "] " & dtaLog & vbCrLf, 20000)
        List1.SelStart = Len(List1.Text)
End Sub
Private Sub AddLog(dtaLog As String)
    If CheckLOG1.value = 1 Then
        List1.Text = Right(List1.Text & "[" & Time & "] " & dtaLog & vbCrLf, 20000)
        List1.SelStart = Len(List1.Text)
    End If
End Sub


Public Sub SetPort(OutNum As Integer)

   'Routine to send guide commands to the EQContrl Driver
   'lxlastcommand variable are used to block duplicate commands.
   
   
    Select Case OutNum
    
        Case 7
               If lxlastcommand = 71 Then GoTo END03
               eqres = EQ_MotorStop(1)
               If eqres = 1 Then GoTo ERROR
            
               Do
        
                  eqres = EQ_GetMotorStatus(1)
                  If eqres = 1 Then GoTo ERROR
        
               Loop While (eqres And &H10) <> 0

            
               If EQ_SendGuideRate(1, 0, Val(VScroll3.value), 0, 0, 0) = 0 Then
                 Call AddLog("EQMOD: Dec- ON")
               Else
                 Call AddLog("EQMOD: Dec- ON: Mount not connected")
               End If
        
               lxlastcommand = 71
               
        Case 6
                If lxlastcommand = 61 Then GoTo END03
                eqres = EQ_MotorStop(1)
                If eqres = 1 Then GoTo ERROR
            
               Do
  
                  eqres = EQ_GetMotorStatus(1)
                  If eqres = 1 Then GoTo ERROR
  
               Loop While (eqres And &H10) <> 0
            
              
               If EQ_SendGuideRate(1, 0, Val(VScroll3.value), 1, 0, 0) = 0 Then
                 Call AddLog("EQMOD: Dec+ ON")
               Else
                 Call AddLog("EQMOD: Dec+ ON: Mount not connected")
               End If
        
               lxlastcommand = 61
               
        Case 5
               If lxlastcommand = 51 Then GoTo END03
               
               If EQ_SendGuideRate(0, 0, Val(VScroll2.value), 0, 0, 0) = 0 Then
                 Call AddLog("EQMOD: RA- ON")
               Else
                 Call AddLog("EQMOD: RA- ON: Mount not connected")
               End If
        
               lxlastcommand = 51
               
        Case 4
               If lxlastcommand = 41 Then GoTo END03
               If EQ_SendGuideRate(0, 0, Val(VScroll2.value), 1, 0, 0) = 0 Then
                 Call AddLog("EQMOD: RA+ ON")
               Else
                 Call AddLog("EQMOD: RA+ ON: Mount not connected")
               End If
        
               lxlastcommand = 41
        
        Case 3, 2, 1, 0
        
              lxlastcommand = 0
        
        Case Else
        
              lxlastcommand = 0
        
    End Select

END03:

    Exit Sub
    
    
ERROR:

    AddLog ("Mount not connected")
    lxlastcommand = 0
End Sub

Public Sub ResetPort(OutNum As Integer)

  
    Select Case OutNum
    
        Case 7
               
               If EQ_MotorStop(1) = 0 Then
                 Call AddLog("EQMOD: Dec- OFF")
               Else
                 Call AddLog("EQMOD: Dec- OFF: Mount not connected")
               End If
    
               lxlastcommand = 70
        
        Case 6
               
               If EQ_MotorStop(1) = 0 Then
                 Call AddLog("EQMOD: Dec+ OFF")
               Else
                 Call AddLog("EQMOD: DEC+ OFF: Mount not connected")
               End If
               
               lxlastcommand = 60
        
        Case 5
               
               If EQ_SendGuideRate(0, 0, 0, 0, 0, 0) = 0 Then
                 Call AddLog("EQMOD: RA- OFF")
               Else
                 Call AddLog("EQMOD: RA- OFF: Mount not connected")
               End If
               
               lxlastcommand = 50
        
        Case 4
        
                
               If EQ_SendGuideRate(0, 0, 0, 0, 0, 0) = 0 Then
                 Call AddLog("EQMOD: RA+ OFF")
               Else
                 Call AddLog("EQMOD: RA+ OFF: Mount not connected")
               End If
        
               lxlastcommand = 40
        
        Case 3, 2, 1, 0

               lxlastcommand = 0

        Case Else
        
               lxlastcommand = 0
        
    End Select



End Sub

