VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form GPSSetup 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "GPS SETUP"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "GPS Data"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   4335
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "NMEA Trace"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   1935
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CheckBox Check1 
            BackColor       =   &H00000000&
            Caption         =   "RAW COMMS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   255
            Left            =   0
            TabIndex        =   41
            Top             =   1680
            Width           =   1815
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H0095C1CB&
            Caption         =   "Clear"
            Height          =   255
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H0095C1CB&
            Caption         =   "Pause"
            Height          =   255
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00000080&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   1530
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   38
            Top             =   120
            Width           =   4095
         End
      End
      Begin VB.TextBox TextLongSec 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   3480
         TabIndex        =   36
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox TextLatSec 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   3480
         TabIndex        =   35
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TextLongMin 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   420
         Left            =   2820
         TabIndex        =   31
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox TextLatMin 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   420
         Left            =   2820
         TabIndex        =   30
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cbWE 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         ItemData        =   "GPSSetup.frx":0000
         Left            =   1440
         List            =   "GPSSetup.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   600
         Width           =   555
      End
      Begin VB.ComboBox cbNS 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         ItemData        =   "GPSSetup.frx":0014
         Left            =   1440
         List            =   "GPSSetup.frx":001E
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   240
         Width           =   555
      End
      Begin VB.TextBox TextDat 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox TextUTC 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox TextElevation 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Text            =   "0"
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox TextLong 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   2160
         TabIndex        =   6
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox TextLat 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "GPS DATE:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   13
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "GPS UTC TIME:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   12
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "GPS ELEVATION:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "GPS LONGITUDE:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Caption         =   "GPS LATITUDE:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton CommandAbort 
      BackColor       =   &H0095C1CB&
      Caption         =   "ABORT"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7440
      Width           =   4335
   End
   Begin VB.Timer DispTimer 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   4080
      Top             =   120
   End
   Begin VB.CommandButton CommandReset 
      BackColor       =   &H0095C1CB&
      Caption         =   "RESET"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   6960
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "GPS Hemisphere"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   855
      Left            =   120
      TabIndex        =   26
      Top             =   4080
      Width           =   4335
      Begin VB.ComboBox cbHemisphere 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         ItemData        =   "GPSSetup.frx":0028
         Left            =   1440
         List            =   "GPSSetup.frx":0032
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   360
         Width           =   2715
      End
      Begin VB.Label Label14 
         BackColor       =   &H00000000&
         Caption         =   "Hemisphere:"
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
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "PC and GPS TIME Comparison"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1935
      Left            =   120
      TabIndex        =   17
      Top             =   4920
      Width           =   4335
      Begin VB.TextBox TextAdjust 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   1440
         TabIndex        =   24
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox Textdelta 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   1440
         TabIndex        =   22
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox Textgpslst 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   1440
         TabIndex        =   20
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox Textpclst 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   1440
         TabIndex        =   19
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         Caption         =   "ADJUSTED LST:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   25
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "TIME DELTA:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   23
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         Caption         =   "GPS LOCAL ST:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   21
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "PC LOCAL ST:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   18
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.ComboBox ComboBaud 
      BackColor       =   &H00000080&
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
      Height          =   330
      ItemData        =   "GPSSetup.frx":0044
      Left            =   120
      List            =   "GPSSetup.frx":0054
      TabIndex        =   16
      Text            =   "baud"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton CommandAccept 
      BackColor       =   &H0095C1CB&
      Caption         =   "ACCEPT"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6960
      Width           =   2175
   End
   Begin VB.CommandButton CommandRetrieve 
      BackColor       =   &H0095C1CB&
      Caption         =   "Retrieve Coordinate and Time Data"
      Height          =   735
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   2895
   End
   Begin VB.ComboBox ComboPort 
      BackColor       =   &H00000080&
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
      Height          =   330
      ItemData        =   "GPSSetup.frx":008A
      Left            =   120
      List            =   "GPSSetup.frx":00BE
      TabIndex        =   2
      Text            =   "Port"
      Top             =   1080
      Width           =   1335
   End
   Begin MSCommLib.MSComm MSComm1_GPS 
      Left            =   3840
      Top             =   8280
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
      InBufferSize    =   2048
      RThreshold      =   1
      BaudRate        =   4800
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   120
      Picture         =   "GPSSetup.frx":0129
      Top             =   75
      Width           =   315
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000040&
      Caption         =   "Connect your GPS and select the COM port settings then click the Retrieve Button"
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
      TabIndex        =   33
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "EQMOD GPS SETUP"
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
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "GPSSetup"
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
' GPSSetup.frm - GPS Setup form
'
'
' Written:  07-Oct-06   Raymund Sarmiento
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' 04-Nov-06 rcs     Initial edit for EQ Mount Driver Function Prototype
' 01-Dec-06 rcs     Add GPS Function
' 01-Dec-06 rcs     Remove "Deactivate Window event" which causes com port to close
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

'Private Const gpsID As String = "EQMOD.Telescope"
'Private Const gpsDESC As String = "EQMOD ASCOM Scope Driver"

Dim WithEvents gps As eqmodnmea
Attribute gps.VB_VarHelpID = -1
Dim inp As String

Private TraceEnable As Boolean

Dim rDelta As Double



Private Sub Command1_Click()
    Text1.Text = ""
End Sub

Private Sub Addlog(dtaLog As String)
    If Frame4.Visible Then
        If TraceEnable Then
            Text1.Text = Right(Text1.Text & vbCrLf & Time$ & " " & dtaLog, 2000)
            Text1.SelStart = Len(Text1.Text)
        End If
    End If
End Sub


Private Sub Command2_Click()
    If TraceEnable Then
        TraceEnable = False
        Command2.Caption = "Go"
    Else
        TraceEnable = True
        Command2.Caption = "Pause"
    End If
End Sub

Private Sub CommandRetrieve_Click()
Dim initstr As String
    
    On Error GoTo handleErr
    Text1.Text = ""
    TraceEnable = True
    If CommandRetrieve.Caption = oLangDll.GetLangString(1305) Then
        CommandRetrieve.Caption = oLangDll.GetLangString(1322)
        inp = ""
        MSComm1_GPS.CommPort = ComboPort.ListIndex + 1
        MSComm1_GPS.Settings = ComboBaud.Text
        MSComm1_GPS.PortOpen = True
        
        If MSComm1_GPS.PortOpen = True Then
            ComboPort.Enabled = False
            ComboBaud.Enabled = False
            initstr = HC.oPersist.ReadIniValue("GPS_INITSTRING1")
            If initstr <> "" Then
                ' an initialiastion string exists - write it to the GPS!
                MSComm1_GPS.Output = initstr
            End If
        Else
            CommandRetrieve.Caption = oLangDll.GetLangString(1305)
            ComboPort.Enabled = True
            ComboBaud.Enabled = True
        End If
    Else
        CommandRetrieve.Caption = oLangDll.GetLangString(1305)
        ComboPort.Enabled = True
        ComboBaud.Enabled = True
        If MSComm1_GPS.PortOpen = True Then MSComm1_GPS.PortOpen = False
    End If
handleErr:
    If Err Then
        MsgBox ("Error " & CStr(Err.Number) & ": " & Err.Description)
    End If
End Sub

Private Sub CommandAccept_Click()
    HC.oPersist.WriteIniValue "GPSPort", CStr(GPSSetup.ComboPort.ListIndex)
    HC.oPersist.WriteIniValue "GPSBaud", CStr(GPSSetup.ComboBaud.ListIndex)
    HC.oPersist.WriteIniValue "LatitudeSec", CStr(GPSSetup.TextLatSec.Text)
    HC.oPersist.WriteIniValue "LatitudeMin", CStr(GPSSetup.TextLatMin.Text)
    HC.oPersist.WriteIniValue "LatitudeDeg", CStr(GPSSetup.TextLat.Text)
    HC.oPersist.WriteIniValue "LatitudeNS", CStr(GPSSetup.cbNS.ListIndex)
    HC.oPersist.WriteIniValue "LongitudeDeg", CStr(GPSSetup.TextLong.Text)
    HC.oPersist.WriteIniValue "LongitudeMin", CStr(GPSSetup.TextLongMin.Text)
    HC.oPersist.WriteIniValue "LongitudeSec", CStr(GPSSetup.TextLongSec.Text)
    HC.oPersist.WriteIniValue "LongitudeEW", CStr(GPSSetup.cbWE.ListIndex)
    HC.oPersist.WriteIniValue "HemisphereNS", CStr(GPSSetup.cbHemisphere.ListIndex)
    HC.oPersist.WriteIniValue "Elevation", CStr(GPSSetup.TextElevation.Text)
    HC.oPersist.WriteIniValue "TimeDelta", CStr(rDelta)
    HC.txtLatSec.Text = GPSSetup.TextLatSec.Text
    HC.txtLatMin.Text = GPSSetup.TextLatMin.Text
    HC.txtLatDeg.Text = GPSSetup.TextLat.Text
    HC.cbNS.ListIndex = GPSSetup.cbNS.ListIndex
    HC.txtLongDeg.Text = GPSSetup.TextLong.Text
    HC.txtLongMin.Text = GPSSetup.TextLongMin.Text
    HC.txtLongSec.Text = GPSSetup.TextLongSec.Text
    HC.cbEW.ListIndex = GPSSetup.cbWE.ListIndex
    HC.cbhem.ListIndex = GPSSetup.cbHemisphere.ListIndex
    HC.txtElevation.Text = GPSSetup.TextElevation.Text

    gLongitude = CDbl(EQFixNum(GPSSetup.TextLong)) + (CDbl(EQFixNum(GPSSetup.TextLongMin)) / 60#) + (CDbl(EQFixNum(GPSSetup.TextLongSec)) / 3600#)
    If GPSSetup.cbWE.Text = oLangDll.GetLangString(115) Then gLongitude = -gLongitude  ' W is neg
    
    gLatitude = CDbl(EQFixNum(GPSSetup.TextLat)) + (CDbl(EQFixNum(GPSSetup.TextLatMin)) / 60#) + (CDbl(EQFixNum(GPSSetup.TextLatSec)) / 3600#)
    If GPSSetup.cbNS.Text = oLangDll.GetLangString(116) Then gLatitude = -gLatitude
       
    gElevation = CDbl(EQFixNum(GPSSetup.TextElevation))
    
    If GPSSetup.cbHemisphere.Text = oLangDll.GetLangString(1110) Then
        gHemisphere = 0
    Else
        gHemisphere = 1
    End If

    gEQTimeDelta = rDelta
    
    Unload GPSSetup

End Sub

Private Sub CommandReset_Click()
    gEQTimeDelta = 0
    rDelta = 0
    HC.oPersist.WriteIniValue "TimeDelta", CStr(rDelta)
    Unload GPSSetup
End Sub

Private Sub CommandAbort_Click()
    Unload GPSSetup
End Sub

Private Sub DispTimer_Timer()
    TextAdjust.Text = FmtSexa(EQnow_lst_time(gLongitude * DEG_RAD, rDelta + CDbl(Now())), False)
End Sub


Private Sub Form_Load()

Dim tmptxt As String
    On Error Resume Next
 
    Call SetText
    
    If HC.HCOnTop.Value = 1 Then Call PutWindowOnTop(GPSSetup)
 
    Set gps = New eqmodnmea
    
    GPSSetup.TextElevation = "0"
    
    tmptxt = HC.oPersist.ReadIniValue("GPSPort")
    If tmptxt = "" Then
        ComboPort.ListIndex = 1
    Else
        ComboPort.ListIndex = val(tmptxt)
    End If
   
    tmptxt = HC.oPersist.ReadIniValue("GPSBaud")
    If tmptxt = "" Then
        ComboBaud.ListIndex = 1
    Else
        ComboBaud.ListIndex = val(tmptxt)
    End If
    
    DispTimer.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rDelta = 0
    If MSComm1_GPS.PortOpen = True Then
        MSComm1_GPS.PortOpen = False
    End If
    DispTimer.Enabled = False
End Sub



Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        Frame4.Visible = True
    Else
        Frame4.Visible = False
    End If
    
End Sub

Private Sub MSComm1_GPS_OnComm()
Dim buf As String
Dim tlines() As String
Dim i As Integer

    On Error Resume Next

     Select Case MSComm1_GPS.CommEvent
        Case comEvReceive
            buf = MSComm1_GPS.Input
            If Err = 0 Then
                If Check1.Value = 1 Then
                    Call Addlog(buf)
                End If
                If Left$(buf, 1) = "$" Then
                    inp = buf
                Else
                    inp = inp + buf
                End If
                tlines = Split(inp, vbCrLf)
                For i = 0 To UBound(tlines) - 1
                    tlines(i) = Trim(tlines(i))
                    If tlines(i) <> "" Then
                        If Check1.Value = 0 Then
                            Call Addlog(tlines(i))
                        End If
                        gps.scan (tlines(i))
                    End If
                Next i
            End If
     End Select
End Sub
Private Sub gps_EQgpsaltitude(ByVal altitude As String)
    TextElevation.Text = EQFixNum(altitude)
End Sub

Private Sub gps_EQgpsdate(ByVal satDate As String)
    TextDat.Text = satDate
End Sub
Private Sub gps_EQgpsposition(ByVal gnmlatitude As String, ByVal gnmlongitude As String, ByVal lathm As String, ByVal lath As String, ByVal latm As String, ByVal lonhm As String, ByVal lonh As String, ByVal lonm As String)

Dim mins As Double
Dim secs As Double

    ' latitude
    TextLat.Text = EQFixNum(lath)
    TextLatMin.Text = EQFixNum(latm)
    mins = CDbl(TextLatMin.Text)
    secs = 60 * (mins - Int(mins))
    TextLatMin.Text = CStr(Int(mins))
    TextLatSec.Text = CStr(secs)
    
    ' longitude
    TextLong.Text = EQFixNum(lonh)
    TextLongMin.Text = EQFixNum(lonm)
    mins = CDbl(TextLongMin.Text)
    secs = 60 * (mins - Int(mins))
    TextLongMin.Text = CStr(Int(mins))
    TextLongSec.Text = CStr(secs)
    
    If lathm = "N" Or lathm = "n" Then
      GPSSetup.cbNS.ListIndex = 0
      GPSSetup.cbHemisphere.ListIndex = 0
    Else
      GPSSetup.cbNS.ListIndex = 1
      GPSSetup.cbHemisphere.ListIndex = 1
    End If
              
    If lonhm = "W" Or lathm = "w" Then
      GPSSetup.cbWE.ListIndex = 1
    Else
      GPSSetup.cbWE.ListIndex = 0
    End If
      
End Sub
Private Sub gps_EQgpstime(ByVal Time As String)
    TextUTC.Text = Time
End Sub

' This event is trigered everytime the date and time from the GPS
' is retrieved.

Private Sub gps_EQgpsnow(ByVal Satnow As String, ByVal hh As String, ByVal mm As String, ByVal ss As String, ByVal mn As String, ByVal dd As String, ByVal yy As String)

Dim tgpsloc As Double
Dim tpcloc As Double
Dim tgpslst As Double
Dim tpclst As Double

    ' Get the GPS DATE and TIME and Convert them to local Time
    tgpsloc = CDbl(CDate(DateSerial(EQFixNum(yy), EQFixNum(mn), EQFixNum(dd))) + CDate(TimeSerial(EQFixNum(hh), EQFixNum(mm), EQFixNum(ss))) - (CDbl(utc_offs()) / 86400))

    ' Get the Local PC Time
    tpcloc = CDbl(Now)
    
    ' Get the Delta Value
    rDelta = tgpsloc - tpcloc
    
    ' Compute for the local sidreal Time
    tgpslst = EQnow_lst_time(gLongitude * DEG_RAD, tgpsloc)
    tpclst = EQnow_lst_time(gLongitude * DEG_RAD, tpcloc)

    Textpclst.Text = FmtSexa(tpclst, False)
    Textgpslst.Text = FmtSexa(tgpslst, False)
    Textdelta.Text = FmtSexa(tgpslst - tpclst, True)

End Sub

Private Sub SetText()
    Dim tmptxt As String
    
    GPSSetup.Caption = oLangDll.GetLangString(1300)
    Label6.Caption = oLangDll.GetLangString(1301)
    Label11.Caption = oLangDll.GetLangString(1302)
    ComboPort.Text = oLangDll.GetLangString(1303)
    ComboBaud.Text = oLangDll.GetLangString(1304)
    CommandRetrieve.Caption = oLangDll.GetLangString(1305)
    Frame1.Caption = oLangDll.GetLangString(1306)
    Label9.Caption = oLangDll.GetLangString(1307)
    Label1.Caption = oLangDll.GetLangString(1308)
    Label2.Caption = oLangDll.GetLangString(1309)
    Label3.Caption = oLangDll.GetLangString(1310)
    Label4.Caption = oLangDll.GetLangString(1311)
    Frame3.Caption = oLangDll.GetLangString(1312)
    Label14.Caption = oLangDll.GetLangString(1313)
    Frame2.Caption = oLangDll.GetLangString(1314)
    Label5.Caption = oLangDll.GetLangString(1315)
    Label7.Caption = oLangDll.GetLangString(1316)
    Label8.Caption = oLangDll.GetLangString(1317)
    Label10.Caption = oLangDll.GetLangString(1318)
    CommandAccept.Caption = oLangDll.GetLangString(1319)
    CommandReset.Caption = oLangDll.GetLangString(1320)
    CommandAbort.Caption = oLangDll.GetLangString(1321)
    
    ComboBaud.Clear
    ComboBaud.AddItem "4800,n,8,1"
    ComboBaud.AddItem "9600,n,8,1"
    ComboBaud.AddItem "19200,n,8,1"
    tmptxt = HC.oPersist.ReadIniValue("GPS_CUSTOM_BAUD")
    If tmptxt <> "" Then
        ComboBaud.AddItem tmptxt
    Else
        HC.oPersist.WriteIniValue "GPS_CUSTOM_BAUD", "38400,n,8,1"
        ComboBaud.AddItem "38400,n,8,1"
    End If


End Sub
