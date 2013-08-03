VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form GPStest 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "EQMOD GPS"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   7800
      Top             =   1560
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   11
      Top             =   120
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00000080&
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
      Height          =   345
      ItemData        =   "GPStest.frx":0000
      Left            =   5520
      List            =   "GPStest.frx":0034
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox TextAlt 
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
      Left            =   4800
      TabIndex        =   7
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox TextLon 
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
      Left            =   3120
      TabIndex        =   4
      Top             =   120
      Width           =   2175
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
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   7800
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
      RThreshold      =   1
      BaudRate        =   4800
   End
   Begin VB.Label Label22 
      BackColor       =   &H00000080&
      Caption         =   "time"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   855
      Left            =   1440
      TabIndex        =   33
      Top             =   6480
      Width           =   6255
   End
   Begin VB.Label Label23 
      BackColor       =   &H00000000&
      Caption         =   "PC TIME + DELTA (Should be equal to GPS TIME)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   1440
      TabIndex        =   34
      Top             =   6240
      Width           =   4095
   End
   Begin VB.Label Label21 
      BackColor       =   &H00000000&
      Caption         =   "LST:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   240
      TabIndex        =   32
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label20 
      BackColor       =   &H00000080&
      Caption         =   "time"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   855
      Left            =   1440
      TabIndex        =   31
      Top             =   5040
      Width           =   6255
   End
   Begin VB.Label Label19 
      BackColor       =   &H00000000&
      Caption         =   "DELTA:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   240
      TabIndex        =   30
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackColor       =   &H00000080&
      Caption         =   "time"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   855
      Left            =   1440
      TabIndex        =   26
      Top             =   3120
      Width           =   6255
   End
   Begin VB.Label Label17 
      BackColor       =   &H00000080&
      Caption         =   "time"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   855
      Left            =   1440
      TabIndex        =   28
      Top             =   4080
      Width           =   6255
   End
   Begin VB.Label Label18 
      BackColor       =   &H00000000&
      Caption         =   "GPS LST:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   240
      TabIndex        =   29
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label16 
      BackColor       =   &H00000000&
      Caption         =   "PC LST:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackColor       =   &H00000080&
      Caption         =   "time"
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
      Left            =   5280
      TabIndex        =   25
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000000&
      Caption         =   "GPS LOCAL time:"
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
      Height          =   375
      Left            =   3840
      TabIndex        =   24
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Labelaa 
      BackColor       =   &H00000080&
      Caption         =   "time"
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
      Left            =   5280
      TabIndex        =   23
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      Caption         =   "GPS UTC time:"
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
      Height          =   375
      Left            =   3840
      TabIndex        =   22
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "GPS EQ LST time:"
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
      Height          =   375
      Left            =   3840
      TabIndex        =   21
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000080&
      Caption         =   "time"
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
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "PC Local Time:"
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
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "PC EQ  lst time:"
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
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000080&
      Caption         =   "time"
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
      Left            =   5280
      TabIndex        =   17
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label lblpctime2 
      BackColor       =   &H00000080&
      Caption         =   "time"
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
      TabIndex        =   16
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "PC LST time:"
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
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblpctime 
      BackColor       =   &H00000080&
      Caption         =   "time"
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
      TabIndex        =   14
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label LabelTime 
      BackColor       =   &H00000080&
      Caption         =   "time"
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
      Left            =   2280
      TabIndex        =   6
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "time:"
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
      Height          =   375
      Left            =   1920
      TabIndex        =   13
      Top             =   600
      Width           =   615
   End
   Begin VB.Label LabelDate 
      BackColor       =   &H00000080&
      Caption         =   "date"
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
      Left            =   480
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Date:"
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
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   615
   End
   Begin VB.Label LabelAltunit 
      BackColor       =   &H00000000&
      Caption         =   "Label3"
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
      Height          =   375
      Left            =   5520
      TabIndex        =   9
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "Altitude:"
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
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Lon:"
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
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Lat:"
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
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "GPStest"
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
' 10-Nov-06 rcs     Initial edit for EQ Mount Driver GPS CLASS
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


Option Explicit



Dim WithEvents gps As eqmodnmea
Attribute gps.VB_VarHelpID = -1

Dim inString As String

Private Sub Command1_Click()
If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
Unload GPStest
End Sub

Private Sub Command2_Click()
    If Command2.Caption = "Connect" Then
        Command2.Caption = "Connecting..."
        MSComm1.CommPort = Combo1.ListIndex + 1
        MSComm1.PortOpen = True
        If MSComm1.PortOpen = True Then
            Command2.Caption = "Disconnect"
            Combo1.Enabled = False
        Else
            Command2.Caption = "Connect"
            Combo1.Enabled = True
        End If
    Else
        Command2.Caption = "Connect"
        Combo1.Enabled = True
        MSComm1.PortOpen = False
    End If
End Sub

Private Sub Form_Load()
Combo1.ListIndex = 1
Set gps = New eqmodnmea
End Sub

Private Sub gps_EQgpsaltitude(ByVal altitude As Double)
TextAlt.Text = CStr(altitude)
End Sub

Private Sub gps_EQgpsunit(ByVal altitudeUnits As String)
LabelAltunit.Caption = altitudeUnits
End Sub


Private Sub gps_EQgpsdate(ByVal satDate As String)
LabelDate.Caption = satDate
End Sub


Private Sub gps_EQgpsposition(ByVal latitude As String, ByVal longitude As String)

      TextLat.Text = latitude
      TextLon.Text = longitude

End Sub


Private Sub gps_EQgpstime(ByVal Time As String)
LabelTime.Caption = Time

End Sub


Private Sub gps_EQgpsnow(ByVal Satnow As String)
 
    
    
Dim h As Double
Dim i As Double
Dim j As Double
Dim k As Double
Dim o As Double
Dim p As Double

Dim x As Double
Dim Y As Double

Dim s As String


Labelaa.Caption = Satnow


x = CDbl(CDate(Satnow)) - (CDbl(utc_offs()) / 86400)

Label14.Caption = CDate(x)

i = vb_mjd(x)
Call utc_gst(mjd_day(i), mjd_hr(i), j)
j = j + radhr(120.044 * DEG_RAD)
Call obliq(i, k)
Call nut(i, o, p)
j = j + radhr(p * Cos(k + o))
Call range(j, 24#)


Label5.Caption = FmtSexa(j, False)
Label17.Caption = FmtSexa(j, False)


Y = CDbl(Now())

i = vb_mjd(Y)
Call utc_gst(mjd_day(i), mjd_hr(i), h)
h = h + radhr(120.044 * DEG_RAD)
Call obliq(i, k)
Call nut(i, o, p)
h = h + radhr(p * Cos(k + o))
Call range(h, 24#)

i = x - Y

Label20.Caption = FmtSexa(j - h, True)

i = vb_mjd(i + Y)
Call utc_gst(mjd_day(i), mjd_hr(i), h)
h = h + radhr(120.044 * DEG_RAD)
Call obliq(i, k)
Call nut(i, o, p)
h = h + radhr(p * Cos(k + o))
Call range(h, 24#)


Label22.Caption = FmtSexa(h, False)


'Get delta here



    
    
End Sub





Private Sub MSComm1_OnComm()
Dim InBuff As String
         Select Case MSComm1.CommEvent

            Case comEvReceive ' Received RThreshold # of chars.
               
               InBuff = MSComm1.Input
               Call HandleInput(InBuff)

         End Select
End Sub

Public Sub HandleInput(Sinput As String)
Dim cluster() As String
Dim counter As Integer
If Left$(Sinput, 1) = "$" Then 'start of string
    inString = Sinput
Else
    inString = inString + Sinput
End If
cluster = Split(inString, vbCrLf)
For counter = 0 To UBound(cluster) - 1
        cluster(counter) = Trim(cluster(counter))
        If cluster(counter) <> vbNullString Then gps.scan (cluster(counter))
    Next counter
End Sub

Public Function FmtSexa(ByVal n As Double, ShowPlus As Boolean) As String
    Dim sg As String
    Dim us As String
    Dim ms As String
    Dim ss As String
    Dim u As Integer
    Dim m As Integer
    Dim fmt

    sg = "+"                                ' Assume positive
    If n < 0 Then                           ' Check neg.
        n = -n                              ' Make pos.
        sg = "-"                            ' Remember sign
    End If

    m = Fix(n)                              ' Units (deg or hr)
    us = Format$(m, "00")

    n = (n - m) * 60#
    m = Fix(n)                              ' Minutes
    ms = Format$(m, "00")

    n = (n - m) * 60#
    m = Fix(n)                              ' Minutes
    ss = Format$(m, "00")

    FmtSexa = us & ":" & ms & ":" & ss
    If ShowPlus Or (sg = "-") Then FmtSexa = sg & FmtSexa
    
End Function

Private Sub Timer1_Timer()

Dim i As Double
Dim j As Double
Dim k As Double
Dim o As Double
Dim p As Double

Dim s As String

s = Str(Now())

Label10.Caption = s

i = vb_mjd(CDbl(Now))
Call utc_gst(mjd_day(i), mjd_hr(i), j)
j = j + radhr(120.044 * DEG_RAD)
Call obliq(i, k)
Call nut(i, o, p)
j = j + radhr(p * Cos(k + o))
Call range(j, 24#)

'Label5.Caption = s
lblpctime2.Caption = FmtSexa(j, False)
lblpctime.Caption = FmtSexa(now_lst(120.044 * DEG_RAD), False)
Label15.Caption = FmtSexa(j, False)


End Sub
