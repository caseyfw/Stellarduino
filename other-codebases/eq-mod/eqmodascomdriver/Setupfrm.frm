VERSION 5.00
Begin VB.Form Setupfrm 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   3465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command11 
      Caption         =   "OK"
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
      Left            =   120
      TabIndex        =   18
      Top             =   5520
      Width           =   3255
   End
   Begin VB.PictureBox picASCOM 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   840
      Left            =   120
      MouseIcon       =   "Setupfrm.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "Setupfrm.frx":0152
      ScaleHeight     =   840
      ScaleWidth      =   720
      TabIndex        =   14
      Top             =   120
      Width           =   720
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "EQMOD Port Details"
      ForeColor       =   &H000080FF&
      Height          =   2295
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   3255
      Begin VB.ComboBox lbRetry 
         Height          =   315
         ItemData        =   "Setupfrm.frx":1016
         Left            =   1320
         List            =   "Setupfrm.frx":1020
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1800
         Width           =   945
      End
      Begin VB.ComboBox lbTimeout 
         Height          =   315
         ItemData        =   "Setupfrm.frx":102A
         Left            =   1320
         List            =   "Setupfrm.frx":1034
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1320
         Width           =   945
      End
      Begin VB.ComboBox lbBaud 
         Height          =   315
         ItemData        =   "Setupfrm.frx":1044
         Left            =   1320
         List            =   "Setupfrm.frx":104E
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   840
         Width           =   945
      End
      Begin VB.ComboBox lbPort 
         Height          =   315
         ItemData        =   "Setupfrm.frx":105E
         Left            =   1320
         List            =   "Setupfrm.frx":107A
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "Retry:"
         ForeColor       =   &H000080FF&
         Height          =   225
         Left            =   240
         TabIndex        =   22
         Top             =   1800
         Width           =   690
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Timeout:"
         ForeColor       =   &H000080FF&
         Height          =   225
         Left            =   240
         TabIndex        =   20
         Top             =   1320
         Width           =   690
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Baud:"
         ForeColor       =   &H000080FF&
         Height          =   225
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   690
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Port:"
         ForeColor       =   &H000080FF&
         Height          =   225
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   690
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Site Information"
      ForeColor       =   &H000080FF&
      Height          =   1980
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   3255
      Begin VB.ComboBox cbhem 
         Height          =   315
         ItemData        =   "Setupfrm.frx":10AE
         Left            =   1320
         List            =   "Setupfrm.frx":10B8
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1560
         Width           =   915
      End
      Begin VB.TextBox txtLatDeg 
         Height          =   315
         Left            =   1920
         TabIndex        =   7
         Text            =   "14"
         Top             =   315
         Width           =   480
      End
      Begin VB.TextBox txtLatMin 
         Height          =   315
         Left            =   2505
         TabIndex        =   6
         Text            =   "35"
         Top             =   315
         Width           =   570
      End
      Begin VB.TextBox txtLongDeg 
         Height          =   315
         Left            =   1920
         TabIndex        =   5
         Text            =   "120"
         Top             =   720
         Width           =   480
      End
      Begin VB.TextBox txtLongMin 
         Height          =   315
         Left            =   2505
         TabIndex        =   4
         Text            =   "57"
         Top             =   720
         Width           =   570
      End
      Begin VB.ComboBox cbEW 
         Height          =   315
         ItemData        =   "Setupfrm.frx":10CA
         Left            =   1320
         List            =   "Setupfrm.frx":10D4
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   555
      End
      Begin VB.ComboBox cbNS 
         Height          =   315
         ItemData        =   "Setupfrm.frx":10DE
         Left            =   1320
         List            =   "Setupfrm.frx":10E8
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   315
         Width           =   555
      End
      Begin VB.TextBox txtElevation 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Text            =   "100"
         Top             =   1125
         Width           =   885
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "Hemisphere:"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   1560
         Width           =   990
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         Caption         =   "Longitude:"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   765
         Width           =   765
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "Latitude:"
         ForeColor       =   &H000080FF&
         Height          =   225
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   690
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "Elevation (m):"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1155
         Width           =   990
      End
   End
   Begin VB.Label MainLabel 
      BackColor       =   &H00000080&
      Caption         =   "EQMOD ASCOM SETUP"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   960
      TabIndex        =   15
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Setupfrm"
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
' Setupfrm.frm - ASCOM EQMOD Setup form
'
'
'
' Written:  07-Oct-06   Raymund Sarmiento
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' 04-Nov-06 rcs     Initial edit for EQ Mount Driver Function Prototype
' 14-Nov-06 rcs     Bug Fix on OK button - gHemispher changed to gHemisphere
' 20-Nov-06 rcs     Bug Fix on Elevation value not being saved to the Registry
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

Private s_Profile As DriverHelper.Profile
Private Const sID As String = "EQMOD.Telescope"
Private Const sDESC As String = "EQMOD ASCOM Scope Driver"

Private Sub Command11_Click()

    s_Profile.WriteValue sID, "Port", CStr(Setupfrm.lbPort.Text)
    s_Profile.WriteValue sID, "Baud", CStr(Setupfrm.lbBaud.Text)
    s_Profile.WriteValue sID, "Timeout", CStr(Setupfrm.lbTimeout.Text)
    s_Profile.WriteValue sID, "Retry", CStr(Setupfrm.lbRetry.Text)
    
    s_Profile.WriteValue sID, "LatitudeMin", CStr(Setupfrm.txtLatMin.Text)
    s_Profile.WriteValue sID, "LatitudeDeg", CStr(Setupfrm.txtLatDeg.Text)
    s_Profile.WriteValue sID, "LatitudeNS", CStr(Setupfrm.cbNS.ListIndex)
    s_Profile.WriteValue sID, "LongitudeDeg", CStr(Setupfrm.txtLongDeg.Text)
    s_Profile.WriteValue sID, "LongitudeMin", CStr(Setupfrm.txtLongMin.Text)
    s_Profile.WriteValue sID, "LongitudeEW", CStr(Setupfrm.cbEW.ListIndex)
    s_Profile.WriteValue sID, "HemisphereNS", CStr(Setupfrm.cbhem.ListIndex)
    s_Profile.WriteValue sID, "Elevation", CStr(Setupfrm.txtElevation.Text)
      
          
    HC.txtLatMin.Text = Setupfrm.txtLatMin.Text
    HC.txtLatDeg.Text = Setupfrm.txtLatDeg.Text
    HC.cbNS.ListIndex = Setupfrm.cbNS.ListIndex
    HC.txtLongDeg.Text = Setupfrm.txtLongDeg.Text
    HC.txtLongMin.Text = Setupfrm.txtLongMin.Text
    HC.cbEW.ListIndex = Setupfrm.cbEW.ListIndex
    HC.cbhem.ListIndex = Setupfrm.cbhem.ListIndex
    HC.txtElevation.Text = Setupfrm.txtElevation.Text
   
    gLongitude = CDbl(Setupfrm.txtLongDeg) + (CDbl(Setupfrm.txtLongMin) / 60#)
    If Setupfrm.cbEW.Text = "W" Then gLongitude = -gLongitude  ' W is neg
    
    gLatitude = CDbl(Setupfrm.txtLatDeg) + (CDbl(Setupfrm.txtLatMin) / 60#)
    If Setupfrm.cbNS.Text = "S" Then gLatitude = -gLatitude
    gElevation = CDbl(Setupfrm.txtElevation)
    
    If Setupfrm.cbhem.Text = "North" Then
        gHemisphere = 0
    Else
        gHemisphere = 1
    End If
       
       
    gPort = Setupfrm.lbPort.Text
    gBaud = val(Setupfrm.lbBaud.Text)
    gTimeout = val(Setupfrm.lbTimeout.Text)
    gRetry = val(Setupfrm.lbRetry.Text)
     
    Unload Me
End Sub




Private Sub Form_Load()

    
    
    Dim tmptxt As String
  
    EnableCloseButton Me.hWnd, False
          
    Setupfrm.cbNS.ListIndex = 0
    Setupfrm.cbEW.ListIndex = 0
    Setupfrm.cbhem.ListIndex = 0
 
    Set s_Profile = New DriverHelper.Profile
 
    tmptxt = s_Profile.GetValue(sID, "Port")
    If tmptxt <> "" Then Setupfrm.lbPort.Text = tmptxt
   
    tmptxt = s_Profile.GetValue(sID, "Baud")
    If tmptxt <> "" Then Setupfrm.lbBaud.Text = tmptxt
    
    tmptxt = s_Profile.GetValue(sID, "Timeout")
    If tmptxt <> "" Then Setupfrm.lbTimeout.Text = tmptxt
    
    tmptxt = s_Profile.GetValue(sID, "Retry")
    If tmptxt <> "" Then Setupfrm.lbRetry.Text = tmptxt
   
    tmptxt = s_Profile.GetValue(sID, "LongitudeDeg")
    If tmptxt <> "" Then Setupfrm.txtLongDeg.Text = tmptxt
     
    tmptxt = s_Profile.GetValue(sID, "LongitudeMin")
    If tmptxt <> "" Then Setupfrm.txtLongMin.Text = tmptxt
     
    tmptxt = s_Profile.GetValue(sID, "LongitudeEW")
    If tmptxt <> "" Then Setupfrm.cbEW.ListIndex = val(tmptxt)
     
    tmptxt = s_Profile.GetValue(sID, "LatitudeDeg")
    If tmptxt <> "" Then Setupfrm.txtLatDeg.Text = tmptxt
   
    tmptxt = s_Profile.GetValue(sID, "LatitudeMin")
    If tmptxt <> "" Then Setupfrm.txtLatMin.Text = tmptxt
     
    tmptxt = s_Profile.GetValue(sID, "LatitudeNS")
    If tmptxt <> "" Then Setupfrm.cbNS.ListIndex = val(tmptxt)
     
    tmptxt = s_Profile.GetValue(sID, "Elevation")
    If tmptxt <> "" Then Setupfrm.txtElevation = tmptxt
     
    tmptxt = s_Profile.GetValue(sID, "HemisphereNS")
    If tmptxt <> "" Then Setupfrm.cbhem.ListIndex = val(tmptxt)
    


End Sub


