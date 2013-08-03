VERSION 5.00
Begin VB.Form Setupfrm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8640
   Icon            =   "Setupfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Guiding"
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   6000
      TabIndex        =   64
      Top             =   3240
      Width           =   2535
      Begin VB.ComboBox ComboGuiding 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         ItemData        =   "Setupfrm.frx":0CCA
         Left            =   240
         List            =   "Setupfrm.frx":0CD4
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   240
         Width           =   2145
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Site Information"
      ForeColor       =   &H000080FF&
      Height          =   2655
      Left            =   2640
      TabIndex        =   0
      Top             =   2160
      Width           =   3255
      Begin VB.TextBox txtLongSec 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   2640
         TabIndex        =   61
         Text            =   "0"
         Top             =   1320
         Width           =   450
      End
      Begin VB.TextBox txtLatSec 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   2640
         TabIndex        =   60
         Text            =   "0"
         Top             =   915
         Width           =   450
      End
      Begin VB.TextBox txtElevation 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   1320
         TabIndex        =   50
         Text            =   "1000"
         Top             =   1725
         Width           =   525
      End
      Begin VB.CommandButton Commandgps 
         BackColor       =   &H0095C1CB&
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
         Left            =   2760
         Picture         =   "Setupfrm.frx":0CF2
         Style           =   1  'Graphical
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1695
         Width           =   375
      End
      Begin VB.ComboBox SitesCombo 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   1320
         TabIndex        =   48
         Text            =   "Sites"
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton CommandLoadSite 
         BackColor       =   &H0095C1CB&
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
         Left            =   240
         Picture         =   "Setupfrm.frx":12F4
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton CommandSaveSite 
         BackColor       =   &H0095C1CB&
         DisabledPicture =   "Setupfrm.frx":1876
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
         Left            =   840
         Picture         =   "Setupfrm.frx":1DF8
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   360
         Width           =   375
      End
      Begin VB.ComboBox cbhem 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Setupfrm.frx":237A
         Left            =   1320
         List            =   "Setupfrm.frx":2384
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2280
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox txtLatDeg 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   1920
         TabIndex        =   6
         Text            =   "14"
         Top             =   915
         Width           =   360
      End
      Begin VB.TextBox txtLatMin 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   2295
         TabIndex        =   5
         Text            =   "35"
         Top             =   915
         Width           =   330
      End
      Begin VB.TextBox txtLongDeg 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   1920
         TabIndex        =   4
         Text            =   "120"
         Top             =   1320
         Width           =   360
      End
      Begin VB.TextBox txtLongMin 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   2295
         TabIndex        =   3
         Text            =   "57"
         Top             =   1320
         Width           =   330
      End
      Begin VB.ComboBox cbEW 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         ItemData        =   "Setupfrm.frx":2396
         Left            =   1320
         List            =   "Setupfrm.frx":23A0
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1320
         Width           =   555
      End
      Begin VB.ComboBox cbNS 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         ItemData        =   "Setupfrm.frx":23AA
         Left            =   1320
         List            =   "Setupfrm.frx":23B4
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   915
         Width           =   555
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "GPS:"
         ForeColor       =   &H000080FF&
         Height          =   225
         Left            =   1920
         TabIndex        =   51
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "Hemisphere:"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   2280
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         Caption         =   "Longitude:"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1365
         Width           =   1005
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "Latitude:"
         ForeColor       =   &H000080FF&
         Height          =   225
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   930
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "Elevation (m):"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1755
         Width           =   990
      End
   End
   Begin VB.Frame Frame12 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Gamepad"
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   6000
      TabIndex        =   44
      Top             =   4080
      Width           =   2535
      Begin VB.CommandButton JStickSetupCommand 
         BackColor       =   &H0095C1CB&
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
         Picture         =   "Setupfrm.frx":23BE
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Mount Options"
      ForeColor       =   &H000080FF&
      Height          =   975
      Left            =   120
      TabIndex        =   36
      Top             =   1080
      Width           =   5775
      Begin VB.CommandButton Commandcustomise 
         BackColor       =   &H0095C1CB&
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
         Left            =   2040
         Picture         =   "Setupfrm.frx":2940
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Customise"
         Top             =   480
         Width           =   375
      End
      Begin VB.ComboBox ComboCoordType 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         ItemData        =   "Setupfrm.frx":2EC2
         Left            =   3960
         List            =   "Setupfrm.frx":2ECC
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   480
         Width           =   1695
      End
      Begin VB.ComboBox ComboMountType 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         ItemData        =   "Setupfrm.frx":2EE0
         Left            =   240
         List            =   "Setupfrm.frx":2EF3
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Caption         =   "Coordinate Type"
         ForeColor       =   &H000080FF&
         Height          =   435
         Left            =   3960
         TabIndex        =   40
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label12 
         BackColor       =   &H00000000&
         Caption         =   "Type"
         ForeColor       =   &H000080FF&
         Height          =   225
         Left            =   240
         TabIndex        =   39
         Top             =   240
         Width           =   930
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "General Options"
      ForeColor       =   &H000080FF&
      Height          =   2000
      Left            =   120
      TabIndex        =   26
      Top             =   4920
      Width           =   8415
      Begin VB.TextBox TextFriendlyName 
         BackColor       =   &H00000040&
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
         Height          =   285
         Left            =   6000
         TabIndex        =   68
         Top             =   1560
         Width           =   2295
      End
      Begin VB.ListBox SlewPresetList 
         BackColor       =   &H00000040&
         ForeColor       =   &H000080FF&
         Height          =   1425
         Left            =   2880
         TabIndex        =   62
         Top             =   480
         Width           =   855
      End
      Begin VB.ComboBox ComboLanguage 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         ItemData        =   "Setupfrm.frx":2F26
         Left            =   120
         List            =   "Setupfrm.frx":2F30
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   1560
         Width           =   2295
      End
      Begin VB.ComboBox ComboUpdateMode 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         ItemData        =   "Setupfrm.frx":2F4F
         Left            =   6000
         List            =   "Setupfrm.frx":2F5C
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   840
         Width           =   2295
      End
      Begin VB.ComboBox ComboProcessPriority 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         ItemData        =   "Setupfrm.frx":2F88
         Left            =   120
         List            =   "Setupfrm.frx":2F98
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   840
         Width           =   2295
      End
      Begin VB.CheckBox CheckAdvanced 
         BackColor       =   &H00000000&
         Caption         =   "Show Advanced Options"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   6000
         TabIndex        =   43
         Top             =   240
         Width           =   2295
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         ItemData        =   "Setupfrm.frx":2FBD
         Left            =   3840
         List            =   "Setupfrm.frx":2FDF
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   720
         Width           =   615
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   135
         Left            =   3840
         Max             =   809
         Min             =   1
         TabIndex        =   31
         Top             =   1560
         Value           =   10
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0095C1CB&
         Caption         =   "Set"
         Height          =   255
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1320
         Width           =   735
      End
      Begin VB.CheckBox ChkAutoFlip 
         BackColor       =   &H00000000&
         Caption         =   "Allow Auto Flip"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label17 
         BackColor       =   &H00000000&
         Caption         =   "Friendly Name"
         ForeColor       =   &H000080FF&
         Height          =   225
         Index           =   1
         Left            =   6000
         TabIndex        =   67
         Top             =   1320
         Width           =   2250
      End
      Begin VB.Label Label18 
         BackColor       =   &H00000000&
         Caption         =   "Language"
         ForeColor       =   &H000080FF&
         Height          =   225
         Left            =   120
         TabIndex        =   59
         Top             =   1320
         Width           =   2250
      End
      Begin VB.Label Label17 
         BackColor       =   &H00000000&
         Caption         =   "Update Notifications"
         ForeColor       =   &H000080FF&
         Height          =   225
         Index           =   0
         Left            =   6000
         TabIndex        =   57
         Top             =   600
         Width           =   2250
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "Windows Process Prioirty"
         ForeColor       =   &H000080FF&
         Height          =   225
         Left            =   120
         TabIndex        =   55
         Top             =   600
         Width           =   2250
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000000&
         Caption         =   "Slew Preset Rates:"
         ForeColor       =   &H000080FF&
         Height          =   225
         Left            =   2880
         TabIndex        =   35
         Top             =   240
         Width           =   1410
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         Caption         =   "No. of Presets"
         ForeColor       =   &H000080FF&
         Height          =   225
         Left            =   3840
         TabIndex        =   33
         Top             =   480
         Width           =   1410
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "1"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4680
         TabIndex        =   32
         Top             =   1320
         Width           =   495
      End
   End
   Begin VB.Frame FrameAscom 
      BackColor       =   &H00000000&
      Caption         =   "ASCOM Options"
      ForeColor       =   &H000080FF&
      Height          =   2055
      Left            =   6000
      TabIndex        =   24
      Top             =   1080
      Width           =   2535
      Begin VB.ComboBox ComboEPOCH 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         ItemData        =   "Setupfrm.frx":3002
         Left            =   120
         List            =   "Setupfrm.frx":3015
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CheckBox CheckSiteWrites 
         BackColor       =   &H00000000&
         Caption         =   "Allow Site Writes"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   960
         Width           =   2295
      End
      Begin VB.CheckBox ChkBlockPark 
         BackColor       =   &H00000000&
         Caption         =   "Synchronous Park"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   720
         Width           =   2295
      End
      Begin VB.CheckBox ChkPulseErrs 
         BackColor       =   &H00000000&
         Caption         =   "PulseGuide Exceptions"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   2295
      End
      Begin VB.CheckBox ChkExceptions 
         BackColor       =   &H00000000&
         Caption         =   "Issue Exceptions"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   2295
      End
      Begin VB.CheckBox ChkPulseguide 
         BackColor       =   &H00000000&
         Caption         =   "Pulseguide Support"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Visible         =   0   'False
         Width           =   2295
      End
   End
   Begin VB.CommandButton CommandOK 
      BackColor       =   &H0095C1CB&
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
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6960
      Width           =   8415
   End
   Begin VB.PictureBox picASCOM 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   840
      Left            =   120
      MouseIcon       =   "Setupfrm.frx":3043
      MousePointer    =   99  'Custom
      Picture         =   "Setupfrm.frx":3195
      ScaleHeight     =   840
      ScaleWidth      =   840
      TabIndex        =   13
      Top             =   120
      Width           =   840
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "EQMOD Port Details"
      ForeColor       =   &H000080FF&
      Height          =   2655
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   2415
      Begin VB.CommandButton Command2 
         BackColor       =   &H0095C1CB&
         Height          =   375
         Left            =   1920
         Picture         =   "Setupfrm.frx":5697
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Find"
         Top             =   1800
         Width           =   375
      End
      Begin VB.ComboBox lbRetry 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         ItemData        =   "Setupfrm.frx":5D09
         Left            =   1320
         List            =   "Setupfrm.frx":5D13
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   720
         Width           =   945
      End
      Begin VB.ComboBox lbTimeout 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         ItemData        =   "Setupfrm.frx":5D1D
         Left            =   1320
         List            =   "Setupfrm.frx":5D27
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   240
         Width           =   945
      End
      Begin VB.ComboBox lbBaud 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         ItemData        =   "Setupfrm.frx":5D37
         Left            =   1320
         List            =   "Setupfrm.frx":5D41
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1200
         Width           =   945
      End
      Begin VB.ComboBox lbPort 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         ItemData        =   "Setupfrm.frx":5D51
         Left            =   1320
         List            =   "Setupfrm.frx":5D85
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2160
         Width           =   945
      End
      Begin VB.Label Label14 
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   240
         TabIndex        =   42
         Top             =   2160
         Width           =   1050
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "Retry:"
         ForeColor       =   &H000080FF&
         Height          =   225
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Width           =   930
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Timeout:"
         ForeColor       =   &H000080FF&
         Height          =   225
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   930
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Baud:"
         ForeColor       =   &H000080FF&
         Height          =   225
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   930
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Port:"
         ForeColor       =   &H000080FF&
         Height          =   225
         Left            =   240
         TabIndex        =   12
         Top             =   1800
         Width           =   930
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
      Left            =   1080
      TabIndex        =   14
      Top             =   120
      Width           =   7455
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
' 25-Nov-06 rcs     Add Regional Setting Functions
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

Private s_Profile As DriverHelper.Profile
Private Const sID As String = "EQMOD.Telescope"
Private Const sDESC As String = "EQMOD ASCOM Scope Driver"

Private Sub cbNS_Change()
    On Error Resume Next
    cbhem.ListIndex = cbNS.ListIndex
End Sub

Private Sub CheckAdvanced_Click()
    On Error Resume Next
    If CheckAdvanced.Value = 1 Then
        Commandcustomise.Visible = True
    Else
        Commandcustomise.Visible = False
    End If
End Sub




Private Sub Combo2_Click()
    On Error Resume Next
    gPresetSlewRatesCount = Combo2.ListIndex + 1
    Call writePresetSlewRates
    Call readPresetSlewRates
    Call refreshrates
End Sub


Private Sub ComboMountType_Click()
    On Error Resume Next

    Select Case ComboMountType.ListIndex + 1
        
        Case EQMOUNT, NXMOUNT
            ComboCoordType.Visible = False
            Label13.Visible = False
        
        Case LXMOUNT, TKMOUNT, HBXMOUNT
            ComboCoordType.Visible = True
            Label13.Visible = True
            ComboCoordType.ListIndex = 0
    
    End Select

End Sub

Private Sub Commandcustomise_Click()
    CustomMountDlg.Show (1)
End Sub

Private Sub CommandOK_Click()
    Dim dllver As Integer
    Dim deg As Double
    Dim min As Double
    Dim sec As Double
    
    On Error Resume Next
    
    If CheckAdvanced.Value = 0 Then
        gAdvanced = 0
        HC.oPersist.WriteIniValue "Advanced", "0"
    Else
        gAdvanced = 1
        HC.oPersist.WriteIniValue "Advanced", "1"
    End If

    HC.oPersist.WriteIniValue "Port", CStr(Setupfrm.lbPort.Text)
    HC.oPersist.WriteIniValue "Baud", CStr(Setupfrm.lbBaud.Text)
    HC.oPersist.WriteIniValue "Timeout", CStr(Setupfrm.lbTimeout.Text)
    HC.oPersist.WriteIniValue "Retry", CStr(Setupfrm.lbRetry.Text)
    
    HC.oPersist.WriteIniValue "LatitudeSec", CStr(Setupfrm.txtLatSec.Text)
    HC.oPersist.WriteIniValue "LatitudeMin", CStr(Setupfrm.txtLatMin.Text)
    HC.oPersist.WriteIniValue "LatitudeDeg", CStr(Setupfrm.txtLatDeg.Text)
    HC.oPersist.WriteIniValue "LatitudeNS", CStr(Setupfrm.cbNS.ListIndex)
    HC.oPersist.WriteIniValue "LongitudeDeg", CStr(Setupfrm.txtLongDeg.Text)
    HC.oPersist.WriteIniValue "LongitudeMin", CStr(Setupfrm.txtLongMin.Text)
    HC.oPersist.WriteIniValue "LongitudeSec", CStr(Setupfrm.txtLongSec.Text)
    HC.oPersist.WriteIniValue "LongitudeEW", CStr(Setupfrm.cbEW.ListIndex)
    HC.oPersist.WriteIniValue "HemisphereNS", CStr(Setupfrm.cbhem.ListIndex)
    HC.oPersist.WriteIniValue "Elevation", CStr(Setupfrm.txtElevation.Text)
    HC.oPersist.WriteIniValue "SiteName", SitesCombo.Text
      
    If ChkPulseguide.Value = 0 Then
        gAscomCompatibility.AllowPulseGuide = False
    Else
        gAscomCompatibility.AllowPulseGuide = True
    End If
    
    Select Case ComboGuiding.ListIndex
        Case 1
            gAscomCompatibility.AllowPulseGuide = False
        Case Else
            gAscomCompatibility.AllowPulseGuide = True
    End Select
    
    gAscomCompatibility.Epoch = ComboEPOCH.ListIndex
    
    If ChkExceptions.Value = 0 Then
        gAscomCompatibility.AllowExceptions = False
    Else
        gAscomCompatibility.AllowExceptions = True
    End If
    
    If ChkPulseErrs.Value = 0 Then
        gAscomCompatibility.AllowPulseGuideExceptions = False
    Else
        gAscomCompatibility.AllowPulseGuideExceptions = True
    End If
    
    If ChkBlockPark.Value = 0 Then
        gAscomCompatibility.BlockPark = False
    Else
        gAscomCompatibility.BlockPark = True
    End If
    
    If CheckSiteWrites.Value = 0 Then
        gAscomCompatibility.AllowSiteWrites = False
    Else
        gAscomCompatibility.AllowSiteWrites = True
    End If
    
    Call WriteAscomCompatibiity
          
    If chkAutoFLip.Value = 0 Then
        ' user doesn't want the auto flip option
        gAutoFlipEnabled = False
        gAutoFlipAllowed = False
    Else
        gAutoFlipAllowed = True
    End If
    Call WriteAutoFlipData
          
    deg = CDbl(EQFixNum(txtLongDeg))
    If deg > 180 Then GoTo inputerror
    min = CDbl(EQFixNum(txtLongMin))
    If min >= 60 Then GoTo inputerror
    sec = CDbl(EQFixNum(txtLongSec))
    If sec > 60 Then GoTo inputerror
          
    HC.txtLongDeg.Text = Setupfrm.txtLongDeg.Text
    HC.txtLongMin.Text = Setupfrm.txtLongMin.Text
    HC.txtLongSec.Text = Setupfrm.txtLongSec.Text
    HC.cbEW.ListIndex = Setupfrm.cbEW.ListIndex
          
    gLongitude = deg + (min / 60#) + (sec / 3600#)
    If Setupfrm.cbEW.Text = oLangDll.GetLangString(115) Then gLongitude = -gLongitude  ' W is neg
          
    deg = CDbl(EQFixNum(txtLatDeg))
    If deg >= 90 Then GoTo inputerror
    min = CDbl(EQFixNum(txtLatMin))
    If min >= 60 Then GoTo inputerror
    sec = CDbl(EQFixNum(txtLatSec))
    If sec >= 60 Then GoTo inputerror
    
    HC.txtLatSec.Text = Setupfrm.txtLatSec.Text
    HC.txtLatMin.Text = Setupfrm.txtLatMin.Text
    HC.txtLatDeg.Text = Setupfrm.txtLatDeg.Text
    
    gLatitude = deg + (min / 60#) + (sec / 3600#)
          
    HC.cbNS.ListIndex = Setupfrm.cbNS.ListIndex
    HC.cbhem.ListIndex = Setupfrm.cbhem.ListIndex
    HC.txtElevation.Text = Setupfrm.txtElevation.Text
    HC.SitesCombo.Text = SitesCombo.Text
    
    If Setupfrm.cbNS.Text = oLangDll.GetLangString(116) Then gLatitude = -gLatitude
    
    gElevation = CDbl(EQFixNum(Setupfrm.txtElevation))
    
    If Setupfrm.cbhem.Text = oLangDll.GetLangString(1110) Then
        gHemisphere = 0
    Else
        gHemisphere = 1
    End If
       
       
    gPort = Setupfrm.lbPort.Text
    gBaud = val(Setupfrm.lbBaud.Text)
    gTimeout = val(Setupfrm.lbTimeout.Text)
    gRetry = val(Setupfrm.lbRetry.Text)
    
    On Error Resume Next
    
    If gDllVer >= 3.04 Then
    
        gMountType = ComboMountType.ListIndex + 1
        
        Select Case gMountType
            
            Case EQMOUNT, NXMOUNT
                ' microstep
                gCoordType = 0
            
            Case LXMOUNT, TKMOUNT, HBXMOUNT
                ' ALTAZ or RADEC
                gCoordType = ComboCoordType.ListIndex + 1
        
        End Select
    
    Else
        ' old dll only supports EQG, microstep
        gMountType = 0
        gCoordType = 0
    
    End If
    
    Call WriteMountType
    
    Call HC.oPersist.WriteIniValue("ProcessPrioirty", CStr(ComboProcessPriority.ListIndex))
    Call HC.oPersist.WriteIniValue("FriendlyName", TextFriendlyName.Text)
    
    gUpdateMode = ComboUpdateMode.ListIndex
    Call HC.oPersist.WriteIniValue("UpdateMode", CStr(gUpdateMode))

    Select Case ComboLanguage.ListIndex
        Case 0
            Call HC.oPersist.WriteIniValue("LANG_DLL", "")
        Case 1
            Call HC.oPersist.WriteIniValue("LANG_DLL", "eqmoden.dll")
        
    End Select
    
    Unload Me
    Exit Sub



inputerror:
    MsgBox (oLangDll.GetLangString(6014))

End Sub

Private Sub Command2_Click()
 Dim i As Integer
 Dim res As Long
 Dim strtmp As String
        
    Label14.Caption = "Searching"
    DoEvents
    
    ' close existing com port
    Call EQ_End
    For i = 1 To 16
        DoEvents
        lbPort.ListIndex = i - 1
        strtmp = "\\.\COM" & CStr(i)
        ' try to open com port send mesages.
        res = EQ_Init(strtmp, val(lbBaud.Text), val(lbTimeout.Text), 1)
        ' close com port
        Call EQ_End
        If res = 0 Then GoTo found
    Next i
    
notfound:
    Label14.Caption = "Not Found"
    Exit Sub
found:
    Label14.Caption = "Found"
End Sub

Private Sub CommandLoadSite_Click()
    On Error Resume Next
    
    HC.SiteIdx = SitesCombo.ListIndex
    LoadSite (HC.SiteIdx)
    
    txtLatSec.Text = HC.txtLatSec.Text
    txtLatMin.Text = HC.txtLatMin.Text
    txtLatDeg.Text = HC.txtLatDeg.Text
    cbNS.ListIndex = HC.cbNS.ListIndex
    txtLongDeg.Text = HC.txtLongDeg.Text
    txtLongMin.Text = HC.txtLongMin.Text
    txtLongSec.Text = HC.txtLongSec.Text
    cbEW.ListIndex = HC.cbEW.ListIndex
    cbhem.ListIndex = HC.cbhem.ListIndex
    txtElevation.Text = HC.txtElevation.Text

    Call WriteSiteValues

End Sub

Private Sub CommandSaveSite_Click()
    Dim N As String

    On Error Resume Next
    'site save
    N = SitesCombo.Text
    HC.txtLatSec.Text = Setupfrm.txtLatSec.Text
    HC.txtLatMin.Text = Setupfrm.txtLatMin.Text
    HC.txtLatDeg.Text = Setupfrm.txtLatDeg.Text
    HC.cbNS.ListIndex = Setupfrm.cbNS.ListIndex
    HC.txtLongDeg.Text = Setupfrm.txtLongDeg.Text
    HC.txtLongMin.Text = Setupfrm.txtLongMin.Text
    HC.txtLongSec.Text = Setupfrm.txtLongSec.Text
    HC.cbEW.ListIndex = Setupfrm.cbEW.ListIndex
    HC.cbhem.ListIndex = Setupfrm.cbhem.ListIndex
    HC.txtElevation.Text = Setupfrm.txtElevation.Text
    SitesCombo.List(HC.SiteIdx) = N
    HC.SitesCombo.List(HC.SiteIdx) = N
    Call SaveSite(HC.SiteIdx, N)
    SitesCombo.ListIndex = HC.SiteIdx
End Sub


Private Sub Commandgps_Click()
    On Error Resume Next
    
    Load GPSSetup
    GPSSetup.Show (1)
    txtLatSec.Text = HC.txtLatSec.Text
    txtLatMin.Text = HC.txtLatMin.Text
    txtLatDeg.Text = HC.txtLatDeg.Text
    cbNS.ListIndex = HC.cbNS.ListIndex
    txtLongDeg.Text = HC.txtLongDeg.Text
    txtLongMin.Text = HC.txtLongMin.Text
    txtLongSec.Text = HC.txtLongSec.Text
    cbEW.ListIndex = HC.cbEW.ListIndex
    cbhem.ListIndex = HC.cbhem.ListIndex
    txtElevation.Text = HC.txtElevation.Text

End Sub

Private Sub Form_Load()
   
    Dim tmptxt As String
    Dim i As Integer
    Dim tmp As Long
    
    On Error Resume Next
    
    Call SetText
    Call LoadSites(SitesCombo)

    EnableCloseButton Me.hwnd, False
          
    Setupfrm.cbNS.ListIndex = 0
    Setupfrm.cbEW.ListIndex = 0
    Setupfrm.cbhem.ListIndex = 0
 
    Set s_Profile = New DriverHelper.Profile
 
    tmptxt = HC.oPersist.ReadIniValue("Port")
    If tmptxt <> "" Then Setupfrm.lbPort.Text = tmptxt
   
    tmptxt = HC.oPersist.ReadIniValue("Baud")
    If tmptxt <> "" Then Setupfrm.lbBaud.Text = tmptxt
    
    tmptxt = HC.oPersist.ReadIniValue("Timeout")
    If tmptxt <> "" Then Setupfrm.lbTimeout.Text = tmptxt
    
    tmptxt = HC.oPersist.ReadIniValue("Retry")
    If tmptxt <> "" Then Setupfrm.lbRetry.Text = tmptxt
   
    tmptxt = HC.oPersist.ReadIniValue("LongitudeDeg")
    If tmptxt <> "" Then Setupfrm.txtLongDeg.Text = tmptxt
     
    tmptxt = HC.oPersist.ReadIniValue("LongitudeMin")
    If tmptxt <> "" Then Setupfrm.txtLongMin.Text = tmptxt
     
    tmptxt = HC.oPersist.ReadIniValue("LongitudeSec")
    If tmptxt <> "" Then Setupfrm.txtLongSec.Text = tmptxt
     
    tmptxt = HC.oPersist.ReadIniValue("LongitudeEW")
    If tmptxt <> "" Then Setupfrm.cbEW.ListIndex = val(tmptxt)
     
    tmptxt = HC.oPersist.ReadIniValue("LatitudeDeg")
    If tmptxt <> "" Then Setupfrm.txtLatDeg.Text = tmptxt
   
    tmptxt = HC.oPersist.ReadIniValue("LatitudeMin")
    If tmptxt <> "" Then Setupfrm.txtLatMin.Text = tmptxt
     
    tmptxt = HC.oPersist.ReadIniValue("LatitudeSec")
    If tmptxt <> "" Then Setupfrm.txtLatSec.Text = tmptxt
     
    tmptxt = HC.oPersist.ReadIniValue("LatitudeNS")
    If tmptxt <> "" Then Setupfrm.cbNS.ListIndex = val(tmptxt)
     
    tmptxt = HC.oPersist.ReadIniValue("Elevation")
    If tmptxt <> "" Then Setupfrm.txtElevation = tmptxt
     
    SitesCombo.Text = HC.oPersist.ReadIniValue("SiteName")
     
    Setupfrm.cbhem.ListIndex = Setupfrm.cbNS.ListIndex
'    tmptxt = HC.oPersist.ReadIniValue("HemisphereNS")
'    If tmptxt <> "" Then Setupfrm.cbhem.ListIndex = val(tmptxt)
    
    Call readAscomCompatibiity
    If gAscomCompatibility.AllowPulseGuide Then
        ComboGuiding.ListIndex = 0
    Else
        ComboGuiding.ListIndex = 1
    End If
    If gAscomCompatibility.AllowPulseGuide Then
        ChkPulseguide.Value = 1
    Else
        ChkPulseguide.Value = 0
    End If
    
    If gAscomCompatibility.AllowExceptions Then
        ChkExceptions.Value = 1
    Else
        ChkExceptions.Value = 0
    End If

    If gAscomCompatibility.AllowPulseGuideExceptions Then
        ChkPulseErrs.Value = 1
    Else
        ChkPulseErrs.Value = 0
    End If

    If gAscomCompatibility.BlockPark Then
        ChkBlockPark.Value = 1
    Else
        ChkBlockPark.Value = 0
    End If

    If gAscomCompatibility.AllowSiteWrites Then
        CheckSiteWrites.Value = 1
    Else
        CheckSiteWrites.Value = 0
    End If
    
    ComboEPOCH.ListIndex = gAscomCompatibility.Epoch


    Call readAutoFlipData
    If gAutoFlipAllowed Then
        chkAutoFLip.Value = 1
    Else
        chkAutoFLip.Value = 0
    End If

    Call readPresetSlewRates
    Combo2.ListIndex = gPresetSlewRatesCount - 1
    Call refreshrates

    Call ReadMountType
    
    If gMountType > 0 Then
        ComboMountType.ListIndex = gMountType - 1
    Else
        ComboMountType.ListIndex = 0
    End If
    
    If gCoordType > 0 Then
        ComboCoordType.ListIndex = gCoordType - 1
    Else
        ComboCoordType.ListIndex = 0
    End If
    
    If gDllVer < 3.04 Then
        ComboMountType.ListIndex = 0
        ComboMountType.Enabled = False
        ComboCoordType.Enabled = False
    Else
        
        Select Case gMountType
            
            Case EQMOUNT, NXMOUNT
                Label13.Visible = False
                ComboCoordType.Visible = False
            
            Case LXMOUNT, TKMOUNT, HBXMOUNT
                Label13.Visible = True
                ComboCoordType.Visible = True
        
        End Select
        
    End If

    Call readDevelopmentOptions
    If gAdvanced = 1 Then
        CheckAdvanced.Value = 1
        Commandcustomise.Visible = True
    Else
        CheckAdvanced.Value = 0
        Commandcustomise.Visible = False
    End If
    
    tmptxt = HC.oPersist.ReadIniValue("ProcessPrioirty")
    Select Case tmptxt
        Case "0", "1", "2", "3"
            Setupfrm.ComboProcessPriority.ListIndex = val(tmptxt)
        Case Else
            Setupfrm.ComboProcessPriority.ListIndex = 0
    End Select
    Setupfrm.ComboProcessPriority.Text = Setupfrm.ComboProcessPriority.List(Setupfrm.ComboProcessPriority.ListIndex)
    
    tmptxt = HC.oPersist.ReadIniValue("FriendlyName")
    Setupfrm.TextFriendlyName.Text = tmptxt
    
    
    Call ReadUpdateParams
    ComboUpdateMode.ListIndex = gUpdateMode
    ComboUpdateMode.Text = ComboUpdateMode.List(gUpdateMode)
    
    tmptxt = HC.oPersist.ReadIniValue("LANG_DLL")
    Select Case tmptxt
        Case "eqmoden.dll"
            ComboLanguage.ListIndex = 1
        Case Else
            ComboLanguage.ListIndex = 0
    End Select
    ComboLanguage.Text = ComboLanguage.List(ComboLanguage.ListIndex)
    
    Call PutWindowOnTop(Setupfrm)
    
End Sub

Private Sub SetText()
    On Error Resume Next
    
    Setupfrm.Caption = oLangDll.GetLangString(700)
    MainLabel.Caption = oLangDll.GetLangString(701)
    Frame2.Caption = oLangDll.GetLangString(702)
    Label1.Caption = oLangDll.GetLangString(703)
    Label2.Caption = oLangDll.GetLangString(704)
    Label3.Caption = oLangDll.GetLangString(705)
    Label4.Caption = oLangDll.GetLangString(706)
    
    Frame1.Caption = oLangDll.GetLangString(126)
    Label6.Caption = oLangDll.GetLangString(127)
    Label7.Caption = oLangDll.GetLangString(128)
    Label5.Caption = oLangDll.GetLangString(129)
    Label8.Caption = oLangDll.GetLangString(130)
    
    FrameAscom.Caption = oLangDll.GetLangString(2200)
    ChkPulseguide.Caption = oLangDll.GetLangString(2201)
    ChkPulseErrs.Caption = oLangDll.GetLangString(2205)
    ChkExceptions.Caption = oLangDll.GetLangString(2206)
    ChkBlockPark.Caption = oLangDll.GetLangString(2207)
    CheckSiteWrites.Caption = oLangDll.GetLangString(2208)
    
    Frame3.Caption = oLangDll.GetLangString(2202)
    chkAutoFLip.Caption = oLangDll.GetLangString(2203)
    Label11.Caption = oLangDll.GetLangString(707)
    Label10.Caption = oLangDll.GetLangString(708)
    Command1.Caption = oLangDll.GetLangString(709)
    CheckAdvanced.Caption = oLangDll.GetLangString(710)
    
    Frame4.Caption = oLangDll.GetLangString(711)
    Label12.Caption = oLangDll.GetLangString(712)
    Label13.Caption = oLangDll.GetLangString(713)
   
   
    CommandOK.Caption = oLangDll.GetLangString(609)
    
    Frame12.Caption = oLangDll.GetLangString(170)
    JStickSetupCommand.ToolTipText = oLangDll.GetLangString(171)
'    Command32.ToolTipText = oLangDll.GetLangString(111)

    Commandgps.ToolTipText = oLangDll.GetLangString(132)
    CommandLoadSite.ToolTipText = oLangDll.GetLangString(184)
    CommandSaveSite.ToolTipText = oLangDll.GetLangString(185)
'    CommandSetSite.Caption = oLangDll.GetLangString(131)

    Label16.Caption = oLangDll.GetLangString(6040)
    ComboProcessPriority.Clear
    ComboProcessPriority.AddItem (oLangDll.GetLangString(6041))
    ComboProcessPriority.AddItem (oLangDll.GetLangString(6042))
    ComboProcessPriority.AddItem (oLangDll.GetLangString(6043))
    ComboProcessPriority.AddItem (oLangDll.GetLangString(6044))

    Label17(0).Caption = oLangDll.GetLangString(6017)
    ComboUpdateMode.Clear
    ComboUpdateMode.AddItem (oLangDll.GetLangString(6018))
    ComboUpdateMode.AddItem (oLangDll.GetLangString(6019))
    ComboUpdateMode.AddItem (oLangDll.GetLangString(6020))
    
    Label18.Caption = oLangDll.GetLangString(6030)
    ComboLanguage.Clear
    ComboLanguage.AddItem (oLangDll.GetLangString(6031))
    ComboLanguage.AddItem (oLangDll.GetLangString(6032))

End Sub

Private Sub HScroll1_Change()
    If HScroll1.Value > 9 Then
        Label9.Caption = CStr(HScroll1.Value - 9)
    Else
        Label9.Caption = "0." & CStr(HScroll1.Value)
    End If
End Sub

Private Sub HScroll1_Scroll()
    HScroll1_Change
End Sub

Private Sub Command1_Click()
Dim val As Double
Dim i As Integer
    
    On Error Resume Next
    
    If SlewPresetList.ListIndex >= 0 Then
        If HScroll1.Value > 9 Then
            val = HScroll1.Value - 9
        Else
            val = HScroll1.Value / 10
        End If
        gPresetSlewRates(SlewPresetList.ListIndex + 1) = val
        Call writePresetSlewRates
        Call refreshrates
    End If
    
End Sub
Private Sub SlewPresetList_Click()
Dim val As Double
    On Error Resume Next
    val = gPresetSlewRates(SlewPresetList.ListIndex + 1)
    If val = 0 Then val = 1
    If val < 1 Then
        val = val * 10
    Else
        val = val + 9
    End If
    HScroll1.Value = val
End Sub


Private Sub refreshrates()
Dim i As Integer
    On Error Resume Next
    SlewPresetList.Clear
    For i = 1 To gPresetSlewRatesCount
        If gPresetSlewRates(i) = 0 Then
            SlewPresetList.AddItem CStr(i) & ": -"
        Else
            SlewPresetList.AddItem CStr(i) & ": " & CStr(gPresetSlewRates(i))
        End If
    Next i
End Sub

Private Sub JStickSetupCommand_Click()
    Load JStickConfigForm
    JStickConfigForm.Show (1)
End Sub

Private Sub SitesCombo_Click()
    HC.SiteIdx = SitesCombo.ListIndex
    CommandLoadSite_Click
End Sub


