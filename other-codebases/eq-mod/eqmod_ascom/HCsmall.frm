VERSION 5.00
Begin VB.Form HC 
   BackColor       =   &H00000000&
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13935
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "HCsmall.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8175
   ScaleWidth      =   13935
   Begin VB.Frame FrameCustomTrack 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Tracking"
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
      Height          =   615
      Left            =   7320
      TabIndex        =   240
      Top             =   5280
      Width           =   3135
      Begin VB.CommandButton CommandRemoveTrackFile 
         BackColor       =   &H0095C1CB&
         Height          =   300
         Left            =   2760
         Picture         =   "HCsmall.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   242
         ToolTipText     =   "Unload"
         Top             =   240
         Width           =   300
      End
      Begin VB.CommandButton CommandLoadTracking 
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
         Height          =   300
         Left            =   0
         Picture         =   "HCsmall.frx":13F0
         Style           =   1  'Graphical
         TabIndex        =   241
         Top             =   240
         Width           =   300
      End
      Begin VB.Label LabelTrackFile 
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
         Height          =   255
         Left            =   360
         TabIndex        =   243
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "TrackFile:"
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
         Left            =   0
         TabIndex        =   244
         Top             =   0
         Width           =   2655
      End
   End
   Begin VB.Timer CustomTrackTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   11160
      Top             =   6720
   End
   Begin VB.Timer PECCapTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   11160
      Top             =   6240
   End
   Begin VB.CommandButton CmdDisplayMode 
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
      Height          =   285
      Left            =   120
      Picture         =   "HCsmall.frx":1972
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton CommandSetup 
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
      Height          =   285
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton CommandUpdate 
      BackColor       =   &H0095C1CB&
      Caption         =   "Updates!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   922
      Style           =   1  'Graphical
      TabIndex        =   200
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Timer AutoParkTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10680
      Top             =   5760
   End
   Begin VB.Frame FrameAdvanced 
      BackColor       =   &H00000000&
      Caption         =   "Advanced"
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
      Height          =   2655
      Left            =   10680
      TabIndex        =   170
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
      Begin VB.HScrollBar HScrollSlewAdjust 
         Height          =   150
         Left            =   120
         Max             =   100
         TabIndex        =   177
         Top             =   1440
         Width           =   2895
      End
      Begin VB.HScrollBar HScrollSlewRetries 
         Height          =   150
         Left            =   120
         Max             =   10
         TabIndex        =   174
         Top             =   960
         Width           =   2895
      End
      Begin VB.HScrollBar HScrollGotoRes 
         Height          =   150
         Left            =   120
         Max             =   50
         TabIndex        =   171
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label LabelSlewAdjust 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "360"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2520
         TabIndex        =   179
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label79 
         BackColor       =   &H00000000&
         Caption         =   "RASlew Adjust"
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
         TabIndex        =   178
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label LabelSlewRetries 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "360"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2520
         TabIndex        =   176
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label78 
         BackColor       =   &H00000000&
         Caption         =   "Slew Retries"
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
         TabIndex        =   175
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label LabelGotoRes 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "360"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2520
         TabIndex        =   173
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label38 
         BackColor       =   &H00000000&
         Caption         =   "Goto Resolution"
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
         TabIndex        =   172
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00000000&
      Caption         =   "Development"
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
      Height          =   1575
      Left            =   10680
      TabIndex        =   153
      Top             =   2760
      Visible         =   0   'False
      Width           =   3135
      Begin VB.ComboBox cbhem 
         BackColor       =   &H00000080&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   315
         ItemData        =   "HCsmall.frx":22B4
         Left            =   1320
         List            =   "HCsmall.frx":22BE
         Style           =   2  'Dropdown List
         TabIndex        =   187
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1125
      End
      Begin VB.CheckBox PierStrict 
         BackColor       =   &H00000000&
         Caption         =   "PierSide Points Only"
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
         Left            =   240
         TabIndex        =   165
         TabStop         =   0   'False
         Top             =   720
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox PolarEnable 
         BackColor       =   &H00000000&
         Caption         =   "AFFINE_TAKI+POLAR"
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
         Left            =   240
         TabIndex        =   154
         TabStop         =   0   'False
         Top             =   360
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "Hemisphere:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   188
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.Frame FrameLimits 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Mount Limits"
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
      Height          =   735
      Left            =   7200
      TabIndex        =   148
      Top             =   2640
      Width           =   3375
      Begin VB.CommandButton CmdEditLimits 
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
         Picture         =   "HCsmall.frx":22D0
         Style           =   1  'Graphical
         TabIndex        =   150
         ToolTipText     =   "Configure"
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox ChkEnableLimits 
         BackColor       =   &H00000000&
         Caption         =   "Enable Limits"
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
         Left            =   600
         TabIndex        =   149
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame TrackingFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Track Rate : Not Tracking "
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
      Height          =   1190
      Left            =   120
      TabIndex        =   28
      Top             =   6230
      Width           =   2895
      Begin VB.TextBox decCustom 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   1920
         TabIndex        =   237
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   720
         Width           =   840
      End
      Begin VB.TextBox raCustom 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   480
         TabIndex        =   236
         TabStop         =   0   'False
         Text            =   "15.04157"
         Top             =   720
         Width           =   825
      End
      Begin VB.CommandButton CmdTrack 
         BackColor       =   &H0095C1CB&
         DisabledPicture =   "HCsmall.frx":2852
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
         Index           =   1
         Left            =   690
         Picture         =   "HCsmall.frx":2DD4
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton CmdTrack 
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
         Index           =   4
         Left            =   2400
         Picture         =   "HCsmall.frx":3356
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton CmdTrack 
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
         Index           =   0
         Left            =   690
         Picture         =   "HCsmall.frx":38D8
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton CmdTrack 
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
         Index           =   5
         Left            =   120
         Picture         =   "HCsmall.frx":3E5A
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton CmdTrack 
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
         Index           =   3
         Left            =   1830
         Picture         =   "HCsmall.frx":43DC
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton CmdTrack 
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
         Index           =   2
         Left            =   1260
         Picture         =   "HCsmall.frx":495E
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label39 
         BackColor       =   &H00000000&
         Caption         =   "DEC:"
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
         Left            =   1440
         TabIndex        =   239
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label24 
         BackColor       =   &H00000000&
         Caption         =   "RA:"
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
         TabIndex        =   238
         Top             =   720
         Width           =   375
      End
   End
   Begin VB.Frame Frame15 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Park Status:"
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
      Height          =   615
      Left            =   120
      TabIndex        =   100
      Top             =   7440
      Width           =   2895
      Begin VB.CommandButton CommandPark 
         BackColor       =   &H0095C1CB&
         Caption         =   "PARK"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   295
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame11 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Other Settings"
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
      Height          =   2415
      Left            =   7200
      TabIndex        =   81
      Top             =   3600
      Width           =   3375
      Begin VB.HScrollBar HScrollSlewLimit 
         Height          =   135
         Left            =   120
         Max             =   400
         Min             =   49
         TabIndex        =   201
         Top             =   1320
         Value           =   50
         Width           =   2175
      End
      Begin VB.CommandButton CommandSounds 
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
         Picture         =   "HCsmall.frx":4EE0
         Style           =   1  'Graphical
         TabIndex        =   152
         Top             =   600
         Width           =   375
      End
      Begin VB.CheckBox HCOnTop 
         BackColor       =   &H00000000&
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   240
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.Label LabelSlewLimit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000040&
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1560
         TabIndex        =   202
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label31 
         BackColor       =   &H00000000&
         Caption         =   "Goto Slew Limit"
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
         TabIndex        =   199
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "Always on Top"
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
         Left            =   480
         TabIndex        =   83
         Top             =   240
         Width           =   1845
      End
   End
   Begin VB.Timer PECTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10680
      Top             =   5280
   End
   Begin VB.Timer ParkTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   11160
      Top             =   5760
   End
   Begin VB.Timer Spiral_Timer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   10680
      Top             =   7680
   End
   Begin VB.Timer JoyStick_Timer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10680
      Top             =   7200
   End
   Begin VB.Timer Pulseguide_Timer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   11160
      Top             =   5280
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Alignment / Sync"
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
      Height          =   3375
      Left            =   3240
      TabIndex        =   58
      Top             =   2625
      Width           =   3855
      Begin VB.CheckBox CheckLocalPier 
         BackColor       =   &H00000000&
         Caption         =   "local to pier"
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
         Left            =   2040
         TabIndex        =   167
         Top             =   2160
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.ComboBox ComboActivePoints 
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
         ItemData        =   "HCsmall.frx":5462
         Left            =   120
         List            =   "HCsmall.frx":546F
         Style           =   2  'Dropdown List
         TabIndex        =   164
         Top             =   2760
         Width           =   1695
      End
      Begin VB.ComboBox Combo3PointAlgorithm 
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
         ItemData        =   "HCsmall.frx":54A5
         Left            =   2040
         List            =   "HCsmall.frx":54AF
         Style           =   2  'Dropdown List
         TabIndex        =   163
         Top             =   2760
         Width           =   1695
      End
      Begin VB.HScrollBar HScrollProximity 
         Height          =   135
         Left            =   120
         Max             =   15
         TabIndex        =   160
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CommandButton CommandAddPoint 
         BackColor       =   &H0095C1CB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         Picture         =   "HCsmall.frx":54D0
         Style           =   1  'Graphical
         TabIndex        =   155
         ToolTipText     =   "Add Point(s)"
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox ListSyncMode 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   345
         ItemData        =   "HCsmall.frx":5A52
         Left            =   120
         List            =   "HCsmall.frx":5A54
         Style           =   2  'Dropdown List
         TabIndex        =   94
         Top             =   1560
         Width           =   1695
      End
      Begin VB.ComboBox ListAlignMode 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   345
         ItemData        =   "HCsmall.frx":5A56
         Left            =   2040
         List            =   "HCsmall.frx":5A58
         Style           =   2  'Dropdown List
         TabIndex        =   92
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CommandButton Edit_Stars_Command 
         BackColor       =   &H0095C1CB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Picture         =   "HCsmall.frx":5A5A
         Style           =   1  'Graphical
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Reset_Align_Command 
         BackColor       =   &H0095C1CB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         Picture         =   "HCsmall.frx":5FDC
         Style           =   1  'Graphical
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton CmdClearSync 
         BackColor       =   &H0095C1CB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Picture         =   "HCsmall.frx":655E
         Style           =   1  'Graphical
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label35 
         BackColor       =   &H00000000&
         Caption         =   "3 Point Selection"
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
         Left            =   2040
         TabIndex        =   169
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label29 
         BackColor       =   &H00000000&
         Caption         =   "Point Filter"
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
         TabIndex        =   168
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label37 
         BackColor       =   &H00000000&
         Caption         =   "Proximity Range"
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
         TabIndex        =   162
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label LabelProximity 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000040&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1320
         TabIndex        =   161
         Top             =   2040
         Width           =   600
      End
      Begin VB.Label Label36 
         BackColor       =   &H00000000&
         Caption         =   "Point Count:"
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
         Height          =   195
         Left            =   2280
         TabIndex        =   159
         Top             =   360
         Width           =   915
      End
      Begin VB.Label AlignmentCountLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000080&
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3240
         TabIndex        =   158
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label34 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Sync Behavior"
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
         Height          =   195
         Left            =   120
         TabIndex        =   95
         Top             =   1320
         Width           =   1635
      End
      Begin VB.Label Label44 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Alignment Behavior"
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
         Height          =   195
         Left            =   2040
         TabIndex        =   93
         Top             =   1320
         Width           =   1635
      End
      Begin VB.Label DxSblbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000080&
         Caption         =   "000000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2880
         TabIndex        =   61
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000000&
         Caption         =   "DxSB:"
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
         Height          =   195
         Left            =   2280
         TabIndex        =   67
         Top             =   960
         Width           =   525
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "DxSA:"
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
         Height          =   195
         Left            =   600
         TabIndex        =   64
         Top             =   960
         Width           =   525
      End
      Begin VB.Label DxSalbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000080&
         Caption         =   "000000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1200
         TabIndex        =   62
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Park / Unpark"
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
      Height          =   1965
      Left            =   3240
      TabIndex        =   59
      Top             =   6040
      Width           =   3855
      Begin VB.TextBox TextAutoParkDuration 
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
         Height          =   315
         Left            =   600
         TabIndex        =   213
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Time to Park (mins)"
         Top             =   1510
         Width           =   1080
      End
      Begin VB.CommandButton CmdParkTimer 
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
         Picture         =   "HCsmall.frx":6AE0
         Style           =   1  'Graphical
         TabIndex        =   212
         TabStop         =   0   'False
         ToolTipText     =   "Start Park Timer"
         Top             =   1480
         Width           =   375
      End
      Begin VB.CommandButton CommandSyncEncoder 
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
         Left            =   3360
         Picture         =   "HCsmall.frx":7062
         Style           =   1  'Graphical
         TabIndex        =   156
         Top             =   480
         Width           =   375
      End
      Begin VB.ComboBox ComboPark 
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
         ItemData        =   "HCsmall.frx":75E4
         Left            =   120
         List            =   "HCsmall.frx":75E6
         Style           =   2  'Dropdown List
         TabIndex        =   103
         Top             =   480
         Width           =   3015
      End
      Begin VB.ComboBox ComboUnPark 
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   102
         Top             =   1080
         Width           =   3015
      End
      Begin VB.CommandButton CommandDefinePark 
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
         Left            =   3360
         Picture         =   "HCsmall.frx":75E8
         Style           =   1  'Graphical
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label12 
         BackColor       =   &H00000000&
         Caption         =   "Minutes Until Park"
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
         Left            =   1800
         TabIndex        =   214
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label50 
         BackColor       =   &H00000000&
         Caption         =   "Unpark Mode"
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
         TabIndex        =   105
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label49 
         BackColor       =   &H00000000&
         Caption         =   "Park Mode"
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
         TabIndex        =   104
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Timer EncoderTimer 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   10680
      Top             =   6720
   End
   Begin VB.Timer DisplayTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10680
      Top             =   6240
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Site Information"
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
      Height          =   2505
      Left            =   3240
      TabIndex        =   37
      Top             =   0
      Width           =   3855
      Begin VB.ComboBox ComboPoleStar 
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
         ItemData        =   "HCsmall.frx":7B6A
         Left            =   2640
         List            =   "HCsmall.frx":7B7D
         Style           =   2  'Dropdown List
         TabIndex        =   211
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtLongSec 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   3000
         TabIndex        =   207
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   945
         Width           =   570
      End
      Begin VB.TextBox txtLatSec 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   3000
         TabIndex        =   206
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   600
         Width           =   570
      End
      Begin VB.CommandButton CommandPolaris 
         BackColor       =   &H0095C1CB&
         Caption         =   "12345"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   295
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   151
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton CommandSaveSite 
         BackColor       =   &H0095C1CB&
         DisabledPicture =   "HCsmall.frx":7BAE
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
         Picture         =   "HCsmall.frx":8130
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   240
         Width           =   375
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
         Picture         =   "HCsmall.frx":86B2
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox SitesCombo 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   1560
         TabIndex        =   40
         Text            =   "Sites"
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton CommandGPS 
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
         Height          =   315
         Left            =   3240
         Picture         =   "HCsmall.frx":8C34
         Style           =   1  'Graphical
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton CommandSetSite 
         BackColor       =   &H0095C1CB&
         Caption         =   "SET"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   295
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txtElevation 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   1560
         TabIndex        =   49
         TabStop         =   0   'False
         Text            =   "1000"
         Top             =   1290
         Width           =   525
      End
      Begin VB.ComboBox cbNS 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   315
         ItemData        =   "HCsmall.frx":9236
         Left            =   1560
         List            =   "HCsmall.frx":9240
         Style           =   2  'Dropdown List
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   600
         Width           =   555
      End
      Begin VB.ComboBox cbEW 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   315
         ItemData        =   "HCsmall.frx":924A
         Left            =   1560
         List            =   "HCsmall.frx":9254
         Style           =   2  'Dropdown List
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   945
         Width           =   555
      End
      Begin VB.TextBox txtLongMin 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   2595
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   "57"
         Top             =   945
         Width           =   330
      End
      Begin VB.TextBox txtLongDeg 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   2160
         TabIndex        =   45
         TabStop         =   0   'False
         Text            =   "120"
         Top             =   945
         Width           =   360
      End
      Begin VB.TextBox txtLatMin 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   2595
         TabIndex        =   43
         TabStop         =   0   'False
         Text            =   "35"
         Top             =   600
         Width           =   330
      End
      Begin VB.TextBox txtLatDeg 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   2160
         TabIndex        =   42
         TabStop         =   0   'False
         Text            =   "14"
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "GPS:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   195
         Index           =   3
         Left            =   2280
         TabIndex        =   189
         Top             =   1380
         Width           =   855
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "Polaris HA:"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   147
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "Elevation (m):"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   52
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "Latitude:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   48
         Top             =   645
         Width           =   1170
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "Longitude:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   46
         Top             =   1020
         Width           =   1125
      End
   End
   Begin VB.PictureBox picASCOM 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   120
      MouseIcon       =   "HCsmall.frx":925E
      Picture         =   "HCsmall.frx":93B0
      ScaleHeight     =   690
      ScaleWidth      =   720
      TabIndex        =   53
      Top             =   80
      Width           =   720
   End
   Begin VB.Frame Frame13 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Drift Compensation"
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
      Height          =   975
      Left            =   7200
      TabIndex        =   96
      Top             =   7080
      Width           =   3375
      Begin VB.HScrollBar DriftScroll 
         Height          =   150
         Left            =   600
         Max             =   20
         Min             =   -20
         TabIndex        =   185
         Top             =   292
         Width           =   2535
      End
      Begin VB.CheckBox CheckRASync 
         BackColor       =   &H00000000&
         Caption         =   "Auto RA Sync"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   184
         TabStop         =   0   'False
         Top             =   600
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.Label Driftlbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000080&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   186
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame12 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Gamepad"
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
      Height          =   735
      Left            =   7200
      TabIndex        =   84
      Top             =   6120
      Width           =   3375
      Begin VB.CommandButton CommandGamepad 
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
         Index           =   1
         Left            =   600
         Picture         =   "HCsmall.frx":ADD2
         Style           =   1  'Graphical
         TabIndex        =   85
         TabStop         =   0   'False
         ToolTipText     =   "Initialize"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton CommandGamepad 
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
         Index           =   0
         Left            =   120
         Picture         =   "HCsmall.frx":B354
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame SlewFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Slew Controls"
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
      Height          =   2620
      Left            =   120
      TabIndex        =   11
      Top             =   3555
      Width           =   2895
      Begin VB.CommandButton Command13 
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
         Index           =   0
         Left            =   120
         Picture         =   "HCsmall.frx":B8D6
         Style           =   1  'Graphical
         TabIndex        =   125
         ToolTipText     =   "Tour"
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton Command13 
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
         Index           =   1
         Left            =   600
         Picture         =   "HCsmall.frx":BF48
         Style           =   1  'Graphical
         TabIndex        =   124
         ToolTipText     =   "Mosaic"
         Top             =   2160
         Width           =   375
      End
      Begin VB.HScrollBar SpiralHScroll1 
         Height          =   150
         Left            =   1560
         Max             =   4000
         Min             =   70
         TabIndex        =   23
         Top             =   2400
         Value           =   2000
         Width           =   1215
      End
      Begin VB.CommandButton CmdSlew 
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
         Index           =   5
         Left            =   1080
         MaskColor       =   &H80000010&
         Picture         =   "HCsmall.frx":C4CA
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2160
         Width           =   375
      End
      Begin VB.VScrollBar VScrollRASlewRate 
         Height          =   1215
         Left            =   1680
         Max             =   809
         Min             =   1
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   480
         Value           =   800
         Width           =   180
      End
      Begin VB.ComboBox PresetRateCombo 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   345
         ItemData        =   "HCsmall.frx":CA4C
         Left            =   1912
         List            =   "HCsmall.frx":CA53
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   840
         Width           =   555
      End
      Begin VB.CheckBox DEC_Inv 
         BackColor       =   &H00000000&
         Caption         =   "DEC Reverse"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CheckBox RA_inv 
         BackColor       =   &H00000000&
         Caption         =   "RA Reverse"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton CmdSlewPad 
         BackColor       =   &H0095C1CB&
         Caption         =   "PAD"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   295
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1400
         Width           =   375
      End
      Begin VB.CommandButton CmdSlew 
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
         Height          =   495
         Index           =   4
         Left            =   600
         MaskColor       =   &H80000010&
         Picture         =   "HCsmall.frx":CA5B
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   720
         Width           =   495
      End
      Begin VB.VScrollBar VScrollDecSlewRate 
         Height          =   1215
         Left            =   2520
         Max             =   809
         Min             =   1
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   480
         Value           =   800
         Width           =   180
      End
      Begin VB.CommandButton CmdSlew 
         BackColor       =   &H0095C1CB&
         Caption         =   "W"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   120
         MaskColor       =   &H80000010&
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton CmdSlew 
         BackColor       =   &H0095C1CB&
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   1080
         MaskColor       =   &H80000010&
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton CmdSlew 
         BackColor       =   &H0095C1CB&
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   600
         MaskColor       =   &H80000010&
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton CmdSlew 
         BackColor       =   &H0095C1CB&
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   600
         MaskColor       =   &H80000010&
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
      End
      Begin VB.Label LabelSpiral 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000040&
         Caption         =   " 2000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2160
         TabIndex        =   88
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "DEC Rate"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   2040
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "RA Rate"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   1200
         TabIndex        =   26
         Top             =   240
         Width           =   735
      End
      Begin VB.Label decslewlbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000040&
         Caption         =   " 800"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2120
         TabIndex        =   25
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label raslewlbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000040&
         Caption         =   " 800"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1280
         TabIndex        =   24
         Top             =   1440
         Width           =   375
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Autoguider Port Rate"
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
      Height          =   2535
      Left            =   7200
      TabIndex        =   73
      Top             =   0
      Width           =   3375
      Begin VB.ComboBox DECGuideRateList 
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
         ItemData        =   "HCsmall.frx":CFDD
         Left            =   1800
         List            =   "HCsmall.frx":CFF0
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox RAGuideRateList 
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
         ItemData        =   "HCsmall.frx":D01A
         Left            =   1800
         List            =   "HCsmall.frx":D02D
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox DummyChk 
         Caption         =   "Check1"
         Height          =   225
         Left            =   1800
         TabIndex        =   87
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "DEC Rate"
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
         TabIndex        =   79
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "RA Rate"
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
         TabIndex        =   78
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "ASCOM PulseGuide Settings"
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
      Height          =   2535
      Left            =   7200
      TabIndex        =   68
      Top             =   0
      Width           =   3375
      Begin VB.Frame FramePGAvanced 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame10"
         Height          =   800
         Left            =   80
         TabIndex        =   190
         Top             =   1680
         Width           =   3250
         Begin VB.CheckBox rafixed_enchk 
            BackColor       =   &H00000000&
            Caption         =   "RA"
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
            Left            =   0
            TabIndex        =   194
            TabStop         =   0   'False
            Top             =   360
            Width           =   975
         End
         Begin VB.CheckBox decfixed_enchk 
            BackColor       =   &H00000000&
            Caption         =   "DEC"
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
            Left            =   1680
            TabIndex        =   193
            TabStop         =   0   'False
            Top             =   360
            Width           =   975
         End
         Begin VB.HScrollBar HScrollRAOride 
            Height          =   135
            Left            =   0
            Max             =   50
            Min             =   1
            TabIndex        =   192
            Top             =   600
            Value           =   1
            Width           =   1455
         End
         Begin VB.HScrollBar HScrollDecOride 
            Height          =   135
            Left            =   1680
            Max             =   50
            Min             =   1
            TabIndex        =   191
            Top             =   600
            Value           =   1
            Width           =   1455
         End
         Begin VB.Label Label16 
            BackColor       =   &H00000000&
            Caption         =   "Pulse Width Overide (msecs)"
            ForeColor       =   &H000080FF&
            Height          =   375
            Left            =   0
            TabIndex        =   197
            Top             =   0
            Width           =   3255
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000040&
            Caption         =   " 0100"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   960
            TabIndex        =   196
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000040&
            Caption         =   " 0100"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   2640
            TabIndex        =   195
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.HScrollBar HScrollBacklashDec 
         Height          =   135
         Left            =   120
         Max             =   200
         TabIndex        =   180
         Top             =   1440
         Width           =   3135
      End
      Begin VB.HScrollBar PltimerHscroll 
         Height          =   135
         Left            =   120
         Max             =   50
         Min             =   20
         TabIndex        =   97
         Top             =   960
         Value           =   50
         Width           =   3135
      End
      Begin VB.HScrollBar HScrollDecRate 
         Height          =   135
         Left            =   1800
         Max             =   9
         Min             =   1
         TabIndex        =   72
         Top             =   480
         Value           =   1
         Width           =   1455
      End
      Begin VB.HScrollBar HScrollRARate 
         Height          =   135
         Left            =   120
         Max             =   9
         Min             =   1
         TabIndex        =   75
         Top             =   480
         Value           =   1
         Width           =   1455
      End
      Begin VB.CheckBox decpulse_enchk 
         BackColor       =   &H00000000&
         Caption         =   "DEC Rate"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   1800
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   240
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox rapulse_enchk 
         BackColor       =   &H00000000&
         Caption         =   "RA Rate"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   240
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.Label Label81 
         BackColor       =   &H00000000&
         Caption         =   "DEC Backlash (msecs)"
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
         Height          =   375
         Left            =   120
         TabIndex        =   182
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label LabelBacklashDec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000040&
         Caption         =   " 0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2880
         TabIndex        =   181
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000040&
         Caption         =   " 50"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2880
         TabIndex        =   99
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Minimum Pulse Width (msecs)"
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
         Height          =   375
         Left            =   120
         TabIndex        =   98
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label15 
         BackColor       =   &H00000040&
         Caption         =   "x0.10"
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
         Left            =   2880
         TabIndex        =   70
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label14 
         BackColor       =   &H00000040&
         Caption         =   "x0.10"
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
         Left            =   1200
         TabIndex        =   69
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame FrameAxis 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Mount Position"
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
      Height          =   2415
      Left            =   120
      TabIndex        =   203
      Top             =   1120
      Width           =   2895
      Begin VB.CheckBox ChkForceFlip 
         BackColor       =   &H00000000&
         Caption         =   "Force Flipped GoTo"
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
         TabIndex        =   208
         Top             =   2040
         Width           =   2055
      End
      Begin VB.PictureBox PictureRA 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   81
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   81
         TabIndex        =   205
         Top             =   240
         Width           =   1215
      End
      Begin VB.PictureBox PictureDEC 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   1560
         ScaleHeight     =   81
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   81
         TabIndex        =   204
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label83 
         Alignment       =   2  'Center
         BackColor       =   &H00000040&
         Caption         =   "80000"
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
         Left            =   1560
         TabIndex        =   210
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label82 
         Alignment       =   2  'Center
         BackColor       =   &H00000040&
         Caption         =   "80000"
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
         TabIndex        =   209
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.Frame PositionFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Mount Position"
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
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   1120
      Width           =   2895
      Begin VB.TextBox TextCommsErr 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2055
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   198
         Text            =   "HCsmall.frx":D057
         Top             =   240
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label lblPier 
         BackColor       =   &H00000080&
         Caption         =   "West, pointing East"
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
         Left            =   1200
         TabIndex        =   55
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "PierSide"
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
         TabIndex        =   54
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label azmrk 
         BackColor       =   &H00000040&
         Caption         =   "AZ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label altmrk 
         BackColor       =   &H00000040&
         Caption         =   "ALT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label decmrk 
         BackColor       =   &H00000040&
         Caption         =   "DEC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
      Begin VB.Label ramrk 
         BackColor       =   &H00000040&
         Caption         =   "RA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lstmrk 
         BackColor       =   &H00000040&
         Caption         =   "LST"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   335
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.Label azlbl 
         BackColor       =   &H00000080&
         Caption         =   "+90:00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   1200
         TabIndex        =   5
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label altlbl 
         BackColor       =   &H00000080&
         Caption         =   "+90:00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   1200
         TabIndex        =   4
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lstlbl 
         BackColor       =   &H00000080&
         Caption         =   "+12:00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   335
         Left            =   1200
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label declbl 
         BackColor       =   &H00000080&
         Caption         =   "+90:00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   1200
         TabIndex        =   2
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label ralbl 
         BackColor       =   &H00000080&
         Caption         =   "+12:00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Frame Slew2Frame 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Slew Controls"
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
      Height          =   1875
      Left            =   120
      TabIndex        =   215
      Top             =   960
      Width           =   2895
      Begin VB.CheckBox Reverse 
         BackColor       =   &H00000000&
         Caption         =   "DEC Reverse"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   223
         TabStop         =   0   'False
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox Reverse 
         BackColor       =   &H00000000&
         Caption         =   "RA Reverse"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   222
         TabStop         =   0   'False
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox PresetRate2Combo 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   345
         ItemData        =   "HCsmall.frx":D06B
         Left            =   1560
         List            =   "HCsmall.frx":D072
         Style           =   2  'Dropdown List
         TabIndex        =   221
         Top             =   1320
         Width           =   555
      End
      Begin VB.CommandButton cmdSlew2 
         BackColor       =   &H0095C1CB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   480
         MaskColor       =   &H80000010&
         Picture         =   "HCsmall.frx":D07A
         Style           =   1  'Graphical
         TabIndex        =   220
         TabStop         =   0   'False
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdSlew2 
         BackColor       =   &H0095C1CB&
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   480
         MaskColor       =   &H80000010&
         Style           =   1  'Graphical
         TabIndex        =   219
         TabStop         =   0   'False
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton cmdSlew2 
         BackColor       =   &H0095C1CB&
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   960
         MaskColor       =   &H80000010&
         Style           =   1  'Graphical
         TabIndex        =   218
         TabStop         =   0   'False
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdSlew2 
         BackColor       =   &H0095C1CB&
         Caption         =   "W"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   0
         MaskColor       =   &H80000010&
         Style           =   1  'Graphical
         TabIndex        =   217
         TabStop         =   0   'False
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdSlew2 
         BackColor       =   &H0095C1CB&
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   480
         MaskColor       =   &H80000010&
         Style           =   1  'Graphical
         TabIndex        =   216
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Message Center"
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
      Height          =   2415
      Left            =   120
      TabIndex        =   56
      Top             =   1120
      Width           =   2895
      Begin VB.CommandButton Command35 
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
         Left            =   1800
         Picture         =   "HCsmall.frx":D5FC
         Style           =   1  'Graphical
         TabIndex        =   183
         ToolTipText     =   "Color"
         Top             =   2000
         Width           =   375
      End
      Begin VB.CommandButton CommandClearMsg 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         Picture         =   "HCsmall.frx":DB7E
         Style           =   1  'Graphical
         TabIndex        =   166
         Top             =   2000
         Width           =   375
      End
      Begin VB.CheckBox CheckLog2File 
         BackColor       =   &H00000000&
         Caption         =   "Log to File"
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
         Height          =   225
         Left            =   120
         TabIndex        =   157
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox HCMessage 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   1695
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   57
         TabStop         =   0   'False
         Text            =   "HCsmall.frx":E2A4
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame16 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "ASCOM PulseGuide Monitor"
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
      Height          =   5060
      Left            =   120
      TabIndex        =   106
      Top             =   1120
      Visible         =   0   'False
      Width           =   2895
      Begin VB.VScrollBar DECdisplay_gain 
         Height          =   1095
         Left            =   2640
         Max             =   900
         TabIndex        =   123
         Top             =   3000
         Value           =   900
         Width           =   135
      End
      Begin VB.VScrollBar RAdisplay_gain 
         Height          =   1095
         Left            =   2640
         Max             =   900
         TabIndex        =   121
         Top             =   600
         Value           =   900
         Width           =   135
      End
      Begin VB.PictureBox Plot_DEC 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000040&
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   240
         ScaleHeight     =   73
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   161
         TabIndex        =   114
         Top             =   3000
         Width           =   2415
      End
      Begin VB.PictureBox Plot_RA 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000040&
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   240
         ScaleHeight     =   73
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   161
         TabIndex        =   113
         Top             =   600
         Width           =   2415
      End
      Begin VB.HScrollBar HScrollRAWidth 
         Height          =   135
         Left            =   240
         Max             =   200
         Min             =   1
         TabIndex        =   108
         Top             =   2160
         Value           =   100
         Width           =   2415
      End
      Begin VB.HScrollBar HScrollDECWidth 
         Height          =   135
         Left            =   240
         Max             =   200
         Min             =   1
         TabIndex        =   107
         Top             =   4560
         Value           =   100
         Width           =   2415
      End
      Begin VB.Label Label60 
         BackColor       =   &H00000000&
         Caption         =   "S"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   120
         Top             =   3840
         Width           =   255
      End
      Begin VB.Label Label59 
         BackColor       =   &H00000000&
         Caption         =   "N"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   119
         Top             =   3000
         Width           =   255
      End
      Begin VB.Label Label58 
         BackColor       =   &H00000000&
         Caption         =   "W"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   118
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label57 
         BackColor       =   &H00000000&
         Caption         =   "E"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   117
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label56 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "DEC"
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
         Height          =   255
         Left            =   240
         TabIndex        =   116
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Label Label55 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "RA"
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
         Height          =   255
         Left            =   240
         TabIndex        =   115
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label54 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000040&
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2040
         TabIndex        =   112
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label53 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000040&
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2040
         TabIndex        =   111
         Top             =   4320
         Width           =   615
      End
      Begin VB.Label Label51 
         BackColor       =   &H00000000&
         Caption         =   "RA Width Gain"
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
         Left            =   240
         TabIndex        =   110
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label52 
         BackColor       =   &H00000000&
         Caption         =   "DEC Width Gain"
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
         Left            =   240
         TabIndex        =   109
         Top             =   4320
         Width           =   1455
      End
   End
   Begin VB.Frame Frame17 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Polar Alignment"
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
      Height          =   5070
      Left            =   120
      TabIndex        =   126
      Top             =   1120
      Visible         =   0   'False
      Width           =   2895
      Begin VB.HScrollBar HScroll4 
         Height          =   150
         Left            =   1440
         Max             =   5000
         Min             =   100
         TabIndex        =   144
         Top             =   3000
         Value           =   1250
         Width           =   1335
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H0095C1CB&
         Caption         =   "CALCULATE"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   143
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1335
      End
      Begin VB.PictureBox PolarPlot 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000040&
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   120
         ScaleHeight     =   2655
         ScaleWidth      =   2655
         TabIndex        =   131
         Top             =   240
         Width           =   2655
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   150
         Left            =   120
         Max             =   10000
         Min             =   1
         TabIndex        =   130
         Top             =   3000
         Value           =   1000
         Width           =   1215
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H0095C1CB&
         Caption         =   "SHIFT MOUNT"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   129
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1335
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   150
         Left            =   120
         Max             =   90
         TabIndex        =   128
         Top             =   3600
         Value           =   45
         Width           =   1215
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   150
         Left            =   1440
         Max             =   90
         Min             =   -90
         TabIndex        =   127
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label67 
         BackColor       =   &H00000000&
         Caption         =   "Zoom"
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
         Height          =   195
         Left            =   120
         TabIndex        =   137
         Top             =   3120
         Width           =   885
      End
      Begin VB.Label Label74 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "18000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2160
         TabIndex        =   146
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label73 
         BackColor       =   &H00000000&
         Caption         =   "ProbeSize"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   1440
         TabIndex        =   145
         Top             =   3120
         Width           =   765
      End
      Begin VB.Label Label68 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "x1000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   840
         TabIndex        =   136
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label Label71 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "45"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   960
         TabIndex        =   133
         Top             =   3360
         Width           =   375
      End
      Begin VB.Label Label70 
         BackColor       =   &H00000000&
         Caption         =   "DEC"
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
         Height          =   315
         Left            =   1440
         TabIndex        =   134
         Top             =   3360
         Width           =   405
      End
      Begin VB.Label Label62 
         BackColor       =   &H00000080&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   720
         TabIndex        =   142
         Top             =   4320
         Width           =   2055
      End
      Begin VB.Label Label63 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "AZx"
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
         Height          =   195
         Left            =   120
         TabIndex        =   141
         Top             =   4320
         Width           =   435
      End
      Begin VB.Label Label64 
         BackColor       =   &H00000080&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   720
         TabIndex        =   140
         Top             =   4680
         Width           =   2055
      End
      Begin VB.Label Label65 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "ALTy"
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
         Height          =   315
         Left            =   120
         TabIndex        =   139
         Top             =   4680
         Width           =   555
      End
      Begin VB.Label Label66 
         BackColor       =   &H00000000&
         Caption         =   "NCP/SCP Offset (in arcminutes)"
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
         Height          =   195
         Left            =   120
         TabIndex        =   138
         Top             =   4080
         Width           =   2565
      End
      Begin VB.Label Label69 
         BackColor       =   &H00000000&
         Caption         =   "VHorizon"
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
         Height          =   315
         Left            =   120
         TabIndex        =   135
         Top             =   3360
         Width           =   765
      End
      Begin VB.Label Label72 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2160
         TabIndex        =   132
         Top             =   3360
         Width           =   615
      End
   End
   Begin VB.Frame Frame9 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "PEC"
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
      Height          =   5055
      Left            =   120
      TabIndex        =   90
      Top             =   1200
      Width           =   2895
      Begin VB.Frame Frame8 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame8"
         Height          =   300
         Left            =   120
         TabIndex        =   227
         Top             =   2640
         Width           =   2655
         Begin VB.CommandButton CommandPecConfig 
            BackColor       =   &H0095C1CB&
            Height          =   300
            Left            =   2355
            Picture         =   "HCsmall.frx":E2AA
            Style           =   1  'Graphical
            TabIndex        =   234
            ToolTipText     =   "Setup"
            Top             =   0
            Width           =   300
         End
         Begin VB.CommandButton CommandPecPlay 
            Appearance      =   0  'Flat
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
            Height          =   300
            Left            =   0
            Picture         =   "HCsmall.frx":E82C
            Style           =   1  'Graphical
            TabIndex        =   233
            ToolTipText     =   "Play/Pause"
            Top             =   0
            Width           =   300
         End
         Begin VB.CommandButton CmdPecClear 
            BackColor       =   &H0095C1CB&
            Height          =   300
            Left            =   1440
            Picture         =   "HCsmall.frx":EB6E
            Style           =   1  'Graphical
            TabIndex        =   232
            ToolTipText     =   "Unload"
            Top             =   0
            Width           =   300
         End
         Begin VB.CommandButton CmdPecLoad 
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
            Height          =   300
            Left            =   840
            Picture         =   "HCsmall.frx":F294
            Style           =   1  'Graphical
            TabIndex        =   231
            ToolTipText     =   "Load"
            Top             =   0
            Width           =   300
         End
         Begin VB.CommandButton CmdPecSave 
            BackColor       =   &H0095C1CB&
            DisabledPicture =   "HCsmall.frx":F816
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1140
            Picture         =   "HCsmall.frx":FD98
            Style           =   1  'Graphical
            TabIndex        =   230
            ToolTipText     =   "Save"
            Top             =   0
            Width           =   300
         End
         Begin VB.CommandButton CommandRecordPec 
            Appearance      =   0  'Flat
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
            Height          =   300
            Left            =   300
            Picture         =   "HCsmall.frx":1031A
            Style           =   1  'Graphical
            TabIndex        =   229
            ToolTipText     =   "Record/Cancel"
            Top             =   0
            Width           =   300
         End
         Begin VB.CommandButton Command23 
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
            Height          =   300
            Left            =   2055
            Picture         =   "HCsmall.frx":1065C
            Style           =   1  'Graphical
            TabIndex        =   228
            ToolTipText     =   "Timestamp"
            Top             =   0
            Width           =   300
         End
      End
      Begin VB.PictureBox PlotCap 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H000000FF&
         Height          =   1935
         Left            =   120
         ScaleHeight     =   129
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   177
         TabIndex        =   224
         Top             =   3000
         Width           =   2655
      End
      Begin VB.PictureBox plot 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   1245
         Left            =   120
         ScaleHeight     =   83
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   177
         TabIndex        =   91
         Top             =   240
         Width           =   2655
      End
      Begin VB.CheckBox CheckPEC 
         BackColor       =   &H00000000&
         Caption         =   "Apply PEC"
         Enabled         =   0   'False
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   720
         TabIndex        =   225
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox CheckCapPec 
         BackColor       =   &H00000000&
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
         Left            =   720
         TabIndex        =   226
         Top             =   1080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox PlotPecStatus 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         FillStyle       =   0  'Solid
         ForeColor       =   &H000000FF&
         Height          =   1005
         Left            =   120
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   175
         TabIndex        =   235
         Top             =   1560
         Width           =   2655
      End
   End
   Begin VB.Image ImageComms 
      Appearance      =   0  'Flat
      Height          =   720
      Left            =   2760
      Picture         =   "HCsmall.frx":10A12
      Top             =   45
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label MainLabel 
      BackColor       =   &H00000080&
      Caption         =   "<set at runtime>"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   660
      Left            =   960
      TabIndex        =   80
      Top             =   75
      Width           =   2055
   End
   Begin VB.Label Label61 
      BackColor       =   &H00000000&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   122
      Top             =   120
      Width           =   1095
   End
   Begin VB.Menu mnuDisplayPopup 
      Caption         =   "DisplayPopup"
      Visible         =   0   'False
      Begin VB.Menu puPosition 
         Caption         =   "Mount Position"
         Tag             =   "0"
      End
      Begin VB.Menu puDials 
         Caption         =   "Axis Position"
         Tag             =   "1"
      End
      Begin VB.Menu puMessageCenter 
         Caption         =   "Message Center"
         Tag             =   "2"
      End
      Begin VB.Menu puPEC 
         Caption         =   "PEC"
         Tag             =   "3"
      End
      Begin VB.Menu puSlew 
         Caption         =   "Slew"
         Tag             =   "5"
      End
      Begin VB.Menu puPulse 
         Caption         =   "Pulse Guide"
         Tag             =   "4"
      End
      Begin VB.Menu puPolar 
         Caption         =   "Polar Alignment"
         Tag             =   "6"
      End
   End
End
Attribute VB_Name = "HC"
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
' HC.frm - Main ASCOM EQMOD Handcontrol form
'
'
' Written:  07-Oct-06   Raymund Sarmiento
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' 04-Nov-06 rcs     Initial edit for EQ Mount Driver Function Prototype
' 20-Nov-06 rcs     Introduced EqNow_lst() to have millisecond granularity
'                   Changed GOTO TIMER interval from 500 to 100 milliseconds
'                   Changes EncoderTimer interval from 500 to 400 milliseconds
' 25-Nov-06 rcs     Add Regional Setting Functions
' 01-Dec-06 rcs     Add GPS Function and Autoguider port save
' 01-Dec-06 rcs     Fix Autoguider port restore bug
' 17-Dec-06 rcs     Added Joystick timer and Joystick Activate Button
' 16-Nov-07 rcs     Added Auto enable sidreal tracking (if disabled) right after a goto
' 19-Mar-07 rcs     Initial Edit for Three star alignment
' 03-Apr-07 rcs     Add RA Slew Limits
' 08-Apr-07 rcs     N-star implementation
' 11-Apr-07 rcs     Add Color Pick
' 25-Jul-07 cs      Added Joystick config btn. Start joystick on form load
' 29-Jul-07 cs      Joystic calib btn. Slew preset combo. Slew presets read on form load
' 30-Jul-07 cs      "sidreal" change to "sidereal"
' 13-Aug-07 cs      Language dll stuff
' 14-Aug-07 cs      Sites list
' 30-Aug-07 cs      Joystick Calibration button removed (calibration no part of jstick config)
'                   On error handling for form resizing via setup button
' 05-Oct-07 cs      Read/Write custom rate on form Load/Unload
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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const CURRENT_VERSION = "V1.27i"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public SiteIdx As Integer
Public EncoderReadErrCount As Integer
Public EncoderTimerFlag As Boolean
Public DisplayMode As Integer
Public oPersist As New Persist

Private m_scope As Object
Private OldHeight As Integer
Private OldScaleHeight As Integer
Private flash As Boolean
Private logcount As Integer
Private logfileindex As Integer
Private lblmode As Integer
Private ScrollFlag As Boolean
Private AutoParkDuration As Long
Private hUtil As DriverHelper.Util
Private gIgnorClick As Boolean
Private JoystickLost As Boolean

Private Sub AutoParkTimer_Timer()
    
    If gEQparkstatus = 1 Then
        ' already parked - cancel autopark
        CmdParkTimer_Click
        AutoParkTimer.Enabled = False
        AutoParkDuration = 0
        TextAutoParkDuration.Text = "0"
        TextAutoParkDuration.BackColor = &H80&
        Exit Sub
    End If
    
    AutoParkDuration = AutoParkDuration - 1
    If AutoParkDuration <= 0 Then
        CommandPark_Click
        AutoParkTimer.Enabled = False
        AutoParkDuration = 0
        TextAutoParkDuration.BackColor = &H80&
    Else
    End If
    
    TextAutoParkDuration.Text = CStr(CInt(AutoParkDuration / 60))
End Sub

Private Sub cbNS_Change()
    cbhem.ListIndex = cbNS.ListIndex
End Sub

Private Sub CheckCapPec_Click()
    If CheckCapPec.Value = 1 Then
        Call PEC_StartCapture
        CommandRecordPec.Picture = LoadResPicture(107, vbResBitmap)
    Else
        Call PEC_StopCapture
        CommandRecordPec.Picture = LoadResPicture(106, vbResBitmap)
    End If
End Sub

Private Sub CheckLocalPier_Click()
    HC.oPersist.WriteIniValue "ALIGN_LOCALTOPIER", CStr(CheckLocalPier.Value)
End Sub

Private Sub CheckLog2File_Click()
    On Error Resume Next
    Close #4
    If CheckLog2File.Value = 1 Then
        Open oPersist.GetIniPath & "\messagelog1.txt" For Output As #4
        logfileindex = 1
        logcount = 0
    End If
End Sub

Private Sub CheckPEC_Click()
    PEC_OnUse
End Sub

Private Sub CheckRASync_Click()
    writeRASyncCheckVal
End Sub


Public Sub Change_Display(displayid As Integer)
        
    On Error Resume Next
    Frame2.Visible = False
    PositionFrame.Visible = False
    Frame9.Visible = False
    Frame16.Visible = False
    Frame17.Visible = False
    FrameAxis.Visible = False
    SlewFrame.Visible = True
    Slew2Frame.Visible = False
    HC.Height = 8640
    HC.ScaleHeight = 8130

    Select Case displayid
        Case 0:
            ' postion
            PositionFrame.Visible = True
        Case 1:
            ' Axis dials
            FrameAxis.Visible = True
        Case 2:
            ' message centre
            Frame2.Visible = True
        Case 3:
            ' PEC
            Frame9.Visible = True
            SlewFrame.Visible = False
        Case 4:
            ' pulse guide
            Frame16.Visible = True
            SlewFrame.Visible = False
        Case 5:
            ' position - short
            Slew2Frame.Visible = True
            HC.Height = 3285
            HC.ScaleHeight = 2775
        Case 6:
            ' Polar alignment
            Frame17.Visible = True
            SlewFrame.Visible = False

    End Select
    
End Sub



Private Sub ChkEnableLimits_Click()
    Call oPersist.WriteIniValue("LIMIT_ENABLE", ChkEnableLimits.Value)
End Sub

Private Sub ChkForceFlip_Click()
    If ChkForceFlip.Value = 1 Then
        gCWUP = True
    Else
        gCWUP = False
    End If
End Sub

Private Sub CmdDisplayMode_Click()
'    If gShowPolarAlign = 1 Then
'        DisplayMode = (DisplayMode + 1) Mod 6
'    Else
'        DisplayMode = (DisplayMode + 1) Mod 5
'    End If
'    Change_Display (DisplayMode)
End Sub

Private Sub CmdDisplayMode_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Select Case Button
        Case vbLeftButton
        
            Select Case DisplayMode
                Case 6
                    DisplayMode = 0
                Case 5
                    If gShowPolarAlign = 1 Then
                        DisplayMode = 6
                    Else
                        DisplayMode = 0
                    End If
                Case 3
                    If gAscomCompatibility.AllowPulseGuide Then
                        DisplayMode = 4
                    Else
                        DisplayMode = 5
                    End If
                Case Else
                    DisplayMode = DisplayMode + 1
            End Select
            
'            If gShowPolarAlign = 1 Then
'                DisplayMode = (DisplayMode + 1) Mod 7
'            Else
'                DisplayMode = (DisplayMode + 1) Mod 6
'            End If
            Change_Display (DisplayMode)
        
        Case vbRightButton
            Me.PopupMenu mnuDisplayPopup
    End Select
End Sub

Private Sub CmdEditLimits_Click()
    Limits_edit
End Sub

Private Sub CmdParkTimer_Click()
    On Error Resume Next
    If AutoParkDuration > 0 Then
        ' already running - stop it
        AutoParkTimer = False
        AutoParkDuration = 0
        TextAutoParkDuration.Text = "0"
        TextAutoParkDuration.BackColor = &H80&
    Else
         AutoParkDuration = 60 * val(TextAutoParkDuration.Text)
         If AutoParkDuration > 0 Then
             AutoParkTimer = True
             TextAutoParkDuration.BackColor = &H4000&
         Else
             AutoParkTimer = False
             AutoParkDuration = 0
             TextAutoParkDuration.Text = "0"
             TextAutoParkDuration.BackColor = &H80&
        End If
    End If
End Sub

Private Sub CmdPecLoad_Click()
    ' load up PEC file
    PEC_Load
    ' bring up pec config form
    PECConfigFrm.Show (0)
End Sub

Private Sub CmdPecSave_Click()
    PEC_Save
End Sub

Private Sub cmdSlew_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Select Case Index
        Case 0
            'N'
            Call North_Down(val(VScrollDecSlewRate.Value))
            Add_Message (oLangDll.GetLangString(5001))
        Case 1
            'E'
            Call East_Down(val(VScrollRASlewRate.Value))
            Add_Message (oLangDll.GetLangString(5009))
        Case 2
            'S'
            Call South_Down(val(VScrollDecSlewRate.Value))
            Add_Message (oLangDll.GetLangString(5011))
        Case 3
            'W'
            Call West_Down(val(VScrollRASlewRate.Value))
            Add_Message (oLangDll.GetLangString(5010))
        Case 4
            'Stop'
            'Call emergency_stop
            Call EmergencyStopPark
        Case 5
            ' spiral search
            Call Spiral_Slew
            Add_Message (oLangDll.GetLangString(5005))
            
            
    End Select
End Sub

Private Sub cmdSlew_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Select Case Index
        Case 0
            'N'
            Call North_Up
        Case 1
            'E'
            Call East_Up
        Case 2
            'S'
            Call South_Up
        Case 3
            'W'
            Call West_Up
        Case 4
            'Stop'
        Case 5
            ' Spiral search
            Call Spiral_Slew_Stop

    End Select
End Sub

Private Sub cmdSlew2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call cmdSlew_MouseDown(Index, Button, Shift, x, Y)
End Sub

Private Sub cmdSlew2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call cmdSlew_MouseUp(Index, Button, Shift, x, Y)
End Sub



Private Sub Combo3PointAlgorithm_Click()
    g3PointAlgorithm = Combo3PointAlgorithm.ListIndex
    HC.oPersist.WriteIniValue "3POINT_ALGORITHM", CStr(g3PointAlgorithm)
End Sub


Private Sub ComboActivePoints_Click()
    gPointFilter = HC.ComboActivePoints.ListIndex
    HC.oPersist.WriteIniValue "ALIGN_SELECTION", CStr(gPointFilter)
End Sub


Private Sub ComboPoleStar_Click()
    gPoleStarIdx = ComboPoleStar.ListIndex
    writePoleStar
End Sub

Private Sub ComboUnPark_Click()
    If ComboUnPark.ListIndex >= 2 Then
        If UserUnparks(ComboUnPark.ListIndex - 1).name = oLangDll.GetLangString(2730) Then
            readParkModes
        End If
    End If
    writeParkMode_unpark
    On Error Resume Next
    DummyChk.SetFocus
    Call SetParkCaption
End Sub

Private Sub ComboPark_Click()
    If ComboPark.ListIndex >= 2 Then
        If UserParks(ComboPark.ListIndex - 1).name = oLangDll.GetLangString(2730) Then
            readParkModes
        End If
    End If
    writeParkMode_park
    On Error Resume Next
    DummyChk.SetFocus
    Call SetParkCaption
End Sub


Private Sub CommandLoadTracking_Click()
    Call LoadTrackingRates
End Sub

Private Sub CommandPecConfig_Click()
    PECConfigFrm.Show (0)
End Sub

Private Sub CommandPecPlay_Click()
    If CheckPEC.Value = 1 Then
        CheckPEC.Value = 0
    Else
        CheckPEC.Value = 1
    End If
End Sub

Private Sub CommandRecordPec_Click()
    If CheckCapPec.Value = 1 Then
        CheckCapPec.Value = 0
        CommandRecordPec.Picture = LoadResPicture(106, vbResBitmap)
    Else
        CheckCapPec.Value = 1
        CommandRecordPec.Picture = LoadResPicture(107, vbResBitmap)
    End If
    
End Sub

Private Sub CommandRemoveTrackFile_Click()
    LabelTrackFile.Caption = ""
    HC.LabelTrackFile.ToolTipText = ""
    gCustomTrackFile = ""
    CmdTrack(4).Picture = LoadResPicture(110, vbResBitmap)
    CmdTrack(4).ToolTipText = oLangDll.GetLangString(189)
End Sub

Private Sub CommandSetup_Click()
    On Error Resume Next
    If CommandSetup.ToolTipText = oLangDll.GetLangString(90) Then
        CommandSetup.ToolTipText = oLangDll.GetLangString(91)
        CommandSetup.Picture = LoadResPicture(102, vbResBitmap)
        HC.ScaleWidth = 10680
        HC.width = 10800
        OldHeight = HC.Height
        OldScaleHeight = HC.ScaleHeight
        HC.Height = 8640
        HC.ScaleHeight = 8130
       
    Else
        CommandSetup.ToolTipText = oLangDll.GetLangString(90)
        CommandSetup.Picture = LoadResPicture(101, vbResBitmap)
        HC.ScaleWidth = 3090
        HC.width = 3210
        HC.Height = OldHeight
        HC.ScaleHeight = OldScaleHeight
    End If
End Sub

Private Sub CmdClearSync_Click()
    Call resetsync
End Sub

Private Sub CommandSetSite_Click()
    Dim deg As Double
    Dim min As Double
    Dim sec As Double
    
    deg = CDbl(EQFixNum(txtLongDeg))
    If deg > 180 Then GoTo inputerror
    min = CDbl(EQFixNum(txtLongMin))
    If min >= 60 Then GoTo inputerror
    sec = CDbl(EQFixNum(txtLongSec))
    If sec > 60 Then GoTo inputerror
    
    gLongitude = deg + (min / 60#) + (sec / 3600#)
    If cbEW.Text = oLangDll.GetLangString(115) Then gLongitude = -gLongitude     ' W is neg
    
    deg = CDbl(EQFixNum(txtLatDeg))
    If deg >= 90 Then GoTo inputerror
    min = CDbl(EQFixNum(txtLatMin))
    If min >= 60 Then GoTo inputerror
    sec = CDbl(EQFixNum(txtLatSec))
    If sec >= 60 Then GoTo inputerror
    
    gLatitude = deg + (min / 60#) + (sec / 3600#)
    
    If cbNS.Text = oLangDll.GetLangString(116) Then
        gLatitude = -gLatitude
        cbhem.ListIndex = 1
    Else
        cbhem.ListIndex = 0
    End If
    gElevation = CDbl(EQFixNum(txtElevation))
    
    If cbhem.Text = oLangDll.GetLangString(1110) Then
        gHemisphere = 0
    Else
        gHemisphere = 1
    End If
    
    Call WriteSiteValues
    Exit Sub

inputerror:
    MsgBox (oLangDll.GetLangString(6014))

    
End Sub


Private Sub Command13_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        'right click - browse for exe
        Call SetUtilityApp(Index)
    Else
        If Button = 1 Then
            ' left click - launch/restore application
            Call LaunchUtilityApp(Index)
        End If
    End If

End Sub

Private Sub Command15_Click()

Dim vh As Double
Dim vy As Double
Dim RA1 As Double
Dim DEC1 As Double
Dim RA2 As Double
Dim DEC2 As Double
Dim obtmp As Coord
Dim raprobe As Double


 '   obtmp = PolarAlign_Map(RAEncoder_Home_pos, gDECEncoder_Home_pos)
 
    vh = HScroll2.Value
    vh = (vh / 360) * gTot_step
    
    vy = 90 + HScroll3.Value
    vy = (vy / 360) * gTot_step
    
    raprobe = HC.HScroll4.Value
    raprobe = raprobe * 100
    
    RA1 = RAEncoder_Home_pos + vh
    DEC1 = gDECEncoder_Home_pos - vy
    RA2 = RAEncoder_Home_pos + vh - (gTot_step / 4)
    DEC2 = gDECEncoder_Home_pos + vy
    
    On Error GoTo End01
    
    obtmp = PolarAlignDrift_Map(RA1, DEC1, RA2, DEC2, raprobe, HScroll1.Value)
    Label62.Caption = Format(obtmp.x * 0.0024, "####0.0000000000")   '.144 * 60
    Label64.Caption = Format(obtmp.Y * 0.0024, "####0.0000000000")
    gPolarAlign_RA = obtmp.x
    gPolarAlign_DEC = obtmp.Y
    ScrollFlag = False
    
    PolarAlign_init (HScroll1.Value)
    Call Plot_PolarAlign(gPolarAlign_RA, gPolarAlign_DEC, HScroll1.Value)
End01:

End Sub

Private Sub CommandAddPoint_Click()
    Load Align
    Align.Show
End Sub

Private Sub CommandClearMsg_Click()
    HCMessage.Text = ""
End Sub

Private Sub CommandDefinePark_Click()
    DefineParkForm.Show (1)
End Sub

Private Sub CmdPecClear_Click()
    PEC_Clear
End Sub

Private Sub Command23_Click()
' timestamp
    Call PEC_Timestamp
End Sub

Private Sub Command35_Click()
    Load ColorPick
    ColorPick.Show
End Sub

Private Sub CommandLoadSite_Click()
    'site load
    LoadSite (SiteIdx)
    ' and set
    Call CommandSetSite_Click
End Sub

Private Sub CommandSaveSite_Click()
    'site save
    HC.SitesCombo.List(SiteIdx) = HC.SitesCombo.Text
    HC.SitesCombo.ListIndex = SiteIdx
    Call SaveSite(SiteIdx, HC.SitesCombo.Text)
End Sub

Private Sub CommandPark_Click()
    If gEQparkstatus = 1 Then
        ' if parked then unpark
        Call ApplyUnParkMode
    Else
        If gEQparkstatus = 0 Then
            'If not parked then park
            Call ApplyParkMode
        End If
    End If

End Sub
Public Sub ApplyUnParkMode()
    ApplyUnParkMode2 (ComboUnPark.ListIndex)
End Sub

Public Sub ApplyParkMode()
    ApplyParkMode2 (ComboPark.ListIndex)
End Sub

Public Sub SyncEncoderToPark()
    
    If MsgBox(oLangDll.GetLangString(61) & vbLf & vbLf & oLangDll.GetLangString(62), vbOKCancel, oLangDll.GetLangString(60)) = vbOK Then
        If EQ_GetMountStatus() = 1 Then     ' mount must be is online
            If gEQparkstatus = 0 Then
                'makes no sense to sync if parked!
                ' stop tracking
                StopTrackingUpdates
                eqres = EQ_MotorStop(0)
                eqres = EQ_MotorStop(1)
            
                Select Case ComboPark.ListIndex
                    Case 0
                        ' Sync to Home
                        eqres = EQSetMotorValues(0, CLng(RAEncoder_Home_pos))
                        eqres = EQSetMotorValues(1, CLng(gDECEncoder_Home_pos))
                    Case 1
                        ' Sync to current
                        ' don't need to do anything!
                    Case Else
                        'Sync to defined
                        With UserParks(ComboPark.ListIndex - 1)
                            ' don't sync if user position is undefined!
                            If .posD <> 0 And .posR <> 0 Then
                                ' sync encoder values
                                eqres = EQSetMotorValues(0, .posR)
                                eqres = EQSetMotorValues(1, .posD)
                            Else
                            
                            End If
                        End With
                End Select
            End If
        End If
    End If
End Sub

Private Sub CommandPolaris_Click()
    polarfrm.Show (0)

End Sub

Private Sub CommandSounds_Click()
    SoundsFrm.Show (1)
End Sub

Private Sub CommandSyncEncoder_Click()
    Call SyncEncoderToPark
End Sub

Private Sub CommandUpdate_Click()
    Select Case gUpdateMode
        Case 1
            OpenBrowser (gUpdateFullUrl)
        Case 2
            OpenBrowser (gUpdateTestUrl)
    End Select

End Sub



Private Sub CustomTrackTimer_Timer()
    Call TrackTimer
End Sub

Private Sub DEC_Inv_Click()
    If DEC_Inv.Value = 1 Then
        EQ_Beep (42)
    Else
        EQ_Beep (43)
    End If
    Reverse(1) = DEC_Inv
   Call writeAxisRevDEC
End Sub

Private Sub declbl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        GotoDialog.Show (1)
    End If

End Sub



Private Sub decmrk_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        GotoDialog.Show (1)
    End If
End Sub

Private Sub HScrollBacklashDec_Change()
    gBacklashDec = HScrollBacklashDec.Value
    LabelBacklashDec.Caption = CStr(gBacklashDec)
    HC.oPersist.WriteIniValue "DEC_BACKLASH", CStr(gBacklashDec)
End Sub

Private Sub HScrollBacklashDec_Scroll()
    HScrollBacklashDec_Change
End Sub

Private Sub HScrollGotoRes_Change()
    gGotoResolution = HScrollGotoRes.Value
    LabelGotoRes.Caption = CStr(gGotoResolution)
    HC.oPersist.WriteIniValue "GOTO_RESOLUTION", CStr(gGotoResolution)
End Sub

Private Sub HScrollGotoRes_Scroll()
    Call HScrollGotoRes_Change
End Sub

Private Sub HScrollProximity_Change()
    LabelProximity.Caption = CStr(HScrollProximity.Value) & " Deg"
    CalcPromximityLimits (HScrollProximity.Value)
    writeAlignProximity
End Sub

Private Sub HScrollProximity_Scroll()
    Call HScrollProximity_Change
End Sub

Private Sub HScrollSlewAdjust_Change()
    gRA_Compensate = HScrollSlewAdjust.Value
    LabelSlewAdjust.Caption = CStr(gRA_Compensate)
    HC.oPersist.WriteIniValue "GOTO_RA_COMPENSATE", CStr(gRA_Compensate)
End Sub

Private Sub HScrollSlewAdjust_Scroll()
    Call HScrollSlewAdjust_Change
End Sub

Private Sub HScrollSlewLimit_Change()
    If HScrollSlewLimit.Value = HScrollSlewLimit.min Then
        gGotoRate = 0
    Else
        gGotoRate = HScrollSlewLimit.Value
    End If
    If gGotoRate = 0 Then
        LabelSlewLimit.Caption = oLangDll.GetLangString(183)
    Else
        LabelSlewLimit.Caption = CStr(gGotoRate)
    End If
    Call writeGotoRate

End Sub

Private Sub HScrollSlewLimit_Scroll()
    Call HScrollSlewLimit_Change
End Sub

Private Sub HScrollSlewRetries_Change()
    gMaxSlewCount = HScrollSlewRetries.Value
    LabelSlewRetries.Caption = CStr(gMaxSlewCount)
    HC.oPersist.WriteIniValue "MAX_GOTO_INTERATIONS", CStr(gMaxSlewCount)
End Sub

Private Sub HScrollSlewRetries_Scroll()
    Call HScrollSlewRetries_Change
End Sub

Private Sub LabelUpdates_Click()
    Select Case gUpdateMode
        Case 1
            OpenBrowser (gUpdateFullUrl)
        Case 2
            OpenBrowser (gUpdateTestUrl)
    End Select
End Sub

Private Sub Label82_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblmode = Button
End Sub
Private Sub Label83_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblmode = Button
End Sub


Private Sub PECCapTimer_Timer()
    Call PEC_CaptureTimer
End Sub




Private Sub PresetRate2Combo_Click()
    On Error Resume Next
    If gIgnorClick = False Then
        gIgnorClick = True
        PresetRateCombo.ListIndex = PresetRate2Combo.ListIndex
        gIgnorClick = False
    End If

End Sub

Private Sub puDials_Click()
    Change_Display (puDials.Tag)
End Sub

Private Sub puMessageCenter_Click()
    Change_Display (puMessageCenter.Tag)
End Sub

Private Sub puPEC_Click()
    Change_Display (puPEC.Tag)
End Sub

Private Sub puPolar_Click()
    Change_Display (puPolar.Tag)
End Sub

Private Sub puPosition_Click()
    Change_Display (puPosition.Tag)
End Sub

Private Sub puPulse_Click()
    Change_Display (puPulse.Tag)
End Sub

Private Sub puSlew_Click()
    Change_Display (puSlew.Tag)
End Sub

Private Sub ralbl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        GotoDialog.Show (1)
    End If
End Sub
Private Sub DriftScroll_Change()
    gDriftComp = DriftScroll.Value
    Driftlbl = CStr(gDriftComp)
    ' readCustomMount
    Call EQSetOffsets
    'write the new value to the ini file
    writeDriftVal
End Sub

Private Sub DriftScroll_Scroll()
    DriftScroll_Change
End Sub

Private Sub Edit_Stars_Command_Click()
    Load StarEditform
    StarEditform.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call writeCustomRa
    Call PEC_Unload
    Call WriteFormPosition
    Close #4
End Sub

Private Sub HCOnTop_Click()
    Call writeOnTop
    If HC.HCOnTop.Value = 1 Then
        Call PutWindowOnTop(HC)
    Else
        Call PutWindowNormal(HC)
    End If
End Sub

Private Sub HScroll1_Change()
    If ScrollFlag = False Then
        PolarAlign_init (HScroll1.Value)
        Call Plot_PolarAlign(gPolarAlign_RA, gPolarAlign_DEC, HScroll1.Value)
    Else
         Call Position_polar(HScroll1.Value)
    End If
    Label68.Caption = "x" & str(HScroll1.Value)
    
End Sub

Private Sub HScroll1_Scroll()
    Call HScroll1_Change
End Sub

Private Sub HScroll2_Change()
    HC.Label71.Caption = HC.HScroll2.Value
    ScrollFlag = True
    Call Position_polar(HScroll1.Value)
End Sub

Private Sub HScroll2_Scroll()
    Call HScroll2_Change
End Sub

Private Sub HScroll3_Change()
    Label72.Caption = HScroll3.Value
    ScrollFlag = True
    Call Position_polar(HScroll1.Value)
End Sub

Private Sub HScroll3_Scroll()
    Call HScroll3_Change
End Sub

Private Sub HScroll4_Change()
Dim x As Double

    x = HScroll4.Value
    x = 0.144 * x * 100
    Label74.Caption = Format$(x, "0000")
    
    ScrollFlag = True
    Call Position_polar(HScroll1.Value)
End Sub


Private Sub HScroll4_Scroll()
    Call HScroll4_Change
End Sub

Private Sub HScrollDECWidth_Change()
    HC.Label53.Caption = Format$(str(HScrollDECWidth.Value), "000") & "%"
End Sub

Private Sub HScrollDECWidth_Scroll()
    Call HScrollDECWidth_Change
End Sub

Private Sub HScrollRAWidth_Change()
    HC.Label54.Caption = Format$(str(HScrollRAWidth.Value), "000") & "%"
End Sub

Private Sub HScrollRAWidth_Scroll()
    Call HScrollRAWidth_Change
End Sub

Private Sub JStickSetupCommand_Click()
End Sub
Private Sub CmdGamepadInit_Click()
    HC.Add_Message (oLangDll.GetLangString(5012))
    HC.JoyStick_Timer.Enabled = True
End Sub

Private Sub CommandGamepad_Click(Index As Integer)
    Select Case Index
        Case 0
            Load JStickConfigForm
            JStickConfigForm.Show
        Case 1
            HC.Add_Message (oLangDll.GetLangString(5012))
            HC.JoyStick_Timer.Enabled = True
    End Select

End Sub


Private Sub ListAlignMode_Click()

    Select Case ListAlignMode.ListIndex
        Case 0
            ' 3-Point+nearset
            gAlignmentMode = 0
            Combo3PointAlgorithm.Enabled = True
        Case 1
            ' nearest
            gAlignmentMode = 2
            gAffine1 = 0
            gAffine2 = 0
            gAffine3 = 0
            Combo3PointAlgorithm.Enabled = False

    End Select

    Call writeAlignCheck1
    On Error Resume Next
    DummyChk.SetFocus
End Sub

Private Sub ListSyncMode_Click()
    Call writeAlignCheck2
    On Error Resume Next
    DummyChk.SetFocus
End Sub


Private Sub ParkTimer_Timer()
    If gParkTimerFlag = True Then
        gParkTimerFlag = False  ' Avoid Overruns
        Call ManagePark
        gParkTimerFlag = True
    End If
End Sub

Private Sub PECTimer_Timer()
    Call PEC_Timer
End Sub

Private Sub PhaseScroll_Change()
    Call PEC_PhaseScroll_Change
End Sub

Private Sub PhaseScroll_Scroll()
    Call PEC_PhaseScroll_Change
End Sub

Private Sub GainScroll_Change()
    Call PEC_GainScroll_Change
End Sub

Private Sub GainScroll_Scroll()
    GainScroll_Change
End Sub

Private Sub picASCOM_Click()
    AscomTrace.Show (0)
End Sub

Private Sub PltimerHscroll_Change()
    HC.Label40.Caption = Format$(str(PltimerHscroll.Value), " 00")
    HC.Pulseguide_Timer.Interval = val(HC.PltimerHscroll.Value)
    gpl_interval = val(HC.PltimerHscroll.Value)
    writePulseguidepwidth
End Sub

Private Sub PltimerHscroll_Scroll()
    Call PltimerHscroll_Change
    End Sub

Private Sub PolarEnable_Click()
    Call ActivateMatrix
End Sub

Private Sub PolarPlot_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
    
        gXmouse = (gXshift - x) * -1
        gYmouse = (gYshift - Y) * -1
        ScrollFlag = True
        Call Position_polar(HScroll1.Value)
    End If
End Sub

Private Sub PolarPlot_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If (Button = 1) Then
        gXshift = (gXmouse - x) * -1
        gYshift = (gYmouse - Y) * -1
        ScrollFlag = True
        Call Position_polar(HScroll1.Value)
    End If

End Sub

Private Sub PresetRateCombo_Click()
Dim rate As Double
    
    gCurrentRatePreset = PresetRateCombo.ListIndex + 1
    rate = gPresetSlewRates(gCurrentRatePreset)
    If rate > 0 Then
        If rate >= 1 Then
            rate = rate + 9
        Else
            rate = rate * 10
        End If
        
        VScrollRASlewRate.Value = CInt(rate)
        VScrollDecSlewRate.Value = CInt(rate)
    End If
    
    If gIgnorClick = False Then
        gIgnorClick = True
        PresetRate2Combo.ListIndex = PresetRateCombo.ListIndex
        gIgnorClick = False
    End If
    
    On Error Resume Next
    DummyChk.SetFocus

End Sub

Private Sub RA_inv_Click()
    If RA_inv.Value = 1 Then
        EQ_Beep (40)
    Else
        EQ_Beep (41)
    End If
    Reverse(0) = RA_inv
    Call writeAxisRevRA
End Sub


Private Sub ramrk_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        GotoDialog.Show (1)
    End If
End Sub

Private Sub Reset_Align_Command_Click()

     gThreeStarEnable = False
     gAlignmentStars_count = 0

     gRA1Star = 0
     gDEC1Star = 0
     
     WriteAlignMap
     Call resetsync
     
     AlignmentStars(1).EncoderRA = 0
     AlignmentStars(1).EncoderDEC = 0
     AlignmentStars(1).TargetRA = 0
     AlignmentStars(1).TargetDEC = 0

     AlignmentStars(2).EncoderRA = 0
     AlignmentStars(2).EncoderDEC = 0
     AlignmentStars(2).TargetRA = 0
     AlignmentStars(2).TargetDEC = 0
     
     AlignmentStars(3).EncoderRA = 0
     AlignmentStars(3).EncoderDEC = 0
     AlignmentStars(3).TargetRA = 0
     AlignmentStars(3).TargetDEC = 0
     
End Sub

Private Sub RAGuideRateList_Click()
    If RAGuideRateList.ListIndex = 4 Then
    Else
        eqres = EQ_SetAutoguiderPortRate(0, RAGuideRateList.ListIndex)
        If eqres <> EQ_OK Then
            HC.Add_Message (oLangDll.GetLangString(5008))
        End If
    End If
    Call writeportrateRa(HC.RAGuideRateList.Text)
    On Error Resume Next
    DummyChk.SetFocus
End Sub

Private Sub DECGuideRateList_Click()
    If DECGuideRateList.ListIndex = 4 Then
    Else
        eqres = EQ_SetAutoguiderPortRate(1, DECGuideRateList.ListIndex)
        If eqres <> EQ_OK Then
            HC.Add_Message (oLangDll.GetLangString(5008))
        End If
    End If
    Call writeportrateDec(HC.DECGuideRateList.Text)
    On Error Resume Next
    DummyChk.SetFocus
End Sub



Private Sub CmdSlewPad_Click()
    Load Slewpad
    Slewpad.Show
End Sub

Private Sub Commandgps_Click()
    Load GPSSetup
    GPSSetup.Show
End Sub



Private Sub CmdTrack_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    ' tracking button handler
    Dim RA As Double
    Dim DEC As Double


    CustomTrackTimer.Enabled = False

    Select Case Index
        Case 0
            'sidereal
            EQStartSidereal2
        Case 1
            ' sidereal + pec
            Call EQStartSidereal2
        Case 2
            ' lunar
            Call Start_Lunar
        Case 3
            ' Solar
            Call Start_Solar
        Case 4
            ' Custom rate
            Call Start_CustomTracking2
            If Button = 2 Then
                If gCustomTrackFile <> "" Then
                    If GetTrackTarget(RA, DEC) = True Then
                        Call goto_TrackTarget(RA, DEC, False)
                    End If
                Else
                End If
            End If
        Case 5
            ' stop tracking
            Call emergency_stop
    End Select


End Sub

Private Sub Form_Load()
    Dim tmptxt As String
    gVersion = CURRENT_VERSION
    gMonitorState = 1

    ScrollFlag = False
    EnableCloseButton Me.hwnd, False
    HC.ScaleWidth = 3090
    HC.width = 3210
    Call ReadFormPosition
    
    
    gPolarAlign_RA = 0
    gPolarAlign_DEC = 0
    EncoderReadErrCount = 0
    lblmode = 1
    
    gXshift = 0
    gYshift = 0
    gXmouse = 0
    gYmouse = 0
    
    Call PlotInit
    
    Call PolarAlign_init(1000)
    
    Call LoadLanguageDll
    Call SetText
    Call ReadSiteValues

    Call readDevelopmentOptions
    Call LoadSites(HC.SitesCombo)
    Call readOnTop
    Call readBeep
    Call readCustomRa
    Call readGotoRate
    Call readFlipGoto
    Call readAxisRev
    Call readAlignCheck
    Call ReadParkOptions
    Call readPoleStar
   
    If HC.HCOnTop.Value = 1 Then Call PutWindowOnTop(HC)
    
    Set hUtil = New DriverHelper.Util
    
    readratebarstateHC
    Call readportrate ' Just call this to update the form (mount may not be connected at this point)
                      ' This will be called again after successful mount communication
                 
     
    EncoderTimerFlag = True
    gJoyTimerFlag = True
    gSpiralTimerFlag = True
    gParkTimerFlag = True
    
    Call readColorDat
    
    Call readPresetSlewRates
    
    ' read joystick definitions
    Call LoadJoystickBtns
    ' read joystick calibration
    Call LoadJoystickCalib
    ' see if there is a joystick out there
    JoyStick_Timer.Enabled = True
    HC.Add_Message (oLangDll.GetLangString(5012))
    
    ' show scope status display
    Change_Display (0)
    
    If gAscomCompatibility.AllowPulseGuide Then
        Frame5.Enabled = True
    Else
        Frame5.Enabled = False
    End If
    
    Call CheckForUpdate
    
    If gUpdateAvailable Then
        HC.CommandUpdate.Visible = True
        HC.CommandUpdate.ToolTipText = gStrUpdateVersion & " " & oLangDll.GetLangString(6016)
    Else
        HC.CommandUpdate.Visible = False
    End If
    
    HC.DisplayTimer.Enabled = True

End Sub

Private Sub EncoderTimer_Timer()

Dim tmpcoord As Coordt
Dim tRA As Double
Dim tDEC As Double
Dim tAlt As Double
Dim tAz As Double
Dim tmpRa, tmpDec As Long

    'Avoid overruns
    If EncoderTimerFlag Then
        EncoderTimerFlag = False
        If (gEmulOneShot = True) Or (gEmulNudge = True) Or (gSlewStatus = True) Or (HC.CheckRASync.Value = 1) Then
            ' Read true motor positions
            tmpRa = EQGetMotorValues(0)
            tmpDec = EQGetMotorValues(1)
            If tmpRa > 16777215 Or tmpDec > 16777215 Then
                ' error reading encoder
                ' use an emulated RA motor position , DEC position will stay as was
                gEmulRA = GetEmulRA()
                
                If EncoderReadErrCount < 5 Then
                    ' keep a count of sucessive failures
                    EncoderReadErrCount = EncoderReadErrCount + 1
                    Call Add_Message("CommErr:ER=" & CStr(tmpRa))
                    Call Add_Message("CommErr:ED=" & CStr(tmpDec))
                    ImageComms.Picture = LoadResPicture(104, vbResBitmap)
                Else
                    ImageComms.Picture = LoadResPicture(105, vbResBitmap)
                    gInitResult = 3
                    If gCommErrorStop <> 0 Then
                        'we've had 5 or more succesive fails so try an emergency stop - it may or may not work but at least we tried.
                        Call Add_Message("CommErr: Stop")
                        Call emergency_stop
                    End If
                    ''''''''''''
                    ' reboot!
                    ''''''''''''
                    ' Close com port
                    EQ_End
                    ' Re-initialise dll (including comports)
                    gInitResult = EQ_Init("\\.\" & gPort, gBaud, gTimeout, gRetry)
                    ' Re-assert tracking offests
                    Call EQSetOffsets
                End If
            Else
                If EncoderReadErrCount >= 5 Then
                    'Comms are now working, but they weren't last time round so
                    'send out one last emergency stop
                    If gCommErrorStop <> 0 Then
                        Call Add_Message("CommErr: Stop")
                        Call emergency_stop
                        gInitResult = 0
                    End If
                Else
                    If EncoderReadErrCount <> 0 Then
                        ' check the motor status
                        ImageComms.Picture = LoadResPicture(104, vbResBitmap)
                        gRAStatus = EQ_GetMotorStatus(0)
                        gDECStatus = EQ_GetMotorStatus(1)
                        Call Add_Message("CommErr:MSR=" & CStr(gRAStatus))
                        Call Add_Message("CommErr:MSD=" & CStr(gDECStatus))
                    End If
                End If
                If EncoderReadErrCount > 0 Then
                    EncoderReadErrCount = EncoderReadErrCount - 1
                Else
                    ImageComms.Picture = LoadResPicture(103, vbResBitmap)
                End If
                gEmulRA = tmpRa
                gEmulDEC = tmpDec
                gEmulRA_Init = gEmulRA
                gLast_time = EQnow_lst_norange()
                gEmulOneShot = False
            End If
        Else
            ' emulate RA motor position
            gEmulRA = GetEmulRA()
        End If

        If gThreeStarEnable = False Then
            gSelectStar = 0
            gRA_Encoder = Delta_RA_Map(gEmulRA)
            gDec_Encoder = Delta_DEC_Map(gEmulDEC)
        Else
            Select Case gAlignmentMode
            
                Case 2
                    tmpcoord = DeltaSync_Matrix_Map(gEmulRA, gEmulDEC)
                    gRA_Encoder = tmpcoord.x
                    gDec_Encoder = tmpcoord.Y
                
                Case 1
                    tmpcoord = Delta_Matrix_Reverse_Map(gEmulRA, gEmulDEC)
                    gRA_Encoder = tmpcoord.x
                    gDec_Encoder = tmpcoord.Y
                    
                Case Else
                    tmpcoord = Delta_Matrix_Reverse_Map(gEmulRA, gEmulDEC)
                    gRA_Encoder = tmpcoord.x
                    gDec_Encoder = tmpcoord.Y
                    If tmpcoord.F = 0 Then
                        tmpcoord = DeltaSync_Matrix_Map(gEmulRA, gEmulDEC)
                        gRA_Encoder = tmpcoord.x
                        gDec_Encoder = tmpcoord.Y
                    End If
            
            End Select
        End If
        
        ' Convert RA_Encoder to Hours
        If (gRA_Encoder < &H1000000) Then gRA_Hours = Get_EncoderHours(gRAEncoder_Zero_pos, gRA_Encoder, gTot_RA, gHemisphere)
        
        ' Convert DEC_Encoder to DEC Degrees
        gDec_DegNoAdjust = Get_EncoderDegrees(gDECEncoder_Zero_pos, gDec_Encoder, gTot_DEC, gHemisphere)
        If (gDec_Encoder < &H1000000) Then
            gDec_Degrees = Range_DEC(gDec_DegNoAdjust)
        End If
        
        tRA = EQnow_lst(gLongitude * DEG_RAD) + gRA_Hours
        If gHemisphere = 0 Then
            If (gDec_DegNoAdjust > 90) And (gDec_DegNoAdjust <= 270) Then
                tRA = tRA - 12
            End If
        Else
            If (gDec_DegNoAdjust <= 90) Or (gDec_DegNoAdjust > 270) Then
                tRA = tRA + 12
            End If
        End If
        tRA = Range24(tRA)
        
        ' assign global RA/Dec
        gRA = tRA
        gDec = gDec_Degrees
        gha = tRA - EQnow_lst(gLongitude * DEG_RAD)
        
        ' calc alt/az poition
        hadec_aa (gLatitude * DEG_RAD), (gha * HRS_RAD), (gDec_Degrees * DEG_RAD), tAlt, tAz
        ' asign global Alt/Az
        gAlt = tAlt * RAD_DEG           ' convert to degrees from Radians
        gAz = 360# - (tAz * RAD_DEG)    '  convert to degrees from Radians
        
        ' Poll the Motor Status while slew is active
        If gSlewStatus = True Then
            gRAStatus = EQ_GetMotorStatus(0)
            gDECStatus = EQ_GetMotorStatus(1)
            If gEQparkstatus = 0 Then
                Call ManageGoto
            End If
        End If
        
        AlignmentCountLbl.Caption = CStr(gAlignmentStars_count)
        
        ' do limit management
        Call Limits_Execute
        
        EncoderTimerFlag = True

    End If

End Sub
Private Sub DisplayTimer_Timer()
    Dim i As Double
    Dim lim1 As Double
    Dim lim2 As Double
    Dim TmpLst As Double
 
    Select Case gInitResult
        Case 0, 5, 10
            ' Success
            TextCommsErr.Visible = False
        Case 1
            'COM Port Not available
            TextCommsErr.Visible = True
            TextCommsErr.Text = oLangDll.GetLangString(6010) & vbCrLf & oLangDll.GetLangString(6011)
            ImageComms.Picture = LoadResPicture(105, vbResBitmap)
            Exit Sub
        Case 2
            'COM Port already Open
            TextCommsErr.Visible = True
            TextCommsErr.Text = oLangDll.GetLangString(6010) & vbCrLf & oLangDll.GetLangString(6012)
            ImageComms.Picture = LoadResPicture(105, vbResBitmap)
            Exit Sub
        Case 3
            'COM Timeout Error
            TextCommsErr.Visible = True
            TextCommsErr.Text = oLangDll.GetLangString(6010) & vbCrLf & oLangDll.GetLangString(6013)
            ImageComms.Picture = LoadResPicture(105, vbResBitmap)
            Exit Sub
        Case 999
            'Invalid parameter
            TextCommsErr.Visible = True
            TextCommsErr.Text = oLangDll.GetLangString(6010) & vbCrLf & oLangDll.GetLangString(6014)
            Exit Sub
        Case Else
            TextCommsErr.Visible = True
            TextCommsErr.Text = oLangDll.GetLangString(6010)
            Exit Sub
    End Select
 
    TmpLst = EQnow_lst(gLongitude * DEG_RAD)
    lstlbl.Caption = FmtSexa(TmpLst, False)
    gPolHa = Range24(TmpLst - gPoleStarRa)
    CommandPolaris.Caption = FmtSexa(gPolHa, False)
    
    ralbl.Caption = FmtSexa(gRA, False)
    declbl.Caption = FmtSexa(gDec, True)

    azlbl.ForeColor = vbRed
    altlbl.ForeColor = vbRed
    declbl.ForeColor = vbRed
    ralbl.ForeColor = vbRed
    azlbl.FontSize = 16
    altlbl.FontSize = 16
    ralbl.FontSize = 16
    declbl.FontSize = 16
    azlbl.Caption = FmtSexa(gAz, False)
    altlbl.Caption = FmtSexa(gAlt, False)
    ralbl.Caption = FmtSexa(gRA, False)
    declbl.Caption = FmtSexa(gDec, True)
    
    If LimitStatus.AtLimit Then
        If flash Then
            flash = False
            azlbl.ForeColor = &H80FF&
            altlbl.ForeColor = &H80FF&
            ralbl.ForeColor = &H80FF&
            declbl.ForeColor = &H80FF&
            altlbl.Caption = oLangDll.GetLangString(2114)
            declbl.Caption = oLangDll.GetLangString(2114)
            If LimitStatus.Horizon Then
                ' Horizon
                ralbl.Caption = oLangDll.GetLangString(2108)
                If LimitStatus.RA Then
                    ' RA
                    azlbl.Caption = oLangDll.GetLangString(2109)
                Else
                    ' Horizon
                    azlbl.Caption = oLangDll.GetLangString(2108)
                End If
            Else
                ' RA
                ralbl.Caption = oLangDll.GetLangString(2109)
                azlbl.Caption = oLangDll.GetLangString(2109)
            End If
        Else
            flash = True
       End If
    Else
        If gEQparkstatus = 1 Then
            If flash Then
                flash = False
                If Len(oLangDll.GetLangString(177)) > 8 Then
                    azlbl.FontSize = 10
                    altlbl.FontSize = 10
                    ralbl.FontSize = 10
                    declbl.FontSize = 10
                End If
                azlbl.ForeColor = &H80FF&
                azlbl.Caption = oLangDll.GetLangString(177)
                altlbl.ForeColor = &H80FF&
                altlbl.Caption = oLangDll.GetLangString(177)
                ralbl.ForeColor = &H80FF&
                ralbl.Caption = oLangDll.GetLangString(177)
                declbl.ForeColor = &H80FF&
                declbl.Caption = oLangDll.GetLangString(177)
            Else
                flash = True
            End If
        End If
    End If
    
    'gSideofPier = SOP_Pointing(gDec_DegNoAdjust)
  
    Select Case SOP_Physical(gRA_Hours)
        Case pierUnknown:    lblPier.Caption = oLangDll.GetLangString(180)
        Case pierEast:       lblPier.Caption = oLangDll.GetLangString(181)
        Case pierWest:       lblPier.Caption = oLangDll.GetLangString(182)
    End Select
  
    If FrameAxis.Visible = True Then
  
        i = (Range24(gRA_Hours - 6) / 24) * 100
        
        If gRA_Limit_East <> 0 And gRA_Limit_West <> 0 And HC.ChkEnableLimits.Value = 1 Then
        
            lim1 = 100 * (gRAEncoder_Zero_pos - gRA_Limit_East) / (gTot_RA)
            If lim1 < 0 Then lim1 = 100 + lim1
            lim2 = 100 * (gRAEncoder_Zero_pos - gRA_Limit_West) / (gTot_RA)
            If lim2 < 0 Then lim2 = 100 + lim2
        Else
            lim1 = -1
            lim2 = -1
        End If
        
        If gHemisphere <> 0 Then
            i = 100 - i
        End If
        Call DrawAxis(PictureRA, 0, i, lim1, lim2)
        
        If lblmode = 1 Then
            Label82.Caption = FmtSexa(gRA_Hours, False)
        Else
            Label82.Caption = printhex(gEmulRA)
        End If
        
    
        If gHemisphere = 0 Then
            i = gDec_DegNoAdjust - 90
        Else
            i = gDec_DegNoAdjust - 270
        End If
        If i < 0 Then i = 360 + i
        If i > 360 Then i = i - 360
        i = (i / 360) * 100
        If gHemisphere <> 0 Then
            i = 100 - i
        End If
        Call DrawAxis(PictureDEC, 1, 100 - i, -1, -1)
        
        If lblmode = 1 Then
            Label83.Caption = FmtSexa(gDec_DegNoAdjust, False)
        Else
            Label83.Caption = printhex(gEmulDEC)
        End If
    End If
  
    If Frame9.Visible = True Then
        Call PEC_DispalyUpdate(PlotPecStatus)
    End If
  
End Sub



Private Sub JoyStick_Timer_Timer()
    
    If gJoyTimerFlag = True Then            'Avoid Overruns
        gJoyTimerFlag = False
        
        If JoystickCal.Enabled Then
            If EQ_JoystickPoller(val(VScrollRASlewRate.Value), val(VScrollDecSlewRate.Value)) = False Then
                ' HC.JoyStick_Timer.Enabled = False
                If JoystickLost = False Then
                    HC.Add_Message (oLangDll.GetLangString(5021))
                End If
                JoystickLost = True
            Else
                JoystickLost = True
            End If
        Else
            Call EQ_JoystickPoller2
        End If
        gJoyTimerFlag = True
    End If
End Sub

Private Sub Pulseguide_Timer_Timer()

    ' This is a timer for Pulse guide - the interval (resolution)is settable by the user

    If gEQPulsetimerflag Then
        
        ' Prevent Overruns
        gEQPulsetimerflag = False
        
        ' don't do anthing if parked or parking!
        If gEQparkstatus = 0 Then
            If gEQDECPulseDuration > 0 Then
                gEQDECPulseDuration = gEQDECPulseDuration - gpl_interval
                If gEQDECPulseDuration <= 0 Then
                    If gTrackingStatus <> 4 Then
                        eqres = EQ_MotorStop(1)
                    Else
                        Call ChangeDEC_by_Rate(gDeclinationRate)
                    End If
                    gEQDECPulseDuration = 0
                    ' Force an update
                    gEmulOneShot = True
                    gPulseguideRateDec = 0
                End If
            End If
            
            ' Guide only when the scope is tracking
            If gTrackingStatus > 0 And gEQRAPulseDuration > 0 Then
                gEQRAPulseDuration = gEQRAPulseDuration - gpl_interval
                If gEQRAPulseDuration <= 0 Then
                    If gTrackingStatus <> 4 And (gRA_LastRate = 0) Then
                        eqres = EQ_SendGuideRate(0, gTrackingStatus - 1, 0, 0, gHemisphere, gHemisphere)
                    Else
                        Call ChangeRA_by_Rate(gRightAscensionRate)
                    End If
                    gEQRAPulseDuration = 0
                    ' Force an update
                    gEmulOneShot = True
                    gPulseguideRateRa = 0
                End If
            Else
                gEQRAPulseDuration = 0
                ' Force an update
                gEmulOneShot = True
                gPulseguideRateRa = 0
            End If
        Else
            gEQRAPulseDuration = 0
            gEQDECPulseDuration = 0
            ' Force an update
            gEmulOneShot = True
            gPulseguideRateRa = 0
            gPulseguideRateDec = 0
        End If
        
        gEQPulsetimerflag = True
    
    End If

End Sub

Private Sub Reverse_Click(Index As Integer)
    Select Case Index
        Case 0
            RA_inv = Reverse(0)
        Case 1
            DEC_Inv = Reverse(1)
    End Select
End Sub

Private Sub SitesCombo_Click()
    SiteIdx = SitesCombo.ListIndex
    ' Force loss of focus
     SendKeys "{TAB}", True
        
    Call CommandLoadSite_Click

End Sub



Private Sub Spiral_Timer_Timer()
    If gSpiralTimerFlag = True Then
        gSpiralTimerFlag = False  ' Avoid Overruns
    
        If gSpiral_AxisFlag = 0 Then
        
            eqres = EQ_GetMotorStatus(0) ' Get the Slew Status of the RA Motor
            If (eqres And EQ_MOTORBUSY) = 0 Then
               
                ' Compute for the next Length
                gRightAscension_Len = gRightAscension_Len + gSPIRAL_JUMP
            
                ' Start the DEC Motor
                If gDeclination_Dir = 0 Then
                    eqres = EQStartMoveMotor(1, 0, 0, gDeclination_Len, GetSlowdown(gDeclination_Len))
                    gDeclination_Dir = 1
                Else
                    eqres = EQStartMoveMotor(1, 0, 1, gDeclination_Len, GetSlowdown(gDeclination_Len))
                    gDeclination_Dir = 0
                End If
                
                ' Activate Sidereal tracking
                eqres = EQ_MotorStop(0)          ' Just Make sure RA motor is stopped at this point
                If eqres <> EQ_OK Then
                    GoTo DDEC01
                End If
                
                Do
                    eqres = EQ_GetMotorStatus(0)
                    If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo DDEC01
                Loop While (eqres And EQ_MOTORBUSY) <> 0
 
DDEC01:
                If gTrackingStatus <> 0 Then
                     If gTrackingStatus = 4 Then
                        Call Restore_CustomTracking
                     Else
                        eqres = EQ_StartRATrack(gTrackingStatus - 1, gHemisphere, gHemisphere)
                     End If
                End If
                
                gSpiral_AxisFlag = 1
            
            End If
        Else
        
            eqres = EQ_GetMotorStatus(1) ' Get the Slew Status of the DEC Motor
            If (eqres And EQ_MOTORBUSY) = 0 Then
   
                ' Compute for the next Length
                gDeclination_Len = gDeclination_Len + gSPIRAL_JUMP
                
                ' Just make RA motor is stopped
                eqres = EQ_MotorStop(0)          ' Just Make sure RA motor is stopped at this point
                If eqres <> EQ_OK Then
                    GoTo DDEC02
                End If
                
                Do
                    eqres = EQ_GetMotorStatus(0)
                    If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo DDEC02
                Loop While (eqres And EQ_MOTORBUSY) <> 0
 
DDEC02:
                'Start RA MOTOR
                If gRightAscension_Dir = 0 Then
                    eqres = EQStartMoveMotor(0, 0, 0, gRightAscension_Len, GetSlowdown(gRightAscension_Len))
                    gRightAscension_Dir = 1
                Else
                    eqres = EQStartMoveMotor(0, 0, 1, gRightAscension_Len, GetSlowdown(gRightAscension_Len))
                    gRightAscension_Dir = 0
                End If
                    
                gSpiral_AxisFlag = 0
            
            End If
    
        End If
        
        gSpiralTimerFlag = True
    End If
End Sub

Private Sub SpiralHScroll1_Change()
    LabelSpiral.Caption = HC.SpiralHScroll1.Value
End Sub

Private Sub SpiralHScroll1_Scroll()
    Call SpiralHScroll1_Change
End Sub

Private Sub VScrollRASlewRate_Change()
    If VScrollRASlewRate.Value >= 10 Then
        raslewlbl.Caption = VScrollRASlewRate.Value - 9
    Else
        raslewlbl.Caption = "0." & VScrollRASlewRate.Value
    End If
    If Slewpad.Visible Then
        Slewpad.VScroll1.Value = VScrollRASlewRate.Value
    End If
End Sub

Private Sub VScrollRASlewRate_Scroll()
    Call VScrollRASlewRate_Change
End Sub

Private Sub VScrollDecSlewRate_Change()
    If VScrollDecSlewRate.Value >= 10 Then
        decslewlbl.Caption = VScrollDecSlewRate.Value - 9
    Else
        decslewlbl.Caption = "0." & VScrollDecSlewRate.Value
    End If
    If Slewpad.Visible Then
        Slewpad.VScroll2.Value = VScrollDecSlewRate.Value
    End If
End Sub

Private Sub VScrollDecSlewRate_Scroll()
    Call VScrollDecSlewRate_Change
End Sub

Public Sub Add_FileMessage(dtaLog As String)

    On Error Resume Next
    
    ' log to file
    If CheckLog2File.Value = 1 Then
        Print #4, "[" & time & "] " & dtaLog
        logcount = logcount + 1
        If logcount > 1000 Then
            Close #4
            If logfileindex = 1 Then
                Open oPersist.GetIniPath & "\messagelog2.txt" For Output As #4
                logfileindex = 2
            Else
                Open oPersist.GetIniPath & "\messagelog1.txt" For Output As #4
                logfileindex = 1
            End If
            logcount = 0
        End If
    End If
    
End Sub


Public Sub Add_Message(dtaLog As String)

    On Error Resume Next
    
    ' log to file
    If CheckLog2File.Value = 1 Then
        Print #4, "[" & time & "] " & dtaLog
        logcount = logcount + 1
        If logcount > 2000 Then
            Close #4
            If logfileindex = 1 Then
                Open oPersist.GetIniPath & "\messagelog2.txt" For Output As #4
                logfileindex = 2
            Else
                Open oPersist.GetIniPath & "\messagelog1.txt" For Output As #4
                logfileindex = 1
            End If
            logcount = 0
        End If
    End If
    
    'Add the message, limit to 1000 characters
    HCMessage.Text = Right(HCMessage.Text & vbCrLf & dtaLog, 1000)
    HCMessage.SelStart = Len(HCMessage.Text)

End Sub

Private Sub HScrollRARate_Change()
    HC.Label14.Caption = Format$(str(HScrollRARate.Value * 0.1), "x0.00")
End Sub

Private Sub HScrollRARate_Scroll()
    Call HScrollRARate_Change
End Sub

Private Sub HScrollDecRate_Change()
    HC.Label15.Caption = Format$(str(HScrollDecRate.Value * 0.1), "x0.00")
End Sub

Private Sub HScrollDecRate_Scroll()
    Call HScrollDecRate_Change
End Sub

Private Sub HScrollRAOride_Change()
    HC.Label19.Caption = Format$(str(HScrollRAOride.Value * 100), " 0000")
End Sub

Private Sub HScrollRAOride_Scroll()
    Call HScrollRAOride_Change
End Sub

Private Sub HScrollDecOride_Change()
    HC.Label20.Caption = Format$(str(HScrollDecOride.Value * 100), " 0000")
End Sub

Private Sub HScrollDecOride_Scroll()
    Call HScrollDecOride_Change
End Sub

Private Sub VScroll7_KeyDown(KeyCode As Integer, Shift As Integer)
    Call keydown(KeyCode, val(VScrollRASlewRate.Value), val(VScrollDecSlewRate.Value))
End Sub


Private Sub VScroll7_KeyUp(KeyCode As Integer, Shift As Integer)
    Call keyup(KeyCode)
End Sub


'Public Sub Add_Message_Align(dtaLog As String)
'
'       HC.HCTextAlign.Text = Right(HC.HCTextAlign.Text & dtaLog & vbCrLf, 500000)
'       HC.HCTextAlign.SelStart = Len(HC.HCTextAlign.Text)
'
'End Sub



Private Sub SetText()
    Dim tmp As Long
    Dim tmptxt As String
    ' load language dll
    oLangDll.LoadLangDll (LanguageDll)

    'HC.Caption = oLangDll.GetLangString(101)
    
    tmptxt = HC.oPersist.ReadIniValue("FriendlyName")
    If tmptxt = "" Then
        HC.Caption = ASCOM_DESC
    '    MainLabel.Caption = oLangDll.GetLangString(101) & " " & CURRENT_VERSION
        MainLabel.Caption = ASCOM_DESC & " " & CURRENT_VERSION
    Else
        HC.Caption = ASCOM_DESC & " (" & tmptxt & ")"
    '    MainLabel.Caption = oLangDll.GetLangString(101) & " " & CURRENT_VERSION
        MainLabel.Caption = tmptxt & " " & CURRENT_VERSION
    End If
'    If gDllVer >= 3.04 Then
'        HC.Caption = oLangDll.GetLangString(101) & " DLL Version = (" & CStr(gddlver) & ")"
'    End If
    
    ' message center
    Frame2.Caption = oLangDll.GetLangString(102)
    CheckLog2File.Caption = oLangDll.GetLangString(2700)
    CommandClearMsg.ToolTipText = oLangDll.GetLangString(2701)
    
    CmdDisplayMode.ToolTipText = oLangDll.GetLangString(200)
    
    'Position
    PositionFrame.Caption = oLangDll.GetLangString(103)
    lstmrk.Caption = oLangDll.GetLangString(104)
    ramrk.Caption = oLangDll.GetLangString(105)
    decmrk.Caption = oLangDll.GetLangString(106)
    azmrk.Caption = oLangDll.GetLangString(107)
    altmrk.Caption = oLangDll.GetLangString(108)
    Label1.Caption = oLangDll.GetLangString(109)
    ChkForceFlip.Caption = oLangDll.GetLangString(6111)
    
    ' Pisition (axis)
    FrameAxis.Caption = oLangDll.GetLangString(103)
    
    SlewFrame.Caption = oLangDll.GetLangString(110)
    CommandGamepad(1).ToolTipText = oLangDll.GetLangString(111)
    CmdSlewPad.Caption = oLangDll.GetLangString(112)
    CmdSlew(0).Caption = oLangDll.GetLangString(113)
    CmdSlew(1).Caption = oLangDll.GetLangString(114)
    CmdSlew(2).Caption = oLangDll.GetLangString(116)
    CmdSlew(3).Caption = oLangDll.GetLangString(115)
    CmdSlew(4).ToolTipText = oLangDll.GetLangString(1100)
    cmdSlew2(0).Caption = oLangDll.GetLangString(113)
    cmdSlew2(1).Caption = oLangDll.GetLangString(114)
    cmdSlew2(2).Caption = oLangDll.GetLangString(116)
    cmdSlew2(3).Caption = oLangDll.GetLangString(115)
    cmdSlew2(4).ToolTipText = oLangDll.GetLangString(1100)
    
    
    Label9.Caption = oLangDll.GetLangString(117)
    Label10.Caption = oLangDll.GetLangString(118)
    RA_inv.Caption = oLangDll.GetLangString(119)
    DEC_Inv.Caption = oLangDll.GetLangString(120)
    
    TrackingFrame.Caption = oLangDll.GetLangString(121)
    CmdTrack(0).ToolTipText = oLangDll.GetLangString(122)
    CmdTrack(2).ToolTipText = oLangDll.GetLangString(123)
    CmdTrack(3).ToolTipText = oLangDll.GetLangString(124)
    CmdTrack(5).ToolTipText = oLangDll.GetLangString(125)
    CmdTrack(1).ToolTipText = oLangDll.GetLangString(188)
    CmdTrack(4).ToolTipText = oLangDll.GetLangString(189)
    SitesCombo.Text = oLangDll.GetLangString(186)

    
    Frame1.Caption = oLangDll.GetLangString(126)
    Label6(0).Caption = oLangDll.GetLangString(127)
    Label6(1).Caption = oLangDll.GetLangString(128)
    Label6(2).Caption = oLangDll.GetLangString(129)
    Label6(4).Caption = oLangDll.GetLangString(1323)
    Label6(5).Caption = oLangDll.GetLangString(130)
    CommandLoadSite.ToolTipText = oLangDll.GetLangString(184)
    CommandSaveSite.ToolTipText = oLangDll.GetLangString(185)
    CommandSetSite.Caption = oLangDll.GetLangString(131)
    Commandgps.ToolTipText = oLangDll.GetLangString(132)
    
    Frame3.Caption = oLangDll.GetLangString(133)
    Reset_Align_Command.ToolTipText = oLangDll.GetLangString(139)
    Edit_Stars_Command.ToolTipText = oLangDll.GetLangString(140)
    CmdClearSync.ToolTipText = oLangDll.GetLangString(141)
    Label8.Caption = oLangDll.GetLangString(142)
    Label11.Caption = oLangDll.GetLangString(143)
    ListAlignMode.AddItem oLangDll.GetLangString(25) & "+" & oLangDll.GetLangString(212) ' 2-Point+nearest
    ListAlignMode.AddItem (oLangDll.GetLangString(212)) ' nearest star
    ListSyncMode.AddItem (oLangDll.GetLangString(213))
    ListSyncMode.AddItem (oLangDll.GetLangString(214))
    Label44.Caption = oLangDll.GetLangString(210)
    Label34.Caption = oLangDll.GetLangString(211)
    Label36.Caption = oLangDll.GetLangString(63)
    Label37.Caption = oLangDll.GetLangString(1490)
    
    ComboActivePoints.Clear
    ComboActivePoints.AddItem (oLangDll.GetLangString(1491))
    ComboActivePoints.AddItem (oLangDll.GetLangString(1492))
    ComboActivePoints.AddItem (oLangDll.GetLangString(1493))
    Combo3PointAlgorithm.Clear
    Combo3PointAlgorithm.AddItem (oLangDll.GetLangString(1494))
    Combo3PointAlgorithm.AddItem (oLangDll.GetLangString(1495))
    CheckLocalPier.Caption = (oLangDll.GetLangString(1496))
    Label29.Caption = oLangDll.GetLangString(1497)
    Label35.Caption = oLangDll.GetLangString(1498)

    ' Park & Unpark
    Frame4.Caption = oLangDll.GetLangString(145)
    CommandDefinePark.ToolTipText = oLangDll.GetLangString(147)
    Label49.Caption = oLangDll.GetLangString(1501)
    Label50.Caption = oLangDll.GetLangString(1502)
    Call LoadParkCombo
    CommandSyncEncoder.ToolTipText = oLangDll.GetLangString(60)
    Frame15.Caption = oLangDll.GetLangString(146)
     
    Frame5.Caption = oLangDll.GetLangString(153)
    rapulse_enchk.Caption = oLangDll.GetLangString(117)
    decpulse_enchk.Caption = oLangDll.GetLangString(118)
    Label16.Caption = oLangDll.GetLangString(154)
    rafixed_enchk.Caption = oLangDll.GetLangString(105)
    decfixed_enchk.Caption = oLangDll.GetLangString(106)
    Label3.Caption = oLangDll.GetLangString(220)
    
    Frame6.Caption = oLangDll.GetLangString(155)
    Label21.Caption = oLangDll.GetLangString(117)
    Label22.Caption = oLangDll.GetLangString(118)
    
    Label24.Caption = oLangDll.GetLangString(105)
    Label39.Caption = oLangDll.GetLangString(106)
   
    Frame11.Caption = oLangDll.GetLangString(164)
    Label4.Caption = oLangDll.GetLangString(165)
    Command35.ToolTipText = oLangDll.GetLangString(167)
    CmdSlew(5).ToolTipText = oLangDll.GetLangString(169)
    CommandSounds.ToolTipText = oLangDll.GetLangString(2800)
    
    Frame12.Caption = oLangDll.GetLangString(170)
    CommandGamepad(0).ToolTipText = oLangDll.GetLangString(171)

    HC.Frame9.Caption = oLangDll.GetLangString(19)
    
    CommandSetup.ToolTipText = oLangDll.GetLangString(90)
    CommandSetup.Picture = LoadResPicture(101, vbResBitmap)

    
    cbNS.Clear
    cbNS.AddItem (oLangDll.GetLangString(113))
    cbNS.AddItem (oLangDll.GetLangString(116))
    cbEW.Clear
    cbEW.AddItem (oLangDll.GetLangString(114))
    cbEW.AddItem (oLangDll.GetLangString(115))
    cbhem.Clear
    cbhem.AddItem (oLangDll.GetLangString(1110))
    cbhem.AddItem (oLangDll.GetLangString(1112))
    
    ' PEC
    CmdPecLoad.ToolTipText = oLangDll.GetLangString(184)
    CmdPecSave.ToolTipText = oLangDll.GetLangString(185)
    CmdPecClear.ToolTipText = oLangDll.GetLangString(21)
'    Label46.Caption = oLangDll.GetLangString(191)
'    Label48.Caption = oLangDll.GetLangString(192)
    Command23.ToolTipText = oLangDll.GetLangString(196)
'    Frame14.Caption = oLangDll.GetLangString(195)
    CheckPEC.Caption = oLangDll.GetLangString(20)
    CommandPecPlay.ToolTipText = oLangDll.GetLangString(6127)
    CommandRecordPec.ToolTipText = oLangDll.GetLangString(6128)
    CommandPecConfig.ToolTipText = oLangDll.GetLangString(6126)

    ' Drift Compensation
    Frame13.Caption = oLangDll.GetLangString(1400)

    ' ASCOM PulseGuide Monitor
    Frame16.Caption = oLangDll.GetLangString(1600)
    Label55.Caption = oLangDll.GetLangString(105)
    Label56.Caption = oLangDll.GetLangString(106)
    Label57.Caption = oLangDll.GetLangString(114)
    Label58.Caption = oLangDll.GetLangString(115)
    Label59.Caption = oLangDll.GetLangString(113)
    Label60.Caption = oLangDll.GetLangString(116)
    Label51.Caption = oLangDll.GetLangString(1601)
    Label52.Caption = oLangDll.GetLangString(1602)
    
    ' Tracking
    FrameCustomTrack.Caption = oLangDll.GetLangString(6200)
    Label2.Caption = oLangDll.GetLangString(6201)
        
    ' limits
    FrameLimits.Caption = oLangDll.GetLangString(2100)
    CmdEditLimits.ToolTipText = oLangDll.GetLangString(171)
    ChkEnableLimits.Caption = oLangDll.GetLangString(2102)

    ' backlash
    Label81.Caption = oLangDll.GetLangString(221)
    
    'Goto form

    ' Update
    CommandUpdate.Caption = oLangDll.GetLangString(6015)
    
    ' Goto Slew Limit
    Label31.Caption = oLangDll.GetLangString(6050)
    
    'popupmenu
'    puPolar.Caption = oLangDll.GetLangString()
    puMessageCenter.Caption = oLangDll.GetLangString(102)
    puPosition.Caption = oLangDll.GetLangString(103) & "_1"
    puDials.Caption = oLangDll.GetLangString(103) & "_2"
    puPEC.Caption = oLangDll.GetLangString(195)
    puPulse.Caption = oLangDll.GetLangString(1600)

    If gAscomCompatibility.AllowPulseGuide Then
        puPulse.Visible = True
    Else
        puPulse.Visible = False
    End If

End Sub

Public Sub LoadParkCombo()
    Dim i As Integer
            
    Call readUserParkPos
 
    ComboPark.Clear
    ComboUnPark.Clear
    
    ComboPark.AddItem (oLangDll.GetLangString(150))
'    ComboPark.AddItem (oLangDll.GetLangString(148))
    ComboPark.AddItem (oLangDll.GetLangString(149))
    For i = 1 To 10
        ComboPark.AddItem (UserParks(i).name)
    Next i
    ComboUnPark.AddItem (oLangDll.GetLangString(151))
    ComboUnPark.AddItem (oLangDll.GetLangString(152))
'    ComboUnPark.AddItem (oLangDll.GetLangString(2000))
    For i = 1 To 10
        ComboUnPark.AddItem (UserUnparks(i).name)
    Next i

End Sub


Public Sub SetParkCombo()
    Dim i As Integer
    
    For i = 1 To 10
        ComboPark.List(1 + i) = UserParks(i).name
        ComboUnPark.List(i + 1) = UserUnparks(i).name
    Next i

End Sub


