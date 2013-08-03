VERSION 5.00
Begin VB.Form JStickConfigForm 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10230
   Icon            =   "JStickConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   10230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Calibration"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   4575
      Left            =   7440
      TabIndex        =   79
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton StartBtn 
         BackColor       =   &H0095C1CB&
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   4080
         Width           =   2295
      End
      Begin VB.Label XminLabel 
         BackColor       =   &H00000080&
         Caption         =   "0"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   960
         TabIndex        =   95
         Top             =   720
         Width           =   615
      End
      Begin VB.Label YminLabel 
         BackColor       =   &H00000080&
         Caption         =   "0"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   960
         TabIndex        =   94
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label ZminLabel 
         BackColor       =   &H00000080&
         Caption         =   "0"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   960
         TabIndex        =   93
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label RminLabel 
         BackColor       =   &H00000080&
         Caption         =   "0"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   960
         TabIndex        =   92
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label XmaxLabel 
         BackColor       =   &H00000080&
         Caption         =   "0"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   1800
         TabIndex        =   91
         Top             =   720
         Width           =   615
      End
      Begin VB.Label YmaxLabel 
         BackColor       =   &H00000080&
         Caption         =   "0"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   1800
         TabIndex        =   90
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label ZmaxLabel 
         BackColor       =   &H00000080&
         Caption         =   "0"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   1800
         TabIndex        =   89
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label RmaxLabel 
         BackColor       =   &H00000080&
         Caption         =   "0"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   1800
         TabIndex        =   88
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "Min"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   960
         TabIndex        =   87
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Caption         =   "Max"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   1800
         TabIndex        =   86
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "X Axis"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "Y Axis"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "Z Axis"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "R Axis"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         Caption         =   "Move the joysick paddles to their extreme limits until the numbers above cease changing."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1215
         Left            =   240
         TabIndex        =   81
         Top             =   2640
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "System"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   1935
      Left            =   7440
      TabIndex        =   96
      Top             =   4560
      Width           =   2655
      Begin VB.CheckBox CheckDualSpeed 
         BackColor       =   &H00000000&
         Caption         =   "Dual Speed Joystick"
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
         TabIndex        =   106
         Top             =   480
         Width           =   2415
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         ItemData        =   "JStickConfig.frx":0CCA
         Left            =   1320
         List            =   "JStickConfig.frx":0CD4
         Style           =   2  'Dropdown List
         TabIndex        =   104
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         ItemData        =   "JStickConfig.frx":0CEC
         Left            =   120
         List            =   "JStickConfig.frx":0CEE
         Style           =   2  'Dropdown List
         TabIndex        =   99
         Top             =   1440
         Width           =   2415
      End
      Begin VB.CheckBox CheckPOV 
         BackColor       =   &H00000000&
         Caption         =   "POV Pad Enabled"
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
         TabIndex        =   98
         Top             =   720
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Joystick Support Enabled"
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
         TabIndex        =   97
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         Caption         =   "MonitorToggle"
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
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3480
      Top             =   480
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7560
      Top             =   360
   End
   Begin VB.CommandButton ApplyBtn 
      BackColor       =   &H0095C1CB&
      Caption         =   "Apply Changes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton CancelBtn 
      BackColor       =   &H0095C1CB&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Buttons"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin VB.CommandButton DefaultBtn 
         BackColor       =   &H0095C1CB&
         Caption         =   "Load Defaults"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   75
         Top             =   6000
         Width           =   1575
      End
      Begin VB.CommandButton ClearBtn 
         BackColor       =   &H0095C1CB&
         Caption         =   "Clear All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   6000
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   38
         Left            =   5760
         TabIndex        =   102
         Top             =   5520
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   37
         Left            =   5760
         TabIndex        =   100
         Top             =   5280
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   36
         Left            =   2040
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   35
         Left            =   5760
         TabIndex        =   74
         Top             =   4920
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   34
         Left            =   2040
         TabIndex        =   40
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   33
         Left            =   2040
         TabIndex        =   38
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   32
         Left            =   2040
         TabIndex        =   36
         Top             =   4920
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   31
         Left            =   2040
         TabIndex        =   34
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   30
         Left            =   5760
         TabIndex        =   72
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   8
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   12
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   14
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   5
         Left            =   5760
         TabIndex        =   42
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   6
         Left            =   5760
         TabIndex        =   44
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   7
         Left            =   2040
         TabIndex        =   24
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   8
         Left            =   5760
         TabIndex        =   46
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   9
         Left            =   5760
         TabIndex        =   48
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   10
         Left            =   5760
         TabIndex        =   50
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   11
         Left            =   5760
         TabIndex        =   52
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   12
         Left            =   5760
         TabIndex        =   66
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   13
         Left            =   2040
         TabIndex        =   26
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   14
         Left            =   2040
         TabIndex        =   28
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   15
         Left            =   2040
         TabIndex        =   30
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   16
         Left            =   2040
         TabIndex        =   32
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   17
         Left            =   5760
         TabIndex        =   68
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   18
         Left            =   2040
         TabIndex        =   20
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   19
         Left            =   2040
         TabIndex        =   10
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   20
         Left            =   2040
         TabIndex        =   16
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   21
         Left            =   2040
         TabIndex        =   18
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   22
         Left            =   5760
         TabIndex        =   54
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   23
         Left            =   5760
         TabIndex        =   56
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   24
         Left            =   5760
         TabIndex        =   70
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   25
         Left            =   5760
         TabIndex        =   58
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   26
         Left            =   5760
         TabIndex        =   60
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   27
         Left            =   5760
         TabIndex        =   62
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   28
         Left            =   5760
         TabIndex        =   64
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Emergency Stop"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Park to Home"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Park to User Defined"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Park to Current Position"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   19
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Unpark"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   1680
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Sidereal Tracking"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Lunar Tracking"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   20
         Left            =   240
         TabIndex        =   15
         Top             =   2400
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Solar Tracking"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   21
         Left            =   240
         TabIndex        =   17
         Top             =   2640
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Custom Tracking"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   18
         Left            =   240
         TabIndex        =   19
         Top             =   2880
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Spiral Search"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   23
         Top             =   3480
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "North"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   25
         Top             =   3720
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "East"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   27
         Top             =   3960
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "South"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   29
         Top             =   4200
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "West"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   16
         Left            =   240
         TabIndex        =   31
         Top             =   4440
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Reverse RA"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   5
         Left            =   3960
         TabIndex        =   41
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Reverse DEC"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   6
         Left            =   3960
         TabIndex        =   43
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Increase RA rate"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   8
         Left            =   3960
         TabIndex        =   45
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Decrease RA rate"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   9
         Left            =   3960
         TabIndex        =   47
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Increase DEC Rate"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   10
         Left            =   3960
         TabIndex        =   49
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Decrease DEC Rate"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   11
         Left            =   3960
         TabIndex        =   51
         Top             =   1800
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Increment Preset"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   22
         Left            =   3960
         TabIndex        =   53
         Top             =   2040
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Decrement Preset"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   23
         Left            =   3960
         TabIndex        =   55
         Top             =   2280
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Rate_1"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   25
         Left            =   3960
         TabIndex        =   57
         Top             =   2520
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Rate_2"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   26
         Left            =   3960
         TabIndex        =   59
         Top             =   2760
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Rate_3"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   27
         Left            =   3960
         TabIndex        =   61
         Top             =   3000
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Rate_4"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   28
         Left            =   3960
         TabIndex        =   63
         Top             =   3240
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Alignment Accept"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   12
         Left            =   3960
         TabIndex        =   65
         Top             =   3720
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Alignment Cancel"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   17
         Left            =   3960
         TabIndex        =   67
         Top             =   3960
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Alignment End"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   24
         Left            =   3960
         TabIndex        =   69
         Top             =   4200
         Width           =   3135
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   29
         Left            =   2040
         TabIndex        =   22
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Sidereal+PEC Tracking"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   29
         Left            =   240
         TabIndex        =   21
         Top             =   3120
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Sync"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   30
         Left            =   3960
         TabIndex        =   71
         Top             =   4440
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "NorthEast"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   31
         Left            =   240
         TabIndex        =   33
         Top             =   4680
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "SouthWest"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   34
         Left            =   240
         TabIndex        =   39
         Top             =   5400
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "SouthEast"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   33
         Left            =   240
         TabIndex        =   37
         Top             =   5160
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "NorthWest"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   32
         Left            =   240
         TabIndex        =   35
         Top             =   4920
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Align PolarScope"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   35
         Left            =   3960
         TabIndex        =   73
         Top             =   4920
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Dead Mans Handle"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   36
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "Gamepad Lock"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   37
         Left            =   3960
         TabIndex        =   101
         Top             =   5280
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "ScreenSaverToggle"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   38
         Left            =   3960
         TabIndex        =   103
         Top             =   5520
         Width           =   3135
      End
   End
End
Attribute VB_Name = "JStickConfigForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------
' EQMOD project Copyright © 2006 Raymund Sarmiento
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
' JStickConfig.frm - ASCOM EQMOD Joystick assignemnt form
'
'
'
' Written:  27-Jul-07   Chris Shillito
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' 27-Jul-07 cs      Initial edit
' 30-Jul-07 cs      Align end added
' 31-Jul-07 cs      Rate buttons added, Right click on label1
' 15-Aug-07 cs      OnTop handling
' 30-Aug-07 cs      Joystick calibration added.
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


    Public LastIdx As Integer
    Public NoofJbtns As Integer
    
Private Sub ApplyBtn_Click()
    BTN_EMERGENCYSTOP = GetBtnData(Label2(0).Caption)
    BTN_HOMEPARK = GetBtnData(Label2(1).Caption)
    BTN_USERPARK = GetBtnData(Label2(2).Caption)
    BTN_UNPARK = GetBtnData(Label2(3).Caption)
    BTN_STARTSIDREAL = GetBtnData(Label2(4).Caption)
    BTN_RAREVERSE = GetBtnData(Label2(5).Caption)
    BTN_DECREVERSE = GetBtnData(Label2(6).Caption)
    BTN_SPIRAL = GetBtnData(Label2(7).Caption)
    BTN_RARATEINC = GetBtnData(Label2(8).Caption)
    BTN_RARATEDEC = GetBtnData(Label2(9).Caption)
    BTN_DECRATEINC = GetBtnData(Label2(10).Caption)
    BTN_DECRATEDEC = GetBtnData(Label2(11).Caption)
    BTN_ALIGNACCEPT = GetBtnData(Label2(12).Caption)
    BTN_NORTH = GetBtnData(Label2(13).Caption)
    BTN_EAST = GetBtnData(Label2(14).Caption)
    BTN_SOUTH = GetBtnData(Label2(15).Caption)
    BTN_WEST = GetBtnData(Label2(16).Caption)
    BTN_ALIGNCANCEL = GetBtnData(Label2(17).Caption)
    BTN_CUSTOMTRACKSTART = GetBtnData(Label2(18).Caption)
    BTN_CURRENTPARK = GetBtnData(Label2(19).Caption)
    BTN_STARTLUNAR = GetBtnData(Label2(20).Caption)
    BTN_STARTSOLAR = GetBtnData(Label2(21).Caption)
    BTN_INCRATEPRESET = GetBtnData(Label2(22).Caption)
    BTN_DECRATEPRESET = GetBtnData(Label2(23).Caption)
    BTN_ALIGNEND = GetBtnData(Label2(24).Caption)
    BTN_RATE1 = GetBtnData(Label2(25).Caption)
    BTN_RATE2 = GetBtnData(Label2(26).Caption)
    BTN_RATE3 = GetBtnData(Label2(27).Caption)
    BTN_RATE4 = GetBtnData(Label2(28).Caption)
    BTN_PEC = GetBtnData(Label2(29).Caption)
    BTN_SYNC = GetBtnData(Label2(30).Caption)
    BTN_NORTHEAST = GetBtnData(Label2(31).Caption)
    BTN_NORTHWEST = GetBtnData(Label2(32).Caption)
    BTN_SOUTHEAST = GetBtnData(Label2(33).Caption)
    BTN_SOUTHWEST = GetBtnData(Label2(34).Caption)
    BTN_POLARSCOPEALIGN = GetBtnData(Label2(35).Caption)
    BTN_DEADMANSHANDLE = GetBtnData(Label2(36).Caption)
    BTN_TOGGLELOCK = GetBtnData(Label2(37).Caption)
    BTN_TOGGLESCREENSAVER = GetBtnData(Label2(38).Caption)
    
    If CheckPOV.Value = 1 Then
        POV_Enabled = 1
    Else
        POV_Enabled = 0
    End If
    
    Call SaveJoystickBtns
    Call SaveJoystickCalib

    Unload JStickConfigForm
    
End Sub

Private Sub CancelBtn_Click()
    Call LoadJoystickCalib
    Unload JStickConfigForm
End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        JoystickCal.Enabled = 1
    Else
        JoystickCal.Enabled = 0
    End If
    SaveJoystickCalib
    
End Sub

Private Sub CheckDualSpeed_Click()
    If CheckDualSpeed.Value = 1 Then
        JoystickCal.DualSpeed = 1
    Else
        JoystickCal.DualSpeed = 0
    End If
    SaveJoystickCalib
End Sub

Private Sub CheckPOV_Click()
    If CheckPOV.Value = 1 Then
        POV_Enabled = 1
    Else
        POV_Enabled = 0
    End If
End Sub

Private Sub ClearBtn_Click()
Dim Index As Integer

    For Index = 0 To Label2.Count - 1
       Label2(Index).ForeColor = &HFF&
       Label2(Index).Caption = "---"
    Next Index

End Sub



Private Sub Combo1_Click()

    If Combo1.ListIndex = 0 Then
        JoystickCal.id = -1
    Else
        JoystickCal.id = Combo1.ListIndex - 1
    End If
    
    Call ReadGampadProperties
    Call SaveJoystickCalib

End Sub

Private Sub Combo2_Click()
   gMonitorMode = Combo2.ListIndex
End Sub

Private Sub DefaultBtn_Click()
    
    Label2(0).Caption = oLangDll.GetLangString(1134) & "11"
    Label2(1).Caption = "---"
    Label2(2).Caption = "---"
    Label2(3).Caption = "---"
    Label2(4).Caption = oLangDll.GetLangString(1134) & "10"
    Label2(5).Caption = "---"
    Label2(6).Caption = "---"
    Label2(7).Caption = oLangDll.GetLangString(1134) & "1"
    Label2(8).Caption = oLangDll.GetLangString(1134) & "5"
    Label2(9).Caption = oLangDll.GetLangString(1134) & "7"
    Label2(10).Caption = oLangDll.GetLangString(1134) & "6"
    Label2(11).Caption = oLangDll.GetLangString(1134) & "8"
    Label2(12).Caption = oLangDll.GetLangString(1134) & "3"
    Label2(13).Caption = "POV_N"
    Label2(14).Caption = "POV_E"
    Label2(15).Caption = "POV_S"
    Label2(16).Caption = "POV_W"
    Label2(17).Caption = oLangDll.GetLangString(1134) & "2"
    Label2(18).Caption = "---"
    Label2(19).Caption = "---"
    Label2(20).Caption = "---"
    Label2(21).Caption = "---"
    Label2(22).Caption = "---"
    Label2(23).Caption = "---"
    Label2(24).Caption = "---"
    Label2(25).Caption = "---"
    Label2(26).Caption = "---"
    Label2(27).Caption = "---"
    Label2(28).Caption = "---"
    Label2(29).Caption = "---"
    Label2(30).Caption = "---"
    Label2(31).Caption = "POV_NE"
    Label2(32).Caption = "POV_NW"
    Label2(33).Caption = "POV_SE"
    Label2(34).Caption = "POV_SW"
    Label2(35).Caption = "---"
    Label2(36).Caption = "---"
    Label2(37).Caption = "---"
    Label2(38).Caption = "---"
    
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim j As Long

    Call SetText
    
    If HC.HCOnTop.Value = 1 Then Call PutWindowOnTop(JStickConfigForm)
    
    Combo1.AddItem "Auto Select"
    j = 0
    While i = JOYERR_NOERROR
        i = joyGetDevCaps(j, JoystickInfo, Len(JoystickInfo))
        If i = JOYERR_NOERROR Then
            Combo1.AddItem (CStr(j) & ": " & JoystickInfo.szPname)
        End If
        j = j + 1
    Wend
    
    If JoystickCal.id = -1 Then
        Combo1.ListIndex = 0
    Else
        i = JoystickCal.id + 1
        If i < Combo1.ListCount Then
            Combo1.ListIndex = JoystickCal.id + 1
        Else
            Combo1.ListIndex = 0
        End If
    End If
    
    Call ReadGampadProperties
        
End Sub
Private Sub ReadGampadProperties()
Dim i As Integer
Dim j As Long
Dim btnval As Long
Dim tmptxt As String
    
    NoofJbtns = 0
    
    JoystickDat.dwSize = Len(JoystickDat)
    JoystickDat.dwFlags = JOY_RETURNALL

    i = JOYERR_NOERROR
    
    If JoystickCal.id = -1 Then
    '    i = joyGetDevCaps(JOYSTICKID1, JoystickInfo, Len(JoystickInfo))
        i = joyGetPosEx(JOYSTICKID1, JoystickDat)
    
        If i <> JOYERR_NOERROR Then
    '        i = joyGetDevCaps(JOYSTICKID2, JoystickInfo, Len(JoystickInfo))
             i = joyGetPosEx(JOYSTICKID2, JoystickDat)
        End If
    Else
'        i = joyGetDevCaps(JOYSTICKcal.ID, JoystickInfo, Len(JoystickInfo))
         i = joyGetPosEx(JoystickCal.id, JoystickDat)
    End If
        
    If i <> JOYERR_NOERROR Then
        ' no joystick don't let user trash any existing data
        ApplyBtn.Enabled = False
        DefaultBtn.Enabled = False
        ClearBtn.Enabled = False
    Else
        XminLabel.Caption = CStr(JoystickCal.dwMinXpos)
        XmaxLabel.Caption = CStr(JoystickCal.dwMaxXpos)
        YminLabel.Caption = CStr(JoystickCal.dwMinYpos)
        YmaxLabel.Caption = CStr(JoystickCal.dwMaxYpos)
        ZminLabel.Caption = CStr(JoystickCal.dwMinZpos)
        ZmaxLabel.Caption = CStr(JoystickCal.dwMaxZpos)
        RminLabel.Caption = CStr(JoystickCal.dwMinRpos)
        RmaxLabel.Caption = CStr(JoystickCal.dwMaxRpos)
        Timer1.Enabled = True
        Call BuildBtns
    End If
    
    If JoystickCal.Enabled = 0 Then
        Check1.Value = 0
    Else
        Check1.Value = 1
    End If
    
    If POV_Enabled = 1 Then
        CheckPOV.Value = 1
    Else
        CheckPOV.Value = 0
    End If

    If JoystickCal.DualSpeed = 0 Then
        CheckDualSpeed.Value = 0
    Else
        CheckDualSpeed = 1
    End If


    Combo2.ListIndex = gMonitorMode

    LastIdx = 0

End Sub


Private Sub Form_Unload(Cancel As Integer)

    HC.JoyStick_Timer.Enabled = True

End Sub

Private Sub Label1_Click(Index As Integer)
    
    ' highlight selcted function
    Label1(LastIdx).BorderStyle = 0
    Label1(Index).BorderStyle = 1
    LastIdx = Index
    
End Sub


Private Sub Label15_Click()
End Sub

Private Sub Label2_Click(Index As Integer)
    Label1_Click (Index)
End Sub

Private Sub BuildBtns()

    Label2(0).Caption = GetBtnStr(BTN_EMERGENCYSTOP)
    Label2(1).Caption = GetBtnStr(BTN_HOMEPARK)
    Label2(2).Caption = GetBtnStr(BTN_USERPARK)
    Label2(3).Caption = GetBtnStr(BTN_UNPARK)
    Label2(4).Caption = GetBtnStr(BTN_STARTSIDREAL)
    Label2(5).Caption = GetBtnStr(BTN_RAREVERSE)
    Label2(6).Caption = GetBtnStr(BTN_DECREVERSE)
    Label2(7).Caption = GetBtnStr(BTN_SPIRAL)
    Label2(8).Caption = GetBtnStr(BTN_RARATEINC)
    Label2(9).Caption = GetBtnStr(BTN_RARATEDEC)
    Label2(10).Caption = GetBtnStr(BTN_DECRATEINC)
    Label2(11).Caption = GetBtnStr(BTN_DECRATEDEC)
    Label2(12).Caption = GetBtnStr(BTN_ALIGNACCEPT)
    Label2(13).Caption = GetBtnStr(BTN_NORTH)
    Label2(14).Caption = GetBtnStr(BTN_EAST)
    Label2(15).Caption = GetBtnStr(BTN_SOUTH)
    Label2(16).Caption = GetBtnStr(BTN_WEST)
    Label2(17).Caption = GetBtnStr(BTN_ALIGNCANCEL)
    Label2(18).Caption = GetBtnStr(BTN_CUSTOMTRACKSTART)
    Label2(19).Caption = GetBtnStr(BTN_CURRENTPARK)
    Label2(20).Caption = GetBtnStr(BTN_STARTLUNAR)
    Label2(21).Caption = GetBtnStr(BTN_STARTSOLAR)
    Label2(22).Caption = GetBtnStr(BTN_INCRATEPRESET)
    Label2(23).Caption = GetBtnStr(BTN_DECRATEPRESET)
    Label2(24).Caption = GetBtnStr(BTN_ALIGNEND)
    Label2(25).Caption = GetBtnStr(BTN_RATE1)
    Label2(26).Caption = GetBtnStr(BTN_RATE2)
    Label2(27).Caption = GetBtnStr(BTN_RATE3)
    Label2(28).Caption = GetBtnStr(BTN_RATE4)
    Label2(29).Caption = GetBtnStr(BTN_PEC)
    Label2(30).Caption = GetBtnStr(BTN_SYNC)
    Label2(31).Caption = GetBtnStr(BTN_NORTHEAST)
    Label2(32).Caption = GetBtnStr(BTN_NORTHWEST)
    Label2(33).Caption = GetBtnStr(BTN_SOUTHEAST)
    Label2(34).Caption = GetBtnStr(BTN_SOUTHWEST)
    Label2(35).Caption = GetBtnStr(BTN_POLARSCOPEALIGN)
    Label2(36).Caption = GetBtnStr(BTN_DEADMANSHANDLE)
    Label2(37).Caption = GetBtnStr(BTN_TOGGLELOCK)
    Label2(38).Caption = GetBtnStr(BTN_TOGGLESCREENSAVER)
End Sub
            
Private Function GetBtnStr(ByVal BtnData As Long) As String
Dim i As Integer
Dim mask As Long

    If BtnData >= 65536 Then
        ' its a point of view button
        BtnData = BtnData - 65536
        Select Case BtnData
            Case 9000
                 GetBtnStr = "POV_E"
            Case 27000
                 GetBtnStr = "POV_W"
            Case 0
                 GetBtnStr = "POV_N"
            Case 18000
                 GetBtnStr = "POV_S"
            Case 31500
                 GetBtnStr = "POV_NW"
            Case 4500
                 GetBtnStr = "POV_NE"
            Case 13500
                 GetBtnStr = "POV_SE"
            Case 22500
                 GetBtnStr = "POV_SW"
            Case Else
                GetBtnStr = "---"
        End Select
    Else
'      For i = 1 To JoystickInfo.wNumButtons
        For i = 1 To 31
            If i = 1 Then
                mask = 1
            Else
                If i = 32 Then
                    ' can't shift a signed long any further so assign number directly
                    mask = &H80000000
                Else
                    mask = mask * 2
                End If
            End If
            If mask = BtnData Then
                GetBtnStr = oLangDll.GetLangString(1134) & CStr(i)
                Exit Function
            End If
        Next i
        ' no natch found
        GetBtnStr = "---"
    End If
End Function

Private Function GetBtnData(BtnStr As String) As Long
Dim tmptxt As String
Dim i As Integer
Dim BtnNo As Integer
Dim pos As Integer

    If BtnStr = "---" Then
        GetBtnData = 0
    Else
        If Left$(BtnStr, 4) = "POV_" Then
            Select Case BtnStr
                Case "POV_N"
                    GetBtnData = 65536 + 0
                Case "POV_E"
                    GetBtnData = 65536 + 9000
                Case "POV_S"
                    GetBtnData = 65536 + 18000
                Case "POV_W"
                    GetBtnData = 65536 + 27000
                Case "POV_NW"
                    GetBtnData = 65536 + 31500
                Case "POV_NE"
                    GetBtnData = 65536 + 4500
                Case "POV_SW"
                    GetBtnData = 65536 + 22500
                Case "POV_SE"
                    GetBtnData = 65536 + 13500
                Case Else
                    GetBtnData = 0
            End Select
        Else
           pos = InStr(BtnStr, "_")
           If pos <> 0 Then
               tmptxt = Right$(BtnStr, Len(BtnStr) - pos)
               BtnNo = val(tmptxt)
    '       If BtnNo > JoystickInfo.wNumButtons Or BtnNo < 1 Then
               If BtnNo > 32 Or BtnNo < 1 Then
                   GetBtnData = 0
               Else
                   If BtnNo = 32 Then
                       ' special case when top bit is set in a signed long
                       GetBtnData = &H80000000
                   Else
                       ' get button data by shifting by the buton number
                       GetBtnData = 1
                       For i = 2 To BtnNo
                           GetBtnData = GetBtnData * 2
                       Next i
                   End If
               End If
           End If
        End If
    End If
    
End Function

Private Sub Label2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
       ' right click = clear assignment
        Label2(Index).ForeColor = &HFF&
        Label2(Index).Caption = "---"
    End If
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
       ' right click = clear assignment
        Label2(Index).ForeColor = &HFF&
        Label2(Index).Caption = "---"
    End If
End Sub

Private Sub StartBtn_Click()
    Dim quarter As Long

    If StartBtn.Caption = oLangDll.GetLangString(1209) Then
        JoystickCal.dwMinXpos = 32767
        JoystickCal.dwMaxXpos = 32767
        JoystickCal.dwMinYpos = 32767
        JoystickCal.dwMaxYpos = 32767
        JoystickCal.dwMinZpos = 32767
        JoystickCal.dwMaxZpos = 32767
        JoystickCal.dwMinRpos = 32767
        JoystickCal.dwMaxRpos = 32767
        JoystickCal.dwX25Left = 32767
        JoystickCal.dwX75left = 32767
        JoystickCal.dwX25Right = 32767
        JoystickCal.dwX75Right = 32767
        JoystickCal.dwY25Left = 32767
        JoystickCal.dwY75left = 32767
        JoystickCal.dwY25Right = 32767
        JoystickCal.dwY75Right = 32767
        Timer2.Enabled = True
        StartBtn.Caption = oLangDll.GetLangString(1208)
    Else
        Timer2.Enabled = False
        StartBtn.Caption = oLangDll.GetLangString(1209)
        quarter = (32767 - JoystickCal.dwMinXpos) / 4
        JoystickCal.dwX25Left = quarter * 3
        JoystickCal.dwX75left = quarter
        quarter = (JoystickCal.dwMinXpos - 32767) / 4
        JoystickCal.dwX25Right = 32767 + quarter
        JoystickCal.dwX75left = 32767 + (quarter * 3)
        quarter = (32767 - JoystickCal.dwMinYpos) / 4
        JoystickCal.dwY25Left = quarter * 3
        JoystickCal.dwY75left = quarter
        quarter = (JoystickCal.dwMinYpos - 32767) / 4
        JoystickCal.dwY25Right = 32767 + quarter
        JoystickCal.dwY75left = 32767 + (quarter * 3)
    End If
    
End Sub

Private Sub Timer1_Timer()
Dim i As Long
Dim BtnNo As Long
Dim Index As Integer
Dim tmptxt As String
Dim mask As Long
    
    If JoystickCal.id = -1 Then
        ' read first joystick
        i = joyGetPosEx(JOYSTICKID1, JoystickDat)
        If i <> JOYERR_NOERROR Then
            ' try second joystick
            i = joyGetPosEx(JOYSTICKID2, JoystickDat)
        End If
    Else
        i = joyGetPosEx(JoystickCal.id, JoystickDat)
    End If
        
    If i = JOYERR_NOERROR Then
'        For BtnNo = 1 To JoystickInfo.wNumButtons
        For BtnNo = 1 To 32
            If BtnNo = 1 Then
                mask = 1
            Else
                If BtnNo = 32 Then
                    ' can't shift a signed long any further so assign number directly
                    mask = &H80000000
                Else
                    ' shift mask to next button
                    mask = mask * 2
                End If
            End If
            ' check if button has been pressed
            If JoystickDat.dwButtons And mask Then
                 ' construct a label
                 tmptxt = oLangDll.GetLangString(1134) & CStr(BtnNo)
                 Label2(LastIdx).Caption = tmptxt
                 ' find any duplicates and highlight them
                 For Index = 0 To Label2.Count - 1
                    Label2(Index).ForeColor = &HFF&
                    If Index <> LastIdx Then
                        If Label2(Index).Caption = tmptxt Then
                            Label2(Index).ForeColor = &H80FFFF
                        End If
                    End If
                 Next Index
                GoTo endtimer
            End If
        Next BtnNo
        
        If POV_Enabled Then
            If JoystickDat.dwPOV <> 65535 Then
                Select Case JoystickDat.dwPOV
                    Case 9000
                         tmptxt = "POV_E"
                    Case 27000
                         tmptxt = "POV_W"
                    Case 0
                         tmptxt = "POV_N"
                    Case 18000
                         tmptxt = "POV_S"
                    Case 31500
                         tmptxt = "POV_NW"
                    Case 4500
                         tmptxt = "POV_NE"
                    Case 22500
                         tmptxt = "POV_SW"
                    Case 13500
                         tmptxt = "POV_SE"
                    Case Else
                        GoTo endtimer
                End Select
                
                Label2(LastIdx).Caption = tmptxt
                ' find any duplicates and highlight them
                For Index = 0 To Label2.Count - 1
                   Label2(Index).ForeColor = &HFF&
                   If Index <> LastIdx Then
                       If Label2(Index).Caption = tmptxt Then
                           Label2(Index).ForeColor = &H80FFFF
                       End If
                   End If
                Next Index
            End If
        End If
        
    End If
endtimer:

End Sub

Private Sub SetText()
    JStickConfigForm.Caption = oLangDll.GetLangString(1133)
    Frame1.Caption = oLangDll.GetLangString(1210)
    Label1(0).Caption = oLangDll.GetLangString(1100)
    Label1(1).Caption = oLangDll.GetLangString(1101)
    Label1(2).Caption = oLangDll.GetLangString(1102)
    Label1(19).Caption = oLangDll.GetLangString(1103)
    Label1(3).Caption = oLangDll.GetLangString(1104)
    Label1(4).Caption = oLangDll.GetLangString(1105)
    Label1(20).Caption = oLangDll.GetLangString(1106)
    Label1(21).Caption = oLangDll.GetLangString(1107)
    Label1(18).Caption = oLangDll.GetLangString(1108)
    Label1(7).Caption = oLangDll.GetLangString(1109)
    Label1(13).Caption = oLangDll.GetLangString(1110)
    Label1(14).Caption = oLangDll.GetLangString(1111)
    Label1(15).Caption = oLangDll.GetLangString(1112)
    Label1(16).Caption = oLangDll.GetLangString(1113)
    Label1(5).Caption = oLangDll.GetLangString(1114)
    Label1(6).Caption = oLangDll.GetLangString(1115)
    Label1(8).Caption = oLangDll.GetLangString(1116)
    Label1(9).Caption = oLangDll.GetLangString(1117)
    Label1(10).Caption = oLangDll.GetLangString(1118)
    Label1(11).Caption = oLangDll.GetLangString(1119)
    Label1(22).Caption = oLangDll.GetLangString(1120)
    Label1(23).Caption = oLangDll.GetLangString(1121)
    Label1(25).Caption = oLangDll.GetLangString(1122)
    Label1(26).Caption = oLangDll.GetLangString(1123)
    Label1(27).Caption = oLangDll.GetLangString(1124)
    Label1(28).Caption = oLangDll.GetLangString(1125)
    Label1(29).Caption = oLangDll.GetLangString(1135)
    Label1(12).Caption = oLangDll.GetLangString(1126)
    Label1(17).Caption = oLangDll.GetLangString(1127)
    Label1(24).Caption = oLangDll.GetLangString(1128)
    Label1(30).Caption = oLangDll.GetLangString(26)
    Label1(31).Caption = oLangDll.GetLangString(1110) & oLangDll.GetLangString(1111)
    Label1(32).Caption = oLangDll.GetLangString(1110) & oLangDll.GetLangString(1113)
    Label1(33).Caption = oLangDll.GetLangString(1112) & oLangDll.GetLangString(1111)
    Label1(34).Caption = oLangDll.GetLangString(1112) & oLangDll.GetLangString(1113)
    Label1(35).Caption = oLangDll.GetLangString(1136)
    Label1(36).Caption = oLangDll.GetLangString(1137)
    Label1(37).Caption = oLangDll.GetLangString(1138)
    Label1(38).Caption = oLangDll.GetLangString(1139)
    DefaultBtn.Caption = oLangDll.GetLangString(1129)
    ClearBtn.Caption = oLangDll.GetLangString(1130)
    CancelBtn.Caption = oLangDll.GetLangString(1131)
    ApplyBtn.Caption = oLangDll.GetLangString(1132)

    Frame2.Caption = oLangDll.GetLangString(1200)
    Label8.Caption = oLangDll.GetLangString(1201)
    Label9.Caption = oLangDll.GetLangString(1202)
    Label3.Caption = oLangDll.GetLangString(1203)
    Label4.Caption = oLangDll.GetLangString(1204)
    Label5.Caption = oLangDll.GetLangString(1205)
    Label6.Caption = oLangDll.GetLangString(1206)
    Label7.Caption = oLangDll.GetLangString(1207)
    StartBtn.Caption = oLangDll.GetLangString(1209)
    
    Frame3.Caption = oLangDll.GetLangString(1190)
    Check1.Caption = oLangDll.GetLangString(1191)
    CheckDualSpeed.Caption = oLangDll.GetLangString(1192)
    CheckPOV.Caption = oLangDll.GetLangString(1193)
    
    Combo2.Clear
    Combo2.AddItem (oLangDll.GetLangString(1212))
    Combo2.AddItem (oLangDll.GetLangString(1213))
    Label10.Caption = oLangDll.GetLangString(1211)
    
End Sub

Private Sub Timer2_Timer()
Dim i As Long
Dim dwXpos As Long
Dim dwYpos As Long
Dim dwZpos As Long
Dim dwRpos As Long
        
    JoystickDat.dwSize = Len(JoystickDat)
    JoystickDat.dwFlags = JOY_RETURNALL

    If JoystickCal.id = -1 Then
        i = joyGetPosEx(JOYSTICKID1, JoystickDat)
        If i <> JOYERR_NOERROR Then
            i = joyGetPosEx(JOYSTICKID2, JoystickDat)
        End If
    Else
        i = joyGetPosEx(JoystickCal.id, JoystickDat)
    End If
    If i <> JOYERR_NOERROR Then
        ' Joystick not found disable joystick scan
    End If
    
    dwXpos = JoystickDat.dwXpos
    dwYpos = JoystickDat.dwYpos
    dwZpos = JoystickDat.dwZpos
    dwRpos = JoystickDat.dwRpos
    
    If dwXpos > JoystickCal.dwMaxXpos Then JoystickCal.dwMaxXpos = dwXpos
    If dwYpos > JoystickCal.dwMaxYpos Then JoystickCal.dwMaxYpos = dwYpos
    If dwZpos > JoystickCal.dwMaxZpos Then JoystickCal.dwMaxZpos = dwZpos
    If dwRpos > JoystickCal.dwMaxRpos Then JoystickCal.dwMaxRpos = dwRpos
    If dwXpos < JoystickCal.dwMinXpos Then JoystickCal.dwMinXpos = dwXpos
    If dwYpos < JoystickCal.dwMinYpos Then JoystickCal.dwMinYpos = dwYpos
    If dwZpos < JoystickCal.dwMinZpos Then JoystickCal.dwMinZpos = dwZpos
    If dwRpos < JoystickCal.dwMinRpos Then JoystickCal.dwMinRpos = dwRpos
    
    XminLabel.Caption = CStr(JoystickCal.dwMinXpos)
    XmaxLabel.Caption = CStr(JoystickCal.dwMaxXpos)
    YminLabel.Caption = CStr(JoystickCal.dwMinYpos)
    YmaxLabel.Caption = CStr(JoystickCal.dwMaxYpos)
    ZminLabel.Caption = CStr(JoystickCal.dwMinZpos)
    ZmaxLabel.Caption = CStr(JoystickCal.dwMaxZpos)
    RminLabel.Caption = CStr(JoystickCal.dwMinRpos)
    RmaxLabel.Caption = CStr(JoystickCal.dwMaxRpos)
End Sub
