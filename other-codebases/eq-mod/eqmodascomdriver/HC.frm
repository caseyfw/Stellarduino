VERSION 5.00
Begin VB.Form HC 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EQMOD ASCOM DRIVER"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10305
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   10305
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      Caption         =   "ASCOM PulseGuide Settings"
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
      Height          =   2055
      Left            =   7080
      TabIndex        =   71
      Top             =   0
      Width           =   3135
      Begin VB.CheckBox decfixed_enchk 
         BackColor       =   &H00000000&
         Caption         =   "Check1"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2640
         TabIndex        =   99
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox rafixed_enchk 
         BackColor       =   &H00000000&
         Caption         =   "Check1"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1920
         TabIndex        =   98
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox decpulse_enchk 
         BackColor       =   &H00000000&
         Caption         =   "Check1"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1200
         TabIndex        =   86
         Top             =   720
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox rapulse_enchk 
         BackColor       =   &H00000000&
         Caption         =   "Check1"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   85
         Top             =   720
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.VScrollBar VScroll6 
         Height          =   975
         Left            =   2280
         Max             =   50
         Min             =   1
         TabIndex        =   79
         Top             =   960
         Value           =   1
         Width           =   255
      End
      Begin VB.VScrollBar VScroll5 
         Height          =   975
         Left            =   1560
         Max             =   50
         Min             =   1
         TabIndex        =   78
         Top             =   960
         Value           =   1
         Width           =   255
      End
      Begin VB.VScrollBar VScroll4 
         Height          =   1215
         Left            =   840
         Max             =   9
         Min             =   1
         TabIndex        =   75
         Top             =   720
         Value           =   1
         Width           =   255
      End
      Begin VB.VScrollBar VScroll3 
         Height          =   1215
         Left            =   120
         Max             =   9
         Min             =   1
         TabIndex        =   72
         Top             =   720
         Value           =   1
         Width           =   255
      End
      Begin VB.Label Label20 
         BackColor       =   &H00000040&
         Caption         =   " 0100"
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
         Left            =   2520
         TabIndex        =   84
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label19 
         BackColor       =   &H00000040&
         Caption         =   " 0100"
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
         Left            =   1800
         TabIndex        =   83
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label18 
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
         Left            =   2280
         TabIndex        =   82
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label17 
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
         Left            =   1560
         TabIndex        =   81
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "RA/DEC pulse width Overide (x100 msecs)"
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
         Height          =   495
         Left            =   1560
         TabIndex        =   80
         Top             =   240
         Width           =   1455
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
         Left            =   1080
         TabIndex        =   77
         Top             =   1680
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
         Left            =   360
         TabIndex        =   76
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Caption         =   "RA RATE"
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
         Height          =   495
         Left            =   120
         TabIndex        =   74
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label12 
         BackColor       =   &H00000000&
         Caption         =   "DEC RATE"
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
         Height          =   495
         Left            =   840
         TabIndex        =   73
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00000000&
      Caption         =   "Autoguider Port RATE"
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
      Height          =   2055
      Left            =   7080
      TabIndex        =   87
      Top             =   2040
      Width           =   3135
      Begin VB.CommandButton Command29 
         Caption         =   "x0.25"
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
         Left            =   2400
         TabIndex        =   97
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton Command28 
         Caption         =   "x0.50"
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
         Left            =   1680
         TabIndex        =   96
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton Command27 
         Caption         =   "x0.75"
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
         Left            =   960
         TabIndex        =   95
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton Command26 
         Caption         =   "x1.0"
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
         TabIndex        =   93
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton Command25 
         Caption         =   "x0.25"
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
         Left            =   2400
         TabIndex        =   91
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command24 
         Caption         =   "x0.50"
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
         Left            =   1680
         TabIndex        =   90
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command23 
         Caption         =   "x0.75"
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
         Left            =   960
         TabIndex        =   89
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command22 
         Caption         =   "x1.0"
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
         TabIndex        =   88
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label22 
         BackColor       =   &H00000000&
         Caption         =   "DEC RATE"
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
         Left            =   240
         TabIndex        =   94
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label21 
         BackColor       =   &H00000000&
         Caption         =   "RA RATE"
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
         Left            =   240
         TabIndex        =   92
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      Caption         =   "Backlash Compensation"
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
      Height          =   1335
      Left            =   7080
      TabIndex        =   101
      Top             =   4080
      Width           =   3135
      Begin VB.Label Label23 
         BackColor       =   &H00000080&
         Caption         =   "             UNDER DEVELOPMENT"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   103
         Top             =   600
         Width           =   2775
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00000000&
      Caption         =   "Custom Track Rate"
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
      Height          =   1215
      Left            =   7080
      TabIndex        =   100
      Top             =   5400
      Width           =   3135
      Begin VB.Label Label24 
         BackColor       =   &H00000080&
         Caption         =   "             UNDER DEVELOPMENT"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   104
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00000000&
      Caption         =   "PEC Training/PulseGuideData"
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
      Height          =   1935
      Left            =   7080
      TabIndex        =   102
      Top             =   6600
      Width           =   3135
      Begin VB.Label Label25 
         BackColor       =   &H00000080&
         Caption         =   "             UNDER DEVELOPMENT"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   105
         Top             =   840
         Width           =   2775
      End
   End
   Begin VB.Timer Pulseguide_Timer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2520
      Top             =   0
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "1-Star Align / Sync"
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
      Height          =   3495
      Left            =   3720
      TabIndex        =   46
      Top             =   2040
      Width           =   3255
      Begin VB.CommandButton Command20 
         Caption         =   "RESET ALIGN Data"
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
         TabIndex        =   60
         Top             =   2160
         Width           =   3015
      End
      Begin VB.CommandButton Command19 
         Caption         =   "ALIGN"
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
         Left            =   1680
         TabIndex        =   57
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton Command18 
         Caption         =   "SLEW to Star"
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
         TabIndex        =   56
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ListBox StarList 
         BackColor       =   &H00000080&
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
         Height          =   960
         ItemData        =   "HC.frx":0000
         Left            =   120
         List            =   "HC.frx":0097
         TabIndex        =   55
         Top             =   240
         Width           =   3015
      End
      Begin VB.CommandButton Command11 
         Caption         =   "RESET SYNC Data"
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
         TabIndex        =   48
         Top             =   3000
         Width           =   3015
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000000&
         Caption         =   "DxSB:"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   1680
         TabIndex        =   68
         Top             =   2640
         Width           =   525
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "DxSA:"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   120
         TabIndex        =   67
         Top             =   2640
         Width           =   525
      End
      Begin VB.Label DxSalbl 
         BackColor       =   &H00000080&
         Caption         =   "000000000"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   720
         TabIndex        =   66
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label DxSblbl 
         BackColor       =   &H00000080&
         Caption         =   "000000000"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2280
         TabIndex        =   65
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "DxB:"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   1680
         TabIndex        =   64
         Top             =   1800
         Width           =   405
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "DxA:"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   120
         TabIndex        =   63
         Top             =   1800
         Width           =   405
      End
      Begin VB.Label DxBlbl 
         BackColor       =   &H00000080&
         Caption         =   "000000000"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2280
         TabIndex        =   62
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label DxAlbl 
         BackColor       =   &H00000080&
         Caption         =   "000000000"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   720
         TabIndex        =   61
         Top             =   1800
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Home/Park/Unpark"
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
      Height          =   3015
      Left            =   3720
      TabIndex        =   49
      Top             =   5520
      Width           =   3255
      Begin VB.CommandButton Command21 
         Caption         =   "PARK to Current Position "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   70
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton Command17 
         Caption         =   "UNPARK then slew to Last Position"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1680
         TabIndex        =   54
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Define Current as NEW Park Position"
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
         Left            =   120
         TabIndex        =   53
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Command15 
         Caption         =   "UNPARK"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1680
         TabIndex        =   52
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Park to Defined Park Location"
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
         Left            =   120
         TabIndex        =   51
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton Command13 
         Caption         =   "PARK to Home Position "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   50
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label parklbl 
         BackColor       =   &H00000080&
         Caption         =   "Mount Park Status:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   69
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Timer GotoTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2040
      Top             =   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Message Center"
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
      Height          =   975
      Left            =   120
      TabIndex        =   44
      Top             =   960
      Width           =   3375
      Begin VB.TextBox HCMessage 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   45
         Text            =   "HC.frx":06B9
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame PositionFrame 
      BackColor       =   &H00000000&
      Caption         =   "Mount Position"
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
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   3375
      Begin VB.Label lblPier 
         BackColor       =   &H00000080&
         Caption         =   "Unknown"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1080
         TabIndex        =   43
         Top             =   3240
         Width           =   2175
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
         TabIndex        =   42
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label azmrk 
         BackColor       =   &H00000040&
         Caption         =   "AZ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label altmrk 
         BackColor       =   &H00000040&
         Caption         =   "ALT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label decmrk 
         BackColor       =   &H00000040&
         Caption         =   "DEC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label ramrk 
         BackColor       =   &H00000040&
         Caption         =   "RA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lstmrk 
         BackColor       =   &H00000040&
         Caption         =   "LST"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
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
         Top             =   240
         Width           =   855
      End
      Begin VB.Label azlbl 
         BackColor       =   &H00000080&
         Caption         =   "+90:00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1080
         TabIndex        =   6
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label altlbl 
         BackColor       =   &H00000080&
         Caption         =   "+90:00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1080
         TabIndex        =   5
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label lstlbl 
         BackColor       =   &H00000080&
         Caption         =   "+12:00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label declbl 
         BackColor       =   &H00000080&
         Caption         =   "+90:00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1080
         TabIndex        =   3
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label ralbl 
         BackColor       =   &H00000080&
         Caption         =   "+12:00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1080
         TabIndex        =   2
         Top             =   840
         Width           =   2175
      End
   End
   Begin VB.Frame SlewFrame 
      BackColor       =   &H00000000&
      Caption         =   "Slew Controls"
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
      Height          =   1815
      Left            =   120
      TabIndex        =   12
      Top             =   5520
      Width           =   3375
      Begin VB.CommandButton Command30 
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
         Height          =   255
         Left            =   120
         TabIndex        =   106
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton Command5 
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
         Left            =   600
         MaskColor       =   &H80000010&
         Picture         =   "HC.frx":06BF
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   720
         Width           =   495
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   1455
         Left            =   2880
         Max             =   800
         Min             =   1
         TabIndex        =   18
         Top             =   240
         Value           =   800
         Width           =   255
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1455
         Left            =   1920
         Max             =   800
         Min             =   1
         TabIndex        =   17
         Top             =   240
         Value           =   800
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Caption         =   "W"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         MaskColor       =   &H80000010&
         TabIndex        =   16
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         MaskColor       =   &H80000010&
         TabIndex        =   15
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         MaskColor       =   &H80000010&
         TabIndex        =   14
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         MaskColor       =   &H80000010&
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         Caption         =   "DEC RATE"
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
         Left            =   2400
         TabIndex        =   22
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Caption         =   "RA RATE"
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
         Left            =   1440
         TabIndex        =   21
         Top             =   240
         Width           =   375
      End
      Begin VB.Label decslewlbl 
         BackColor       =   &H00000040&
         Caption         =   " 800"
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
         Left            =   2400
         TabIndex        =   20
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label raslewlbl 
         BackColor       =   &H00000040&
         Caption         =   " 800"
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
         Left            =   1440
         TabIndex        =   19
         Top             =   1440
         Width           =   375
      End
   End
   Begin VB.Timer EncoderTimer 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   1560
      Top             =   0
   End
   Begin VB.Timer DisplayTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Top             =   0
   End
   Begin VB.PictureBox picASCOM 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   840
      Left            =   120
      MouseIcon       =   "HC.frx":10B5
      MousePointer    =   99  'Custom
      Picture         =   "HC.frx":1207
      ScaleHeight     =   840
      ScaleWidth      =   720
      TabIndex        =   41
      Top             =   120
      Width           =   720
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Site Information"
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
      Height          =   2025
      Left            =   3720
      TabIndex        =   29
      Top             =   0
      Width           =   3255
      Begin VB.ComboBox cbhem 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         ItemData        =   "HC.frx":20CB
         Left            =   1200
         List            =   "HC.frx":20D5
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   1440
         Width           =   1155
      End
      Begin VB.CommandButton Command12 
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
         Height          =   735
         Left            =   2400
         TabIndex        =   40
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtElevation 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   1200
         TabIndex        =   36
         Text            =   "100"
         Top             =   1080
         Width           =   1125
      End
      Begin VB.ComboBox cbNS 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         ItemData        =   "HC.frx":20E7
         Left            =   1200
         List            =   "HC.frx":20F1
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   360
         Width           =   555
      End
      Begin VB.ComboBox cbEW 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         ItemData        =   "HC.frx":20FB
         Left            =   1200
         List            =   "HC.frx":2105
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   720
         Width           =   555
      End
      Begin VB.TextBox txtLongMin 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   2505
         TabIndex        =   33
         Text            =   "57"
         Top             =   720
         Width           =   570
      End
      Begin VB.TextBox txtLongDeg 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   1920
         TabIndex        =   32
         Text            =   "120"
         Top             =   720
         Width           =   600
      End
      Begin VB.TextBox txtLatMin 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   2505
         TabIndex        =   31
         Text            =   "35"
         Top             =   360
         Width           =   570
      End
      Begin VB.TextBox txtLatDeg 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   1920
         TabIndex        =   30
         Text            =   "14"
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Hemisphere:"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   120
         TabIndex        =   58
         Top             =   1560
         Width           =   990
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "Elevation (m):"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   1155
         Width           =   990
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "Latitude:"
         ForeColor       =   &H000080FF&
         Height          =   225
         Left            =   165
         TabIndex        =   38
         Top             =   360
         Width           =   690
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         Caption         =   "Longitude:"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   765
         Width           =   765
      End
   End
   Begin VB.CommandButton Command10 
      Caption         =   "SETUP >>>>>>>"
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
      TabIndex        =   26
      Top             =   8160
      Width           =   3375
   End
   Begin VB.Frame TrackingFrame 
      BackColor       =   &H00000000&
      Caption         =   "Track Rate : Not Tracking "
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
      Height          =   735
      Left            =   120
      TabIndex        =   23
      Top             =   7320
      Width           =   3375
      Begin VB.CommandButton Command6 
         Caption         =   "Sidreal"
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
         TabIndex        =   47
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command9 
         Caption         =   "STOP"
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
         Left            =   2640
         TabIndex        =   27
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Solar"
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
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Lunar"
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
         Left            =   960
         TabIndex        =   24
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label MainLabel 
      BackColor       =   &H00000080&
      Caption         =   "EQMOD ASCOM DRIVER ver1.02"
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
      TabIndex        =   0
      Top             =   120
      Width           =   2535
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

Public GotoTimerFlag As Boolean
Public EncoderTimerFlag As Boolean
Private hUtil As DriverHelper.Util




Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If gEQparkstatus = 1 Then
            HC.Add_Message ("Cannot slew, mount still parked. Please unpark.")
            Exit Sub
    End If
    
    eqres = EQ_MotorStop(1)          ' Stop DEC Motor
    If eqres <> EQ_OK Then
            GoTo END13
    End If

   ' Wait for motor stop

   Do
        
       eqres = EQ_GetMotorStatus(1)
       If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo END13
     
   Loop While (eqres And EQ_MOTORBUSY) <> 0
        
    
   eqres = EQ_Slew(1, 0, 0, val(VScroll2.Value))
   Add_Message ("Slewing North ...")
    
END13:

End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    eqres = EQ_MotorStop(1)
   
End Sub

Private Sub Command10_Click()
    If Command10.Caption = "SETUP >>>>>>>" Then
        Command10.Caption = "SETUP <<<<<<<"
        HC.ScaleWidth = 10305
        HC.Width = 10395
        
    Else
        Command10.Caption = "SETUP >>>>>>>"
        HC.ScaleWidth = 3570
        HC.Width = 3660
    End If
End Sub


Private Sub Command11_Click()

    Call resetsync
    
End Sub

Private Sub Command12_Click()
    gLongitude = CDbl(txtLongDeg) + (CDbl(txtLongMin) / 60#)
    If cbEW.Text = "W" Then gLongitude = -gLongitude  ' W is neg
    gLatitude = CDbl(txtLatDeg) + (CDbl(txtLatMin) / 60#)
    If cbNS.Text = "S" Then gLatitude = -gLatitude
    gElevation = CDbl(txtElevation)
    
    If cbhem.Text = "North" Then
        gHemisphere = 0
    Else
        gHemisphere = 1
    End If
    
End Sub


Private Sub Command13_Click()

    TrackingFrame.Caption = "Track Rate : Not Tracking"
    ParkHome
    HC.Add_Message ("Scope parked to Home. You may turn it off after slewing")
  
    gEQparkstatus = 1
    HC.parklbl.Caption = "Mount Park Status: PARKED"
  
End Sub

Private Sub Command14_Click()

    TrackingFrame.Caption = "Track Rate : Not Tracking"
    ParktoUserDefine (True)
    HC.Add_Message ("Scope parked. You may turn it off after slewing")
    gEQparkstatus = 1
    HC.parklbl.Caption = "Mount Park Status: PARKED"

End Sub

Private Sub Command15_Click()

    If EQ_GetMountStatus() = 1 Then     ' Make sure that we unpark only if the mount is online
    
        TrackingFrame.Caption = "Track Rate : Not Tracking"
        Unparkscope
        gEQparkstatus = 0
        HC.parklbl.Caption = "Mount Park Status: NOT PARKED"
    
    End If
    
End Sub

Private Sub Command16_Click()

    DefinePark (True)
    HC.Add_Message ("Current location set as new park position.")

End Sub

Private Sub Command17_Click()

    If EQ_GetMountStatus() = 1 Then     ' Make sure that we unpark only if the mount is online
        
        TrackingFrame.Caption = "Track Rate : Not Tracking"
        UnparkscopeToLastPos
        gEQparkstatus = 0
        HC.parklbl.Caption = "Mount Park Status: NOT PARKED"
    
    End If
    
End Sub

Private Sub Command18_Click()

    gTargetRA = hUtil.HMSToHours(Mid(StarList.Text, 3, 10))
    gTargetDec = hUtil.DMSToDegrees(Mid(StarList.Text, 14, 10))

    HC.Add_Message ("SyncCSlew: RA[ " & FmtSexa(gTargetRA, False) & " ] DEC[ " & FmtSexa(gTargetDec, True) & " ]")
   
    gSlewCount = NUM_SLEW_RETRIES               'Set initial iterative slew count
    Call radecAsyncSlew(False)                  'Call the slew
    

End Sub

Private Sub Command19_Click()

'    MsgBox "Please center your scope to " & Mid(StarList.Text, 25) & "using the dir buttons", _
'                (vbOKOnly + vbCritical + vbMsgBoxSetForeground), App.FileDescription

   Load Align
   Align.Show
      
End Sub

Private Sub Command20_Click()

     gRA1Star = 0
     gDEC1Star = 0
     
     WriteAlignMap
     
     HC.DxAlbl.Caption = Format$(Str(gRA1Star), "000000000")
     HC.DxBlbl.Caption = Format$(Str(gDEC1Star), "000000000")
     
End Sub

Private Sub Command21_Click()

    DefinePark (True)
    TrackingFrame.Caption = "Track Rate : Not Tracking"
    ParktoUserDefine (False) ' No need to slew mount
    HC.Add_Message ("Scope parked. You may turn it off after slewing")
    gEQparkstatus = 1
    HC.parklbl.Caption = "Mount Park Status: PARKED"

End Sub

Private Sub Command22_Click()
    eqres = EQ_SetAutoguiderPortRate(0, 3)
    If eqres = EQ_OK Then
        HC.Add_Message ("RA guider port rate set to x1.00")
    Else
        HC.Add_Message ("EQMOD COM Port error: Cannot set port rate")
    End If
End Sub

Private Sub Command23_Click()

    eqres = EQ_SetAutoguiderPortRate(0, 2)
    If eqres = EQ_OK Then
        HC.Add_Message ("RA guider port rate set to x0.75")
    Else
        HC.Add_Message ("EQMOD COM Port error: Cannot set port rate")
    End If

End Sub

Private Sub Command24_Click()
    eqres = EQ_SetAutoguiderPortRate(0, 1)
    If eqres = EQ_OK Then
        HC.Add_Message ("RA guider port rate set to x0.50")
    Else
        HC.Add_Message ("EQMOD COM Port error: Cannot set port rate")
    End If
End Sub

Private Sub Command25_Click()
    eqres = EQ_SetAutoguiderPortRate(0, 0)
    If eqres = EQ_OK Then
        HC.Add_Message ("RA guider port rate set to x0.25")
    Else
        HC.Add_Message ("EQMOD COM Port error: Cannot set port rate")
    End If
End Sub

Private Sub Command26_Click()
    eqres = EQ_SetAutoguiderPortRate(1, 3)
    If eqres = EQ_OK Then
        HC.Add_Message ("DEC guider port rate set to x1.00")
    Else
        HC.Add_Message ("EQMOD COM Port error: Cannot set port rate")
    End If
End Sub

Private Sub Command27_Click()
    eqres = EQ_SetAutoguiderPortRate(1, 2)
    If eqres = EQ_OK Then
        HC.Add_Message ("DEC guider port rate set to x0.75")
    Else
        HC.Add_Message ("EQMOD COM Port error: Cannot set port rate")
    End If
End Sub

Private Sub Command28_Click()
    eqres = EQ_SetAutoguiderPortRate(1, 1)
    If eqres = EQ_OK Then
        HC.Add_Message ("DEC guider port rate set to x0.50")
    Else
        HC.Add_Message ("EQMOD COM Port error: Cannot set port rate")
    End If
End Sub

Private Sub Command29_Click()
    eqres = EQ_SetAutoguiderPortRate(1, 0)
    If eqres = EQ_OK Then
        HC.Add_Message ("DEC guider port rate set to x0.25")
    Else
        HC.Add_Message ("EQMOD COM Port error: Cannot set port rate")
    End If
End Sub

Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
   If gEQparkstatus = 1 Then
            HC.Add_Message ("Cannot slew, mount still parked. Please unpark.")
            Exit Sub
   End If
  
   eqres = EQ_MotorStop(0)          ' Stop RA Motor

   If eqres <> EQ_OK Then
            GoTo END05
   End If

    'Wait until RA motor is stable

   Do
       eqres = EQ_GetMotorStatus(0)
       If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo END05
   Loop While (eqres And EQ_MOTORBUSY) <> 0


   eqres = EQ_Slew(0, 0, 1, val(VScroll1.Value))
   Add_Message ("Slewing East ...")

END05:
End Sub

Private Sub Command3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    eqres = EQ_MotorStop(0)          ' Stop RA Motor
    If eqres <> 0 Then
            GoTo END02
    End If


    Do
        
       eqres = EQ_GetMotorStatus(0)
       If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo END02
  
    Loop While (eqres And EQ_MOTORBUSY) <> 0
    
   
    If gTrackingStatus > 0 Then
    
        eqres = EQ_StartRATrack(gTrackingStatus - 1, gHemisphere, gHemisphere)           ' Track RA Motor at Sidreal Rate
        Add_Message ("RA motor set back to normal tracking rate.")
        
    End If

END02:
End Sub



Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If gEQparkstatus = 1 Then
            HC.Add_Message ("Cannot slew, mount still parked. Please unpark.")
            Exit Sub
    End If
    
    eqres = EQ_MotorStop(1)          ' Stop DEC Motor
    If eqres <> EQ_OK Then
            GoTo END12
    End If

   ' Wait for motor stop

    Do
        
       eqres = EQ_GetMotorStatus(1)
       If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo END12
    
    Loop While (eqres And EQ_MOTORBUSY) <> 0
    
    eqres = EQ_Slew(1, 0, 1, val(VScroll2.Value))
    Add_Message ("Slewing South ...")

END12:
End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     eqres = EQ_MotorStop(1)
'    Add_Message ("Done.")

End Sub


Private Sub Command30_Click()

    Load Slewpad
    Slewpad.Show
    
End Sub

Private Sub Command4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If gEQparkstatus = 1 Then
            HC.Add_Message ("Cannot slew, mount still parked. Please unpark.")
            Exit Sub
    End If
    
    eqres = EQ_MotorStop(0)          ' Stop RA Motor
    If eqres <> EQ_OK Then
            GoTo END04
    End If

    'Wait until RA motor is stable

    Do
       eqres = EQ_GetMotorStatus(0)
       If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo END04

    Loop While (eqres And EQ_MOTORBUSY) <> 0
   
    eqres = EQ_Slew(0, 0, 0, val(VScroll1.Value))
    Add_Message ("Slewing West ...")
    
END04:
End Sub

Private Sub Command4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    eqres = EQ_MotorStop(0)          ' Stop RA Motor
    If eqres <> 0 Then
            GoTo END01
    End If

    'Wait until RA motor is stable

    Do
       eqres = EQ_GetMotorStatus(0)
       If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo END01

    Loop While (eqres And EQ_MOTORBUSY) <> 0

    If gTrackingStatus <> 0 Then
    
        eqres = EQ_StartRATrack(gTrackingStatus - 1, gHemisphere, gHemisphere)
        Add_Message ("RA motor set back to normal tracking rate.")
    
    End If
  
END01:
End Sub

Private Sub Command5_Click()
    eqres = EQ_MotorStop(0)
    eqres = EQ_MotorStop(1)
    gSlewStatus = False
    gRAStatus_slew = False
    HC.GotoTimer.Enabled = False
    gTrackingStatus = 0
    HC.TrackingFrame.Caption = "Track Rate : Not Tracking"
    Add_Message ("RA and DEC Motors stopped.")
    
 End Sub



Private Sub Command6_Click()
                
        If gEQparkstatus = 1 Then
        
            HC.Add_Message ("Cannot start tracking, mount still parked. Please unpark.")
            Exit Sub
        End If
                
        eqres = EQ_StartRATrack(0, gHemisphere, gHemisphere)
        gTrackingStatus = 1                 'Sidreal rate tracking'
        HC.TrackingFrame.Caption = "Track Rate : Tracking at Sidreal"
        Add_Message ("RA Motor set at Sidereal rate")
    
End Sub

Private Sub Command7_Click()
        
        If gEQparkstatus = 1 Then
            HC.Add_Message ("Cannot start tracking, mount still parked. Please unpark.")
            Exit Sub
        End If
        
        eqres = EQ_StartRATrack(1, gHemisphere, gHemisphere)
        gTrackingStatus = 2                 'Lunar rate tracking'
        HC.TrackingFrame.Caption = "Track Rate : Tracking at Lunar"
        Add_Message ("RA Motor set at Lunar rate")
    
End Sub

Private Sub Command8_Click()

        If gEQparkstatus = 1 Then
            HC.Add_Message ("Cannot start tracking, mount still parked. Please unpark.")
            Exit Sub
        End If
        
        eqres = EQ_StartRATrack(2, gHemisphere, gHemisphere)
        gTrackingStatus = 3                 'Solar rate tracking'
        HC.TrackingFrame.Caption = "Track Rate : Tracking at Solar"
        Add_Message ("RA Motor set at Solar rate")
End Sub

Private Sub Command9_Click()
        eqres = EQ_MotorStop(0)
        gTrackingStatus = 0
        HC.TrackingFrame.Caption = "Track Rate : Not Tracking"
        Add_Message ("RA Motor stopped.")
End Sub

Private Sub EncoderTimer_Timer()


    If EncoderTimerFlag Then
    
        EncoderTimerFlag = False            'Avoid overruns
        gRA_Encoder = Delta_RA_Map(EQ_GetMotorValues(0))
        If (gRA_Encoder < &H1000000) Then gRA_Hours = Get_EncoderHours(gRAEncoder_Zero_pos, gRA_Encoder, gTot_RA, gHemisphere)

        ' Convert DEC_Encoder to DEC Degrees

        gDec_Encoder = Delta_DEC_Map(EQ_GetMotorValues(1))
        gDec_DegNoAdjust = Get_EncoderDegrees(gDECEncoder_Zero_pos, gDec_Encoder, gTot_DEC, gHemisphere)
        If (gDec_Encoder < &H1000000) Then gDec_Degrees = Range_DEC(gDec_DegNoAdjust)

        gRAStatus = EQ_GetMotorStatus(0)
        gDECStatus = EQ_GetMotorStatus(1)
    
        EncoderTimerFlag = True

    End If

End Sub
Private Sub DisplayTimer_Timer()
    
    Dim tha As Double
    Dim tRA As Double
    Dim tDEC As Double
    Dim tAlt As Double
    Dim tAz As Double
    Dim i As Double
 
    EncoderTimer.Enabled = False
 
    ' convert dHA which is in Radians to true RA and scale it

    tRA = EQnow_lst(gLongitude * DEG_RAD) + gRA_Hours

    If gHemisphere = 0 Then
        If (gDec_DegNoAdjust > 90) And (gDec_DegNoAdjust <= 270) Then tRA = tRA - 12
    Else
        If (gDec_DegNoAdjust <= 90) Or (gDec_DegNoAdjust > 270) Then tRA = tRA + 12
    End If
    
    tRA = Range24(tRA)


    gRA = tRA
    gDec = gDec_Degrees
    
 
    hadec_aa (gLatitude * DEG_RAD), ((tRA - EQnow_lst(gLongitude * DEG_RAD)) * HRS_RAD), (gDec_Degrees * DEG_RAD), tAlt, tAz

    gAlt = tAlt * RAD_DEG ' tAlt was in Radians
    gAz = 360# - (tAz * RAD_DEG) ' tAz was in Radians
    
       
    lstlbl.Caption = FmtSexa(EQnow_lst(gLongitude * DEG_RAD), False)
    
    ralbl.Caption = FmtSexa(gRA, False)
    declbl.Caption = FmtSexa(gDec, True)

    azlbl.Caption = FmtSexa(gAz, False)
    altlbl.Caption = FmtSexa(gAlt, False)
  
    Select Case SOP_RAHours(gRA_Hours)
                Case pierUnknown:    lblPier.Caption = "Unknown"
                Case PierEast:       lblPier.Caption = "East, Scope pointing West"
                Case PierWest:       lblPier.Caption = "West, Scope pointing East"
    End Select
  
    EncoderTimer.Enabled = True

End Sub



Private Sub Form_Load()
    EnableCloseButton Me.hWnd, False
    HC.ScaleWidth = 3570
    HC.Width = 3660
    Set hUtil = New DriverHelper.Util
    
    readratebarstateHC
    
    GotoTimerFlag = True
    EncoderTimerFlag = True
    StarList.Text = "A 01:37:42.8 -57:14:12 Achernar"
End Sub



'Timer Tick Routine to Monitor Asynchronous Slew Functions
'RA motor should be set to tracking mode right after a goto function

Private Sub GotoTimer_Timer()
    

    Dim tRA As Double
    Dim tha As Double
    Dim tdiff As Double
    
    If GotoTimerFlag Then
        GotoTimerFlag = False               'Avoid Overruns
    

        If gSlewStatus Then
            
              If (gRAStatus And EQ_MOTORBUSY) = 0 Then
            
                'At This point RA motor has completed the slew
            
                gRAStatus_slew = True
                
                ' Should we track back to the usual tracking rate ?
                If gTrackingStatus > 0 Then
                    eqres = EQ_StartRATrack(gTrackingStatus - 1, gHemisphere, gHemisphere)           ' Track RA Motor at Sidreal Rate
                End If
                
            End If

        
     
            If (gDECStatus And EQ_MOTORBUSY) = 0 And gRAStatus_slew Then
        
                'DEC and RA motors are not moving at this point
                'First iteration slew is complete
                'We need to check if it requires a new slew to reduce the difference
                'Caused by the Movement of the earth during the slew process
        
                gSlewCount = gSlewCount - 1                 ' Allow only NUM_SLEW_RETRIES
        
                If (gSlewCount > 0) And (gTrackingStatus > 0) Then  ' Retry only if tracking is enabled
                    tha = RangeHA(gTargetRA - EQnow_lst(gLongitude * DEG_RAD))
                    If tha < 0 Then
                        tRA = Range24(gTargetRA - 12)
                    Else
                        tRA = gTargetRA
                    End If

                    'Compute for Target RA Encoder Difference

                    tdiff = Abs(Get_RAEncoderfromRA(tRA, 0, gLongitude, gRAEncoder_Zero_pos, gTot_RA, gHemisphere) - gRA_Encoder)
                    
                    If tdiff < gRA_Allowed_diff Then
                        Add_Message ("Goto Slew Complete. Diff at " & Str(tdiff))
                        gSlewStatus = False
                        gRAStatus_slew = False
                        GotoTimer.Enabled = False
                    Else
                        'Re Execute a new RA-Only slew here
                        Add_Message ("SlewRetry[" & Str(gSlewCount) & "]: Diff at :" & Str(tdiff) & " Target is: <" & Str(gRA_Allowed_diff))
                        Call radecAsyncSlew(True)
                        
                    End If
               Else
                    Add_Message ("Goto Slew Complete.")
                    gSlewStatus = False
                    gRAStatus_slew = False
                   GotoTimer.Enabled = False
                
               End If
            End If
        End If

        GotoTimerFlag = True
    End If
End Sub




Private Sub Pulseguide_Timer_Timer()

    ' This is a 100 millisecond time ticker interval for Pulse guide

    If gTrackingStatus > 0 And gEQPulsetimerflag Then  ' Guide only when the scope is tracking
        gEQPulsetimerflag = False                      ' Prevent Overruns
        
       
        If gEQDECPulseDuration > 0 Then
            gEQDECPulseDuration = gEQDECPulseDuration - 100
            If gEQDECPulseDuration <= 0 Then
                eqres = EQ_MotorStop(1)
                gEQDECPulseDuration = 0
            End If
                
        End If
        If gEQRAPulseDuration > 0 Then
            gEQRAPulseDuration = gEQRAPulseDuration - 100
            If gEQRAPulseDuration <= 0 Then
                eqres = EQ_SendGuideRate(0, gTrackingStatus - 1, 0, 0, gHemisphere, gHemisphere)
                gEQRAPulseDuration = 0
            End If
        End If
        gEQPulsetimerflag = True
    End If

End Sub





Private Sub VScroll1_Change()
    raslewlbl.Caption = VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
    raslewlbl.Caption = VScroll1.Value
End Sub

Private Sub VScroll2_Change()
    decslewlbl.Caption = VScroll2.Value
End Sub

Private Sub VScroll2_Scroll()
    decslewlbl.Caption = VScroll2.Value
End Sub

Public Sub Add_Message(dtaLog As String)

    'Add the message, limit to 1000 characters

    HCMessage.Text = Right(HCMessage.Text & vbCrLf & dtaLog, 1000)
    HCMessage.SelStart = Len(HCMessage.Text)

End Sub

Private Sub VScroll3_Change()
    HC.Label14.Caption = Format$(Str(VScroll3.Value * 0.1), "x0.00")
End Sub

Private Sub VScroll3_Scroll()
    HC.Label14.Caption = Format$(Str(VScroll3.Value * 0.1), "x0.00")
End Sub

Private Sub VScroll4_Change()
    HC.Label15.Caption = Format$(Str(VScroll4.Value * 0.1), "x0.00")
End Sub

Private Sub VScroll4_Scroll()
    HC.Label15.Caption = Format$(Str(VScroll4.Value * 0.1), "x0.00")
End Sub

Private Sub VScroll5_Change()
    HC.Label19.Caption = Format$(Str(VScroll5.Value * 100), " 0000")
End Sub

Private Sub VScroll5_Scroll()
    HC.Label19.Caption = Format$(Str(VScroll5.Value * 100), " 0000")
End Sub

Private Sub VScroll6_Change()
    HC.Label20.Caption = Format$(Str(VScroll6.Value * 100), " 0000")
End Sub

Private Sub VScroll6_Scroll()
    HC.Label20.Caption = Format$(Str(VScroll6.Value * 100), " 0000")
End Sub
