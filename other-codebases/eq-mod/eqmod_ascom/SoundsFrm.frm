VERSION 5.00
Begin VB.Form SoundsFrm 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Set Sounds"
   ClientHeight    =   7125
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   9870
   Icon            =   "SoundsFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   475
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   658
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdLoadDecReverseOff 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   7800
      Picture         =   "SoundsFrm.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   123
      Top             =   2220
      Width           =   350
   End
   Begin VB.CheckBox ChkReverse 
      BackColor       =   &H00000000&
      Caption         =   "Slew Reverse"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   7440
      TabIndex        =   122
      TabStop         =   0   'False
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton CmdLoadRaReverseOn 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   7800
      Picture         =   "SoundsFrm.frx":124C
      Style           =   1  'Graphical
      TabIndex        =   121
      Top             =   600
      Width           =   350
   End
   Begin VB.CommandButton CmdLoadRaReverseOff 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   7800
      Picture         =   "SoundsFrm.frx":17CE
      Style           =   1  'Graphical
      TabIndex        =   120
      Top             =   1155
      Width           =   350
   End
   Begin VB.CommandButton CmdLoadDecReverseOn 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   7800
      Picture         =   "SoundsFrm.frx":1D50
      Style           =   1  'Graphical
      TabIndex        =   119
      Top             =   1680
      Width           =   350
   End
   Begin VB.CommandButton CmdMonitorOn 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   7920
      Picture         =   "SoundsFrm.frx":22D2
      Style           =   1  'Graphical
      TabIndex        =   113
      Top             =   3960
      Width           =   350
   End
   Begin VB.CommandButton CmdMonitorOff 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   7920
      Picture         =   "SoundsFrm.frx":2854
      Style           =   1  'Graphical
      TabIndex        =   112
      Top             =   4320
      Width           =   350
   End
   Begin VB.CheckBox ChkGPL 
      BackColor       =   &H00000000&
      Caption         =   "GamePad Lock"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   7560
      TabIndex        =   107
      TabStop         =   0   'False
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton CmdGPLOn 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   7920
      Picture         =   "SoundsFrm.frx":2DD6
      Style           =   1  'Graphical
      TabIndex        =   106
      Top             =   2880
      Width           =   350
   End
   Begin VB.CommandButton CmdGPLOff 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   7920
      Picture         =   "SoundsFrm.frx":3358
      Style           =   1  'Graphical
      TabIndex        =   105
      Top             =   3240
      Width           =   350
   End
   Begin VB.CommandButton CmdLoadDMS2 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   5280
      Picture         =   "SoundsFrm.frx":38DA
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   5820
      Width           =   350
   End
   Begin VB.CommandButton CmdLoadDMS 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   5280
      Picture         =   "SoundsFrm.frx":3E5C
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   5280
      Width           =   350
   End
   Begin VB.CheckBox ChkDMS 
      BackColor       =   &H00000000&
      Caption         =   "Dead Man's Switch"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   4920
      TabIndex        =   98
      TabStop         =   0   'False
      Top             =   4800
      Width           =   2295
   End
   Begin VB.CommandButton CmdLoadPComplete 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   5280
      Picture         =   "SoundsFrm.frx":43DE
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   4260
      Width           =   350
   End
   Begin VB.CommandButton CmdLoadPAlign 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   5280
      Picture         =   "SoundsFrm.frx":4960
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   3690
      Width           =   350
   End
   Begin VB.CommandButton CmdLoadPHome 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   5280
      Picture         =   "SoundsFrm.frx":4EE2
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   3120
      Width           =   350
   End
   Begin VB.CommandButton CmdLoadEnd 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   5280
      Picture         =   "SoundsFrm.frx":5464
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   1680
      Width           =   350
   End
   Begin VB.CommandButton CmdLoadCancel 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   5280
      Picture         =   "SoundsFrm.frx":59E6
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   1155
      Width           =   350
   End
   Begin VB.CommandButton CmdLoadAccept 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   5280
      Picture         =   "SoundsFrm.frx":5F68
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   600
      Width           =   350
   End
   Begin VB.CheckBox ChkPolar 
      BackColor       =   &H00000000&
      Caption         =   "Polar Scope Alignment"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   4920
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CheckBox ChkAlign 
      BackColor       =   &H00000000&
      Caption         =   "Alignment"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   4920
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton CmdLoadGotoStart 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   120
      Picture         =   "SoundsFrm.frx":64EA
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   2895
      Width           =   350
   End
   Begin VB.CheckBox ChkGotoStart 
      BackColor       =   &H00000000&
      Caption         =   "Goto Start"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton CmdLoadCustom 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   2760
      Picture         =   "SoundsFrm.frx":6A6C
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   2220
      Width           =   350
   End
   Begin VB.CommandButton CmdLoadLunar 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   2760
      Picture         =   "SoundsFrm.frx":6FEE
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   1680
      Width           =   350
   End
   Begin VB.CommandButton CmdLoadSolar 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   2760
      Picture         =   "SoundsFrm.frx":7570
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   1155
      Width           =   350
   End
   Begin VB.CommandButton CmdLoadSidereal 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   2760
      Picture         =   "SoundsFrm.frx":7AF2
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   600
      Width           =   350
   End
   Begin VB.CheckBox ChkTracking 
      BackColor       =   &H00000000&
      Caption         =   "Tracking"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2400
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton CmdLoadUnpark 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   120
      Picture         =   "SoundsFrm.frx":8074
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   5295
      Width           =   350
   End
   Begin VB.CheckBox ChkUnpark 
      BackColor       =   &H00000000&
      Caption         =   "Unpark"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton CmdApply 
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton CmdLoadParked 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   120
      Picture         =   "SoundsFrm.frx":85F6
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   4680
      Width           =   350
   End
   Begin VB.CheckBox ChkParked 
      BackColor       =   &H00000000&
      Caption         =   "Parked"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CheckBox ChkStop 
      BackColor       =   &H00000000&
      Caption         =   "Stop"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   5625
      Width           =   1935
   End
   Begin VB.CommandButton CmdLoadStop 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   120
      Picture         =   "SoundsFrm.frx":8B78
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   5880
      Width           =   350
   End
   Begin VB.CheckBox ChkPark 
      BackColor       =   &H00000000&
      Caption         =   "Parking"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   3825
      Width           =   1935
   End
   Begin VB.CommandButton CmdLoadPark 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   120
      Picture         =   "SoundsFrm.frx":90FA
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   4080
      Width           =   350
   End
   Begin VB.CheckBox ChkGoto 
      BackColor       =   &H00000000&
      Caption         =   "Goto Complete"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   3225
      Width           =   1935
   End
   Begin VB.CommandButton CmdLoadGoto 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   120
      Picture         =   "SoundsFrm.frx":967C
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   3480
      Width           =   350
   End
   Begin VB.CommandButton CmdLoadSync 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   5280
      Picture         =   "SoundsFrm.frx":9BFE
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   2220
      Width           =   350
   End
   Begin VB.CommandButton CmdLoadRate 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Index           =   9
      Left            =   2760
      Picture         =   "SoundsFrm.frx":A180
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6120
      Width           =   350
   End
   Begin VB.CommandButton CmdLoadRate 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Index           =   8
      Left            =   2760
      Picture         =   "SoundsFrm.frx":A702
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5760
      Width           =   350
   End
   Begin VB.CommandButton CmdLoadRate 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Index           =   7
      Left            =   2760
      Picture         =   "SoundsFrm.frx":AC84
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5400
      Width           =   350
   End
   Begin VB.CommandButton CmdLoadRate 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Index           =   6
      Left            =   2760
      Picture         =   "SoundsFrm.frx":B206
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5040
      Width           =   350
   End
   Begin VB.CommandButton CmdLoadRate 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Index           =   5
      Left            =   2760
      Picture         =   "SoundsFrm.frx":B788
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4680
      Width           =   350
   End
   Begin VB.CommandButton CmdLoadRate 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Index           =   4
      Left            =   2760
      Picture         =   "SoundsFrm.frx":BD0A
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4320
      Width           =   350
   End
   Begin VB.CommandButton CmdLoadRate 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Index           =   3
      Left            =   2760
      Picture         =   "SoundsFrm.frx":C28C
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3960
      Width           =   350
   End
   Begin VB.CommandButton CmdLoadRate 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Index           =   2
      Left            =   2760
      Picture         =   "SoundsFrm.frx":C80E
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3600
      Width           =   350
   End
   Begin VB.CommandButton CmdLoadRate 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Index           =   1
      Left            =   2760
      Picture         =   "SoundsFrm.frx":CD90
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3240
      Width           =   350
   End
   Begin VB.CheckBox ChkRate 
      BackColor       =   &H00000000&
      Caption         =   "Rate Presets"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton CmdLoadRate 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Index           =   0
      Left            =   2760
      Picture         =   "SoundsFrm.frx":D312
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2880
      Width           =   350
   End
   Begin VB.ComboBox ComboMode 
      BackColor       =   &H00000080&
      ForeColor       =   &H000080FF&
      Height          =   315
      ItemData        =   "SoundsFrm.frx":D894
      Left            =   120
      List            =   "SoundsFrm.frx":D89E
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton CmdLoadAlarm 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   120
      Picture         =   "SoundsFrm.frx":D8BA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2040
      Width           =   350
   End
   Begin VB.CommandButton CmdLoadClick 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   120
      Picture         =   "SoundsFrm.frx":DE3C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1440
      Width           =   350
   End
   Begin VB.CommandButton CmdLoadBeep 
      BackColor       =   &H0095C1CB&
      Height          =   250
      Left            =   120
      Picture         =   "SoundsFrm.frx":E3BE
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   840
      Width           =   350
   End
   Begin VB.CheckBox CheckAlarm 
      BackColor       =   &H00000000&
      Caption         =   "Alarm"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1785
      Width           =   1935
   End
   Begin VB.CheckBox ChkClick 
      BackColor       =   &H00000000&
      Caption         =   "Button Click"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1185
      Width           =   1935
   End
   Begin VB.CheckBox ChkBeep 
      BackColor       =   &H00000000&
      Caption         =   "Beep"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   585
      Width           =   1935
   End
   Begin VB.CommandButton CancelButton 
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H0095C1CB&
      Caption         =   "OK"
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
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CheckBox ChkMonitor 
      BackColor       =   &H00000000&
      Caption         =   "Monitor"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   7560
      TabIndex        =   114
      TabStop         =   0   'False
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label LabelDecReverseOff 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   8160
      TabIndex        =   131
      Top             =   2220
      Width           =   1575
   End
   Begin VB.Label Label25 
      BackColor       =   &H00000000&
      Caption         =   "RA On"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   7800
      TabIndex        =   130
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label LabelRaReverseOn 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   8160
      TabIndex        =   129
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label30 
      BackColor       =   &H00000000&
      Caption         =   "RA Off"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   7800
      TabIndex        =   128
      Top             =   900
      Width           =   1935
   End
   Begin VB.Label LabelRaReverseOff 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   8160
      TabIndex        =   127
      Top             =   1155
      Width           =   1575
   End
   Begin VB.Label LabelReverseDecOn 
      BackColor       =   &H00000000&
      Caption         =   "DEC On"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   7800
      TabIndex        =   126
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label LabelDecReverseOn 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   8160
      TabIndex        =   125
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label LabelReverseDecOff 
      BackColor       =   &H00000000&
      Caption         =   "DEC Off"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   7800
      TabIndex        =   124
      Top             =   1980
      Width           =   1935
   End
   Begin VB.Label LabelMonitorOn 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   8280
      TabIndex        =   118
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label19 
      BackColor       =   &H00000000&
      Caption         =   "On"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   7560
      TabIndex        =   117
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label LabelMonitorOff 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   8280
      TabIndex        =   116
      Top             =   4335
      Width           =   1455
   End
   Begin VB.Label Label14 
      BackColor       =   &H00000000&
      Caption         =   "Off"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   7560
      TabIndex        =   115
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label LabelGPLOn 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   8280
      TabIndex        =   111
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label16 
      BackColor       =   &H00000000&
      Caption         =   "On"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   7560
      TabIndex        =   110
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label LabelGPLOff 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   8280
      TabIndex        =   109
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "Off"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   7560
      TabIndex        =   108
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "Disarmed"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5280
      TabIndex        =   104
      Top             =   5580
      Width           =   1935
   End
   Begin VB.Label LabelDMS2 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5640
      TabIndex        =   103
      Top             =   5820
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "Armed"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5280
      TabIndex        =   101
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label LabelDMS 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5640
      TabIndex        =   100
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label LabelPComplete 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5640
      TabIndex        =   97
      Top             =   4260
      Width           =   1575
   End
   Begin VB.Label Label17 
      BackColor       =   &H00000000&
      Caption         =   "Alignment Complete"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5280
      TabIndex        =   96
      Top             =   3990
      Width           =   1935
   End
   Begin VB.Label LabelPAlign 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5640
      TabIndex        =   94
      Top             =   3690
      Width           =   1575
   End
   Begin VB.Label Label15 
      BackColor       =   &H00000000&
      Caption         =   "Aligning Polar Scope"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5280
      TabIndex        =   93
      Top             =   3420
      Width           =   1935
   End
   Begin VB.Label LabelPHome 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5640
      TabIndex        =   91
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000000&
      Caption         =   "Move to Home"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5280
      TabIndex        =   90
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      Caption         =   "Sync"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5280
      TabIndex        =   88
      Top             =   1980
      Width           =   1935
   End
   Begin VB.Label LabelEnd 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5640
      TabIndex        =   87
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Caption         =   "End"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5280
      TabIndex        =   86
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label LabelCancel 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5640
      TabIndex        =   84
      Top             =   1155
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "Cancel"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5280
      TabIndex        =   83
      Top             =   900
      Width           =   1935
   End
   Begin VB.Label LabelAccept 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5640
      TabIndex        =   81
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Accept"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5280
      TabIndex        =   80
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label LabelGotoStart 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   480
      TabIndex        =   76
      Top             =   2895
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Sidereal"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2760
      TabIndex        =   73
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Custom"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2760
      TabIndex        =   72
      Top             =   1980
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Lunar"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2760
      TabIndex        =   71
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Solar"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2760
      TabIndex        =   70
      Top             =   900
      Width           =   1935
   End
   Begin VB.Label LabelCustom 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   3120
      TabIndex        =   69
      Top             =   2220
      Width           =   1575
   End
   Begin VB.Label LabelLunar 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   3120
      TabIndex        =   67
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label LabelSolar 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   3120
      TabIndex        =   65
      Top             =   1155
      Width           =   1575
   End
   Begin VB.Label LabelSidereal 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   3120
      TabIndex        =   63
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label LabelUnpark 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   480
      TabIndex        =   60
      Top             =   5295
      Width           =   1575
   End
   Begin VB.Label LabelParked 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   240
      Left            =   480
      TabIndex        =   56
      Top             =   4695
      Width           =   1575
   End
   Begin VB.Label LabelStop 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   480
      TabIndex        =   53
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label LabelPark 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   480
      TabIndex        =   50
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label LabelGoto 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   480
      TabIndex        =   47
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label LabelSync 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5640
      TabIndex        =   44
      Top             =   2220
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "10"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   9
      Left            =   2400
      TabIndex        =   42
      Top             =   6120
      Width           =   285
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "9"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   8
      Left            =   2400
      TabIndex        =   41
      Top             =   5760
      Width           =   285
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "8"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   7
      Left            =   2400
      TabIndex        =   40
      Top             =   5400
      Width           =   285
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "7"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   6
      Left            =   2400
      TabIndex        =   39
      Top             =   5040
      Width           =   285
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "6"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   5
      Left            =   2400
      TabIndex        =   38
      Top             =   4680
      Width           =   285
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "5"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   4
      Left            =   2400
      TabIndex        =   37
      Top             =   4320
      Width           =   285
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "4"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   36
      Top             =   3960
      Width           =   285
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "3"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   35
      Top             =   3600
      Width           =   285
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "2"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   34
      Top             =   3240
      Width           =   285
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "1"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   33
      Top             =   2880
      Width           =   285
   End
   Begin VB.Label LabelRate 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   9
      Left            =   3120
      TabIndex        =   32
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label LabelRate 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   8
      Left            =   3120
      TabIndex        =   30
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label LabelRate 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   7
      Left            =   3120
      TabIndex        =   28
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label LabelRate 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   6
      Left            =   3120
      TabIndex        =   26
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label LabelRate 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   5
      Left            =   3120
      TabIndex        =   24
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label LabelRate 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   4
      Left            =   3120
      TabIndex        =   22
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label LabelRate 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   20
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label LabelRate 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   18
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label LabelRate 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   16
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label LabelRate 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   14
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label LabelAlarm 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label LabelClick 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label LabelBeep 
      BackColor       =   &H00000040&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "SoundsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
        Unload SoundsFrm
End Sub


Private Sub CmdApply_Click()
    Dim i As Integer
    
    With EQSounds
        .mode = ComboMode.ListIndex
        If .mode = 1 Then
            .BeepWav = LabelBeep.ToolTipText
            .AlarmWav = LabelAlarm.ToolTipText
            .ClickWav = LabelClick.ToolTipText
            .SyncWav = LabelSync.ToolTipText
            .ParkWav = LabelPark.ToolTipText
            .ParkedWav = LabelParked.ToolTipText
            .GotoWav = LabelGoto.ToolTipText
            .GotoStartWav = LabelGotoStart.ToolTipText
            .StopWav = LabelStop.ToolTipText
            .Unparkwav = LabelUnpark.ToolTipText
            .SiderealWav = LabelSidereal.ToolTipText
            .SolarWav = LabelSolar.ToolTipText
            .LunarWav = LabelLunar.ToolTipText
            .CustomWav = LabelCustom.ToolTipText
            .AcceptWav = LabelAccept.ToolTipText
            .CancelWav = LabelCancel.ToolTipText
            .EndWav = LabelEnd.ToolTipText
            .PHomeWav = LabelPHome.ToolTipText
            .PAlignwav = LabelPAlign.ToolTipText
            .PAlignedwav = LabelPComplete.ToolTipText
            .DMSwav = LabelDMS.ToolTipText
            .DMS2wav = LabelDMS2.ToolTipText
            .GPLOnwav = LabelGPLOn.ToolTipText
            .GPLOffwav = LabelGPLOff.ToolTipText
            .MonitorOnwav = LabelMonitorOn.ToolTipText
            .MonitorOffwav = LabelMonitorOff.ToolTipText
            .RaReverseOffwav = LabelRaReverseOff.ToolTipText
            .RAReverseOnwav = LabelRaReverseOn.ToolTipText
            .DecReverseOffwav = LabelDecReverseOff.ToolTipText
            .DecReverseOnwav = LabelDecReverseOn.ToolTipText
            
            For i = 1 To 10
                .RateWav(i) = LabelRate(i - 1).ToolTipText
            Next i
            
        End If
        .PositionBeep = ChkBeep.Value
        .ButtonClick = ChkClick.Value
        .FlipWarning = CheckAlarm.Value
        .RateClick = ChkRate.Value
        .ParkClick = ChkPark.Value
        .ParkedClick = ChkParked.Value
        .GotoClick = ChkGoto.Value
        .GotoStartClick = ChkGotoStart.Value
        .Stopclick = ChkStop.Value
        .Unparkclick = ChkUnpark.Value
        .TrackClick = ChkTracking.Value
        .AlignClick = ChkAlign.Value
        .PolarClick = ChkPolar.Value
        .DMSClick = ChkDMS.Value
        .GPLClick = ChkGPL.Value
        .MonitorClick = ChkMonitor.Value
        .ReverseClick = ChkReverse.Value
        
    End With
    ' write to ini
    Call writeBeep

End Sub

Private Sub CmdLoadAccept_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelAccept.Caption = FileDlg.filename2
        LabelAccept.ToolTipText = FileDlg.filename
    End If

End Sub

Private Sub CmdLoadAlarm_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelAlarm.Caption = FileDlg.filename2
        LabelAlarm.ToolTipText = FileDlg.filename
    End If
End Sub

Private Sub CmdLoadBeep_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelBeep.Caption = FileDlg.filename2
        LabelBeep.ToolTipText = FileDlg.filename
    End If
End Sub

Private Sub CmdLoadCancel_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelCancel.Caption = FileDlg.filename2
        LabelCancel.ToolTipText = FileDlg.filename
    End If

End Sub

Private Sub CmdLoadClick_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelClick.Caption = FileDlg.filename2
        LabelClick.ToolTipText = FileDlg.filename
    End If

End Sub

Private Sub CmdLoadCustom_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelCustom.Caption = FileDlg.filename2
        LabelCustom.ToolTipText = FileDlg.filename
    End If

End Sub

Private Sub CmdGPLOn_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelGPLOn.Caption = FileDlg.filename2
        LabelGPLOn.ToolTipText = FileDlg.filename
    End If
End Sub

Private Sub CmdGPLOff_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelGPLOff.Caption = FileDlg.filename2
        LabelGPLOff.ToolTipText = FileDlg.filename
    End If
End Sub

Private Sub CmdLoadDMS_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelDMS.Caption = FileDlg.filename2
        LabelDMS.ToolTipText = FileDlg.filename
    End If
End Sub

Private Sub CmdLoadDMS2_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelDMS2.Caption = FileDlg.filename2
        LabelDMS2.ToolTipText = FileDlg.filename
    End If
End Sub

Private Sub CmdLoadEnd_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelEnd.Caption = FileDlg.filename2
        LabelEnd.ToolTipText = FileDlg.filename
    End If

End Sub

Private Sub CmdLoadGoto_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelGoto.Caption = FileDlg.filename2
        LabelGoto.ToolTipText = FileDlg.filename
    End If

End Sub

Private Sub CmdLoadGotoStart_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelGotoStart.Caption = FileDlg.filename2
        LabelGotoStart.ToolTipText = FileDlg.filename
    End If

End Sub

Private Sub CmdLoadLunar_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelLunar.Caption = FileDlg.filename2
        LabelLunar.ToolTipText = FileDlg.filename
    End If

End Sub

Private Sub CmdLoadPAlign_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelPAlign.Caption = FileDlg.filename2
        LabelPAlign.ToolTipText = FileDlg.filename
    End If

End Sub

Private Sub CmdLoadPark_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelPark.Caption = FileDlg.filename2
        LabelPark.ToolTipText = FileDlg.filename
    End If

End Sub

Private Sub CmdLoadPComplete_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelPComplete.Caption = FileDlg.filename2
        LabelPComplete.ToolTipText = FileDlg.filename
    End If

End Sub

Private Sub CmdLoadPHome_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelPHome.Caption = FileDlg.filename2
        LabelPHome.ToolTipText = FileDlg.filename
    End If

End Sub

Private Sub CmdLoadSidereal_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelSidereal.Caption = FileDlg.filename2
        LabelSidereal.ToolTipText = FileDlg.filename
    End If

End Sub

Private Sub CmdLoadSolar_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelSolar.Caption = FileDlg.filename2
        LabelSolar.ToolTipText = FileDlg.filename
    End If

End Sub

Private Sub CmdLoadSync_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelSync.Caption = FileDlg.filename2
        LabelSync.ToolTipText = FileDlg.filename
    End If

End Sub

Private Sub CmdLoadParked_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelParked.Caption = FileDlg.filename2
        LabelParked.ToolTipText = FileDlg.filename
    End If

End Sub

Private Sub CmdLoadStop_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelStop.Caption = FileDlg.filename2
        LabelStop.ToolTipText = FileDlg.filename
    End If

End Sub

Private Sub CmdLoadUnpark_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelUnpark.Caption = FileDlg.filename2
        LabelUnpark.ToolTipText = FileDlg.filename
    End If

End Sub

Private Sub CmdMonitorOff_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelMonitorOff.Caption = FileDlg.filename2
        LabelMonitorOff.ToolTipText = FileDlg.filename
    End If
End Sub

Private Sub CmdMonitorOn_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelMonitorOn.Caption = FileDlg.filename2
        LabelMonitorOn.ToolTipText = FileDlg.filename
    End If
End Sub

Private Sub ComboMode_Click()
    Call RefreshControls
End Sub

Private Sub CmdLoadRate_Click(Index As Integer)
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelRate(Index).Caption = FileDlg.filename2
        LabelRate(Index).ToolTipText = FileDlg.filename
    End If
End Sub

Private Sub CmdLoadRaReverseOff_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelRaReverseOff.Caption = FileDlg.filename2
        LabelRaReverseOff.ToolTipText = FileDlg.filename
    End If
End Sub

Private Sub CmdLoadRaReverseOn_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelRaReverseOn.Caption = FileDlg.filename2
        LabelRaReverseOn.ToolTipText = FileDlg.filename
    End If
End Sub

Private Sub CmdLoadDecReverseOff_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelDecReverseOff.Caption = FileDlg.filename2
        LabelDecReverseOff.ToolTipText = FileDlg.filename
    End If
End Sub

Private Sub CmdLoadDecReverseOn_Click()
    FileDlg.filter = "*.wav*"
    FileDlg.Show (1)
    If FileDlg.filename <> "" Then
        LabelDecReverseOn.Caption = FileDlg.filename2
        LabelDecReverseOn.ToolTipText = FileDlg.filename
    End If
End Sub

Private Sub Form_Load()
    If HC.HCOnTop.Value = 1 Then Call PutWindowOnTop(SoundsFrm)
    Call SetText
    
    ComboMode.ListIndex = EQSounds.mode
    ComboMode.Text = ComboMode.List(ComboMode.ListIndex)
    
    Call RefreshControls
    
    
End Sub

Private Sub SetText()
    ComboMode.Clear
    ComboMode.AddItem (oLangDll.GetLangString(2601))
    ComboMode.AddItem (oLangDll.GetLangString(2602))
    ComboMode.AddItem (oLangDll.GetLangString(2606))
    ChkBeep.Caption = oLangDll.GetLangString(2605)
    ChkClick.Caption = oLangDll.GetLangString(2603)
    CheckAlarm.Caption = oLangDll.GetLangString(2604)
  
    ChkGotoStart.Caption = oLangDll.GetLangString(2802)
    ChkGoto.Caption = oLangDll.GetLangString(2803)
    ChkPark.Caption = oLangDll.GetLangString(2804)
    ChkParked.Caption = oLangDll.GetLangString(2805)
    ChkUnpark.Caption = oLangDll.GetLangString(2809)
    ChkStop.Caption = oLangDll.GetLangString(2807)
    ChkTracking.Caption = oLangDll.GetLangString(121)
    Label5.Caption = oLangDll.GetLangString(122)
    Label2.Caption = oLangDll.GetLangString(123)
    Label1.Caption = oLangDll.GetLangString(124)
    Label4.Caption = oLangDll.GetLangString(2806)
    ChkRate.Caption = oLangDll.GetLangString(2810)
    ChkAlign.Caption = oLangDll.GetLangString(2815)
    Label6.Caption = oLangDll.GetLangString(504)
    Label8.Caption = oLangDll.GetLangString(402)
    Label10.Caption = oLangDll.GetLangString(505)
    Label12.Caption = oLangDll.GetLangString(2816)
    ChkPolar.Caption = oLangDll.GetLangString(2811)
    Label13.Caption = oLangDll.GetLangString(2812)
    Label15.Caption = oLangDll.GetLangString(2813)
    Label17.Caption = oLangDll.GetLangString(2814)
    
    ChkDMS.Caption = oLangDll.GetLangString(1137)
    Label7.Caption = oLangDll.GetLangString(6109)
    Label11.Caption = oLangDll.GetLangString(6110)
    
    ChkGPL.Caption = oLangDll.GetLangString(6112)
    Label16.Caption = oLangDll.GetLangString(6114)
    Label9.Caption = oLangDll.GetLangString(6115)
    
    ChkMonitor.Caption = oLangDll.GetLangString(6113)
    Label19.Caption = oLangDll.GetLangString(6114)
    Label14.Caption = oLangDll.GetLangString(6115)
    
    
    SoundsFrm.Caption = oLangDll.GetLangString(2801)
    CancelButton.Caption = oLangDll.GetLangString(402)
    OKButton.Caption = oLangDll.GetLangString(401)
    CmdApply.Caption = oLangDll.GetLangString(1132)
  
    
    
End Sub

Private Sub RefreshControls()
    Dim i As Integer
    
    If ComboMode.ListIndex = 0 Then
        CmdLoadBeep.Enabled = False
        CmdLoadClick.Enabled = False
        CmdLoadAlarm.Enabled = False
        CmdLoadSync.Enabled = False
        CmdLoadPark.Enabled = False
        CmdLoadParked.Enabled = False
        CmdLoadGoto.Enabled = False
        CmdLoadGotoStart.Enabled = False
        CmdLoadStop.Enabled = False
        CmdLoadUnpark.Enabled = False
        CmdLoadSidereal.Enabled = False
        CmdLoadSolar.Enabled = False
        CmdLoadLunar.Enabled = False
        CmdLoadCustom.Enabled = False
        CmdLoadAccept.Enabled = False
        CmdLoadCancel.Enabled = False
        CmdLoadEnd.Enabled = False
        CmdLoadPHome.Enabled = False
        CmdLoadPAlign.Enabled = False
        CmdLoadPComplete.Enabled = False
        CmdLoadDMS.Enabled = False
        CmdLoadDMS2.Enabled = False
        CmdGPLOn.Enabled = False
        CmdGPLOff.Enabled = False
        CmdMonitorOn.Enabled = False
        CmdMonitorOff.Enabled = False
        CmdLoadRaReverseOn.Enabled = False
        CmdLoadRaReverseOff.Enabled = False
        CmdLoadDecReverseOn.Enabled = False
        CmdLoadDecReverseOff.Enabled = False
        LabelBeep.Caption = ""
        LabelAlarm.Caption = ""
        LabelClick.Caption = ""
        LabelSync.Caption = ""
        LabelPark.Caption = ""
        LabelParked.Caption = ""
        LabelGoto.Caption = ""
        LabelGotoStart.Caption = ""
        LabelStop.Caption = ""
        LabelUnpark.Caption = ""
        LabelSidereal.Caption = ""
        LabelSolar.Caption = ""
        LabelLunar.Caption = ""
        LabelCustom.Caption = ""
        LabelEnd.Caption = ""
        LabelAccept.Caption = ""
        LabelCancel.Caption = ""
        LabelPHome.Caption = ""
        LabelPAlign.Caption = ""
        LabelPComplete.Caption = ""
        LabelDMS.Caption = ""
        LabelDMS2.Caption = ""
        LabelGPLOn.Caption = ""
        LabelGPLOff.Caption = ""
        LabelMonitorOn.Caption = ""
        LabelMonitorOff.Caption = ""

        For i = 1 To 10
            LabelRate(i - 1).Caption = ""
            CmdLoadRate(i - 1).Enabled = False
        Next i
    Else
        CmdLoadBeep.Enabled = True
        CmdLoadClick.Enabled = True
        CmdLoadAlarm.Enabled = True
        CmdLoadSync.Enabled = True
        CmdLoadPark.Enabled = True
        CmdLoadParked.Enabled = True
        CmdLoadGoto.Enabled = True
        CmdLoadGotoStart.Enabled = True
        CmdLoadStop.Enabled = True
        CmdLoadUnpark.Enabled = True
        CmdLoadSidereal.Enabled = True
        CmdLoadSolar.Enabled = True
        CmdLoadLunar.Enabled = True
        CmdLoadCustom.Enabled = True
        CmdLoadAccept.Enabled = True
        CmdLoadCancel.Enabled = True
        CmdLoadEnd.Enabled = True
        CmdLoadPHome.Enabled = True
        CmdLoadPAlign.Enabled = True
        CmdLoadPComplete.Enabled = True
        CmdLoadDMS.Enabled = True
        CmdLoadDMS2.Enabled = True
        CmdGPLOn.Enabled = True
        CmdGPLOff.Enabled = True
        CmdMonitorOn.Enabled = True
        CmdMonitorOff.Enabled = True
        CmdLoadRaReverseOn.Enabled = True
        CmdLoadRaReverseOff.Enabled = True
        CmdLoadDecReverseOn.Enabled = True
        CmdLoadDecReverseOff.Enabled = True
        
        LabelBeep.ToolTipText = EQSounds.BeepWav
        LabelBeep.Caption = StripPath(EQSounds.BeepWav)
        LabelAlarm.ToolTipText = EQSounds.AlarmWav
        LabelAlarm.Caption = StripPath(EQSounds.AlarmWav)
        LabelClick.ToolTipText = EQSounds.ClickWav
        LabelClick.Caption = StripPath(EQSounds.ClickWav)
        LabelSync.Caption = StripPath(EQSounds.SyncWav)
        LabelSync.ToolTipText = EQSounds.SyncWav
        LabelPark.Caption = StripPath(EQSounds.ParkWav)
        LabelPark.ToolTipText = EQSounds.ParkWav
        LabelParked.Caption = StripPath(EQSounds.ParkedWav)
        LabelParked.ToolTipText = EQSounds.ParkedWav
        LabelGoto.Caption = StripPath(EQSounds.GotoWav)
        LabelGoto.ToolTipText = EQSounds.GotoWav
        LabelGotoStart.Caption = StripPath(EQSounds.GotoStartWav)
        LabelGotoStart.ToolTipText = EQSounds.GotoStartWav
        LabelStop.Caption = StripPath(EQSounds.StopWav)
        LabelStop.ToolTipText = EQSounds.StopWav
        LabelUnpark.Caption = StripPath(EQSounds.Unparkwav)
        LabelUnpark.ToolTipText = EQSounds.Unparkwav
        LabelSidereal.Caption = StripPath(EQSounds.SiderealWav)
        LabelSidereal.ToolTipText = EQSounds.SiderealWav
        LabelSolar.Caption = StripPath(EQSounds.SolarWav)
        LabelSolar.ToolTipText = EQSounds.SolarWav
        LabelLunar.Caption = StripPath(EQSounds.LunarWav)
        LabelLunar.ToolTipText = EQSounds.LunarWav
        LabelCustom.Caption = StripPath(EQSounds.CustomWav)
        LabelCustom.ToolTipText = EQSounds.CustomWav
        LabelEnd.Caption = StripPath(EQSounds.EndWav)
        LabelEnd.ToolTipText = EQSounds.EndWav
        LabelAccept.Caption = StripPath(EQSounds.AcceptWav)
        LabelAccept.ToolTipText = EQSounds.AcceptWav
        LabelCancel.Caption = StripPath(EQSounds.CancelWav)
        LabelCancel.ToolTipText = EQSounds.CancelWav
        LabelPHome.Caption = StripPath(EQSounds.PHomeWav)
        LabelPHome.ToolTipText = EQSounds.PHomeWav
        LabelPAlign.Caption = StripPath(EQSounds.PAlignwav)
        LabelPAlign.ToolTipText = EQSounds.PAlignwav
        LabelPComplete.Caption = StripPath(EQSounds.PAlignedwav)
        LabelPComplete.ToolTipText = EQSounds.PAlignedwav
        LabelDMS.Caption = StripPath(EQSounds.DMSwav)
        LabelDMS.ToolTipText = EQSounds.DMSwav
        LabelDMS2.Caption = StripPath(EQSounds.DMS2wav)
        LabelDMS2.ToolTipText = EQSounds.DMS2wav
        LabelGPLOn.Caption = StripPath(EQSounds.GPLOnwav)
        LabelGPLOn.ToolTipText = EQSounds.GPLOnwav
        LabelGPLOff.Caption = StripPath(EQSounds.GPLOffwav)
        LabelGPLOff.ToolTipText = EQSounds.GPLOffwav
        LabelMonitorOn.Caption = StripPath(EQSounds.MonitorOnwav)
        LabelMonitorOn.ToolTipText = EQSounds.MonitorOnwav
        LabelMonitorOff.Caption = StripPath(EQSounds.MonitorOffwav)
        LabelMonitorOff.ToolTipText = EQSounds.MonitorOffwav
        LabelRaReverseOff.ToolTipText = EQSounds.RaReverseOffwav
        LabelRaReverseOff.Caption = StripPath(EQSounds.RaReverseOffwav)
        LabelRaReverseOn.ToolTipText = EQSounds.RAReverseOnwav
        LabelRaReverseOn.Caption = StripPath(EQSounds.RAReverseOnwav)
        LabelDecReverseOff.ToolTipText = EQSounds.DecReverseOffwav
        LabelDecReverseOff.Caption = StripPath(EQSounds.DecReverseOffwav)
        LabelDecReverseOn.ToolTipText = EQSounds.DecReverseOnwav
        LabelDecReverseOn.Caption = StripPath(EQSounds.DecReverseOnwav)
        
        For i = 1 To 10
            LabelRate(i - 1).Caption = StripPath(EQSounds.RateWav(i))
            LabelRate(i - 1).ToolTipText = EQSounds.RateWav(i)
            CmdLoadRate(i - 1).Enabled = True
        Next i
    End If
    
    If EQSounds.PositionBeep Then
        ChkBeep.Value = 1
    Else
        ChkBeep.Value = 0
    End If
    If EQSounds.ButtonClick Then
        ChkClick.Value = 1
    Else
        ChkClick.Value = 0
    End If
    If EQSounds.FlipWarning Then
        CheckAlarm.Value = 1
    Else
        CheckAlarm.Value = 0
    End If
    If EQSounds.ParkClick Then
        ChkPark.Value = 1
    Else
        ChkPark.Value = 0
    End If
    If EQSounds.ParkedClick Then
        ChkParked.Value = 1
    Else
        ChkParked.Value = 0
    End If
    If EQSounds.GotoClick Then
        ChkGoto.Value = 1
    Else
        ChkGoto.Value = 0
    End If
    If EQSounds.GotoStartClick Then
        ChkGotoStart.Value = 1
    Else
        ChkGotoStart.Value = 0
    End If
    If EQSounds.Stopclick Then
        ChkStop.Value = 1
    Else
        ChkStop.Value = 0
    End If
    If EQSounds.RateClick Then
        ChkRate.Value = 1
    Else
        ChkRate.Value = 0
    End If
    If EQSounds.Unparkclick Then
        ChkUnpark.Value = 1
    Else
        ChkUnpark.Value = 0
    End If
    If EQSounds.TrackClick Then
        ChkTracking.Value = 1
    Else
        ChkTracking.Value = 0
    End If
    If EQSounds.AlignClick Then
        ChkAlign.Value = 1
    Else
        ChkAlign.Value = 0
    End If
    If EQSounds.PolarClick Then
        ChkPolar.Value = 1
    Else
        ChkPolar.Value = 0
    End If
    If EQSounds.DMSClick Then
        ChkDMS.Value = 1
    Else
        ChkDMS.Value = 0
    End If
    If EQSounds.GPLClick Then
        ChkGPL.Value = 1
    Else
        ChkGPL.Value = 0
    End If
    If EQSounds.MonitorClick Then
        ChkMonitor.Value = 1
    Else
        ChkMonitor.Value = 0
    End If
    If EQSounds.ReverseClick Then
        ChkReverse.Value = 1
    Else
        ChkReverse.Value = 0
    End If

End Sub

Private Sub OKButton_Click()
    CmdApply_Click
    Unload SoundsFrm
End Sub



