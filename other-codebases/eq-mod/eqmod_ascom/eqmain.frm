VERSION 5.00
Begin VB.Form EQMOD 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EQMOD"
   ClientHeight    =   10170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14460
   LinkTopic       =   "eqmain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10170
   ScaleWidth      =   14460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command42 
      Caption         =   "Autoread DEC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   150
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command41 
      Caption         =   "Autoread RA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   149
      Top             =   2040
      Width           =   975
   End
   Begin VB.Timer Timer_readmotor 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   7080
   End
   Begin VB.CommandButton Command40 
      Caption         =   "Set DEC 0.75x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      TabIndex        =   142
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command39 
      Caption         =   "Set DEC 0.25x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      TabIndex        =   141
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command38 
      Caption         =   "Set DEC 0.50x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      TabIndex        =   140
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command37 
      Caption         =   "Set DEC 1.00x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      TabIndex        =   139
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Text33 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   12120
      TabIndex        =   136
      Text            =   "0"
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton Command36 
      Caption         =   "Set RA  0.25x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      TabIndex        =   134
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton Command35 
      Caption         =   "Set RA 0.50x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      TabIndex        =   133
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton Command34 
      Caption         =   "Set RA 0.75x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      TabIndex        =   132
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command33 
      Caption         =   "Set RA  1.00x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      TabIndex        =   131
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command32 
      Caption         =   "Get Mount Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   130
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text32 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   1320
      TabIndex        =   120
      Text            =   "&H800000"
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox Text31 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   1320
      TabIndex        =   119
      Text            =   "&H800000"
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command31 
      Caption         =   "Get Mount Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   108
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Get RA microstep"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   107
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Get DEC microstep"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   106
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Clear LOG Window"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   9720
      TabIndex        =   105
      Top             =   7920
      Width           =   4695
   End
   Begin VB.ListBox Loglst 
      BackColor       =   &H80000000&
      Height          =   2205
      ItemData        =   "eqmain.frx":0000
      Left            =   120
      List            =   "eqmain.frx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   104
      Top             =   7920
      Width           =   9495
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Track DEC Custom"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   103
      Top             =   3480
      Width           =   975
   End
   Begin VB.HScrollBar HScroll4 
      Height          =   255
      Left            =   9720
      Max             =   300
      Min             =   -300
      TabIndex        =   100
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Track RA Custom"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   99
      Top             =   2760
      Width           =   975
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      Left            =   10920
      Max             =   9
      TabIndex        =   98
      Top             =   4800
      Value           =   5
      Width           =   975
   End
   Begin VB.CommandButton Command25 
      Caption         =   "STOP MOTORS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   13320
      TabIndex        =   96
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text30 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   6120
      TabIndex        =   93
      Text            =   "0"
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox Text29 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   4920
      TabIndex        =   92
      Text            =   "0"
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command24 
      Caption         =   "SET DEC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   91
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command23 
      Caption         =   "SET RA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   90
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text28 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   10920
      TabIndex        =   88
      Text            =   "0"
      Top             =   7200
      Width           =   975
   End
   Begin VB.TextBox Text27 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   3720
      TabIndex        =   81
      Text            =   "0"
      Top             =   7200
      Width           =   975
   End
   Begin VB.TextBox Text26 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   2520
      TabIndex        =   80
      Text            =   "0"
      Top             =   7200
      Width           =   975
   End
   Begin VB.CommandButton Command22 
      Caption         =   "STOP DEC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3720
      TabIndex        =   79
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command21 
      Caption         =   "STOP RA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2520
      TabIndex        =   78
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text25 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   9720
      TabIndex        =   60
      Text            =   "0"
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox Text24 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   8520
      TabIndex        =   58
      Text            =   "0"
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox Text23 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   7320
      TabIndex        =   56
      Text            =   "0"
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox Text22 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   13320
      TabIndex        =   54
      Text            =   "0"
      Top             =   6360
      Width           =   975
   End
   Begin VB.TextBox Text21 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   10920
      TabIndex        =   53
      Text            =   "0"
      Top             =   6360
      Width           =   975
   End
   Begin VB.TextBox Text20 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   9720
      TabIndex        =   52
      Text            =   "0"
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Guide DEC(-)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10920
      TabIndex        =   51
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Guide DEC(+)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10920
      TabIndex        =   50
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Guide  RA(-)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10920
      TabIndex        =   49
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Guide RA(+)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10920
      TabIndex        =   48
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text19 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   8520
      TabIndex        =   47
      Text            =   "0"
      Top             =   7200
      Width           =   975
   End
   Begin VB.TextBox Text18 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   8520
      TabIndex        =   46
      Text            =   "0"
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Track Lunar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   45
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Track Solar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   44
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Track Sidreal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   43
      Top             =   600
      Width           =   975
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   8520
      Max             =   800
      Min             =   1
      TabIndex        =   42
      Top             =   2640
      Value           =   300
      Width           =   975
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Slew DEC(-)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8520
      TabIndex        =   41
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Slew DEC(+)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8520
      TabIndex        =   40
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text17 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   7320
      TabIndex        =   39
      Text            =   "0"
      Top             =   7200
      Width           =   975
   End
   Begin VB.TextBox Text16 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   7320
      TabIndex        =   38
      Text            =   "0"
      Top             =   6360
      Width           =   975
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   7320
      Max             =   800
      Min             =   1
      TabIndex        =   37
      Top             =   2640
      Value           =   400
      Width           =   975
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Slew RA(-)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      TabIndex        =   36
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Slew RA(+)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      TabIndex        =   35
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text15 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   6120
      TabIndex        =   34
      Text            =   "0"
      Top             =   7200
      Width           =   975
   End
   Begin VB.TextBox Text14 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   4920
      TabIndex        =   33
      Text            =   "0"
      Top             =   7200
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Read DEC Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   32
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Read RA Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   31
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   6120
      TabIndex        =   30
      Text            =   "0"
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Read DEC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   29
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   4920
      TabIndex        =   28
      Text            =   "0"
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Read RA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   27
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   3720
      TabIndex        =   26
      Text            =   "0"
      Top             =   6360
      Width           =   975
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   2520
      TabIndex        =   25
      Text            =   "0"
      Top             =   6360
      Width           =   975
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   3720
      TabIndex        =   24
      Text            =   "0"
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   3720
      TabIndex        =   21
      Text            =   "123456"
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Move DEC Steps(-)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   20
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Move DEC Steps(+)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   19
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   2520
      TabIndex        =   18
      Text            =   "0"
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   2520
      TabIndex        =   16
      Text            =   "123456"
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Move RA Steps(-)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   14
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Move RA Steps(+)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   13
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   1320
      TabIndex        =   12
      Text            =   "0"
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Initialize and Activate Stepper Motors"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   1320
      TabIndex        =   11
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Text            =   "1"
      Top             =   5520
      Width           =   975
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Text            =   "1000"
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Text            =   "9600"
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "COM1"
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txt001 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Text            =   "0"
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "CLOSE    EQ   DRIVER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   13320
      TabIndex        =   1
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdConnect 
      Appearance      =   0  'Flat
      Caption         =   "CONNECT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label72 
      BackColor       =   &H0080FF80&
      Caption         =   "  1 sec Diff"
      Height          =   255
      Left            =   6120
      TabIndex        =   154
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label71 
      BackColor       =   &H0080FF80&
      Caption         =   "  1 sec Diff"
      Height          =   255
      Left            =   4920
      TabIndex        =   153
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label70 
      BackColor       =   &H0080FF80&
      Caption         =   " - - - - - - - - - -"
      Height          =   255
      Left            =   6120
      TabIndex        =   152
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label69 
      BackColor       =   &H0080FF80&
      Caption         =   " - - - - - - - - - -"
      Height          =   255
      Left            =   4920
      TabIndex        =   151
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label68 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Computed as"
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
      Left            =   10920
      TabIndex        =   148
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label67 
      BackColor       =   &H00E0E0E0&
      Caption         =   "x% of the"
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
      Left            =   10920
      TabIndex        =   147
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label66 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sidereal rate"
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
      Left            =   10920
      TabIndex        =   146
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label Label65 
      BackColor       =   &H00E0E0E0&
      Caption         =   "DEC port Rate"
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
      Left            =   12120
      TabIndex        =   145
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label64 
      BackColor       =   &H00E0E0E0&
      Caption         =   "RA port Rate"
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
      Left            =   12120
      TabIndex        =   144
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label63 
      BackColor       =   &H0080FF80&
      Caption         =   "  - - - - -"
      Height          =   255
      Left            =   12240
      TabIndex        =   143
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label Label62 
      BackColor       =   &H00E0E0E0&
      Caption         =   " AUTOGUIDER"
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
      Left            =   12120
      TabIndex        =   138
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label61 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Result"
      Height          =   255
      Left            =   13320
      TabIndex        =   137
      Top             =   6120
      Width           =   855
   End
   Begin VB.Line Line16 
      X1              =   13200
      X2              =   13200
      Y1              =   480
      Y2              =   7800
   End
   Begin VB.Label Label60 
      BackColor       =   &H0080FF80&
      Caption         =   "  - - - - -"
      Height          =   255
      Left            =   12240
      TabIndex        =   135
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Label59 
      BackColor       =   &H00E0E0E0&
      Caption         =   " TRACKING"
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
      Left            =   9720
      TabIndex        =   129
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label58 
      BackColor       =   &H00E0E0E0&
      Caption         =   " TERMINATE"
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
      Left            =   13320
      TabIndex        =   128
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label57 
      BackColor       =   &H00E0E0E0&
      Caption         =   "GUIDING/PEC"
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
      Left            =   10920
      TabIndex        =   127
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label56 
      BackColor       =   &H00E0E0E0&
      Caption         =   " SLEW FUNCTIONS"
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
      Left            =   7680
      TabIndex        =   126
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label55 
      BackColor       =   &H00E0E0E0&
      Caption         =   " MOTOR PARAMETERS"
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
      Left            =   5160
      TabIndex        =   125
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label54 
      BackColor       =   &H00E0E0E0&
      Caption         =   " GOTO/PARK FUNCTIONS"
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
      Left            =   2640
      TabIndex        =   124
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label53 
      BackColor       =   &H00E0E0E0&
      Caption         =   "INITIALIZE"
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
      Left            =   840
      TabIndex        =   123
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label51 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Initial RA val"
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
      Left            =   1320
      TabIndex        =   122
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label52 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Initial DEC val"
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
      Left            =   1320
      TabIndex        =   121
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label50 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1 - South"
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
      Left            =   8520
      TabIndex        =   118
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label49 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1 - South"
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
      TabIndex        =   117
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label48 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1- South"
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
      Left            =   3720
      TabIndex        =   116
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label47 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0 - North"
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
      Left            =   8520
      TabIndex        =   115
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label46 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0 - North"
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
      TabIndex        =   114
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label45 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0 - North"
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
      Left            =   3720
      TabIndex        =   113
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label44 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1 - South"
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
      Left            =   2520
      TabIndex        =   112
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label43 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0 - North"
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
      Left            =   2520
      TabIndex        =   111
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label42 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0.144 arcsecs"
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
      Left            =   3720
      TabIndex        =   110
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label41 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0.144 arcsecs"
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
      Left            =   2520
      TabIndex        =   109
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label40 
      BackColor       =   &H0080FF80&
      Caption         =   "0"
      Height          =   255
      Left            =   9840
      TabIndex        =   102
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label39 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Custom Rate"
      Height          =   255
      Left            =   9720
      TabIndex        =   101
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label38 
      BackColor       =   &H0080FF80&
      Caption         =   " 50%"
      Height          =   255
      Left            =   11040
      TabIndex        =   97
      Top             =   4200
      Width           =   615
   End
   Begin VB.Line Line15 
      X1              =   14400
      X2              =   14400
      Y1              =   480
      Y2              =   7800
   End
   Begin VB.Line Line13 
      X1              =   0
      X2              =   14400
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Label Label37 
      BackColor       =   &H00E0E0E0&
      Caption         =   "New DEC Val"
      Height          =   255
      Left            =   6120
      TabIndex        =   95
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label36 
      BackColor       =   &H00E0E0E0&
      Caption         =   "New RA Val"
      Height          =   255
      Left            =   4920
      TabIndex        =   94
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label35 
      BackColor       =   &H00E0E0E0&
      Caption         =   "End Result"
      Height          =   255
      Left            =   10920
      TabIndex        =   89
      Top             =   6960
      Width           =   855
   End
   Begin VB.Label Label34 
      BackColor       =   &H0080FF80&
      Caption         =   "300"
      Height          =   255
      Left            =   8640
      TabIndex        =   87
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label33 
      BackColor       =   &H0080FF80&
      Caption         =   "400"
      Height          =   255
      Left            =   7560
      TabIndex        =   86
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label32 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Slew Rate"
      Height          =   255
      Left            =   8520
      TabIndex        =   85
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label31 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Slew Rate"
      Height          =   255
      Left            =   7320
      TabIndex        =   84
      Top             =   2400
      Width           =   975
   End
   Begin VB.Line Line14 
      X1              =   0
      X2              =   9600
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label30 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Stop Result"
      Height          =   255
      Left            =   3720
      TabIndex        =   83
      Top             =   6960
      Width           =   855
   End
   Begin VB.Label Label29 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Stop Result"
      Height          =   255
      Left            =   2520
      TabIndex        =   82
      Top             =   6960
      Width           =   855
   End
   Begin VB.Label Label28 
      BackColor       =   &H00E0E0E0&
      Caption         =   "EQMOD  EQCONTRL.DLL DRIVER TESTER  V1.00c"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   77
      Top             =   0
      Width           =   5895
   End
   Begin VB.Line Line12 
      X1              =   0
      X2              =   14400
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label27 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Result"
      Height          =   255
      Left            =   12120
      TabIndex        =   76
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label26 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Start Result"
      Height          =   255
      Left            =   10920
      TabIndex        =   75
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label25 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Result"
      Height          =   255
      Left            =   9720
      TabIndex        =   74
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label24 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Stat Result"
      Height          =   255
      Left            =   6120
      TabIndex        =   73
      Top             =   6960
      Width           =   855
   End
   Begin VB.Label Label23 
      BackColor       =   &H00E0E0E0&
      Caption         =   "End Result"
      Height          =   255
      Left            =   8520
      TabIndex        =   72
      Top             =   6960
      Width           =   855
   End
   Begin VB.Label Label22 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Start Result"
      Height          =   255
      Left            =   8520
      TabIndex        =   71
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label21 
      BackColor       =   &H00E0E0E0&
      Caption         =   "End Result"
      Height          =   255
      Left            =   7320
      TabIndex        =   70
      Top             =   6960
      Width           =   855
   End
   Begin VB.Label Label20 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Start Result"
      Height          =   255
      Left            =   7320
      TabIndex        =   69
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label19 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Result"
      Height          =   255
      Left            =   3720
      TabIndex        =   68
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label18 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Result"
      Height          =   255
      Left            =   6120
      TabIndex        =   67
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label17 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Stat Result"
      Height          =   255
      Left            =   4920
      TabIndex        =   66
      Top             =   6960
      Width           =   855
   End
   Begin VB.Label Label16 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Result"
      Height          =   255
      Left            =   4920
      TabIndex        =   65
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label15 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Result"
      Height          =   255
      Left            =   2520
      TabIndex        =   64
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label14 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Result"
      Height          =   255
      Left            =   1320
      TabIndex        =   63
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label13 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Result"
      Height          =   255
      Left            =   120
      TabIndex        =   62
      Top             =   6120
      Width           =   855
   End
   Begin VB.Line Line11 
      X1              =   0
      X2              =   14400
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Label Label12 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Guide Rate"
      Height          =   255
      Left            =   10920
      TabIndex        =   61
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Hemisphere"
      Height          =   255
      Left            =   9720
      TabIndex        =   59
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label10 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Hemisphere"
      Height          =   255
      Left            =   8520
      TabIndex        =   57
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Hemisphere"
      Height          =   255
      Left            =   7320
      TabIndex        =   55
      Top             =   4680
      Width           =   975
   End
   Begin VB.Line Line10 
      X1              =   12000
      X2              =   12000
      Y1              =   480
      Y2              =   7800
   End
   Begin VB.Line Line9 
      X1              =   10800
      X2              =   10800
      Y1              =   480
      Y2              =   7800
   End
   Begin VB.Line Line8 
      X1              =   9600
      X2              =   9600
      Y1              =   480
      Y2              =   7800
   End
   Begin VB.Line Line7 
      X1              =   8400
      X2              =   8400
      Y1              =   480
      Y2              =   7800
   End
   Begin VB.Line Line6 
      X1              =   7200
      X2              =   7200
      Y1              =   480
      Y2              =   7800
   End
   Begin VB.Line Line5 
      X1              =   6000
      X2              =   6000
      Y1              =   480
      Y2              =   7800
   End
   Begin VB.Line Line4 
      X1              =   4800
      X2              =   4800
      Y1              =   480
      Y2              =   7800
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Hemisphere"
      Height          =   255
      Left            =   3720
      TabIndex        =   23
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Microsteps"
      Height          =   255
      Left            =   3720
      TabIndex        =   22
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Line Line3 
      X1              =   3600
      X2              =   3600
      Y1              =   480
      Y2              =   7800
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Hemisphere"
      Height          =   255
      Left            =   2520
      TabIndex        =   17
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Microsteps"
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Line Line2 
      X1              =   2400
      X2              =   2400
      Y1              =   480
      Y2              =   7800
   End
   Begin VB.Line Line1 
      X1              =   1200
      X2              =   1200
      Y1              =   480
      Y2              =   7800
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Retry"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Timeout"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Baud"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Port"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   1095
   End
End
Attribute VB_Name = "EQMOD"
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
' 24-Oct-03 rcs     Initial edit for EQ Mount Driver Tester
'---------------------------------------------------------------------
'
'
'  SYNOPSIS:
'
'  This is a demonstration of a EQ6/ATLAS/EQG direct stepper motor control access
'  using the EQCONTRL.DLL driver file.
'
'  File EQCONTROL.bas contains all the function prototypes of all subroutines
'  encoded in the EQCONTRL.dll
'
'  The EQCONTRL.DLL simplifies execution of the Mount controller board stepper
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

'  CREDITS:
'
'  Portions of the information on this code should be attributed
'  to Mr. John Archbold from his initial observations and analysis
'  of the interface circuits and of the ASCII data stream between
'  the Hand Controller (HC) and the Go To Controller.
'

Dim lasttrackrate As Long
Dim g_raread As Integer
Dim g_decread As Integer
Dim g_timerblock As Integer
Dim g_radif As Long
Dim g_decdif As Long




Private Sub cmdConnect_Click()

    ' Initilize Mount

    txt001.Text = Trim(Str(EQ_Init(Text3.Text, val(Text4.Text), val(Text5.Text), val(Text6.Text))))

    Call AddLogComment("EQ_Init(" & Text3.Text & "," & Text4.Text & "," & Text5.Text & "," & Text6.Text & ")", txt001.Text)
    

End Sub

Private Sub cmdClose_Click()

    ' Disconnect from mount

   Text22.Text = Trim(Str(EQ_End))
   
   Call AddLogComment("EQ_End()", Text22.Text)

End Sub


Private Sub Command10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' Slew RA +

    Text16.Text = Trim(Str(EQ_Slew(0, val(Text23.Text), 0, val(HScroll1.Value))))
    
    Call AddLogComment("EQ_Slew(0," & Text23.Text & ",0," & HScroll1.Value & ")", Text16.Text)
    
End Sub


Private Sub Command10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' Stop Slewing RA+

    Text17.Text = Trim(Str(EQ_MotorStop(0)))

    Call AddLogComment("EQ_MotorStop(0)", Text17.Text)


End Sub




Private Sub Command11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Slew RA -
    
    Text16.Text = Trim(Str(EQ_Slew(0, val(Text23.Text), 1, val(HScroll1.Value))))
    
    Call AddLogComment("EQ_Slew(0," & Text23.Text & ",1," & HScroll1.Value & ")", Text16.Text)
    
End Sub

Private Sub Command11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' Stop Slewing RA -

   Text17.Text = Trim(Str(EQ_MotorStop(0)))
   
   Call AddLogComment("EQ_MotorStop(0)", Text17.Text)

End Sub



Private Sub Command12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   ' Slew DEC +

    Text18.Text = Trim(Str(EQ_Slew(1, val(Text24.Text), 0, val(HScroll2.Value))))
   
   Call AddLogComment("EQ_Slew(1," & Text24.Text & ",0," & HScroll2.Value & ")", Text18.Text)

End Sub

Private Sub Command12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' Stop Slewing DEC +

   Text19.Text = Trim(Str(EQ_MotorStop(1)))
   
   Call AddLogComment("EQ_MotorStop(1)", Text19.Text)
    
End Sub


Private Sub Command13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   ' Slew DEC -

    Text18.Text = Trim(Str(EQ_Slew(1, val(Text24.Text), 1, val(HScroll2.Value))))
   
    Call AddLogComment("EQ_Slew(1," & Text24.Text & ",1," & HScroll2.Value & ")", Text18.Text)
   
End Sub

Private Sub Command13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' Stop Slew DEC -

    Text19.Text = Trim(Str(EQ_MotorStop(1)))

    Call AddLogComment("EQ_MotorStop(1)", Text19.Text)


End Sub

Private Sub Command14_Click()

    ' Start RA Lunar Tracking

    Text20.Text = Trim(Str(EQ_StartRATrack(0, val(Text25.Text), 0)))
    lasttrackrate = 0
    
    Call AddLogComment("EQ_StartRATrack(0," & Text25.Text & ",0)", Text20.Text)

End Sub


Private Sub Command15_Click()

    ' Start RA Solar Tracking
    
    Text20.Text = Trim(Str(EQ_StartRATrack(2, val(Text25.Text), 0)))
    lasttrackrate = 2
    
    Call AddLogComment("EQ_StartRATrack(2," & Text25.Text & ",0)", Text20.Text)
    
End Sub

Private Sub Command16_Click()

    ' Start RA Lunar Tracking

    Text20.Text = Trim(Str(EQ_StartRATrack(1, val(Text25.Text), 0)))
    lasttrackrate = 1
    
    Call AddLogComment("EQ_StartRATrack(1," & Text25.Text & ",0)", Text20.Text)
    
End Sub



Private Sub Command17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' Send RA Guide Command +

    Text21.Text = Trim(Str(EQ_SendGuideRate(0, lasttrackrate, val(HScroll3.Value), 0, 0, 0)))

    Call AddLogComment("EQ_SendGuideRate(0," & Str(lasttrackrate) & "," & HScroll3.Value & ",0,0,0)", Text21.Text)

End Sub

Private Sub Command17_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' Restore RA normal Track rate

    Text28.Text = Trim(Str(EQ_SendGuideRate(0, lasttrackrate, 0, 0, 0, 0)))

    Call AddLogComment("EQ_SendGuideRate(0," & Str(lasttrackrate) & ",0,0,0,0)", Text28.Text)

End Sub



Private Sub Command18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Send RA Guide Command -
    
    Text21.Text = Trim(Str(EQ_SendGuideRate(0, lasttrackrate, val(HScroll3.Value), 1, 0, 0)))

    Call AddLogComment("EQ_SendGuideRate(0," & Str(lasttrackrate) & "," & HScroll3.Value & ",1,0,0)", Text21.Text)

End Sub

Private Sub Command18_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' Restore RA normal Track rate

    Text28.Text = Trim(Str(EQ_SendGuideRate(0, lasttrackrate, 0, 0, 0, 0)))

    Call AddLogComment("EQ_SendGuideRate(0," & Str(lasttrackrate) & ",0,0,0,0)", Text28.Text)

End Sub


Private Sub Command19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' Send DEC Guide Command +

    Text21.Text = Trim(Str(EQ_SendGuideRate(1, lasttrackrate, val(HScroll3.Value), 0, 0, 0)))

    Call AddLogComment("EQ_SendGuideRate(1," & Str(lasttrackrate) & "," & HScroll3.Value & ",0,0,0)", Text21.Text)
    

End Sub

Private Sub Command19_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' Stop DEC Motor

    Text28.Text = Trim(Str(EQ_MotorStop(1)))

    Call AddLogComment("EQ_MotorStop(1)", Text28.Text)


End Sub

Private Sub Command2_Click()

    ' Initialize and activate motors

    Text7.Text = Trim(Str(EQ_InitMotors(val(Text31.Text), val(Text32.Text))))
        
    Call AddLogComment("EQ_InitMotors(" & Text31.Text & "," & Text32.Text & ")", Text7.Text)
 
End Sub



Private Sub Command20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Send DEC Guide command -
    
    Text21.Text = Trim(Str(EQ_SendGuideRate(1, lasttrackrate, val(HScroll3.Value), 1, 0, 0)))

    Call AddLogComment("EQ_SendGuideRate(1," & Str(lasttrackrate) & "," & HScroll3.Value & ",1,0,0)", Text21.Text)

End Sub

Private Sub Command20_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' Stop DEC Motor

    Text28.Text = Trim(Str(EQ_MotorStop(1)))

    Call AddLogComment("EQ_MotorStop(1)", Text28.Text)


End Sub

Private Sub Command21_Click()

    'Stop RA Motor

    Text26.Text = Trim(Str(EQ_MotorStop(0)))
    
    Call AddLogComment("EQ_MotorStop(0)", Text26.Text)

End Sub

Private Sub Command22_Click()

    'Stop DEC Motor

    Text27.Text = Trim(Str(EQ_MotorStop(1)))
    
    Call AddLogComment("EQ_MotorStop(1)", Text27.Text)

End Sub

Private Sub Command23_Click()

    ' SET RA Motor encoder/counter values

    Text1.Text = Trim(Str(EQ_SetMotorValues(0, val(Text29.Text))))
    
    Call AddLogComment("EQ_SetMotorValues(0," & Text29.Text & ")", Text1.Text)
    

End Sub

Private Sub Command24_Click()

    ' SET DEC Motor encoder/counter values

    Text2.Text = Trim(Str(EQ_SetMotorValues(1, val(Text30.Text))))
    
    Call AddLogComment("EQ_SetMotorValues(1," & Text30.Text & ")", Text2.Text)
    
End Sub

Private Sub Command25_Click()

    'Stop BOTH RA and DEC Motors

    Dim i As Double
    
    i = EQ_MotorStop(0)
    
    Call AddLogComment("EQ_MotorStop(0)", Str(i))
    
    i = EQ_MotorStop(1)
    
    Call AddLogComment("EQ_MotorStop(1)", Str(i))
    
End Sub

Private Sub Command26_Click()

Dim trackoffset As Long
Dim trackdir As Long

   
    trackoffset = val(HScroll4.Value)
    
    If trackoffset > 0 Then
        trackdir = 0            ' Set to RA+ Rate
    Else
        trackdir = 1            ' Set to RA- Rate
    End If
    
    ' Set trackoffset parameter always positive as dicatated by the function
    
    trackoffset = Abs(trackoffset)


    ' Send new RA track rate

    Text20.Text = Trim(Str(EQ_SendCustomTrackRate(0, 0, trackoffset, trackdir, val(Text25.Text), 0)))

    Call AddLogComment("EQ_SendCustomTrackRate(0,0," & Str(trackoffset) & "," & Str(trackdir) & "," & Text25.Text & ",0)", val(Text20.Text))
    

End Sub

Private Sub Command27_Click()

Dim trackoffset As Long
Dim trackdir As Long

   
    trackoffset = val(HScroll4.Value)
    
    If trackoffset > 0 Then
        trackdir = 0                ' Set rate to DEC+
    Else
        trackdir = 1                ' Set rate to DEC-
    End If
    
    ' Set trackoffset parameter always positive as dictated by the function
    
    trackoffset = Abs(trackoffset)

    ' Send new DEC track rate

    Text20.Text = Trim(Str(EQ_SendCustomTrackRate(1, 0, trackoffset, trackdir, val(Text25.Text), 0)))

    Call AddLogComment("EQ_SendCustomTrackRate(1,0," & Str(trackoffset) & "," & Str(trackdir) & "," & Text25.Text & ",0)", val(Text20.Text))

End Sub

Private Sub Command28_Click()
    Loglst.Clear
End Sub

Private Sub Command29_Click()

    'Get DEC 360 Microstep count

    txt001.Text = Trim(Str(EQ_GetTotal360microstep(1)))
  
    Call AddLogComment("EQ_GetTotal360microstep(1)", val(txt001.Text))
    
End Sub

Private Sub Command3_Click()

    Dim j As Long
    
    j = val(Text8.Text) * 0.9   'Compute for motor de-accelaration !Warning! dont set to 100% of step count
    
    If ((val(Text8.Text) - j) < 50000) Then j = val(Text8.Text) * 0.5  'Further reduce if de-acceleration space is too small

    'Move RA motor with the given number of steps RA+

    Text12.Text = Trim(Str(EQ_StartMoveMotor(0, val(Text9.Text), 0, val(Text8.Text), j)))
  
  
    Call AddLogComment("EQ_StartMoveMotor(0," & Text9.Text & ",0," & Text8.Text & "," & Str(j) & ")", val(Text12.Text))
  
  
End Sub

Private Sub Command30_Click()

    'Get RA 360 Microstep count

    txt001.Text = Trim(Str(EQ_GetTotal360microstep(0)))
  
    Call AddLogComment("EQ_GetTotal360microstep(0)", val(txt001.Text))

End Sub

Private Sub Command31_Click()

    txt001.Text = Trim(Str(EQ_GetMountVersion()))
    Call AddLogComment("EQ_GetMountVersion()", val(txt001.Text))

End Sub

Private Sub Command32_Click()

    Text7.Text = Trim(Str(EQ_GetMountStatus()))
    Call AddLogComment3("EQ_GetMountStatus()", val(Text7.Text))

End Sub

Private Sub Command33_Click()

    Text33.Text = Trim(Str(EQ_SetAutoguiderPortRate(0, 3)))
    If (val(Text33.Text) = 0) Then Label60.Caption = "1.00x"
    
    Call AddLogComment("EQ_SetAutoguiderPortRate(0,3)", val(Text33.Text))

End Sub

Private Sub Command34_Click()

    Text33.Text = Trim(Str(EQ_SetAutoguiderPortRate(0, 2)))
    If (val(Text33.Text) = 0) Then Label60.Caption = "0.75x"
    
    Call AddLogComment("EQ_SetAutoguiderPortRate(0,2)", val(Text33.Text))

End Sub

Private Sub Command35_Click()

    Text33.Text = Trim(Str(EQ_SetAutoguiderPortRate(0, 1)))
    If (val(Text33.Text) = 0) Then Label60.Caption = "0.50x"
    
    Call AddLogComment("EQ_SetAutoguiderPortRate(0,1)", val(Text33.Text))

End Sub

Private Sub Command36_Click()

    Text33.Text = Trim(Str(EQ_SetAutoguiderPortRate(0, 0)))
    If (val(Text33.Text) = 0) Then Label60.Caption = "0.25x"
    
    Call AddLogComment("EQ_SetAutoguiderPortRate(0,0)", val(Text33.Text))

End Sub

Private Sub Command37_Click()

    Text33.Text = Trim(Str(EQ_SetAutoguiderPortRate(1, 3)))
    If (val(Text33.Text) = 0) Then Label63.Caption = "1.00x"
    
    Call AddLogComment("EQ_SetAutoguiderPortRate(1,3)", val(Text33.Text))
    
End Sub

Private Sub Command38_Click()

    Text33.Text = Trim(Str(EQ_SetAutoguiderPortRate(1, 1)))
    If (val(Text33.Text) = 0) Then Label63.Caption = "0.50x"
    
    Call AddLogComment("EQ_SetAutoguiderPortRate(1,1)", val(Text33.Text))
    
End Sub

Private Sub Command39_Click()

    Text33.Text = Trim(Str(EQ_SetAutoguiderPortRate(1, 0)))
    If (val(Text33.Text) = 0) Then Label63.Caption = "0.25x"
 
    Call AddLogComment("EQ_SetAutoguiderPortRate(1,0)", val(Text33.Text))

End Sub


Private Sub Command4_Click()

    Dim j As Long
    
    j = val(Text8.Text) * 0.9  'Compute for motor de-accelaration !Warning! dont set to 100% of step count

    If ((val(Text8.Text) - j) < 50000) Then j = val(Text8.Text) * 0.5 'Further reduce if de-acceleration space is too small

    ' Move RA motor with the given number of steps RA-

    Text12.Text = Trim(Str(EQ_StartMoveMotor(0, val(Text9.Text), 1, val(Text8.Text), j)))

    Call AddLogComment("EQ_StartMoveMotor(0," & Text9.Text & ",1," & Text8.Text & "," & Str(j) & ")", val(Text12.Text))

End Sub

Private Sub Command40_Click()

    Text33.Text = Trim(Str(EQ_SetAutoguiderPortRate(1, 2)))
    If (val(Text33.Text) = 0) Then Label63.Caption = "0.75x"
    
    Call AddLogComment("EQ_SetAutoguiderPortRate(1,2)", val(Text33.Text))

End Sub

Private Sub Command41_Click()

    If g_raread = 0 Then
        g_raread = 1
        Command41.Caption = "Autoread Disable"
    Else
        g_raread = 0
        Command41.Caption = "Autoread RA"
        Label69.Caption = " - - - - - - - - - -"
        Label71.Caption = "  1 sec Diff"
    End If
    
End Sub

Private Sub Command42_Click()

    If g_decread = 0 Then
        g_decread = 1
        Command42.Caption = "Autoread Disable"
    Else
        g_decread = 0
        Command42.Caption = "Autoread DEC"
        Label70.Caption = " - - - - - - - - - -"
        Label72.Caption = "  1 sec Diff"
    End If

End Sub

Private Sub Command5_Click()

    Dim j As Long
    
    j = val(Text10.Text) * 0.9 'Compute for motor de-accelaration !Warning! dont set to 100% of step count

    If ((val(Text10.Text) - j) < 50000) Then j = val(Text10.Text) * 0.5 'Further reduce if de-acceleration space is too small

    'Move DEC motor with the given number of steps DEC+

    Text13.Text = Trim(Str(EQ_StartMoveMotor(1, val(Text11.Text), 0, val(Text10.Text), j)))
  
    Call AddLogComment("EQ_StartMoveMotor(1," & Text11.Text & ",0," & Text10.Text & "," & Str(j) & ")", val(Text13.Text))
  
  
End Sub

Private Sub Command6_Click()

    Dim j As Long
    
    j = val(Text10.Text) * 0.9 'Compute for motor de-accelaration !Warning! dont set to 100% of step count

    If ((val(Text10.Text) - j) < 50000) Then j = val(Text10.Text) * 0.5 'Further reduce if de-acceleration space is too small

    'Move DEC motor with the given number of steps DEC-

    Text13.Text = Trim(Str(EQ_StartMoveMotor(1, val(Text11.Text), 1, val(Text10.Text), j)))
  
    Call AddLogComment("EQ_StartMoveMotor(1," & Text11.Text & ",1," & Text10.Text & "," & Str(j) & ")", val(Text13.Text))

End Sub
Private Sub Command1_Click()

    ' Read RA Step value
    
    Text1.Text = Trim(Str(EQ_GetMotorValues(0)))
    
    Call AddLogComment2("EQ_GetMotorValues(0)", val(Text1.Text))

End Sub

Private Sub Command7_Click()

    ' Read DEC Step value
    
    Text2.Text = Trim(Str(EQ_GetMotorValues(1)))
    
    Call AddLogComment2("EQ_GetMotorValues(1)", val(Text2.Text))

End Sub

Private Sub Command8_Click()

    'Get RA Status
    
    Text14.Text = Trim(EQ_GetMotorStatus(0))
    
    Call AddLogComment("EQ_GetMotorStatus(0)", val(Text14.Text))

End Sub

Private Sub Command9_Click()

    'Get DEC Status
    
    Text15.Text = Trim(EQ_GetMotorStatus(1))
    
    Call AddLogComment("EQ_GetMotorStatus(1)", val(Text15.Text))

End Sub

Private Sub Form_Load()
    lasttrackrate = 0
    g_raread = 0
    g_decread = 0
    g_timerblock = 0
    g_radif = 0
    g_decdif = 0
    Timer_readmotor.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Timer_readmotor.Enabled = False
    
Dim i As Long

    ' Disconnect from mount

   i = EQ_End


End Sub

Private Sub HScroll1_Change()

    Label33.Caption = HScroll1.Value

End Sub

Private Sub HScroll2_Change()

    Label34.Caption = HScroll2.Value

End Sub

Private Sub HScroll3_Change()

    Label38.Caption = Str(val(HScroll3.Value) * 10) + "%"

End Sub

Private Sub HScroll4_Change()
    Label40.Caption = HScroll4.Value
End Sub

Private Sub AddLogComment(datFunction As String, dtacode As String)

Dim code_desc As String
Dim code_val As Double


    code_val = val(dtacode)

    Select Case code_val
    
        Case 1, &H1000000
        
            code_desc = "Specified COM Port not available"
            
        Case 2
        
            code_desc = "Specified COM Port already opened"
            
        Case 3, &H1000005
        
            code_desc = "COM Port Communication timeout error"
        
        Case 4
        
            code_desc = "Motor still busy"
        
        Case 5
        
            code_desc = "Mount Initialized on non-standard parameters"
        
        Case 6
        
            code_desc = "RA Stepper Motor still running"
        
        Case 7
        
            code_desc = "DEC Stepper Motor still running"
        
        Case 8
        
            code_desc = "Cannot Initialize RA Motor"
        
        Case 9
        
            code_desc = "Cannot Initialize DEC Motor"
        
        Case 10
        
            code_desc = "Cannot execute command at the current state, issue motor stop command first"
        
        Case 11
        
            code_desc = "Motor Coils not initialized"
            
        Case 128
        
            code_desc = "Motor not rotating, Gear Teeth at front contact"
        
        Case 144
        
            code_desc = "Motor  rotating, Gear Teeth at front contact"
        
        Case 160
        
            code_desc = "Motor not rotating, Gear Teeth at rear contact"
        
        Case 176
        
            code_desc = "Motor rotating, Gear Teeth at rear contact"
        
        Case 200
        
            code_desc = "Motor Coils not initialized"
        
              
        Case 999, &H3000000
        
            code_desc = "Invalid function parameters"
            
        Case &H10000FF
        
            code_desc = "Illegal/Unknown mount reply"
        
   
        Case 0
        
            code_desc = "Success"
        
        Case Else
        
            code_desc = "Returned value: " & Str(code_val) & "   (0x" & printhex(code_val) & ")"
            
         
    End Select
    
    
    Loglst.AddItem "[" & Time & "]: " & datFunction & " =" & Str(code_val) & " : " & code_desc
End Sub
Private Sub AddLogComment2(datFunction As String, dtacode As String)

Dim code_desc As String
Dim code_val As Double


    code_val = val(dtacode)

    Select Case code_val
    
        Case &H1000000
        
            code_desc = "Specified COM Port not available"
            
              
        Case &H1000005
        
            code_desc = "COM Port Communication timeout error"
        
              
        Case &H3000000
        
            code_desc = "Invalid function parameters"
            
        Case &H10000FF
        
            code_desc = "Illegal/Unknown mount reply"
        
  
        Case Else
        
            
            code_desc = "Returned value: " & Str(code_val) & "   (0x" & printhex(code_val) & ")"
            
         
    End Select
    
    
    Loglst.AddItem "[" & Time & "]: " & datFunction & " =" & Str(code_val) & " : " & code_desc
End Sub
Private Sub AddLogComment3(datFunction As String, dtacode As String)

Dim code_desc As String
Dim code_val As Double


    code_val = val(dtacode)

    Select Case code_val
    
        Case 0
        
            code_desc = "Mount not yet connected"
 
  
        Case Else
        
            
            code_desc = "Mount Connected"
            
         
    End Select
    
    
    Loglst.AddItem "[" & Time & "]: " & datFunction & " =" & Str(code_val) & " : " & code_desc
End Sub
Private Function printhex(inpval As Double)

    printhex = Hex$((inpval And &HF00000) / 1048576 And &HF) + Hex$((inpval And &HF0000) / 65536 And &HF) + Hex$((inpval And &HF000) / 4096 And &HF) + Hex$((inpval And &HF00) / 256 And &HF) + Hex$((inpval And &HF0) / 16 And &HF) + Hex$(inpval And &HF)

End Function

Private Sub Timer_readmotor_Timer()

' Timer routine to regularly poll Motor Encoder values for autodisplay

Dim j As Double

   If (g_timerblock = 0) Then
        g_timerblock = 1        ' Block Other successive calls
        
        
        If (g_raread = 1) Then
        
            j = EQ_GetMotorValues(0)        ' Read RA
            
            
            If (j < &H1000000) Then
                Label69.Caption = "  H" & printhex(j)
                Label71.Caption = "  H" & printhex(Abs(j - g_radif))
                g_radif = j
            Else
                Label69.Caption = " - - - - - - - - - -"
                Label71.Caption = "  1 sec Diff"
            End If
        End If
    
        If (g_decread = 1) Then
        
            j = EQ_GetMotorValues(1)        ' Read DEC
            
            If (j < &H1000000) Then
            
                Label70.Caption = "  H" & printhex(j)
                Label72.Caption = "  H" & printhex(Abs(j - g_decdif))
                
                g_decdif = j
                
            Else
                Label70.Caption = " - - - - - - - - - -"
                Label72.Caption = "  1 sec Diff"
            End If
    
        End If

        g_timerblock = 0
   End If
   
End Sub
