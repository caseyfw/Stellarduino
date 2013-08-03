VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "RADEC"
   ClientHeight    =   10380
   ClientLeft      =   5190
   ClientTop       =   4005
   ClientWidth     =   10545
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10380
   ScaleWidth      =   10545
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      Caption         =   "Init Motor Coils"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      TabIndex        =   73
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Clear Logs"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      TabIndex        =   71
      Top             =   3600
      Width           =   3375
   End
   Begin VB.TextBox List1 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   70
      Top             =   9240
      Width           =   10215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "STOP MOTORS"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      TabIndex        =   68
      Top             =   1440
      Width           =   3375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      TabIndex        =   67
      Top             =   720
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   9240
      TabIndex        =   65
      Text            =   "COM1"
      Top             =   240
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   120
      Picture         =   "radec.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   2835
      TabIndex        =   63
      Top             =   600
      Width           =   2895
   End
   Begin MSChart20Lib.MSChart dec_pie 
      Height          =   2490
      Left            =   120
      OleObjectBlob   =   "radec.frx":0ECA
      TabIndex        =   1
      Top             =   5520
      Width           =   2955
   End
   Begin MSChart20Lib.MSChart ra_pie 
      Height          =   2490
      Left            =   120
      OleObjectBlob   =   "radec.frx":2A46
      TabIndex        =   0
      Top             =   2160
      Width           =   2955
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3120
      Top             =   480
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Site Information"
      ForeColor       =   &H80000007&
      Height          =   1620
      Left            =   3360
      TabIndex        =   12
      Top             =   2520
      Width           =   3255
      Begin VB.TextBox txtLatDeg 
         Height          =   315
         Left            =   1920
         TabIndex        =   19
         Text            =   "14"
         Top             =   315
         Width           =   480
      End
      Begin VB.TextBox txtLatMin 
         Height          =   315
         Left            =   2505
         TabIndex        =   18
         Text            =   "35"
         Top             =   315
         Width           =   570
      End
      Begin VB.TextBox txtLongDeg 
         Height          =   315
         Left            =   1920
         TabIndex        =   17
         Text            =   "120"
         Top             =   720
         Width           =   480
      End
      Begin VB.TextBox txtLongMin 
         Height          =   315
         Left            =   2505
         TabIndex        =   16
         Text            =   "57"
         Top             =   720
         Width           =   570
      End
      Begin VB.ComboBox cbEW 
         Height          =   315
         ItemData        =   "radec.frx":45BF
         Left            =   1275
         List            =   "radec.frx":45C9
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   720
         Width           =   555
      End
      Begin VB.ComboBox cbNS 
         Height          =   315
         ItemData        =   "radec.frx":45D3
         Left            =   1275
         List            =   "radec.frx":45DD
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   315
         Width           =   555
      End
      Begin VB.TextBox txtElevation 
         Height          =   315
         Left            =   1275
         TabIndex        =   13
         Text            =   "100"
         Top             =   1125
         Width           =   885
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Longitude:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   22
         Top             =   765
         Width           =   765
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Latitude:"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   165
         TabIndex        =   21
         Top             =   360
         Width           =   690
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Elevation (m):"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   1155
         Width           =   990
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Start Tracking"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6960
      TabIndex        =   11
      Top             =   2160
      Width           =   3375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "DEC-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   9
      Top             =   1680
      Width           =   1215
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1335
      Left            =   6360
      Max             =   800
      Min             =   1
      TabIndex        =   6
      Top             =   1080
      Value           =   800
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "RA-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RA+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3480
      Top             =   480
   End
   Begin VB.CommandButton Command3 
      Caption         =   "DEC+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   8
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label44 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SLEW Rate"
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
      Left            =   6360
      TabIndex        =   72
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label43 
      BackColor       =   &H00C0FFC0&
      Caption         =   " 800"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   69
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label42 
      BackColor       =   &H00FFC0C0&
      Caption         =   "EQMOD COM Interface:"
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
      Left            =   6960
      TabIndex        =   66
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label41 
      BackColor       =   &H00FFC0C0&
      Caption         =   "EQMOD RADEC/ALTAZ ROUTINES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   64
      Top             =   0
      Width           =   3855
   End
   Begin VB.Label Label40 
      BackColor       =   &H00FFC0C0&
      Caption         =   "DEC Encoder Value Computed from  ALT/AZ-EPier"
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
      Left            =   7200
      TabIndex        =   62
      Top             =   8520
      Width           =   3255
   End
   Begin VB.Label Label39 
      BackColor       =   &H00FFC0C0&
      Caption         =   "DEC Encoder Value Computed from  ALT/AZ-WPier"
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
      Left            =   7200
      TabIndex        =   61
      Top             =   7920
      Width           =   3255
   End
   Begin VB.Label Label38 
      BackColor       =   &H00000000&
      Caption         =   " 000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   7200
      TabIndex        =   60
      Top             =   8760
      Width           =   3135
   End
   Begin VB.Label Label37 
      BackColor       =   &H00000000&
      Caption         =   " 000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   7200
      TabIndex        =   59
      Top             =   8160
      Width           =   3135
   End
   Begin VB.Label Label25 
      BackColor       =   &H00000000&
      Caption         =   " 000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   7200
      TabIndex        =   47
      Top             =   4440
      Width           =   3135
   End
   Begin VB.Label Label34 
      BackColor       =   &H00000000&
      Caption         =   " 000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   7200
      TabIndex        =   56
      Top             =   7560
      Width           =   3135
   End
   Begin VB.Label Label33 
      BackColor       =   &H00000000&
      Caption         =   " 000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   7200
      TabIndex        =   55
      Top             =   6960
      Width           =   3135
   End
   Begin VB.Label Label32 
      BackColor       =   &H00000000&
      Caption         =   " 000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   7200
      TabIndex        =   54
      Top             =   6360
      Width           =   3135
   End
   Begin VB.Label Label31 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Current DEC Encoder Value"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   53
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Label Label29 
      BackColor       =   &H00000000&
      Caption         =   " 000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   7200
      TabIndex        =   51
      Top             =   5640
      Width           =   3135
   End
   Begin VB.Label Label27 
      BackColor       =   &H00000000&
      Caption         =   " 000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   7200
      TabIndex        =   49
      Top             =   5040
      Width           =   3135
   End
   Begin VB.Label Label26 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Current RA Encoder Value"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   48
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label24 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Computed from the PC Clock"
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
      Left            =   4320
      TabIndex        =   46
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Computed from the ALT/AZ values"
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
      Left            =   4320
      TabIndex        =   45
      Top             =   7920
      Width           =   2655
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Computed from the RA/DEC values"
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
      Left            =   4320
      TabIndex        =   44
      Top             =   6480
      Width           =   2655
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Computed from the Encoder values"
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
      Left            =   4320
      TabIndex        =   43
      Top             =   5040
      Width           =   2655
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFC0C0&
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
      Height          =   495
      Left            =   3480
      TabIndex        =   42
      Top             =   8640
      Width           =   615
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFC0C0&
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
      Height          =   495
      Left            =   3480
      TabIndex        =   41
      Top             =   8160
      Width           =   615
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ADJ DEC DEG"
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
      TabIndex        =   40
      Top             =   8760
      Width           =   1455
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFC0C0&
      Caption         =   "DEC DEG"
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
      TabIndex        =   39
      Top             =   8400
      Width           =   1455
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFC0C0&
      Caption         =   "DEC ENCODER"
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
      TabIndex        =   38
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFC0C0&
      Caption         =   "HOUR ANGLE"
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
      TabIndex        =   37
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFC0C0&
      Caption         =   "RA ENCODER"
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
      TabIndex        =   36
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   1680
      TabIndex        =   35
      Top             =   8760
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFC0C0&
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
      Height          =   495
      Left            =   3480
      TabIndex        =   34
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label lstlbl 
      BackColor       =   &H00C0FFC0&
      Caption         =   " 000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   33
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label deccmplbl 
      BackColor       =   &H00C0FFC0&
      Caption         =   " 000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   32
      Top             =   8640
      Width           =   2415
   End
   Begin VB.Label racmplbl 
      BackColor       =   &H00C0FFC0&
      Caption         =   " 000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   31
      Top             =   8160
      Width           =   2415
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFC0C0&
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
      Height          =   495
      Left            =   3480
      TabIndex        =   30
      Top             =   7200
      Width           =   615
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFC0C0&
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
      Height          =   495
      Left            =   3480
      TabIndex        =   29
      Top             =   6720
      Width           =   615
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFC0C0&
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
      Height          =   495
      Left            =   3480
      TabIndex        =   28
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label azlbl 
      BackColor       =   &H00C0FFC0&
      Caption         =   " 000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   27
      Top             =   7200
      Width           =   2415
   End
   Begin VB.Label altlbl 
      BackColor       =   &H00C0FFC0&
      Caption         =   " 000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   26
      Top             =   6720
      Width           =   2415
   End
   Begin VB.Label declbl 
      BackColor       =   &H00C0FFC0&
      Caption         =   " 000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   25
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
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
      Height          =   495
      Left            =   3480
      TabIndex        =   24
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label ralbl 
      BackColor       =   &H00C0FFC0&
      Caption         =   " 000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   23
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label28 
      BackColor       =   &H00FFC0C0&
      Caption         =   " RA Encoder Value Computed from RA vlaue"
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
      Left            =   7200
      TabIndex        =   50
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Label Label30 
      BackColor       =   &H00FFC0C0&
      Caption         =   " RA Encoder Value Computed from ALT/AZ"
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
      Left            =   7200
      TabIndex        =   52
      Top             =   5400
      Width           =   2775
   End
   Begin VB.Label Label36 
      BackColor       =   &H00FFC0C0&
      Caption         =   "DEC Encoder Value Computed from DEC-EPier"
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
      Left            =   7200
      TabIndex        =   58
      Top             =   7320
      Width           =   3015
   End
   Begin VB.Label Label35 
      BackColor       =   &H00FFC0C0&
      Caption         =   "DEC Encoder Value Computed from  DEC-WPier"
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
      Left            =   7200
      TabIndex        =   57
      Top             =   6720
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
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
' 21-Oct-06 rcs     Initial edit for EQ Mount Driver RADEC/ALTAZ Program
'---------------------------------------------------------------------
'
'
'  SYNOPSIS:
'
'  This is a demonstration of a EQ6/ATLAS/EQG direct stepper motor control access
'  using the EQCONTRL.DLL driver code. The code read RA/DEC motor encoder values
'  and converts them to RA/DEC/ALT/AZ Coordinates.
'  There are two Asynchronous timers that were used here;
'
'  The first one is used to read encoder values from the mount at a regular interval
'  It also computes for Hour angle and DEC coordinate in degrees. The timer is also used
'  to update the PIE charts to represent motor position
'
'  The second timer is used to compute and display the RA/DEC/ALT/AZ coordinates
'  It is also used to verify if the computation is correct by re-computing for the RA and DEC
'  Encoder values.
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

Dim eqres As Double
Dim gtimer0 As Double


Dim RAEncoder_Zero_pos As Double        'RA Encoder 0 Hour Position
Dim DECEncoder_Zero_pos As Double       'DEC Encoder 0 Degree Position

Dim RA_Park_pos As Double               'RA Encoder Park Position (0 hour)
Dim DEC_Park_pos As Double              'DEC Encoder Park Position (90 Degrees)


Dim Tot_RA As Double                    'Total RA Microstep count
Dim Tot_DEC As Double                   'Total DEC Microstep count

Dim tRA As Double



Dim RA_Encoder As Double                'RA Encoder global value
Dim DEC_Encoder As Double               'DEC Encoder global value



Dim RA_Hours As Double                  'RA Hours Angle (HA)
Dim DEC_Degrees As Double               'DEC Coordinate in degrees
Dim Latitude As Double                  'Site Latitude
Dim Longitude As Double                 'Site Longitude
Dim Elevation As Double                 'Site Elevation (not used)

Dim RA As Double                        'Global RA Coordinate
Dim DEC As Double                       'Global DEC Coordinate
Dim Alt As Double                       'Global ALT Coordinate
Dim Az As Double                        'Global AZ Coordinate
Dim ha As Double                        'Globak Hour Angle value


' Button press to stop the motor

Private Sub Command8_Click()

    eqres = EQ_MotorStop(0)
    eqres = EQ_MotorStop(1)
    AddLog ("Both RA/DEC Motors stopped")
    
End Sub

Private Sub Command9_Click()

        'Intialize Motor Coils
      
        eqres = EQ_MotorStop(0)
        eqres = EQ_MotorStop(1)
      
        eqres = EQ_InitMotors(RA_Park_pos, DEC_Park_pos)
    
        If eqres <> 0 Then
    
            Call AddLog("Cannot Connect Initialize EQ Motors, Check COM port settings. Return value: " & Str(eqres))
            Command7.Caption = "Connect"
            eqres = EQ_End()
    
        Else
            Call AddLog("Motors initialized. Return value: " & Str(eqres))
        End If

End Sub

Private Sub Form_Load()

    gtimer0 = 0        'Intialize Motor Coils


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
    
    '0 Hour / 0 Degree Encoder position

    RAEncoder_Zero_pos = &H800000
    DECEncoder_Zero_pos = &H800000

    cbNS.ListIndex = 0
    cbEW.ListIndex = 0

    ' Initial RA/DEC Park Position Encoder values

    RA_Park_pos = &H800000
    DEC_Park_pos = &HA26C80
    
    ' Initialize Encoder Global Variables
    
    RA_Encoder = RA_Park_pos
    DEC_Encoder = DEC_Park_pos

    ' Total Step Count value

    Tot_RA = 9024000
    Tot_DEC = 9024000


End Sub

' RA+ Button Pressed

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
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
   
    eqres = EQ_Slew(0, 0, 0, val(VScroll1.value))
    If eqres <> 0 Then
       AddLog ("Error[RA+ ]: Mount not connected")
    Else
        AddLog ("Slewing at RA+")
    End If
    
END04:
End Sub

'RA+ Button Released

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

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

    If Command5.Caption = "Stop Tracking" Then
    
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

' RA- Button Pressed

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  eqres = EQ_MotorStop(0)          ' Stop RA Motor
    If eqres <> 0 Then
            AddLog ("Error[RA- ]: Cannot stop RA motor, mount not connected")
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


    eqres = EQ_Slew(0, 0, 1, val(VScroll1.value))
    If eqres <> 0 Then
        AddLog ("Error[RA- ]: Mount not connected")
    Else
        AddLog ("Slewing at RA-")
    End If

END05:

End Sub

'RA- Button Released

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
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
    
    If Command5.Caption = "Stop Tracking" Then
    
        eqres = EQ_StartRATrack(0, 0, 0)             ' Track RA Motor at Sidreal Rate
        If eqres <> 0 Then
            AddLog ("Error[RA- ]: Cannot resume tracking, mount not connected")
        Else
            AddLog ("Restoring RA tracking rate to Sidreal")
        End If
    End If

    
END02:
End Sub

'DEC+ Button Pressed

Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    
    eqres = EQ_Slew(1, 0, 0, val(VScroll1.value))
    If eqres <> 0 Then
        AddLog ("Error[DEC+ ]: Mount not connected")
    Else
        AddLog ("Slewing at DEC+")
    End If
    
End Sub

'DEC+ Button Released

Private Sub Command3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    eqres = EQ_MotorStop(1)
    If eqres <> 0 Then
        AddLog ("Error[DEC+ ]: Cannot Stop DEC Motor, Mount not connected")
    Else
        AddLog ("Stopping DEC Motor")
    End If
    
End Sub

'DEC- Button Pressed

Private Sub Command4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    eqres = EQ_Slew(1, 0, 1, val(VScroll1.value))
    If eqres <> 0 Then
        AddLog ("Error[DEC- ]: Mount not connected")
    Else
        AddLog ("Slewing at DEC-")
    End If
End Sub

'DEC- Button Released

Private Sub Command4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    eqres = EQ_MotorStop(1)
    If eqres <> 0 Then
        AddLog ("Error[DEC- ]: Cannto Stop DEC motor, Mount not connected")
    Else
        AddLog ("Stopping DEC Motor")
    End If
    
End Sub

'Routine to process initiation of Sidreal tracking

Private Sub Command5_Click()

    If Command5.Caption = "Start Tracking" Then
        eqres = EQ_StartRATrack(0, 0, 0)             ' Star RA Motor at Sidreal Rate
        If eqres <> 0 Then
            AddLog ("Cannot Start tracking, mount not connected")
        Else
           AddLog ("Mount now tracking at the sidreal rate")
           Command5.Caption = "Stop Tracking"
        End If
    Else
        eqres = EQ_MotorStop(0)                      ' Stop RA Motor
        Command5.Caption = "Start Tracking"
        AddLog ("Mount RA tracking disabled")
        
    End If

End Sub

'Clear Log Window

Private Sub Command6_Click()
    List1.Text = ""
End Sub

'Routine to Initialize EQMOD Mount

Private Sub Command7_Click()

 If Command7.Caption <> "Stop/Disconnect" Then
   
    Command7.Caption = "Connecting .."
    
    'Initiate Connection to the modded mount
    
    eqres = EQ_Init(Combo1.Text, "9600", 1000, 1)
    
    If eqres <> 0 Then
        Call AddLog("Cannot Connect to the EQModded mount, Check COM port settings. Return value:" & Str(eqres))
        Command7.Caption = "Connect"


    Else
      
        'Get Mount Parameters
      
        Tot_RA = EQ_GetTotal360microstep(0)
        Tot_DEC = EQ_GetTotal360microstep(1)
      
        'Intialize Motor Coils
      
        eqres = EQ_MotorStop(0)
        eqres = EQ_MotorStop(1) + eqres
          
        If eqres <> 0 Then
    
            Call AddLog("Cannot Connect Stop EQ Motors, Check COM port settings. Return value: " & Str(eqres))
            Command7.Caption = "Connect"
            eqres = EQ_End()
    
        Else
        
            Command7.Caption = "Stop/Disconnect"
        
            'Compute for Site Location Data
    
            Longitude = CDbl(txtLongDeg) + (CDbl(txtLongMin) / 60#)
            If cbEW.Text = "W" Then Longitude = -Longitude  ' W is neg
    
            Latitude = CDbl(txtLatDeg) + (CDbl(txtLatMin) / 60#)
            If cbNS.Text = "S" Then Latitude = -Latitude
    
            Elevation = CDbl(txtElevation)
          
          
            Timer.Enabled = True
            Timer1.Enabled = True
            Call AddLog("EQModded mount found at port " & Combo1.Text)
            Combo1.Locked = True
            Combo1.Enabled = False
        End If
        
    End If
    
  Else 'Caption = "Stop/Disconnect"
  
    
  
    Command7.Caption = "Connect"
    Timer.Enabled = False
    Timer1.Enabled = False
    
    eqres = EQ_End()
    Call AddLog("Mount Disconnected from port " & Combo1.Text)
    Combo1.Locked = False
    Combo1.Enabled = True

  End If
    
End Sub

'Timer to read RA/DEC Encoder values and to update PIE CHART DATA

Private Sub Timer_Timer()

Dim i, dectemp As Double

    If gtimer0 = 0 Then

        gtimer0 = 1     ' Block Timer overruns
                
        i = EQ_GetMotorValues(0)
    
        'Make sure global RA Encoder variable is updated upon successful read
    
        If i < &H1000000 Then RA_Encoder = i

        i = EQ_GetMotorValues(1)

       'Make sure global DEC Encoder variable is updated upon successful read

        If i < &H1000000 Then DEC_Encoder = i

        ' Convert  RA_Encoder to RA Hours

        RA_Hours = Get_EncoderHours(RAEncoder_Zero_pos, RA_Encoder, Tot_RA)

        ' Convert DEC_Encoder to DEC Degrees

        DEC_Degrees = Get_EncoderDegrees(DECEncoder_Zero_pos, DEC_Encoder, Tot_DEC)

        ' Re-adjust DEC for a -90/+90 range or else astrobas dll will do funky things

        dectemp = DEC_Degrees

        DEC_Degrees = Range_DEC(DEC_Degrees)    'Do the adjustment here

        'Draw PIE CHART for RA
    
        i = (RA_Hours / 24) * 100
        ra_pie.Column = 1
        ra_pie.Data = i
        ra_pie.Column = 2
        ra_pie.Data = 100 - i
        
        'Update labels as well
        
        Label1.Caption = printhex(RA_Encoder)
        Label3.Caption = Format$(Str(RA_Hours), " 000.000000000")

        'Draw PIE CHAR for DEC
    
        i = dectemp - 90
        If i < 0 Then i = 360 + i
        If i > 360 Then i = i - 360
        i = (i / 360) * 100
     
        dec_pie.Column = 1
        dec_pie.Data = 100 - i
        dec_pie.Column = 2
        dec_pie.Data = i
        
        'Update labels as well
        
        Label2.Caption = printhex(DEC_Encoder)
        Label4.Caption = Format$(Str(dectemp), " 000.000000000")
        Label10.Caption = Format$(Str(DEC_Degrees), " 000.000000000")
        
        ' Allow Timer to re-execute
        
        gtimer0 = 0
    Else
        AddLog ("Warning Timer overflow, please increase the delay")
    End If

End Sub

'Generate string function that will convert  Double value to HHH:MM:SS or DEG:MM:SS

Public Function FmtSexa(ByVal n As Double, ShowPlus As Boolean) As String
    Dim sg As String
    Dim us As String, ms As String, ss As String
    Dim u As Integer, m As Integer
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


'Routine to Convert all RA/DEC ENcoder Data to RA/DEC/ALT/AZ Coordinates

Private Sub Timer1_Timer()

    Dim tHA, tDEC, i As Double


    Timer.Enabled = False       ' Disable routine that reads encoder values from the mount

    ' Compute for RA/DEC Coordinates using RA_Hour/DEC_Degrees Global


    tRA = now_lst(Longitude * DEG_RAD) + RA_Hours
    
    ' Make sure value is within range
    
    If tRA < 0 Then tRA = tRA + 24#
    If tRA >= 24 Then tRA = tRA - 24#

    RA = tRA
    DEC = DEC_Degrees
    
    'Compute for ALT/AZ using astro32.dll function
 
    hadec_aa (Latitude * DEG_RAD), ((RA - now_lst(Longitude * DEG_RAD)) * HRS_RAD), (DEC_Degrees * DEG_RAD), Alt, Az

    Alt = Alt * RAD_DEG ' tAlt was in Radians
    Az = 360# - (Az * RAD_DEG) ' tAz was in Radians
    
    
    ' Routines below will just recompute back the original values
    ' This will confirm if the original computation is correc
    ' Verification is done visually on the program itself
        
    
    'Convert ALT/AZ Back to RA/DEC
            
    aa_hadec (Latitude * DEG_RAD), (Alt * DEG_RAD), ((360# - Az) * DEG_RAD), tHA, tDEC

    ' convert HA which is in Radians to true RA and scale it
    
    tRA = (tHA * RAD_HRS) + now_lst(Longitude * DEG_RAD)
    If tRA < 0 Then tRA = tRA + 24#
    If tRA >= 24 Then tRA = tRA - 24#
    tDEC = tDEC * RAD_DEG ' tDec was in Radians

    'Display computed values
    
    lstlbl.Caption = FmtSexa(now_lst(Longitude * DEG_RAD), False)   'LST
    ralbl.Caption = FmtSexa(RA, False)                              'RA
    declbl.Caption = FmtSexa(DEC, True)                             'DEC
    
    azlbl.Caption = FmtSexa(Az, False)                              'AZ
    altlbl.Caption = FmtSexa(Alt, False)                            'ALT
    
    racmplbl.Caption = FmtSexa(tRA, False)                          'Computed RA
    deccmplbl.Caption = FmtSexa(tDEC, True)                         'Computed DEC
    
    
    ' The following routines will attempt to recompute back the RA and DEC
    ' Encoder values from the RA/DEC/ALT/AZ Coordinates and current Location
    ' and Sidereal time settings
    ' You can actually use these routines for your GOTO / SYNC functions
    
    
    ' Display actual RA/DEC Encoder values
    
    Label25.Caption = printhex(RA_Encoder)
    Label32.Caption = printhex(DEC_Encoder)
    
    ' Compute and Display RA Encoder value from RA/ALT/AZ
  
    Label27.Caption = printhex(Get_RAEncoderfromRA(RA, Longitude, RAEncoder_Zero_pos, Tot_RA))

    Label29.Caption = printhex(Get_RAEncoderfromAltAz(Alt, Az, Longitude, Latitude, RAEncoder_Zero_pos, Tot_RA))
    
    ' Compute and Display DEC Encoder values from DEC/ALT/AZ
    
    
    ' There are two possible DEC encoder values that can be generated
    ' Both values actually point to the same physical position except
    ' that the other one will require you to initiate a complete 360 degree turn
    ' To which turn will be used will depend on the position of the pier (west or east side)
    
    ' Derive DEC Encoder value from DEC Coordinate using EAST and WESTSIDE PIER position
    Label33.Caption = printhex(Get_DECEncoderfromDEC(DEC, 0, DECEncoder_Zero_pos, Tot_DEC)) ' West PIER
    Label34.Caption = printhex(Get_DECEncoderfromDEC(DEC, 1, DECEncoder_Zero_pos, Tot_DEC)) ' East PIER

    ' Derive DEC Encoder value from DEC Coordinate using EAST and WESTSIDE PIER position
    Label37.Caption = printhex(Get_DECEncoderfromAltAz(Alt, Az, Longitude, Latitude, DECEncoder_Zero_pos, Tot_DEC, 0)) ' West PIER
    Label38.Caption = printhex(Get_DECEncoderfromAltAz(Alt, Az, Longitude, Latitude, DECEncoder_Zero_pos, Tot_DEC, 1)) ' East PIER

    Timer.Enabled = True
    

End Sub

' Compute for the HA value using the stepper motor encoder value
' You also need to specify the 0 position encoder value and the Total microstep counts for a complete 360 revolution
' encOffset0 - 0 Hour Encoder position
' encoderval - Current encoder value
' Tot_enc - Total 360 degree microstepping count

Public Function Get_EncoderHours(encOffset0 As Double, encoderval As Double, Tot_enc As Double) As Double

Dim i As Double

    ' Compute in Hours the encoder value based on 0 position value (RAOffset0)
    ' and Total 360 degree rotation microstep count (Tot_Enc

    If encoderval > encOffset0 Then
        i = ((encoderval - encOffset0) / Tot_enc) * 24
        i = 24 - i
        If i < 0 Then i = 24 + i
    Else
        i = ((encOffset0 - encoderval) / Tot_enc) * 24
        If i > 24 Then i = i - 24

    End If
    
    Get_EncoderHours = i

End Function

' Compute for the degree value using the stepper motor encoder value
' You also need to specify the 0 position encoder value and the Total microstep counts for a complete 360 revolution
' encOffset0 - 0 Hour Encoder position
' encoderval - Current encoder value
' Tot_enc - Total 360 degree microstepping count

Public Function Get_EncoderDegrees(encOffset0 As Double, encoderval As Double, Tot_enc As Double) As Double

Dim i As Double

    ' Compute in Hours the encoder value based on 0 position value (EncOffset0)
    ' and Total 360 degree rotation microstep count (Tot_Enc


    If encoderval > encOffset0 Then

       i = ((DEC_Encoder - encOffset0) / Tot_enc) * 360
       If i > 360 Then i = i - 360

    Else

        i = ((encOffset0 - encoderval) / Tot_enc) * 360
        i = 360 - i
        If i < 0 Then i = 360 + i

    End If

        Get_EncoderDegrees = i

End Function

' Function to Convert back an HA value to Encoder position values
' You also need to specify the 0 position encoder value and the Total microstep counts for a complete 360 revolution
' encOffset0 - 0 Hour Encoder position
' hourval - Current HA value
' Tot_enc - Total 360 degree microstepping count

Public Function Get_EncoderfromHours(encOffset0 As Double, hourval As Double, Tot_enc As Double) As Long

    If (hourval < 12) Then
    
        Get_EncoderfromHours = encOffset0 - ((hourval / 24) * Tot_enc)
    
    Else
    
        Get_EncoderfromHours = (((24 - hourval) / 24) * Tot_enc) + encOffset0
    
    End If

End Function

' Function to Convert back a degree value to Encoder position values
' You also need to specify the 0 position encoder value and the Total microstep counts for a complete 360 revolution
' encOffset0 - 0 Hour Encoder position
' degval - Current degree position value
' Tot_enc - Total 360 degree microstepping count

Public Function Get_EncoderfromDegrees(encOffset0 As Double, degval As Double, Tot_enc As Double, Pier As Double) As Long

    If (degval > 180) And (Pier = 0) Then
    
        Get_EncoderfromDegrees = encOffset0 - (((360 - degval) / 360) * Tot_enc)
    
    Else
    
        Get_EncoderfromDegrees = ((degval / 360) * Tot_enc) + encOffset0
    
    End If

End Function


' Function that will ensure that the DEC value will be between -90 to 90
' Even if it is set at the other side of the pier

Public Function Range_DEC(decdegrees As Double) As Double

    If (decdegrees >= 270) And (decdegrees <= 360) Then
        Range_DEC = decdegrees - 360
        Exit Function
    End If
    
    If (decdegrees >= 180) And (decdegrees < 270) Then
        Range_DEC = 180 - decdegrees
        Exit Function
    End If
    
    If (decdegrees >= 90) And (decdegrees < 180) Then
        Range_DEC = 180 - decdegrees
        Exit Function
    End If
    
    Range_DEC = decdegrees

End Function

'Function that will generate the RA Encoder value based on the RA Coordinate
'ra_in_hours - Current RA coordinate value
'pLongitude - Site's Longitude value
'encOffset0 - RA Encoder 0 Hour position value
'Tot_Enc - Total 360 degree microstepping count

Private Function Get_RAEncoderfromRA(ra_in_hours As Double, pLongitude As Double, encOffset0 As Double, Tot_enc As Double) As Long

Dim i As Double
  
    i = ra_in_hours - now_lst(pLongitude * DEG_RAD)
    If i < 0 Then i = 24 + i
    Get_RAEncoderfromRA = Get_EncoderfromHours(encOffset0, i, Tot_enc)
   
End Function

'Function that will generate the RA Encoder value based on the ALT/AZ Coordinate
'Alt_in_deg - Current ALT coordinate value in degrees
'Az_in_deg - Current AZ coordinate value in degrees
'pLongitude - Site's Longitude value
'encOffset0 - RA Encoder 0 Hour position value
'Tot_Enc - Total 360 degree microstepping count

Private Function Get_RAEncoderfromAltAz(Alt_in_deg As Double, Az_in_deg As Double, pLongitude As Double, pLatitude As Double, encOffset0 As Double, Tot_enc As Double) As Long

Dim i As Double
Dim ttha, ttdec As Double

    
    'Use the Astro32.dll function for the conversion
    
    aa_hadec (pLatitude * DEG_RAD), (Alt_in_deg * DEG_RAD), ((360# - Az_in_deg) * DEG_RAD), ttha, ttdec
    i = (ttha * RAD_HRS) + now_lst(pLongitude * DEG_RAD)
    If i < 0 Then i = i + 24#
    If i >= 24 Then i = i - 24#
    i = i - now_lst(pLongitude * DEG_RAD)
    If i < 0 Then i = 24 + i

 
    Get_RAEncoderfromAltAz = Get_EncoderfromHours(encOffset0, i, Tot_enc)
   
End Function

'Function that will generate the DEC Encoder value based on the DEC Coordinate
'ra_in_hours - Current RA coordinate value
'pLongitude - Site's Longitude value
'encOffset0 - RA Encoder 0 Hour position value
'Tot_Enc - Total 360 degree microstepping count
'Pier - Pierside position 0 - West, 1 East


Private Function Get_DECEncoderfromDEC(dec_in_degrees As Double, Pier As Double, encOffset0 As Double, Tot_enc As Double) As Long

Dim i As Double

    i = dec_in_degrees
  
    If Pier = 1 Then i = 180 - i

    Get_DECEncoderfromDEC = Get_EncoderfromDegrees(encOffset0, i, Tot_enc, Pier)
   
End Function

'Function that will generate the DEC Encoder value based on the ALT/AZ Coordinate
'Alt_in_deg - Current ALT coordinate value in degrees
'Az_in_deg - Current AZ coordinate value in degrees
'pLongitude - Site's Longitude value
'encOffset0 - RA Encoder 0 Hour position value
'Tot_Enc - Total 360 degree microstepping count
'Pier - Pierside position 0 - West, 1 East

Private Function Get_DECEncoderfromAltAz(Alt_in_deg As Double, Az_in_deg As Double, pLongitude As Double, pLatitude As Double, encOffset0 As Double, Tot_enc As Double, Pier As Double) As Long

Dim i As Double
Dim ttha, ttdec As Double

  
    aa_hadec (pLatitude * DEG_RAD), (Alt_in_deg * DEG_RAD), ((360# - Az_in_deg) * DEG_RAD), ttha, ttdec
    
    i = ttdec * RAD_DEG ' tDec was in Radians

    If Pier = 1 Then i = 180 - i    ' Use the other side

     Get_DECEncoderfromAltAz = Get_EncoderfromDegrees(encOffset0, i, Tot_enc, Pier)
   
End Function

' Function to print the hexadecimal values

Private Function printhex(inpval As Double) As String

    printhex = " " & Hex$((inpval And &HF00000) / 1048576 And &HF) + Hex$((inpval And &HF0000) / 65536 And &HF) + Hex$((inpval And &HF000) / 4096 And &HF) + Hex$((inpval And &HF00) / 256 And &HF) + Hex$((inpval And &HF0) / 16 And &HF) + Hex$(inpval And &HF)

End Function

' Routine to update the slew rate label

Private Sub VScroll1_Change()
    Label43.Caption = Str(VScroll1.value)
End Sub

' Routine to update the slew rate label

Private Sub VScroll1_Scroll()
    Label43.Caption = Str(VScroll1.value)
End Sub

' Logger

Private Sub AddLog(dtaLog As String)
        List1.Text = Right(List1.Text & "[" & Time & "] " & dtaLog & vbCrLf, 20000)
        List1.SelStart = Len(List1.Text)
End Sub
