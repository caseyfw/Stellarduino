VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form EQSIM 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "EQCONTRL.DLL SIMULATOR"
   ClientHeight    =   8715
   ClientLeft      =   5175
   ClientTop       =   3870
   ClientWidth     =   4500
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H00808080&
      Caption         =   "Check1"
      Height          =   255
      Left            =   2400
      TabIndex        =   71
      Top             =   1800
      Width           =   255
   End
   Begin MSChart20Lib.MSChart dec_pie 
      Height          =   2490
      Left            =   120
      OleObjectBlob   =   "radec.frx":0000
      TabIndex        =   1
      Top             =   5400
      Width           =   1995
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2040
      Top             =   5640
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
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
      ForeColor       =   &H80000007&
      Height          =   1380
      Left            =   2160
      TabIndex        =   11
      Top             =   360
      Width           =   2295
      Begin VB.TextBox txtLatDeg 
         Height          =   315
         Left            =   1320
         TabIndex        =   18
         Text            =   "33"
         Top             =   240
         Width           =   480
      End
      Begin VB.TextBox txtLatMin 
         Height          =   315
         Left            =   1800
         TabIndex        =   17
         Text            =   "58"
         Top             =   240
         Width           =   450
      End
      Begin VB.TextBox txtLongDeg 
         Height          =   315
         Left            =   1320
         TabIndex        =   16
         Text            =   "151"
         Top             =   600
         Width           =   480
      End
      Begin VB.TextBox txtLongMin 
         Height          =   315
         Left            =   1800
         TabIndex        =   15
         Text            =   "7"
         Top             =   600
         Width           =   450
      End
      Begin VB.ComboBox cbEW 
         Height          =   315
         ItemData        =   "radec.frx":1B7C
         Left            =   720
         List            =   "radec.frx":1B86
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   600
         Width           =   555
      End
      Begin VB.ComboBox cbNS 
         Height          =   315
         ItemData        =   "radec.frx":1B90
         Left            =   720
         List            =   "radec.frx":1B9A
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   555
      End
      Begin VB.TextBox txtemulElevation 
         Height          =   315
         Left            =   720
         TabIndex        =   12
         Text            =   "100"
         Top             =   960
         Width           =   885
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808080&
         Caption         =   "LONG:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   765
      End
      Begin VB.Label Label6 
         BackColor       =   &H00808080&
         Caption         =   "LAT:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   690
      End
      Begin VB.Label Label5 
         BackColor       =   &H00808080&
         Caption         =   "Elev (m):"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   990
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Track Sidreal"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "DEC-"
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
      Left            =   480
      TabIndex        =   8
      Top             =   1200
      Width           =   735
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1095
      Left            =   1800
      Max             =   32000
      Min             =   1
      TabIndex        =   6
      Top             =   480
      Value           =   32000
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "RA-"
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
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RA+"
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
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.Timer Timer 
      Interval        =   100
      Left            =   2040
      Top             =   6000
   End
   Begin VB.CommandButton Command3 
      Caption         =   "DEC+"
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
      Left            =   480
      TabIndex        =   7
      Top             =   480
      Width           =   735
   End
   Begin MSChart20Lib.MSChart ra_pie 
      Height          =   2370
      Left            =   120
      OleObjectBlob   =   "radec.frx":1BA4
      TabIndex        =   0
      Top             =   2280
      Width           =   1995
   End
   Begin VB.Label Label49 
      BackColor       =   &H00808080&
      Caption         =   "FAST GOTO"
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
      Left            =   1680
      TabIndex        =   72
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label48 
      BackColor       =   &H00808080&
      Caption         =   "Pier"
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
      Left            =   2400
      TabIndex        =   70
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label47 
      BackColor       =   &H00000040&
      Caption         =   "Label8"
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
      TabIndex        =   69
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label45 
      BackColor       =   &H00808080&
      Caption         =   " - "
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
      TabIndex        =   67
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label42 
      BackColor       =   &H00000040&
      Caption         =   "32000"
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
      Left            =   1320
      TabIndex        =   64
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
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
      Left            =   960
      TabIndex        =   63
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label41 
      BackColor       =   &H00C0C0C0&
      Caption         =   "EQMOD EQCONTRL.DLL SIMULATOR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   62
      Top             =   0
      Width           =   4215
   End
   Begin VB.Label Label38 
      BackColor       =   &H00000000&
      Caption         =   "Label8"
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
      Left            =   2400
      TabIndex        =   59
      Top             =   7920
      Width           =   2055
   End
   Begin VB.Label Label37 
      BackColor       =   &H00000000&
      Caption         =   "Label8"
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
      Left            =   2400
      TabIndex        =   58
      Top             =   8400
      Width           =   2055
   End
   Begin VB.Label Label25 
      BackColor       =   &H00000000&
      Caption         =   "Label8"
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
      Left            =   2400
      TabIndex        =   46
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label34 
      BackColor       =   &H00000000&
      Caption         =   "Label8"
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
      Left            =   2400
      TabIndex        =   55
      Top             =   6960
      Width           =   2055
   End
   Begin VB.Label Label33 
      BackColor       =   &H00000000&
      Caption         =   "Label8"
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
      Left            =   2400
      TabIndex        =   54
      Top             =   7440
      Width           =   2055
   End
   Begin VB.Label Label32 
      BackColor       =   &H00000000&
      Caption         =   "Label8"
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
      Left            =   2400
      TabIndex        =   53
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Label Label29 
      BackColor       =   &H00000000&
      Caption         =   "Label8"
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
      Left            =   2400
      TabIndex        =   50
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Label Label27 
      BackColor       =   &H00000000&
      Caption         =   "Label8"
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
      Left            =   2400
      TabIndex        =   48
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label Label26 
      BackColor       =   &H00808080&
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
      Left            =   2400
      TabIndex        =   47
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label20 
      BackColor       =   &H00808080&
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
      Height          =   255
      Left            =   2400
      TabIndex        =   41
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label19 
      BackColor       =   &H00808080&
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
      Height          =   255
      Left            =   2400
      TabIndex        =   40
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label18 
      BackColor       =   &H00808080&
      Caption         =   "ADJ DEC DEG"
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
      TabIndex        =   39
      Top             =   8400
      Width           =   735
   End
   Begin VB.Label Label17 
      BackColor       =   &H00808080&
      Caption         =   "DEC DEG"
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
      TabIndex        =   38
      Top             =   8160
      Width           =   615
   End
   Begin VB.Label Label16 
      BackColor       =   &H00808080&
      Caption         =   "DEC ENCODER"
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
      TabIndex        =   37
      Top             =   7920
      Width           =   735
   End
   Begin VB.Label Label12 
      BackColor       =   &H00808080&
      Caption         =   "HOUR ANGLE"
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
      TabIndex        =   36
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label11 
      BackColor       =   &H00808080&
      Caption         =   "RA ENCODER"
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
      TabIndex        =   35
      Top             =   4680
      Width           =   735
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
      Left            =   960
      TabIndex        =   34
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00808080&
      Caption         =   "LST"
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
      Left            =   2400
      TabIndex        =   33
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label lstlbl 
      BackColor       =   &H00000040&
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2880
      TabIndex        =   32
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label deccmplbl 
      BackColor       =   &H00000040&
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2880
      TabIndex        =   31
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label racmplbl 
      BackColor       =   &H00000040&
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2880
      TabIndex        =   30
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label15 
      BackColor       =   &H00808080&
      Caption         =   "AZ"
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
      Left            =   2400
      TabIndex        =   29
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label14 
      BackColor       =   &H00808080&
      Caption         =   "ALT"
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
      Left            =   2400
      TabIndex        =   28
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label13 
      BackColor       =   &H00808080&
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
      Height          =   255
      Left            =   2400
      TabIndex        =   27
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label azlbl 
      BackColor       =   &H00000040&
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2880
      TabIndex        =   26
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label altlbl 
      BackColor       =   &H00000040&
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2880
      TabIndex        =   25
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label declbl 
      BackColor       =   &H00000040&
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2880
      TabIndex        =   24
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackColor       =   &H00808080&
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
      Height          =   255
      Left            =   2400
      TabIndex        =   23
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label ralbl 
      BackColor       =   &H00000040&
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2880
      TabIndex        =   22
      Top             =   2520
      Width           =   1575
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
      Left            =   960
      TabIndex        =   9
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   " 800000"
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
      Left            =   960
      TabIndex        =   3
      Top             =   7920
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   " 800000"
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
      Left            =   960
      TabIndex        =   2
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label28 
      BackColor       =   &H00808080&
      Caption         =   " RA Encoder  from RA"
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
      Left            =   2400
      TabIndex        =   49
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label30 
      BackColor       =   &H00808080&
      Caption         =   " RA Encoder from ALT/AZ"
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
      Left            =   2400
      TabIndex        =   51
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label Label36 
      BackColor       =   &H00808080&
      Caption         =   "DEC Encoder from DEC-EPier"
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
      Left            =   2400
      TabIndex        =   57
      Top             =   7200
      Width           =   1815
   End
   Begin VB.Label Label35 
      BackColor       =   &H00808080&
      Caption         =   "DEC Encoder from DEC-WPier"
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
      Left            =   2400
      TabIndex        =   56
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Label Label40 
      BackColor       =   &H00808080&
      Caption         =   "DEC Encoder from ALTAZ-EPier"
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
      Left            =   2400
      TabIndex        =   61
      Top             =   8160
      Width           =   1815
   End
   Begin VB.Label Label24 
      BackColor       =   &H00808080&
      Caption         =   "PC Clock"
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
      Left            =   2880
      TabIndex        =   45
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label31 
      BackColor       =   &H00808080&
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
      Left            =   2400
      TabIndex        =   52
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label Label39 
      BackColor       =   &H00808080&
      Caption         =   "DEC Encoder from ALTAZ-WPier"
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
      Left            =   2400
      TabIndex        =   60
      Top             =   7680
      Width           =   1815
   End
   Begin VB.Label Label22 
      BackColor       =   &H00808080&
      Caption         =   "From RA/DEC values"
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
      Left            =   2880
      TabIndex        =   43
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label21 
      BackColor       =   &H00808080&
      Caption         =   "From Encoder values"
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
      Left            =   2880
      TabIndex        =   42
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label23 
      BackColor       =   &H00808080&
      Caption         =   "From ALT/AZ values"
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
      Left            =   2880
      TabIndex        =   44
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label43 
      BackColor       =   &H00808080&
      Caption         =   " - "
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
      Left            =   120
      TabIndex        =   65
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label44 
      BackColor       =   &H00808080&
      Caption         =   " - "
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
      Left            =   120
      TabIndex        =   66
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label46 
      BackColor       =   &H00808080&
      Caption         =   " - "
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
      TabIndex        =   68
      Top             =   5160
      Width           =   855
   End
End
Attribute VB_Name = "EQSIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Check1_Click()
    If Check1.Value = 1 Then
        emulRA_gotorate = GMS * 10000        'Outrageously fast slew
        emulDEC_gotorate = GMS * 10000
    Else
        emulRA_gotorate = GMS * 800        'Full Speed x800 slew
        emulDEC_gotorate = GMS * 800
    End If
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
emulRA_shift = val(VScroll1.Value)
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
emulRA_shift = 0
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    emulRA_shift = -val(VScroll1.Value)
End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    emulRA_shift = 0
End Sub



Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    emulDEC_Shift = val(VScroll1.Value)
End Sub

Private Sub Command3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    emulDEC_Shift = 0
End Sub



Private Sub Command4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    emulDEC_Shift = -val(VScroll1.Value)
End Sub

Private Sub Command4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    emulDEC_Shift = 0
End Sub

Private Sub Command5_Click()

    If Command5.Caption = "Track Sidreal" Then
        Command5.Caption = "Stop Tracking"
    Else
        Command5.Caption = "Track Sidreal"
    End If
    
    If Command5.Caption = "Stop Tracking" Then
        emulRA_track = GMS
    Else
        emulRA_track = 0
    End If


End Sub

Private Sub Command6_Click()
    emulRA_Encoder = &H800010
    emulDEC_Encoder = &HA26C80
End Sub

Private Sub Form_Load()

EnableCloseButton Me.hWnd, False
emulHemisphere = 0
emulpieCounter = 0

emulRAEncoder_Zero_pos = &H800000
emulDECEncoder_Zero_pos = &H800000

cbNS.ListIndex = 1
cbEW.ListIndex = 0

emulRA_Encoder = &H800000
emulDEC_Encoder = &HA26C80

emulTot_RA = 9024000
emulTot_DEC = 9024000

emulRA_gotorate = GMS * 800        'Full Speed slew
emulDEC_gotorate = GMS * 800       'Full Speed slew

emulRA_track = 0
emulRA_shift = 0
emulDEC_track = 0
emulDEC_Shift = 0
emulRA_target = 0
emulDEC_target = 0

End Sub






Private Sub Timer_Timer()

Dim i As Double


If emulRA_target = 0 Then

    ' Slew based on Slew and track rate state

    Label43.Caption = Format(Str(emulRA_track), "000.000000000")
    Label45.Caption = Format(Str(emulRA_shift), "0000.0000")
    

    emulRA_Encoder = emulRA_Encoder + emulRA_track
    emulRA_Encoder = emulRA_Encoder + emulRA_shift
Else

    ' Slew based on target encoder value

    If emulRA_target > emulRA_Encoder Then
    
        emulRA_Encoder = emulRA_Encoder + emulRA_gotorate       'Slew forward
        
        If emulRA_Encoder >= emulRA_target Then
            emulRA_Encoder = emulRA_target
            emulRA_target = 0
        End If
    Else
        emulRA_Encoder = emulRA_Encoder - emulRA_gotorate       'Slew reverse
        
        If emulRA_Encoder <= emulRA_target Then
            emulRA_Encoder = emulRA_target
            emulRA_target = 0
        End If
        
    End If
End If

If emulDEC_target = 0 Then

    ' Slew based on Slew buttons
     
    Label44.Caption = Format(Str(emulDEC_track), "000.000000000")
    Label46.Caption = Format(Str(emulDEC_Shift), "0000.0000")
    
    emulDEC_Encoder = emulDEC_Encoder + emulDEC_Shift
    emulDEC_Encoder = emulDEC_Encoder + emulDEC_track
    
Else
    If emulDEC_target > emulDEC_Encoder Then
    
    'Slew based on Target encoder position
    
        emulDEC_Encoder = emulDEC_Encoder + emulDEC_gotorate    'forward movement
        
        If emulDEC_Encoder >= emulDEC_target Then
            emulDEC_Encoder = emulDEC_target
            emulDEC_target = 0
        End If
    Else
        emulDEC_Encoder = emulDEC_Encoder - emulDEC_gotorate    'reverse movement
        
        If emulDEC_Encoder <= emulDEC_target Then
            emulDEC_Encoder = emulDEC_target
            emulDEC_target = 0
        End If
    End If
End If

' Convert  emulRA_Encoder to RA Hours

emulRA_Hours = Get_EncoderHours(emulRAEncoder_Zero_pos, emulRA_Encoder, emulTot_RA, emulHemisphere)


' Convert emulDEC_Encoder to DEC Degrees


emulDEC_Degrees = Get_EncoderDegrees(emulDECEncoder_Zero_pos, emulDEC_Encoder, emulTot_DEC, emulHemisphere)


' Re adjust DEC for a -90/+90 range

emulDec_DegNoAdjust = emulDEC_Degrees

emulDEC_Degrees = Range_DEC(emulDEC_Degrees)

emulpieCounter = emulpieCounter + 1

If emulpieCounter > PIEDISP Then
    
    emulpieCounter = 0
    
    i = (Range24(emulRA_Hours - 6) / 24) * 100
    ra_pie.Column = 1
    If emulHemisphere = 0 Then
        ra_pie.Data = i
    Else
        ra_pie.Data = 100 - i
    End If
    ra_pie.Column = 2
    If emulHemisphere = 0 Then
        ra_pie.Data = 100 - i
    Else
        ra_pie.Data = i
    End If

    If emulHemisphere = 0 Then
        i = emulDec_DegNoAdjust - 90
    Else
        i = emulDec_DegNoAdjust - 270
    End If
    If i < 0 Then i = 360 + i
    If i > 360 Then i = i - 360
    i = (i / 360) * 100
    dec_pie.Column = 1
    If emulHemisphere = 0 Then
        dec_pie.Data = 100 - i
    Else
        dec_pie.Data = i
    End If
    dec_pie.Column = 2
    If emulHemisphere = 0 Then
        dec_pie.Data = i
    Else
        dec_pie.Data = 100 - i
    End If

End If

Label1.Caption = printhex(emulRA_Encoder)
Label3.Caption = Format$(Str(emulRA_Hours), " 000.000000000")


Label2.Caption = printhex(emulDEC_Encoder)
Label4.Caption = Format$(Str(emulDec_DegNoAdjust), " 000.000000000")
Label10.Caption = Format$(Str(emulDEC_Degrees), " 000.000000000")



End Sub


Private Sub Timer1_Timer()

    Dim tha As Double
    Dim tRA As Double
    Dim tDEC As Double
    Dim tAz As Double
    Dim tAlt As Double
    Dim i As Double
    

    emulLongitude = CDbl(txtLongDeg) + (CDbl(txtLongMin) / 60#)
    If cbEW.Text = "W" Then emulLongitude = -emulLongitude  ' W is neg
    
    emulLatitude = CDbl(txtLatDeg) + (CDbl(txtLatMin) / 60#)
    If cbNS.Text = "S" Then emulLatitude = -emulLatitude
    
    emulElevation = CDbl(txtemulElevation)
    
    
     ' convert dHA which is in Radians to true RA and scale it

    tRA = now_lst(emulLongitude * DEG_RAD) + emulRA_Hours
    
    If emulHemisphere = 0 Then
        If (emulDec_DegNoAdjust > 90) And (emulDec_DegNoAdjust <= 270) Then tRA = tRA - 12
    Else
        If (emulDec_DegNoAdjust <= 90) Or (emulDec_DegNoAdjust > 270) Then tRA = tRA + 12

    End If
    
    tRA = Range24(tRA)

    emulRA = tRA
    emulDEC = emulDEC_Degrees
    
    Select Case SOP_RAHours(emulRA_Hours)
                Case pierUnknown:    Label47.Caption = "Unknown"
                Case PierEast:       Label47.Caption = "East, Scope to West"
                Case PierWest:       Label47.Caption = "West, Scope to East"
    End Select
 
    hadec_aa (emulLatitude * DEG_RAD), ((emulRA - now_lst(emulLongitude * DEG_RAD)) * HRS_RAD), (emulDEC_Degrees * DEG_RAD), emulAlt, emulAz

    emulAlt = emulAlt * RAD_DEG ' tAlt was in Radians
    emulAz = 360# - (emulAz * RAD_DEG) ' tAz was in Radians
    
    
    ' Check for re-computation
    
    
    
    aa_hadec (emulLatitude * DEG_RAD), (emulAlt * DEG_RAD), ((360# - emulAz) * DEG_RAD), tha, tDEC

    ' convert dHA which is in Radians to true RA and scale it
    
    tRA = (tha * RAD_HRS) + now_lst(emulLongitude * DEG_RAD)
    tRA = Range24(tRA)

    tDEC = tDEC * RAD_DEG ' tDec was in Radians

    
    lstlbl.Caption = FmtSexa(now_lst(emulLongitude * DEG_RAD), False)
    
    ralbl.Caption = FmtSexa(emulRA, False)
    declbl.Caption = FmtSexa(emulDEC, True)
    
    azlbl.Caption = FmtSexa(emulAz, False)
    altlbl.Caption = FmtSexa(emulAlt, False)
    
    racmplbl.Caption = FmtSexa(tRA, False)
    deccmplbl.Caption = FmtSexa(tDEC, True)
    
    Label25.Caption = printhex(emulRA_Encoder)
    Label32.Caption = printhex(emulDEC_Encoder)
    
    ' Routine Below will attempt to compute back for the ENCODER Values
    ' This will be very useful for GOTO implementations as it converts RA/DEC and ALT/AZ
    ' back to Motor Encoder values based on Local location and local sidereal time
  
    Label27.Caption = printhex(Get_RAEncoderfromRA(emulRA, emulDec_DegNoAdjust, emulLongitude, emulRAEncoder_Zero_pos, emulTot_RA, emulHemisphere))


    hadec_aa (emulLatitude * DEG_RAD), ((emulRA - now_lst(emulLongitude * DEG_RAD)) * HRS_RAD), (emulDec_DegNoAdjust * DEG_RAD), tAlt, tAz

    tAlt = tAlt * RAD_DEG ' tAlt was in Radians
    tAz = 360# - (tAz * RAD_DEG) ' tAz was in Radians

    Label29.Caption = printhex(Get_RAEncoderfromAltAz(tAlt, tAz, emulLongitude, emulLatitude, emulRAEncoder_Zero_pos, emulTot_RA, emulHemisphere))
    
    Label33.Caption = printhex(Get_DECEncoderfromDEC(emulDEC, 0, emulDECEncoder_Zero_pos, emulTot_DEC, emulHemisphere)) ' West PIER
    
    Label34.Caption = printhex(Get_DECEncoderfromDEC(emulDEC, 1, emulDECEncoder_Zero_pos, emulTot_DEC, emulHemisphere)) ' East PIER

    Label37.Caption = printhex(Get_DECEncoderfromAltAz(emulAlt, emulAz, emulLongitude, emulLatitude, emulDECEncoder_Zero_pos, emulTot_DEC, 0, emulHemisphere)) ' West PIER

    Label38.Caption = printhex(Get_DECEncoderfromAltAz(emulAlt, emulAz, emulLongitude, emulLatitude, emulDECEncoder_Zero_pos, emulTot_DEC, 1, emulHemisphere)) ' East PIER

End Sub

Private Sub VScroll1_Change()
    Label42.Caption = VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
    Label42.Caption = VScroll1.Value
End Sub
