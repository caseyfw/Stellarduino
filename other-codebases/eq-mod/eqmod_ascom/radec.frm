VERSION 5.00
Begin VB.Form EQSIM 
   BackColor       =   &H00000040&
   Caption         =   "EQASCOM TELESCOPE SIMULATOR"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   3660
   ClipControls    =   0   'False
   Icon            =   "radec.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   3660
   WindowState     =   1  'Minimized
   Begin VB.CommandButton Command1 
      BackColor       =   &H0095C1CB&
      Caption         =   "StarSim"
      Height          =   255
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   2760
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000040&
      Caption         =   "Developer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   4215
      Left            =   120
      TabIndex        =   21
      Top             =   3120
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Label Label23 
         BackColor       =   &H00000040&
         Caption         =   "From ALT/AZ values"
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
         Height          =   255
         Left            =   2760
         TabIndex        =   47
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label22 
         BackColor       =   &H00000040&
         Caption         =   "From RA/DEC values"
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
         Height          =   255
         Left            =   600
         TabIndex        =   46
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label altlbl 
         BackColor       =   &H00000040&
         Caption         =   "Label8"
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
         Left            =   600
         TabIndex        =   39
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label azlbl 
         BackColor       =   &H00000040&
         Caption         =   "Label8"
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
         Left            =   600
         TabIndex        =   38
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackColor       =   &H00000040&
         Caption         =   "ALT"
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
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label Label15 
         BackColor       =   &H00000040&
         Caption         =   "AZ"
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
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label racmplbl 
         BackColor       =   &H00000040&
         Caption         =   "Label8"
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
         Left            =   2760
         TabIndex        =   35
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label deccmplbl 
         BackColor       =   &H00000040&
         Caption         =   "Label8"
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
         Left            =   2760
         TabIndex        =   34
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label lstlbl 
         BackColor       =   &H00000040&
         Caption         =   "Label8"
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
         Left            =   600
         TabIndex        =   33
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000040&
         Caption         =   "LST"
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
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   3840
         Width           =   375
      End
      Begin VB.Label Label19 
         BackColor       =   &H00000040&
         Caption         =   "RA"
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
         Height          =   255
         Left            =   2280
         TabIndex        =   31
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label Label20 
         BackColor       =   &H00000040&
         Caption         =   "DEC"
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
         Height          =   255
         Left            =   2280
         TabIndex        =   30
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label Label27 
         BackColor       =   &H00000040&
         Caption         =   "Label8"
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
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label29 
         BackColor       =   &H00000040&
         Caption         =   "Label8"
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
         Left            =   120
         TabIndex        =   28
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label33 
         BackColor       =   &H00000040&
         Caption         =   "Label8"
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
         Left            =   3120
         TabIndex        =   27
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label34 
         BackColor       =   &H00000040&
         Caption         =   "Label8"
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
         Left            =   3120
         TabIndex        =   26
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label37 
         BackColor       =   &H00000040&
         Caption         =   "Label8"
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
         Left            =   3120
         TabIndex        =   25
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label38 
         BackColor       =   &H00000040&
         Caption         =   "Label8"
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
         Left            =   3120
         TabIndex        =   24
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label47 
         BackColor       =   &H00000040&
         Caption         =   "Label8"
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
         Left            =   600
         TabIndex        =   23
         Top             =   3600
         Width           =   2895
      End
      Begin VB.Label Label48 
         BackColor       =   &H00000040&
         Caption         =   "Pier"
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
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   3600
         Width           =   375
      End
      Begin VB.Label Label28 
         BackColor       =   &H00000040&
         Caption         =   " RA Encoder  from RA"
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
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label35 
         BackColor       =   &H00000040&
         Caption         =   "DEC Encoder from DEC-WPier"
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
         Height          =   255
         Left            =   3120
         TabIndex        =   43
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label36 
         BackColor       =   &H00000040&
         Caption         =   "DEC Encoder from DEC-EPier"
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
         Height          =   255
         Left            =   3120
         TabIndex        =   42
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label39 
         BackColor       =   &H00000040&
         Caption         =   "DEC Encoder from ALTAZ-WPier"
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
         Height          =   255
         Left            =   3120
         TabIndex        =   45
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label40 
         BackColor       =   &H00000040&
         Caption         =   "DEC Encoder from ALTAZ-EPier"
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
         Height          =   255
         Left            =   3120
         TabIndex        =   44
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label30 
         BackColor       =   &H00000040&
         Caption         =   " RA Encoder from ALT/AZ"
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
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   960
         Width           =   2055
      End
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000040&
      Caption         =   "Fast Goto"
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
      TabIndex        =   0
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1560
      Top             =   2760
   End
   Begin VB.Timer Timer 
      Interval        =   200
      Left            =   2040
      Top             =   2760
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000040&
      Caption         =   "DEC AXIS"
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
      Left            =   1800
      TabIndex        =   7
      Top             =   0
      Width           =   1815
      Begin VB.PictureBox PictureDEC 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   360
         ScaleHeight     =   81
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   81
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label44 
         BackColor       =   &H00000040&
         Caption         =   "88888.8888"
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
         Left            =   720
         TabIndex        =   18
         ToolTipText     =   "Steps/Sec"
         Top             =   1800
         Width           =   1020
      End
      Begin VB.Label Label18 
         BackColor       =   &H00000040&
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   20
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label declbl 
         BackColor       =   &H00000040&
         Caption         =   "Label8"
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
         Left            =   720
         TabIndex        =   16
         Top             =   2280
         Width           =   1020
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000040&
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
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000040&
         Caption         =   "800000"
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
         Left            =   720
         TabIndex        =   12
         ToolTipText     =   "Encoder Count"
         Top             =   1560
         Width           =   1020
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000040&
         Caption         =   "000.0000"
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
         Left            =   720
         TabIndex        =   11
         ToolTipText     =   "Degrees"
         Top             =   2040
         Width           =   1020
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000040&
         Caption         =   "ENC"
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
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label17 
         BackColor       =   &H00000040&
         Caption         =   "DEG"
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
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000040&
      Caption         =   "RA AXIS"
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
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1815
      Begin VB.PictureBox PictureRA 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   360
         ScaleHeight     =   81
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   81
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label43 
         BackColor       =   &H00000040&
         Caption         =   "88888.8888"
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
         Left            =   720
         TabIndex        =   17
         ToolTipText     =   "Steps/Sec"
         Top             =   1800
         Width           =   1020
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000040&
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   19
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label ralbl 
         BackColor       =   &H00000040&
         Caption         =   "Label8"
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
         Left            =   720
         TabIndex        =   14
         Top             =   2280
         Width           =   1020
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000040&
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
         Left            =   120
         TabIndex        =   13
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000040&
         Caption         =   "800000"
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
         Left            =   720
         TabIndex        =   6
         ToolTipText     =   "Encoder Count"
         Top             =   1560
         Width           =   1025
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000040&
         Caption         =   "ENC"
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
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label12 
         BackColor       =   &H00000040&
         Caption         =   "HA"
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
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000040&
         Caption         =   "000.0000"
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
         Left            =   720
         TabIndex        =   3
         ToolTipText     =   "Hour Angle"
         Top             =   2040
         Width           =   975
      End
   End
End
Attribute VB_Name = "EQSIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function Arc Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function AngleArc Lib "GDI32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal dwRadius As Long, ByVal eStartAngle As Single, ByVal eSweepAngle As Single) As Long
Private Declare Function MoveToEx Lib "GDI32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByRef lpPoint As Any) As Long
Dim timer1flag As Boolean
Dim timer2flag As Boolean

Dim EMUCurrent_time As Double
Dim EMULast_time As Double



Private Sub Check1_Click()
    If Check1.Value = 1 Then
        emulRA_gotorate = GMS * 10000        'Outrageously fast slew
        emulDEC_gotorate = GMS * 10000
    Else
        emulRA_gotorate = GMS * 800        'Full Speed x800 slew
        emulDEC_gotorate = GMS * 800
    End If
End Sub





Private Sub Command1_Click()
    If StarSim.Visible = False Then
        StarSim.Show
    End If
End Sub

Private Sub Form_Load()

Call SetText

EnableCloseButton Me.hwnd, False
emulHemisphere = 0
emulpieCounter = 0

emulRAEncoder_Zero_pos = gRAEncoder_Zero_pos
emulDECEncoder_Zero_pos = gDECEncoder_Zero_pos

emulRA_Encoder = gRAEncoder_Zero_pos
emulDEC_Encoder = DECEncoder_Home_pos

emulCurrent_time = 0
emulLast_time = 0
emulEmulRA_Init = emulRA_Encoder

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

If HC.oPersist.ReadIniValue("ShowStarSim") = "1" Then
    Command1.Visible = True
Else
    Command1.Visible = False
End If

EQSIM.Height = 3540


End Sub



Private Sub Timer_Timer()

Dim i As Double
Dim emulinc As Double
Dim elapsed As Double

If Not timer1flag Then
    timer1flag = True
    
    EMUCurrent_time = EQnow_lst_norange()
    If EMULast_time = 0 Then EMUCurrent_time = 0.000002
    
    If EMULast_time > EMUCurrent_time Then      ' Counter wrap around ?
        EMULast_time = EQnow_lst_norange()
        EMUCurrent_time = EMULast_time
    End If
    elapsed = EMUCurrent_time - EMULast_time
    
    If emulRA_target = 0 Then
        ' Slew based on Slew and track rate state
        
        ' Compute Elapsed stepper count based on Elapsed Local Sidreal time (PC time)
        emulinc = (emulRA_shift + emulRA_track) * 10
        Label43.Caption = FormatNumber(emulinc, 4, , , False)
        emulinc = emulinc * elapsed
        emulRA_Encoder = emulRA_Encoder + emulinc
    Else
        ' Slew based on target encoder value
        If emulRA_target > emulRA_Encoder Then
            emulinc = emulRA_gotorate * 10
            Label43.Caption = FormatNumber(emulinc, 4, , , False)
            emulinc = emulinc * elapsed
            emulRA_Encoder = emulRA_Encoder + emulinc       'Slew forward
            If emulRA_Encoder >= emulRA_target Then
                emulRA_Encoder = emulRA_target
                emulRA_target = 0
                emulRA_shift = 0
                emulRA_track = 0
            End If
        Else
            emulinc = -emulRA_gotorate * 10
            Label43.Caption = FormatNumber(emulinc, 4, , , False)
            emulinc = emulinc * elapsed
            emulRA_Encoder = emulRA_Encoder + emulinc
            If emulRA_Encoder <= emulRA_target Then
                emulRA_Encoder = emulRA_target
                emulRA_target = 0
                emulRA_shift = 0
                emulRA_track = 0
            End If
        End If
    End If
    
    If emulDEC_target = 0 Then
        ' Slew based on Slew buttons
        emulinc = (emulDEC_Shift + emulDEC_track) * 10
        Label44.Caption = FormatNumber(emulinc, 4, , , False)
        emulinc = emulinc * elapsed
        emulDEC_Encoder = emulDEC_Encoder + emulinc
    Else
        If emulDEC_target > emulDEC_Encoder Then
            'Slew based on Target encoder position
            emulinc = emulDEC_gotorate * 10
            Label44.Caption = FormatNumber(emulinc, 4, , , False)
            emulinc = emulinc * elapsed
            emulDEC_Encoder = emulDEC_Encoder + emulinc
            If emulDEC_Encoder >= emulDEC_target Then
                emulDEC_Encoder = emulDEC_target
                emulDEC_target = 0
                emulDEC_Shift = 0
                emulDEC_track = 0
            End If
        Else
            emulinc = -emulDEC_gotorate * 10
            Label44.Caption = FormatNumber(emulinc, 4, , , False)
            emulinc = emulinc * elapsed
            emulDEC_Encoder = emulDEC_Encoder + emulinc
            If emulDEC_Encoder <= emulDEC_target Then
                emulDEC_Encoder = emulDEC_target
                emulDEC_target = 0
                emulDEC_Shift = 0
                emulDEC_track = 0
            End If
        End If
    End If
    
    EMULast_time = EMUCurrent_time
    
    ' Convert  emulRA_Encoder to RA Hours
    emulRA_Hours = Get_EncoderHours(emulRAEncoder_Zero_pos, emulRA_Encoder, emulTot_RA, emulHemisphere)
    
    ' Convert emulDEC_Encoder to DEC Degrees
    emulDEC_Degrees = Get_EncoderDegrees(emulDECEncoder_Zero_pos, emulDEC_Encoder, emulTot_DEC, emulHemisphere)
   
    ' Re adjust DEC for a -90/+90 range
    emulDec_DegNoAdjust = emulDEC_Degrees
    Label4.Caption = Format$(str(emulDec_DegNoAdjust), "000.0000")
    emulDEC_Degrees = Range_DEC(emulDEC_Degrees)
    
    emulpieCounter = emulpieCounter + 1
    
    If emulpieCounter > PIEDISP Then
        
        emulpieCounter = 0
        
        i = (Range24(emulRA_Hours - 6) / 24) * 100
        If emulHemisphere <> 0 Then
            i = 100 - i
        End If
        Call DrawAxis(PictureRA, 0, i, -1, -1)
    
        If emulHemisphere = 0 Then
            i = emulDec_DegNoAdjust - 90
        Else
            i = emulDec_DegNoAdjust - 270
        End If
        If i < 0 Then i = 360 + i
        If i > 360 Then i = i - 360
        i = (i / 360) * 100
        If emulHemisphere <> 0 Then
            i = 100 - i
        End If
        Call DrawAxis(PictureDEC, 1, 100 - i, -1, -1)
    
    End If
    
    Label1.Caption = printhex(emulRA_Encoder)
    Label3.Caption = Format$(str(emulRA_Hours), "000.0000")
    Label2.Caption = printhex(emulDEC_Encoder)
    
    timer1flag = False
End If


End Sub


Private Sub Timer1_Timer()

    Dim tha As Double
    Dim tRA As Double
    Dim tDEC As Double
    Dim tAz As Double
    Dim tAlt As Double
    Dim i As Double
    
    If Not timer2flag Then
        timer2flag = True
        emulLongitude = CDbl(EQFixNum(HC.txtLongDeg)) + (CDbl(EQFixNum(HC.txtLongMin)) / 60#) + (CDbl(EQFixNum(HC.txtLongSec)) / 3600#)
        If HC.cbEW.Text = "W" Then emulLongitude = -emulLongitude  ' W is neg
        
        emulLatitude = CDbl(EQFixNum(HC.txtLatDeg)) + (CDbl(EQFixNum(HC.txtLatMin)) / 60#) + (CDbl(EQFixNum(HC.txtLatSec)) / 3600#)
        If HC.cbNS.Text = "S" Then emulLatitude = -emulLatitude
        
        emulElevation = CDbl(EQFixNum(HC.txtElevation))
        
        
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
        
        Select Case SOP_Physical(emulRA_Hours)
            Case pierUnknown:    Label47.Caption = "Unknown"
            Case pierEast:       Label47.Caption = "East, Scope to West"
            Case pierWest:       Label47.Caption = "West, Scope to East"
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
        timer2flag = False
    End If
End Sub


Private Sub SetText()
EQSIM.Caption = oLangDll.GetLangString(6100)
Frame2.Caption = oLangDll.GetLangString(105)
Frame3.Caption = oLangDll.GetLangString(106)
Label9.Caption = oLangDll.GetLangString(105)
Label13.Caption = oLangDll.GetLangString(106)
Label12.Caption = oLangDll.GetLangString(2112)

Label11.Caption = oLangDll.GetLangString(6101)
Label1.ToolTipText = oLangDll.GetLangString(6105)
Label16.Caption = oLangDll.GetLangString(6101)
Label2.ToolTipText = oLangDll.GetLangString(6105)

Label10.Caption = oLangDll.GetLangString(6104)
Label43.ToolTipText = oLangDll.GetLangString(6106)
Label18.Caption = oLangDll.GetLangString(6104)
Label44.ToolTipText = oLangDll.GetLangString(6106)

Label3.ToolTipText = oLangDll.GetLangString(6107)
Label4.ToolTipText = oLangDll.GetLangString(6108)

Check1.Caption = oLangDll.GetLangString(6103)

End Sub



