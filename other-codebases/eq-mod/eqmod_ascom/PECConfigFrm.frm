VERSION 5.00
Begin VB.Form PECConfigFrm 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PEC Configuration"
   ClientHeight    =   4470
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "AutoPEC"
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
      Height          =   2775
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   4815
      Begin VB.CheckBox CheckCaptureOnly 
         BackColor       =   &H00000000&
         Caption         =   "Auto Apply PEC"
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
         Left            =   2640
         TabIndex        =   21
         Top             =   360
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CommandButton CommandFileDir 
         Height          =   375
         Left            =   120
         Picture         =   "PECConfigFrm.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Set Working directory"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox Text1 
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
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   2280
         Width           =   4215
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Timestamp Files"
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
         TabIndex        =   18
         Top             =   1920
         Width           =   1695
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   150
         Left            =   120
         Max             =   50
         TabIndex        =   15
         Top             =   1560
         Value           =   10
         Width           =   2055
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   150
         Left            =   120
         Max             =   100
         Min             =   9
         TabIndex        =   12
         Top             =   1080
         Value           =   30
         Width           =   2055
      End
      Begin VB.ComboBox ComboPecCap 
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
         ItemData        =   "PECConfigFrm.frx":0582
         Left            =   1680
         List            =   "PECConfigFrm.frx":0592
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
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
         Left            =   1440
         TabIndex        =   17
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Mag. Filter"
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
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
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
         TabIndex        =   14
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "LoPassFilter"
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
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Capture Cycles"
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
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Playback"
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
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.ComboBox PECMethodCombo 
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
         ItemData        =   "PECConfigFrm.frx":05A2
         Left            =   2640
         List            =   "PECConfigFrm.frx":05AC
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin VB.HScrollBar PhaseScroll 
         Height          =   150
         Left            =   120
         Max             =   479
         TabIndex        =   3
         Top             =   960
         Width           =   2055
      End
      Begin VB.HScrollBar GainScroll 
         Height          =   150
         Left            =   120
         Max             =   50
         TabIndex        =   2
         Top             =   480
         Value           =   10
         Width           =   2055
      End
      Begin VB.CheckBox CheckTracePec 
         BackColor       =   &H00000000&
         Caption         =   "Debug Trace"
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
         Left            =   2640
         TabIndex        =   1
         Top             =   840
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label48 
         BackColor       =   &H00000000&
         Caption         =   "Phaseverschuiving"
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
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label46 
         BackColor       =   &H00000000&
         Caption         =   "Gain"
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
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "360 deg"
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
         TabIndex        =   5
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Gain"
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
         Left            =   1440
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "PECConfigFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Check1_Click()
    gPEC_TimeStampFiles = Check1.Value
    Call PEC_WriteParams
End Sub

Private Sub CheckCaptureOnly_Click()
    gPEC_AutoApply = CheckCaptureOnly.Value
    Call PEC_WriteParams
End Sub

Private Sub CheckTracePec_Click()
    gPEC_trace = CheckTracePec.Value
End Sub

Private Sub ComboPecCap_Click()
    gPEC_Capture_Cycles = ComboPecCap.ItemData(ComboPecCap.ListIndex)
    Call PEC_WriteParams
End Sub

Private Sub CommandFileDir_Click()
    FileDlg.Show (1)
    Text1.Text = FileDlg.Dir1
    gPEC_FileDir = FileDlg.Dir1

End Sub

Private Sub Form_Load()
    On Error Resume Next

    Call SetText
    
    Call PEC_ReadParams
    
    Select Case gPEC_Capture_Cycles
        Case 5
            ComboPecCap.ListIndex = 0
        Case 4
            ComboPecCap.ListIndex = 1
        Case 3
            ComboPecCap.ListIndex = 2
        Case 2
            ComboPecCap.ListIndex = 3
        Case Else
    End Select
    
    If gPEC_filter_lowpass < HScroll1.min Then
        HScroll1.Value = HScroll1.min
    Else
        HScroll1.Value = gPEC_filter_lowpass
    End If
    HScroll2.Value = gPEC_mag
    GainScroll.Value = gPEC_Gain * 10
    PhaseScroll.Value = gPEC_PhaseAdjust
    
    Call PEC_PhaseScroll_Change
    Call PEC_GainScroll_Change
    
    Check1.Value = gPEC_TimeStampFiles
    Text1.Text = gPEC_FileDir
    
    PECMethodCombo.ListIndex = gPEC_DynamicRateAdjust
    CheckTracePec.Value = gPEC_trace
    CheckCaptureOnly.Value = gPEC_AutoApply


    Call PutWindowOnTop(PECConfigFrm)

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call PEC_WriteParams
End Sub

Private Sub HScroll1_Change()
    Call PEC_LoPassScroll_Change
End Sub

Private Sub HScroll1_Scroll()
    Call HScroll1_Change
End Sub

Private Sub HScroll2_Change()
    Call PEC_MagScroll_Change
End Sub

Private Sub HScroll2_Scroll()
    Call HScroll2_Change
End Sub



Private Sub PECMethodCombo_Click()
    gPEC_DynamicRateAdjust = PECMethodCombo.ListIndex
    Call PECMode_click
End Sub

Private Sub PhaseScroll_Change()
    Call PEC_PhaseScroll_Change
End Sub

Private Sub PhaseScroll_Scroll()
    PhaseScroll_Change
End Sub

Private Sub GainScroll_Change()
    Call PEC_GainScroll_Change
End Sub

Private Sub GainScroll_Scroll()
    GainScroll_Change
End Sub

Private Sub SetText()
    PECConfigFrm.Caption = oLangDll.GetLangString(6126)
    PECMethodCombo.Clear
    PECMethodCombo.AddItem oLangDll.GetLangString(6124)
    PECMethodCombo.AddItem oLangDll.GetLangString(6125)
    
    Label1(0).Caption = oLangDll.GetLangString(6120)
    Label1(1).Caption = oLangDll.GetLangString(6121)
    Label1(2).Caption = oLangDll.GetLangString(6122)
    Label46.Caption = oLangDll.GetLangString(191)
    Label48.Caption = oLangDll.GetLangString(192)
    Check1.Caption = oLangDll.GetLangString(6123)
    Frame1.Caption = oLangDll.GetLangString(6129)
    Frame2.Caption = oLangDll.GetLangString(6130)
    CheckCaptureOnly.Caption = oLangDll.GetLangString(6131)
    'ComboPecCap.ListIndex = 0
End Sub
