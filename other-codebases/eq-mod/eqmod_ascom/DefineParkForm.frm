VERSION 5.00
Begin VB.Form DefineParkForm 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DefinePark"
   ClientHeight    =   5205
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   4455
   Icon            =   "DefineParkForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "User Defined Unpark Positions"
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
      TabIndex        =   4
      Top             =   2640
      Width           =   4215
      Begin VB.PictureBox PictureRA 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   1
         Left            =   120
         ScaleHeight     =   81
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   81
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.PictureBox PictureDEC 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   1
         Left            =   1440
         ScaleHeight     =   81
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   81
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton CommandResetUnparkPos 
         BackColor       =   &H0095C1CB&
         Caption         =   "Reset"
         Height          =   375
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton CommandSetUnparkPos 
         BackColor       =   &H0095C1CB&
         Caption         =   "Set"
         Height          =   375
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1920
         Width           =   1215
      End
      Begin VB.ComboBox ComboUparkPos 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   1920
         Width           =   2535
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
         Index           =   1
         Left            =   1440
         TabIndex        =   15
         Top             =   1560
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
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "User Defined Park Positions"
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
      Top             =   120
      Width           =   4215
      Begin VB.PictureBox PictureRA 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   0
         Left            =   120
         ScaleHeight     =   81
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   81
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.PictureBox PictureDEC 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   0
         Left            =   1440
         ScaleHeight     =   81
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   81
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton CommandResetParkPos 
         BackColor       =   &H0095C1CB&
         Caption         =   "Reset"
         Height          =   375
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton CommandSetParkPos 
         BackColor       =   &H0095C1CB&
         Caption         =   "Set"
         Height          =   375
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1920
         Width           =   1215
      End
      Begin VB.ComboBox ComboParkPos 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   1920
         Width           =   2535
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
         Index           =   0
         Left            =   1440
         TabIndex        =   11
         Top             =   1560
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
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   1215
      End
   End
End
Attribute VB_Name = "DefineParkForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private selectedPark As Integer
Private selectedUnPark As Integer
Private lblmode As Integer

Private Sub ComboParkPos_Click()
    selectedPark = ComboParkPos.ListIndex
    update_park_display (selectedPark + 1)
End Sub
Private Sub ComboUParkPos_Click()
    selectedUnPark = ComboUparkPos.ListIndex
    update_unpark_display (selectedUnPark + 1)
End Sub

Private Sub CommandResetParkPos_Click()
    Dim Index As Integer
    Dim name As String
    Index = ComboParkPos.ListIndex + 1
    If Index <= 10 Then
        ComboParkPos.List(ComboParkPos.ListIndex) = oLangDll.GetLangString(2730)
        UserParks(Index).name = oLangDll.GetLangString(2730)
        UserParks(Index).posR = 0
        UserParks(Index).posD = 0
        Call writeUserParkPos
        Call HC.SetParkCombo
        update_park_display (Index)
    End If
End Sub

Private Sub CommandResetUnparkPos_Click()
    Dim Index As Integer
    Dim name As String
    Index = ComboUparkPos.ListIndex + 1
    If Index <= 10 Then
        ComboUparkPos.List(ComboUparkPos.ListIndex) = oLangDll.GetLangString(2730)
        UserUnparks(Index).name = oLangDll.GetLangString(2730)
        UserUnparks(Index).posR = 0
        UserUnparks(Index).posD = 0
        Call writeUserParkPos
        Call HC.SetParkCombo
        update_unpark_display (Index)
    End If
End Sub

Private Sub CommandSetParkPos_Click()
    Dim Index As Integer
    Dim aname As String
    
    Index = selectedPark + 1
    If Index <= 10 Then
        aname = ComboParkPos.Text
        If aname = oLangDll.GetLangString(2730) Then
            aname = "UserPark" & CStr(Index)
        End If
        ComboParkPos.List(selectedPark) = aname
        ComboParkPos.ListIndex = selectedPark
        Call DefineUserPark(True, Index, aname)
        Call HC.SetParkCombo
        update_park_display (Index)
    End If
End Sub

Private Sub CommandSetUnparkPos_Click()
    Dim Index As Integer
    Dim aname As String
    
    Index = selectedUnPark + 1
    If Index <= 10 Then
        aname = ComboUparkPos.Text
        If aname = oLangDll.GetLangString(2730) Then
            aname = "UserUnpark" & CStr(Index)
        End If
        ComboUparkPos.List(selectedUnPark) = aname
        ComboUparkPos.ListIndex = selectedUnPark
        Call DefineUserUnPark(True, Index, aname)
        Call HC.SetParkCombo
        update_unpark_display (Index)
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    lblmode = 1
    
    SetText
    Call PutWindowOnTop(DefineParkForm)
    
    ComboParkPos.Clear
    ComboUparkPos.Clear
    
    For i = 1 To 10
        ComboParkPos.AddItem (UserParks(i).name)
        ComboUparkPos.AddItem (UserUnparks(i).name)
    Next i
    
    ComboParkPos.ListIndex = 0
    ComboUparkPos.ListIndex = 0
End Sub

Private Sub update_park_display(Index As Integer)
    Dim i As Double
    Dim Dec_DegNoAdjust As Double
    Dim RA_Hours As Double

    If UserParks(Index).posR = 0 Or UserParks(Index).posR = 0 Then
        Call DrawAxis(PictureRA(0), -1, 0, -1, -1)
        Call DrawAxis(PictureDEC(0), -1, 0, -1, -1)
        Label83(0).Caption = ""
        Label82(0).Caption = ""
        Exit Sub
    End If

    RA_Hours = Get_EncoderHours(gRAEncoder_Zero_pos, CDbl(UserParks(Index).posR), gTot_RA, gHemisphere)
    Dec_DegNoAdjust = Get_EncoderDegrees(gDECEncoder_Zero_pos, CDbl(UserParks(Index).posD), gTot_DEC, gHemisphere)
    
    i = (Range24(RA_Hours - 6) / 24) * 100
    If gHemisphere <> 0 Then
        i = 100 - i
    End If
    Call DrawAxis(PictureRA(0), 0, i, -1, -1)
    If lblmode = 1 Then
        Label82(0).Caption = FmtSexa(RA_Hours, False)
    Else
        Label82(0).Caption = printhex(CDbl(UserParks(Index).posR))
    End If

    If gHemisphere = 0 Then
        i = Dec_DegNoAdjust - 90
    Else
        i = Dec_DegNoAdjust - 270
    End If
    If i < 0 Then i = 360 + i
    If i > 360 Then i = i - 360
    i = (i / 360) * 100
    If gHemisphere <> 0 Then
        i = 100 - i
    End If
    Call DrawAxis(PictureDEC(0), 1, 100 - i, -1, -1)
    If lblmode = 1 Then
        Label83(0).Caption = FmtSexa(Dec_DegNoAdjust, False)
    Else
        Label83(0).Caption = printhex(CDbl(UserParks(Index).posD))
    End If
End Sub

Private Sub update_unpark_display(Index As Integer)
    Dim i As Double
    Dim Dec_DegNoAdjust As Double
    Dim RA_Hours As Double

    If UserUnparks(Index).posR = 0 Or UserUnparks(Index).posR = 0 Then
        Call DrawAxis(PictureRA(1), -1, 0, -1, -1)
        Call DrawAxis(PictureDEC(1), -1, 0, -1, -1)
        Label83(1).Caption = ""
        Label82(1).Caption = ""
        Exit Sub
    End If

    RA_Hours = Get_EncoderHours(gRAEncoder_Zero_pos, CDbl(UserUnparks(Index).posR), gTot_RA, gHemisphere)
    Dec_DegNoAdjust = Get_EncoderDegrees(gDECEncoder_Zero_pos, CDbl(UserUnparks(Index).posD), gTot_DEC, gHemisphere)
    
    i = (Range24(RA_Hours - 6) / 24) * 100
    If gHemisphere <> 0 Then
        i = 100 - i
    End If
    Call DrawAxis(PictureRA(1), 0, i, -1, -1)
    If lblmode = 1 Then
        Label82(1).Caption = FmtSexa(RA_Hours, False)
    Else
        Label82(1).Caption = printhex(CDbl(UserUnparks(Index).posR))
    End If

    If gHemisphere = 0 Then
        i = Dec_DegNoAdjust - 90
    Else
        i = Dec_DegNoAdjust - 270
    End If
    If i < 0 Then i = 360 + i
    If i > 360 Then i = i - 360
    i = (i / 360) * 100
    If gHemisphere <> 0 Then
        i = 100 - i
    End If
    Call DrawAxis(PictureDEC(1), 1, 100 - i, -1, -1)
    If lblmode = 1 Then
        Label83(1).Caption = FmtSexa(Dec_DegNoAdjust, False)
    Else
        Label83(1).Caption = printhex(CDbl(UserUnparks(Index).posD))
    End If
End Sub


Private Sub SetText()
    DefineParkForm.Caption = oLangDll.GetLangString(2500)
    Frame1.Caption = oLangDll.GetLangString(2501)
    Frame2.Caption = oLangDll.GetLangString(2502)
    CommandSetParkPos.Caption = oLangDll.GetLangString(2503)
    CommandSetUnparkPos.Caption = oLangDll.GetLangString(2503)
    CommandResetParkPos.Caption = oLangDll.GetLangString(2504)
    CommandResetUnparkPos.Caption = oLangDll.GetLangString(2504)
End Sub

Private Sub Label82_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblmode = Button
    update_park_display (selectedPark + 1)
    update_unpark_display (selectedUnPark + 1)
End Sub

Private Sub Label83_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblmode = Button
    update_park_display (selectedPark + 1)
    update_unpark_display (selectedUnPark + 1)
End Sub
