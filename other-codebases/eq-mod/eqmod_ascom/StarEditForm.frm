VERSION 5.00
Begin VB.Form StarEditform 
   BackColor       =   &H00000000&
   Caption         =   "Points List Editor"
   ClientHeight    =   5550
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11940
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "StarEditForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   11940
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Alignment List"
      ForeColor       =   &H000080FF&
      Height          =   3255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   6375
      Begin VB.ListBox StarList 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         ForeColor       =   &H000080FF&
         Height          =   2130
         Left            =   120
         TabIndex        =   13
         Top             =   570
         Width           =   6135
      End
      Begin VB.CommandButton Delete_Star_Command 
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
         Picture         =   "StarEditForm.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2760
         Width           =   375
      End
      Begin VB.CommandButton Goto_Command 
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
         Left            =   675
         Picture         =   "StarEditForm.frx":13F0
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2760
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   345
         Width           =   6135
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Triangle"
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
         Height          =   255
         Left            =   3360
         TabIndex        =   19
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000040&
         Caption         =   " 0"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5160
         TabIndex        =   18
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000040&
         Caption         =   " 0"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5760
         TabIndex        =   17
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000040&
         Caption         =   " 0"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4560
         TabIndex        =   16
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label Nstarbl1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000040&
         Caption         =   " 0"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2715
         TabIndex        =   15
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Nearest"
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
         Height          =   255
         Left            =   1200
         TabIndex        =   14
         Top             =   2880
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Load / Save Alignment Preset"
      ForeColor       =   &H000080FF&
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   6375
      Begin VB.CheckBox Check3 
         BackColor       =   &H00000000&
         Caption         =   "Save Alignment Points to Preset on append"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1560
         Width           =   5895
      End
      Begin VB.ComboBox PresetCombo 
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
         Left            =   840
         TabIndex        =   8
         Text            =   "Presets"
         Top             =   360
         Width           =   4575
      End
      Begin VB.CommandButton Save_Command 
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
         Left            =   5640
         Picture         =   "StarEditForm.frx":1B16
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton LoadCommand 
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
         Picture         =   "StarEditForm.frx":2098
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Load Alignment Stars from Preset on Unpark"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   5775
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00000000&
         Caption         =   "Save Alignment Stars to Preset on Park"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   5895
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   2
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton CommandTransform 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      Picture         =   "StarEditForm.frx":261A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   375
   End
   Begin VB.PictureBox SkyPlot 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   6600
      ScaleHeight     =   5115
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   240
      Width           =   5175
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   120
         Top             =   120
      End
   End
   Begin VB.Menu popupmenu1 
      Caption         =   "MyMenu"
      Visible         =   0   'False
      Begin VB.Menu GotoPoint 
         Caption         =   "Goto Selected"
         Visible         =   0   'False
      End
      Begin VB.Menu sep 
         Caption         =   "--------------------"
         Visible         =   0   'False
      End
      Begin VB.Menu deletePoint 
         Caption         =   "Delete Selected"
      End
   End
End
Attribute VB_Name = "StarEditform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private PresetIdx As Integer

Private Declare Function SendMessage Lib "user32" Alias _
        "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As _
        Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LB_SETTABSTOPS = &H192

' The number of tabs
Const OBorder = 5
 
Dim Tabulator(1 To OBorder) As Long
Public RA_Zpos As Double
Public DEC_Zpos As Double
Public MotorTotPos As Double
Public TotSize As Double
Public tmpx1 As Double
Public tmpy1 As Double
Public tmpx2 As Double
Public tmpy2 As Double
Public RefreshDisplay As Boolean
Dim xr As Double
Dim yr As Double

Dim HelpOffsetX As Double
Dim HelpOffsetY As Double
Dim StdFormWidth As Double
Dim StdFormHeight As Double


Dim plotcentre As Double
Dim timerflag As Boolean
Dim lastNearest As Long
Dim lastCount As Long
Dim af1 As Long
Dim af2 As Long
Dim af3 As Long
Dim DisplayMode As Integer
Dim RubOut As Boolean
Dim IngnorResize As Boolean

Private Sub Check1_Click()
Dim Alignini As String
   ' set up a file path for the align.ini file
    gLoadAPresetOnUnpark = Check1.Value
    Alignini = HC.oPersist.GetIniPath & "\ALIGN.ini"
    Call HC.oPersist.WriteIniValueEx("LOAD_APRESET_ON_UNPARK", CStr(Check1.Value), "[default]", Alignini)
End Sub

Private Sub Check2_Click()
Dim Alignini As String
   ' set up a file path for the align.ini file
    gSaveAPresetOnPark = Check2.Value
    Alignini = HC.oPersist.GetIniPath & "\ALIGN.ini"
    Call HC.oPersist.WriteIniValueEx("SAVE_APRESET_ON_UNPARK", CStr(Check2.Value), "[default]", Alignini)
End Sub


Private Sub Check3_Click()
Dim Alignini As String
   ' set up a file path for the align.ini file
    gSaveAPresetOnAppend = Check3.Value
    Alignini = HC.oPersist.GetIniPath & "\ALIGN.ini"
    Call HC.oPersist.WriteIniValueEx("SAVE_APRESET_ON_APPEND", CStr(Check3.Value), "[default]", Alignini)
End Sub

Private Sub CommandTransform_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Select Case DisplayMode
        Case 0
            DisplayMode = 3
            Call ShowHelp1
        Case 1
            DisplayMode = 4
            Call ShowHelp2
    End Select
End Sub

Private Sub CommandTransform_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    Select Case DisplayMode
        Case 3
            Call PlotAlignmentPoints
            DisplayMode = 0
        Case 4
            Call PlotTransform
            DisplayMode = 1
    End Select
    
End Sub

Private Sub Delete_Star_Command_Click()

Dim Index, DeleteThisOne As Integer
    
    If StarList.ListIndex <> -1 Then
        DeleteThisOne = StarList.ListIndex + 1 ' listindex is 0 based, starlist is 1 based
        Call EQ_NPointDelete(DeleteThisOne)
        FillStarList
        
        ' send to matrix will initialise the catalog and measured points arrays
        Call SendtoMatrix
    End If
End Sub



Private Sub Form_Load()
    
    RA_Zpos = RAEncoder_Home_pos
    DEC_Zpos = gDECEncoder_Home_pos
    MotorTotPos = gTot_step
    TotSize = 7000
    tmpx1 = 0
    tmpy1 = 0
    Call PlotInit(SkyPlot)
    
    Call SetText
    
    ' set tabs values
    Tabulator(1) = 20
    Tabulator(2) = 100
    Tabulator(3) = 140
    Tabulator(4) = 180
    Tabulator(5) = 220
    SendMessage StarList.hwnd, LB_SETTABSTOPS, OBorder, Tabulator(1)
    SendMessage List1.hwnd, LB_SETTABSTOPS, OBorder, Tabulator(1)
    
    ' poplate the star list
    
    If HC.HCOnTop.Value = 1 Then Call PutWindowOnTop(StarEditform)
    Call FillStarList
    ' Initialise the Preset Combo
    Call ReadPresets
    Call ReadParkOptions
    
    Check1.Value = gLoadAPresetOnUnpark
    Check2.Value = gSaveAPresetOnPark
    Check3.Value = gSaveAPresetOnAppend
    
    DisplayMode = 0
    
    HelpOffsetX = Me.width - CommandTransform.Left
    HelpOffsetY = Me.Height - CommandTransform.Top
    StdFormWidth = Me.width
    StdFormHeight = Me.Height
    
    Call PlotAlignmentPoints
    
    lastCount = gAlignmentStars_count
    timerflag = False
    Timer1.Enabled = True
    
End Sub

Public Sub FillStarList()
Dim Index As Integer
Dim tmpstr As String
    StarList.Clear
    On Error Resume Next
    For Index = 1 To gAlignmentStars_count
        With AlignmentStars(Index)
            tmpstr = CStr(Index) & vbTab & CStr(.AlignTime) & vbTab & FmtSexa(.OrigTargetRA, False) & vbTab & FmtSexa(.OrigTargetDEC, False) & vbTab & Format$(str(.TargetRA - .EncoderRA), "0000000") & vbTab & Format$(str(.TargetDEC - .EncoderDEC), "0000000")
            StarList.AddItem (tmpstr)
'            StarList.AddItem (CStr(Index) + ") " + CStr(.AlignTime) + " : RA[" + FmtSexa(.OrigTargetRA, False) + "] , DEC[" + FmtSexa(.OrigTargetDEC, False) + "]" + "  RA:" + Format$(str(.TargetRA - .EncoderRA), "0000000") + "  DEC:" + Format$(str(.TargetDEC - .EncoderDEC), "0000000"))
        End With
    Next Index
End Sub


Private Sub Form_Resize()

    On Error Resume Next

    If Me.Height < StdFormHeight Or Me.width < StdFormWidth Then
        Me.Height = StdFormHeight
    End If
    
    If Me.width < StdFormWidth Then
        SkyPlot.width = 5175
        Me.width = StdFormWidth
        CommandTransform.Left = Me.width - HelpOffsetX
        CommandTransform.Top = Me.Height - HelpOffsetY
    Else
        If Me.width > StdFormWidth Then
            SkyPlot.width = Me.width - SkyPlot.Left - 300
            Me.Height = SkyPlot.width + 800
            CommandTransform.Left = Me.width - HelpOffsetX
            CommandTransform.Top = Me.Height - HelpOffsetY
        End If
    End If
    
    SkyPlot.Height = SkyPlot.width
    Call PlotInit(SkyPlot)
    Select Case DisplayMode
        Case 0
            Call PlotAlignmentPoints
        Case 1
            Call PlotTransform
    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' save current preset index to ini
    Call SavePresetIdx(PresetIdx)
    Timer1.Enabled = False
    Unload StarEditform
End Sub

Private Sub Goto_Command_Click()
    With AlignmentStars(StarList.ListIndex + 1)
        If (.OrigTargetRA + .OrigTargetDEC) <> 0 Then
            Call m_telescope.SlewToCoordinates(.OrigTargetRA, .OrigTargetDEC)
        End If
    End With
End Sub

Private Sub deletePoint_Click()
    Delete_Star_Command_Click
End Sub

Private Sub GotoPoint_Click()
    Goto_Command_Click
End Sub

Private Sub List1_Click()
    On Error Resume Next
    List1.ListIndex = -1
End Sub


Private Sub LoadCommand_Click()
    If LoadAlignmentPreset(PresetIdx) = False Then
        'attempt to load an empty preset - action aborted
        MsgBox (oLangDll.GetLangString(2300))
    End If
    
    FillStarList
    Call SavePresetIdx(PresetIdx)
    
    PlotAlignmentPoints
End Sub


Private Sub PresetCombo_Click()
    PresetIdx = PresetCombo.ListIndex
    ' Force loss of focus
     SendKeys "{TAB}", True

End Sub

Private Sub Save_Command_Click()
    PresetCombo.List(PresetIdx) = PresetCombo.Text
    ' save the preset to file
    Call SaveAlignmentStars(PresetIdx, PresetCombo.List(PresetIdx))
    ' load it from file
    PresetCombo.ListIndex = PresetIdx
    Call SavePresetIdx(PresetIdx)
End Sub



Private Sub ReadPresets()
Dim Index As Integer
Dim keyStr As String
Dim tmptxt As String
Dim Alignini As String

    ' set up a file path for the align.ini file
    Alignini = HC.oPersist.GetIniPath & "\ALIGN.ini"

    For Index = 0 To 9
    
        keyStr = "[alignment_preset" & CStr(Index) & "]"
        
        tmptxt = HC.oPersist.ReadIniValueEx("NAME", keyStr, Alignini)
        If tmptxt = "" Then
            ' create a preset place holder
            tmptxt = "Empty_" & CStr(Index)
            Call HC.oPersist.WriteIniValueEx("NAME", tmptxt, keyStr, Alignini)
        End If
        
        PresetCombo.AddItem (tmptxt)
        
        tmptxt = HC.oPersist.ReadIniValueEx("STAR_COUNT", keyStr, Alignini)
        If tmptxt = "" Then
            ' create a preset place holder
            tmptxt = "0"
            Call HC.oPersist.WriteIniValueEx("STAR_COUNT", tmptxt, keyStr, Alignini)
        End If
    
    Next Index

    PresetCombo.ListIndex = GetPresetIdx
    
End Sub





'------------------------
' Public functions follow
'------------------------




Private Sub SetText()
    Dim str As String
    StarEditform.Caption = oLangDll.GetLangString(900)
    Delete_Star_Command.ToolTipText = oLangDll.GetLangString(902)
    deletePoint.Caption = oLangDll.GetLangString(902)
    Goto_Command.ToolTipText = oLangDll.GetLangString(903)
    GotoPoint.Caption = oLangDll.GetLangString(903)
    LoadCommand.ToolTipText = oLangDll.GetLangString(904)
    Save_Command.ToolTipText = oLangDll.GetLangString(905)
    PresetCombo.Text = oLangDll.GetLangString(907)
    Check1.Caption = oLangDll.GetLangString(908)
    Check2.Caption = oLangDll.GetLangString(909)
    Check3.Caption = oLangDll.GetLangString(910)

    str = oLangDll.GetLangString(1013)
    str = Replace(str, " ", vbTab)
    List1.AddItem (str)
    
    Frame2.Caption = oLangDll.GetLangString(1014)
    Frame1.Caption = oLangDll.GetLangString(1015)
    Label1.Caption = oLangDll.GetLangString(1016)
    Label5.Caption = oLangDll.GetLangString(1017)

End Sub


Public Sub PlotAlignmentPoints()
Dim i As Integer
Dim tmpobj As Coord
Dim tmpobj2 As Coord
Dim tmpobj3 As Coord
Dim col As Long
Dim size As Long
Dim x As Double
Dim x1 As Double
Dim Y As Double
Dim y1 As Double
Dim idx As Integer

    Call draw_lines(SkyPlot)

    ' Draw the Reference Points (catalog and measured stars)
    idx = 0
    For i = 1 To gAlignmentStars_count
        ' plot catalog
        ' tmpobj = ct_PointsC(i)
        tmpobj = EQ_sp2Cs2(ct_Points(i))
        Call NStarPlotCircle(SkyPlot, tmpobj.x, tmpobj.Y, 30, vbBlue, 1)
        If gAlignmentStars_count >= 3 Then
            If i = gAffine1 Or i = gAffine2 Or i = gAffine3 Then
                idx = idx + 1
                If idx = 3 Then
'                    tmpobj = my_PointsC(gAffine1)
'                    tmpobj2 = my_PointsC(gAffine2)
'                    tmpobj3 = my_PointsC(gAffine3)
                    tmpobj = EQ_sp2Cs2(my_Points(gAffine1))
                    tmpobj2 = EQ_sp2Cs2(my_Points(gAffine2))
                    tmpobj3 = EQ_sp2Cs2(my_Points(gAffine3))
                    Call NStarPlotLine2(SkyPlot, tmpobj.x, tmpobj.Y, tmpobj2.x, tmpobj2.Y, vbRed)
                    Call NStarPlotLine2(SkyPlot, tmpobj.x, tmpobj.Y, tmpobj3.x, tmpobj3.Y, vbRed)
                    Call NStarPlotLine2(SkyPlot, tmpobj2.x, tmpobj2.Y, tmpobj3.x, tmpobj3.Y, vbRed)
                End If
            End If
        End If
        size = 30
        col = vbRed
        If gSelectStar <> 0 Then
            If i = gSelectStar Then
                col = vbYellow
                size = 60
            End If
        End If
        ' plot measured
        tmpobj = EQ_sp2Cs2(my_Points(i))
        If i = StarList.ListIndex + 1 Then
            Call NStarPlotCircle(SkyPlot, tmpobj.x, tmpobj.Y, 90, vbGreen, 1)
        End If
        Call NStarPlotCircle(SkyPlot, tmpobj.x, tmpobj.Y, size, col, 1)
    Next i
    
    
End Sub


Public Sub drawmount(ByVal x1 As Double, ByVal y1 As Double)

Dim i As Integer
Dim tmpobj As Coord
Dim tmpobj3 As Coord
Dim tmpobj2 As Coordt
Dim col As Long

'    If gAlignmentStars_count <> 0 Then
    
         ' remove last position
         If RubOut Then
            Call NStarPlotCircleXor(SkyPlot, tmpx1, tmpy1, 100, vbWhite, 1)
            Call NStarPlotCircleXor(SkyPlot, tmpx2, tmpy2, 30, vbWhite, 1)
        Else
            RubOut = True
        End If
        
        tmpobj.x = x1 - gRASync01
        tmpobj.Y = y1 - gDECSync01
        
        ' Draw the mount's current position
        tmpobj3 = EQ_sp2Cs2(tmpobj)
        tmpx1 = tmpobj3.x
        tmpy1 = tmpobj3.Y
        Call NStarPlotCircleXor(SkyPlot, tmpx1, tmpy1, 100, vbWhite, 1)
                
    '     tmpobj = EQ_Transform_Taki(EQ_sp2Cs(tmpobj))
    '     Call NStarPlotCircle(tmpobj.x, tmpobj.y, 0, 0, 30, vbWhite)
                 
        Select Case gAlignmentMode
           Case 0
               ' nstar+nearest
               tmpobj2 = Delta_Matrix_Map(tmpobj.x, tmpobj.Y)
               If tmpobj2.f = 0 Then
                 tmpobj2 = DeltaSyncReverse_Matrix_Map(tmpobj.x, tmpobj.Y)
               End If
           Case 1
               ' nstar
               tmpobj2 = Delta_Matrix_Map(tmpobj.x, tmpobj.Y)
           Case Else
               ' nearest
               tmpobj2 = DeltaSyncReverse_Matrix_Map(tmpobj.x, tmpobj.Y)
        End Select
                
        tmpobj.x = tmpobj2.x
        tmpobj.Y = tmpobj2.Y
        
        tmpobj3 = EQ_sp2Cs2(tmpobj)
        tmpx2 = tmpobj3.x
        tmpy2 = tmpobj3.Y
         
         ' plot mounts transformed position
         Call NStarPlotCircleXor(SkyPlot, tmpx2, tmpy2, 30, vbWhite, 1)
                 
'    End If



End Sub

Private Sub PlotTransform()
Dim RA As Double
Dim DEC As Double
Dim tmpcoord As Coordt
Dim tmpc1 As Coord
Dim tmpc2 As Coord
Dim tmpsp As SphereCoord
Dim Onestar As Integer

Dim DECStart As Double
Dim DECEnd As Double
Dim RAStart As Double
Dim RAEnd As Double
Dim SkyLimit As Double

Dim ra_step As Double
Dim dec_step As Double

    If gAlignmentStars_count < 1 Then Exit Sub
    
    Call draw_lines(SkyPlot)


    ra_step = 180000
    dec_step = -180000
    RAStart = (gRAEncoder_Zero_pos - (gTot_step / 4))
    RAEnd = (gRAEncoder_Zero_pos + (gTot_step / 4) - ra_step)
    DECStart = (gDECEncoder_Zero_pos + (gTot_step / 4))
    DECEnd = (gDECEncoder_Zero_pos - (gTot_step / 4))
    SkyLimit = ((gTot_step / 2) + gDECEncoder_Home_pos)
    
    For RA = RAStart To RAEnd Step ra_step
        DoEvents
        
        For DEC = DECStart To DECEnd Step dec_step
        
            tmpsp = EQ_SphericalPolar(RA, DEC, gTot_step, gRAEncoder_Zero_pos, gDECEncoder_Home_pos, Abs(gLatitude))
            ' Check if sky visible
            If tmpsp.Y < SkyLimit Then
            
                Onestar = 0
                
                Select Case gAlignmentMode
                    
                    Case 0
                        ' nstar+nearest
                        tmpcoord = Delta_Matrix_Map(RA - gRASync01, DEC - gDECSync01)
                        If tmpcoord.f = 0 Then
                            Onestar = 1
                            tmpcoord = DeltaSyncReverse_Matrix_Map(RA - gRASync01, DEC - gDECSync01)
                        End If
                    
                    Case 1
                        ' nstar
                        tmpcoord = Delta_Matrix_Map(RA - gRASync01, DEC - gDECSync01)
                    
                    Case 2
                        ' nearest
                        Onestar = 1
                        tmpcoord = DeltaSyncReverse_Matrix_Map(RA - gRASync01, DEC - gDECSync01)
                
                End Select
                
                tmpc1.x = tmpcoord.x
                tmpc1.Y = tmpcoord.Y
                'tmpc1 = EQ_sp2Cs(tmpc1)
                tmpc1 = EQ_sp2Cs2(tmpc1)
                tmpc2.x = RA - gRASync01
                tmpc2.Y = DEC - gDECSync01
                'tmpc2 = EQ_sp2Cs(tmpc2)
                tmpc2 = EQ_sp2Cs2(tmpc2)
            
                Call NStarPlotCross(SkyPlot, tmpc1.x, tmpc1.Y, 100, &H808080)
                If Onestar = 1 Then
                     Call NStarPlotLine(SkyPlot, tmpc1.x, tmpc1.Y, tmpc2.x, tmpc2.Y, vbGreen)
                    Call NStarPlotCross(SkyPlot, tmpc2.x, tmpc2.Y, 75, vbGreen)
                Else
                    Call NStarPlotLine(SkyPlot, tmpc1.x, tmpc1.Y, tmpc2.x, tmpc2.Y, vbRed)
                    Call NStarPlotCross(SkyPlot, tmpc2.x, tmpc2.Y, 75, vbRed)
                End If
            End If
        
        Next DEC
    Next RA

    
    ra_step = -180000
    dec_step = 180000
    RAStart = (gRAEncoder_Zero_pos + (gTot_step / 4))
    RAEnd = (gRAEncoder_Zero_pos - (gTot_step / 4))
    DECStart = (gDECEncoder_Zero_pos + (gTot_step / 4))
    DECEnd = (gDECEncoder_Zero_pos + (gTot_step / 2) + (gTot_step / 4))
    
    For RA = RAStart To RAEnd Step ra_step
       
        DoEvents
        For DEC = DECStart To DECEnd Step dec_step

            tmpsp = EQ_SphericalPolar(RA, DEC, gTot_step, gRAEncoder_Zero_pos, gDECEncoder_Home_pos, Abs(gLatitude))

            ' Check if sky visible
            If tmpsp.Y < SkyLimit Then
                
                Onestar = 0
            
                Select Case gAlignmentMode
                    
                    Case 0
                        ' nstar+nearest
                        tmpcoord = Delta_Matrix_Map(RA - gRASync01, DEC - gDECSync01)
                        If tmpcoord.f = 0 Then
                            Onestar = 1
                            tmpcoord = DeltaSyncReverse_Matrix_Map(RA - gRASync01, DEC - gDECSync01)
                        End If
                     
                    Case 1
                        ' nstar
                        tmpcoord = Delta_Matrix_Map(RA - gRASync01, DEC - gDECSync01)
                    
                    Case Else
                        ' nearest
                        Onestar = 1
                        tmpcoord = DeltaSyncReverse_Matrix_Map(RA - gRASync01, DEC - gDECSync01)
                
                End Select
                
                tmpc1.x = tmpcoord.x
                tmpc1.Y = tmpcoord.Y
                tmpc2.x = RA - gRASync01
                tmpc2.Y = DEC - gDECSync01
'                tmpc1 = EQ_sp2Cs(tmpc1)
'                tmpc2 = EQ_sp2Cs(tmpc2)
                tmpc1 = EQ_sp2Cs2(tmpc1)
                tmpc2 = EQ_sp2Cs2(tmpc2)
                
                Call NStarPlotCross(SkyPlot, tmpc1.x, tmpc1.Y, 100, &H808080)
                If Onestar = 1 Then
                    Call NStarPlotLine(SkyPlot, tmpc1.x, tmpc1.Y, tmpc2.x, tmpc2.Y, vbGreen)
                    Call NStarPlotCross(SkyPlot, tmpc2.x, tmpc2.Y, 75, vbGreen)
                Else
                    Call NStarPlotLine(SkyPlot, tmpc1.x, tmpc1.Y, tmpc2.x, tmpc2.Y, vbRed)
                    Call NStarPlotCross(SkyPlot, tmpc2.x, tmpc2.Y, 75, vbRed)
                End If
            End If
        Next DEC
    Next RA

End Sub


Public Sub draw_lines(ByRef plot As PictureBox)
    Dim tmpobj As Coord

    Call ClearPlot(plot)
    plot.Circle (plotcentre, plotcentre), plotcentre, &H808080
    plot.Line (plotcentre, 0)-(plotcentre, plot.Height), &H7000
    plot.Line (0, plotcentre)-(plot.Height, plotcentre), &H707070

    ' Draw the center DOT (NCP/SCP)
    tmpobj.x = RAEncoder_Home_pos
    tmpobj.Y = gDECEncoder_Home_pos
    tmpobj = EQ_sp2Cs2(tmpobj)
    Call NStarPlotCircle(SkyPlot, tmpobj.x, tmpobj.Y, 100, &H707070, 1)
    Call NStarPlotCircle(SkyPlot, tmpobj.x, tmpobj.Y, 30, &H707070, 1)
End Sub

Public Sub NStarPlotCircle(ByRef plot As PictureBox, ByVal y1 As Double, ByVal x1 As Double, size As Long, dcolor As Long, linewidth As Long)
    x1 = x1 * -1 * xr
    y1 = y1 * yr
    plot.DrawMode = 13
    plot.DrawWidth = linewidth
    plot.Circle (plotcentre + x1, plotcentre + y1), size, dcolor
End Sub
Public Sub NStarPlotCircleXor(ByRef plot As PictureBox, ByVal y1 As Double, ByVal x1 As Double, size As Long, dcolor As Long, linewidth As Long)
    x1 = x1 * -1 * xr
    y1 = y1 * yr
    plot.DrawMode = 7
    plot.DrawWidth = linewidth
    plot.Circle (plotcentre + x1, plotcentre + y1), size, dcolor
End Sub
Public Sub NStarPlotCrossXor(ByRef plot As PictureBox, ByVal y1 As Double, ByVal x1 As Double, size As Long, dcolor As Long)
Dim mid As Double
    mid = size / 2
    x1 = plotcentre - (x1 * xr)
    y1 = plotcentre + y1 * yr
    plot.DrawMode = 7
    plot.Line (x1 - mid, y1)-(x1 + mid, y1), dcolor
    plot.Line (x1, y1 - mid)-(x1, y1 + mid), dcolor
End Sub


Public Sub NStarPlotCross(ByRef plot As PictureBox, ByVal y1 As Double, ByVal x1 As Double, size As Long, dcolor As Long)
Dim mid As Double
    mid = size / 2
    x1 = plotcentre - (x1 * xr)
    y1 = plotcentre + y1 * yr
    plot.DrawMode = 13
    plot.Line (x1 - mid, y1)-(x1 + mid, y1), dcolor
    plot.Line (x1, y1 - mid)-(x1, y1 + mid), dcolor
End Sub
Public Sub NStarPlotLine2(ByRef plot As PictureBox, ByVal y1 As Double, ByVal x1 As Double, ByVal y2 As Double, ByVal x2 As Double, dcolor As Long)
    x1 = plotcentre - x1 * xr
    y1 = plotcentre + y1 * yr
    x2 = plotcentre - x2 * xr
    y2 = plotcentre + y2 * yr
    plot.DrawMode = 13
    plot.Line (x1, y1)-(x2, y2), dcolor
End Sub

Public Sub NStarPlotLine(ByRef plot As PictureBox, ByVal y1 As Double, ByVal x1 As Double, ByVal y2 As Double, ByVal x2 As Double, dcolor As Long)
'    x1 = x1 * -1 * xr
'    x2 = x2 * -1 * xr
'    y1 = x1 * -1 * yr
'    y2 = y2 * -1 * yr
'    plot.DrawMode = 13
'    plot.Line (plotcentre + x1, plotcentre + y1)-(plotcentre + x2, plotcentre + y2), dcolor

Dim xs As Double
Dim ys As Double
Dim xr As Double
Dim yr As Double
Dim i As Long
Dim idiv2 As Double

    If plot.Height < plot.width Then
        i = plot.Height
    Else
        i = plot.width
    End If

    xs = 0
    ys = 0
    
    x1 = x1 * -1
    x2 = x2 * -1
   
    xr = i / 9024000
    yr = i / 9024000
    
    plot.DrawMode = 13
    idiv2 = i / 2
    
    plot.Line ((xs + idiv2 + (x1 * xr)), ys + idiv2 + (y1 * yr))-((xs + idiv2 + (x2 * xr)), ys + idiv2 + (y2 * yr)), dcolor
    

End Sub
Public Sub PlotInit(ByRef plot As PictureBox)
    plotcentre = plot.Height / 2
    xr = plot.Height / 9024000
    yr = plot.Height / 9024000
End Sub

Public Sub ClearPlot(ByRef plot As PictureBox)
    plot.Cls
    tmpx1 = 0
    tmpx2 = 0
    tmpy1 = 0
    tmpy2 = 0
    ' don't need to rub out position
    RubOut = False
End Sub

Private Sub ShowHelp1()

    Dim fontheight As Double
    
    Call ClearPlot(SkyPlot)
    
    fontheight = TextHeight("0") / 2
    SkyPlot.DrawMode = 13
    SkyPlot.DrawWidth = 1
    
    SkyPlot.Circle (200, 200), 30, vbBlue
    SkyPlot.ForeColor = vbBlue
    SkyPlot.CurrentX = 400
    SkyPlot.CurrentY = 200 - fontheight
    SkyPlot.Print oLangDll.GetLangString(1004)
    
    SkyPlot.Circle (200, 500), 30, vbRed
    SkyPlot.ForeColor = vbRed
    SkyPlot.CurrentX = 400
    SkyPlot.CurrentY = 500 - fontheight
    SkyPlot.Print oLangDll.GetLangString(1005)
    
    SkyPlot.Circle (200, 800), 60, vbWhite
    SkyPlot.ForeColor = vbWhite
    SkyPlot.CurrentX = 400
    SkyPlot.CurrentY = 800 - fontheight
    SkyPlot.Print oLangDll.GetLangString(1006)
    
    SkyPlot.Circle (200, 1100), 30, vbWhite
    SkyPlot.ForeColor = vbWhite
    SkyPlot.CurrentX = 400
    SkyPlot.CurrentY = 1100 - fontheight
    SkyPlot.Print oLangDll.GetLangString(1009)
    
    SkyPlot.Circle (200, 1400), 60, vbYellow
    SkyPlot.ForeColor = vbYellow
    SkyPlot.CurrentX = 400
    SkyPlot.CurrentY = 1400 - fontheight
    SkyPlot.Print oLangDll.GetLangString(1007)
    
    SkyPlot.Circle (200, 1700), 90, vbGreen
    SkyPlot.ForeColor = vbGreen
    SkyPlot.CurrentX = 400
    SkyPlot.CurrentY = 1700 - fontheight
    SkyPlot.Print oLangDll.GetLangString(1008)
    
    SkyPlot.Circle (200, 2000), 100, &H707070
    SkyPlot.Circle (200, 2000), 40, &H707070
    SkyPlot.ForeColor = &H707070
    SkyPlot.CurrentX = 400
    SkyPlot.CurrentY = 2000 - fontheight
    SkyPlot.Print oLangDll.GetLangString(1012)
    
    SkyPlot.ForeColor = vbWhite
    SkyPlot.CurrentX = 200
    SkyPlot.CurrentY = 2600 - fontheight
    SkyPlot.Print oLangDll.GetLangString(1010)
    SkyPlot.CurrentX = 200
    SkyPlot.CurrentY = 2900 - fontheight
    SkyPlot.Print oLangDll.GetLangString(1011)

End Sub
Private Sub ShowHelp2()
    Dim fontheight As Double
    
    Call ClearPlot(SkyPlot)
    
    fontheight = TextHeight("0") / 2
    SkyPlot.DrawMode = 13
    SkyPlot.DrawWidth = 1
    
    ' mount
    SkyPlot.Circle (200, 200), 30, &HFF8080
    SkyPlot.ForeColor = &HFF8080
    SkyPlot.CurrentX = 400
    SkyPlot.CurrentY = 200 - fontheight
    SkyPlot.Print oLangDll.GetLangString(1006)
    
    ' Mount transformed
    SkyPlot.Circle (200, 500), 30, vbWhite
    SkyPlot.ForeColor = vbWhite
    SkyPlot.CurrentX = 400
    SkyPlot.CurrentY = 500 - fontheight
    SkyPlot.Print oLangDll.GetLangString(1009)
    
    ' NCP
    SkyPlot.Circle (200, 800), 100, &H707070
    SkyPlot.Circle (200, 800), 40, &H707070
    SkyPlot.ForeColor = &H707070
    SkyPlot.CurrentX = 400
    SkyPlot.CurrentY = 800 - fontheight
    SkyPlot.Print oLangDll.GetLangString(1012)
    
    ' reference point
    SkyPlot.Line (150, 1300)-(250, 1300), &H707070
    SkyPlot.Line (200, 1250)-(200, 1370), &H707070
    SkyPlot.ForeColor = &H707070
    SkyPlot.CurrentX = 400
    SkyPlot.CurrentY = 1300 - fontheight
    SkyPlot.Print oLangDll.GetLangString(1018)
    
    ' Transformation (Nearest)
    SkyPlot.Line (150, 1600)-(250, 1600), vbGreen
    SkyPlot.Line (200, 1550)-(200, 1670), vbGreen
    SkyPlot.ForeColor = vbGreen
    SkyPlot.CurrentX = 400
    SkyPlot.CurrentY = 1600 - fontheight
    SkyPlot.Print oLangDll.GetLangString(1019)
    
    ' transformation (3-Point)
    SkyPlot.Line (150, 1900)-(250, 1900), vbRed
    SkyPlot.Line (200, 1850)-(200, 1970), vbRed
    SkyPlot.ForeColor = vbRed
    SkyPlot.CurrentX = 400
    SkyPlot.CurrentY = 1900 - fontheight
    SkyPlot.Print oLangDll.GetLangString(1020)
    
    ' Mouse click options
    SkyPlot.ForeColor = vbWhite
    SkyPlot.CurrentX = 200
    SkyPlot.CurrentY = 2600 - fontheight
    SkyPlot.Print oLangDll.GetLangString(1010)
    SkyPlot.CurrentX = 200
    SkyPlot.CurrentY = 2900 - fontheight
    SkyPlot.Print oLangDll.GetLangString(1011)

End Sub

Private Sub SkyPlot_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        DisplayMode = 1
        Call PlotTransform
        
    Else
        If Button = 1 Then
            DisplayMode = 0
            Call PlotAlignmentPoints
        End If
    End If
End Sub

Private Sub StarList_Keyup(KeyCode As Integer, Shift As Integer)
    Select Case DisplayMode
        Case 0
            PlotAlignmentPoints
    End Select
End Sub


Private Sub StarList_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Select Case Button
        Case 1
            Select Case DisplayMode
                Case 0
                    PlotAlignmentPoints
            End Select
        Case 2
            Me.PopupMenu popupmenu1
    End Select
End Sub

Private Sub Timer1_Timer()

    If Not timerflag Then
        timerflag = True

        If StarEditform.Visible Then
            If gSelectStar > 0 Then
                Nstarbl1.Caption = gSelectStar
            Else
                Nstarbl1.Caption = "---"
            End If
            If gAffine1 > 0 And gAlignmentStars_count >= 3 Then
                Label2.Caption = gAffine1
            Else
                Label2.Caption = "---"
            End If
            If gAffine2 > 0 And gAlignmentStars_count >= 3 Then
                Label3.Caption = gAffine2
            Else
                Label3.Caption = "---"
            End If
            If gAffine3 > 0 And gAlignmentStars_count >= 3 Then
                Label4.Caption = gAffine3
            Else
                Label4.Caption = "---"
            End If
            
            If lastCount <> gAlignmentStars_count Or RefreshDisplay = True Then
                ' alignment list has changed!
                RefreshDisplay = False
                lastCount = gAlignmentStars_count
                Call FillStarList
                Call PlotAlignmentPoints
                
            Else
                Select Case DisplayMode
                    Case 0
                    If lastNearest <> gSelectStar Or af1 <> gAffine1 Or af2 <> gAffine2 Or af3 <> gAffine3 Then
                        Call PlotAlignmentPoints
                        lastNearest = gSelectStar
                        af1 = gAffine1
                        af2 = gAffine2
                        af3 = gAffine3
                    End If
                End Select
                If DisplayMode < 3 Then
                    Call drawmount(gEmulRA, gEmulDEC)
                End If
            End If
        End If
    End If
    timerflag = False

End Sub
