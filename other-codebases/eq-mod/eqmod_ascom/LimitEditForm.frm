VERSION 5.00
Begin VB.Form LimitEditForm 
   BackColor       =   &H00000000&
   Caption         =   "Mount Limits Editor"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8700
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LimitEditForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   433
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   580
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Options"
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
      Height          =   1575
      Left            =   3360
      TabIndex        =   28
      Top             =   0
      Width           =   5175
      Begin VB.CheckBox chkAutoFLip 
         BackColor       =   &H00000000&
         Caption         =   "Auto Merdian Flip"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox chkSlew 
         BackColor       =   &H00000000&
         Caption         =   "Apply Limits to Gotos"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   600
         Width           =   1935
      End
      Begin VB.CheckBox chkPark 
         BackColor       =   &H00000000&
         Caption         =   "Park on Limit"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Horizon"
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
      Height          =   4095
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   8415
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   330
         Left            =   3240
         TabIndex        =   32
         Text            =   "Combo2"
         Top             =   382
         Width           =   2055
      End
      Begin VB.CommandButton CmdAddAuto 
         BackColor       =   &H0095C1CB&
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
         Left            =   2160
         Picture         =   "LimitEditForm.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Add Automtic"
         Top             =   360
         Width           =   375
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   2640
         Top             =   360
      End
      Begin VB.CommandButton CmdUserAdd 
         BackColor       =   &H0095C1CB&
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
         Left            =   3480
         Picture         =   "LimitEditForm.frx":124C
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Add Manual"
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Text            =   "0"
         Top             =   3600
         Width           =   495
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   600
         TabIndex        =   22
         Text            =   "0"
         Top             =   3600
         Width           =   495
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   1200
         TabIndex        =   21
         Text            =   "0"
         Top             =   3600
         Width           =   495
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   1680
         TabIndex        =   20
         Text            =   "0"
         Top             =   3600
         Width           =   495
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   330
         ItemData        =   "LimitEditForm.frx":17CE
         Left            =   2280
         List            =   "LimitEditForm.frx":17D8
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   3600
         Width           =   975
      End
      Begin VB.PictureBox PicX 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   3480
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   297
         TabIndex        =   18
         Top             =   3240
         Width           =   4455
      End
      Begin VB.PictureBox PicY 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H000080FF&
         Height          =   2415
         Left            =   7920
         ScaleHeight     =   161
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   17
         Top             =   840
         Width           =   375
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2415
         Left            =   3240
         Max             =   11
         Min             =   90
         TabIndex        =   16
         Top             =   840
         Value           =   11
         Width           =   255
      End
      Begin VB.PictureBox plot 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000040&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   3480
         ScaleHeight     =   159
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   295
         TabIndex        =   14
         Top             =   840
         Width           =   4455
         Begin VB.PictureBox Picture1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3255
            Left            =   6720
            ScaleHeight     =   3255
            ScaleWidth      =   15
            TabIndex        =   15
            Top             =   0
            Width           =   15
         End
      End
      Begin VB.CommandButton CmdSave 
         BackColor       =   &H0095C1CB&
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
         Left            =   600
         Picture         =   "LimitEditForm.frx":17EC
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Save"
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton CmdLoad 
         BackColor       =   &H0095C1CB&
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
         Left            =   120
         Picture         =   "LimitEditForm.frx":1D6E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Load"
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton CmdDel 
         BackColor       =   &H0095C1CB&
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
         Left            =   1680
         Picture         =   "LimitEditForm.frx":22F0
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Delete"
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton CmdClear 
         BackColor       =   &H0095C1CB&
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
         Left            =   1200
         Picture         =   "LimitEditForm.frx":2A16
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Defaults"
         Top             =   360
         Width           =   375
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00000040&
         ForeColor       =   &H000000FF&
         Height          =   2370
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   2895
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
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
         Left            =   2040
         TabIndex        =   25
         Top             =   2640
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Hour Angle"
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
         TabIndex        =   27
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
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
         Left            =   1200
         TabIndex        =   26
         Top             =   3360
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Meridian"
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
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.PictureBox PictureRA 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   600
         ScaleHeight     =   81
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   81
         TabIndex        =   38
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command19 
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
         Picture         =   "LimitEditForm.frx":2F98
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton Command20 
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
         Picture         =   "LimitEditForm.frx":351A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command38 
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
         Picture         =   "LimitEditForm.frx":3A9C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   660
         Width           =   375
      End
      Begin VB.Label Label36 
         BackColor       =   &H00000080&
         Caption         =   "05D5500"
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
         Left            =   1920
         TabIndex        =   7
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label37 
         BackColor       =   &H00000000&
         Caption         =   "West:"
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
         Height          =   195
         Left            =   1920
         TabIndex        =   6
         Top             =   840
         Width           =   915
      End
      Begin VB.Label Label35 
         BackColor       =   &H00000080&
         Caption         =   "0A2AB00"
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
         Left            =   1920
         TabIndex        =   5
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label38 
         BackColor       =   &H00000000&
         Caption         =   "East:"
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
         Height          =   195
         Left            =   1920
         TabIndex        =   4
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "--:--:--"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5520
      TabIndex        =   37
      Top             =   5940
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Time To Meridian"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   3600
      TabIndex        =   36
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "--:--:--"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2040
      TabIndex        =   35
      Top             =   5940
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Time To Horizon"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   6000
      Width           =   1815
   End
End
Attribute VB_Name = "LimitEditForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias _
        "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As _
        Long, ByVal wParam As Long, lParam As Any) As Long

Private Const LB_SETTABSTOPS = &H192
' The number of tabs
Const OBorder = 4
Private TAB0_Caption As String
Private TAB1_Caption As String

Dim Tabulator(1 To OBorder) As Long

Private Sub chkAutoFLip_Click()
    gAutoFlipEnabled = chkAutoFLip.Value
    Call WriteAutoFlipData
End Sub

Private Sub chkPark_Click()
    gLimitPark = ChkPark.Value
    Call HC.oPersist.WriteIniValue("LIMIT_PARK", CStr(gLimitPark))
End Sub

Private Sub chkSlew_Click()
    gLimitSlews = chkSlew.Value
    Call HC.oPersist.WriteIniValue("LIMIT_SLEWS", CStr(gLimitSlews))
End Sub

Private Sub CmdClear_Click()
    Call Limits_Clear
    RefreshList
End Sub


Private Sub CmdDel_Click()
Dim idx As Integer
    idx = List1.ListIndex - 1
    Call Limits_DeleteIdx(idx)
    RefreshList
End Sub

Private Sub CmdLoad_Click()
    Limits_Load
    RefreshList
End Sub

Private Sub CmdSave_Click()
    Limits_Save
End Sub

Private Sub CmdUserAdd_Click()
Dim lim As LIMIT
    On Error GoTo endsub
    If Combo1.ListIndex = 0 Then
        lim.ha = CDbl(Text1.Text) + CDbl(Text2.Text) / 60
        If lim.ha < 0 Or lim.ha >= 24 Then
            GoTo endsub
        End If
        lim.DEC = CDbl(Text3.Text) + CDbl(Text4.Text) / 60
'        If lim.dec < -90 Or lim.dec > 90 Then
        If lim.DEC < -360 Or lim.DEC > 360 Then
            GoTo endsub
        End If
        hadec_aa (gLatitude * DEG_RAD), (lim.ha * HRS_RAD), (lim.DEC * DEG_RAD), lim.Alt, lim.Az
        lim.Alt = lim.Alt * RAD_DEG           ' convert to degrees from Radians
        lim.Az = 360# - (lim.Az * RAD_DEG)    '  convert to degrees from Radians
    Else
        lim.Az = CDbl(Text3.Text) + CDbl(Text4.Text) / 60
        If lim.Az < 0 Or lim.Az >= 360 Then
            GoTo endsub
        End If
        lim.Alt = CDbl(Text1.Text) + CDbl(Text2.Text) / 60
        If lim.Alt < -90 Or lim.Alt > 90 Then
            GoTo endsub
        End If
        aa_hadec gLatitude * DEG_RAD, lim.Alt * DEG_RAD, lim.Az * DEG_RAD, lim.ha, lim.DEC
        lim.ha = Range24(lim.ha * RAD_HRS)
        lim.DEC = lim.DEC * RAD_DEG
    End If
    Call Limits_Add(lim)
    RefreshList
    
endsub:
End Sub

Private Sub Combo1_Click()
    If Combo1.ListIndex = 0 Then
        Label1.Caption = oLangDll.GetLangString(2112)
        Label2.Caption = oLangDll.GetLangString(2113)
    Else
        Label1.Caption = oLangDll.GetLangString(2110)
        Label2.Caption = oLangDll.GetLangString(2111)
    End If
    ' loose focus
    On Error Resume Next
    Check1.SetFocus

End Sub


Private Sub Combo2_Click()
    ' get selection
    gHorizonAlgorithm = Combo2.ListIndex
    ' apply algorithm
    Call Limits_BuildLimitDef
    
    ' write to ini file
    Call HC.oPersist.WriteIniValue("LIMIT_HORIZON_ALGORITHM", CStr(gHorizonAlgorithm))
    
    ' loose focus
    On Error Resume Next
    Check1.SetFocus
    
    ' redraw using selected algorithm
    PlotLimits
    
End Sub

Private Sub CmdAddAuto_Click()
Dim lim As LIMIT
    lim.Alt = gAlt
    lim.Az = gAz
    lim.ha = gRA_Hours
    lim.DEC = gDec
    Call Limits_Add(lim)
    RefreshList
End Sub

Private Function FmtSexa2(ByVal n As Double, ShowPlus As Boolean) As String
    Dim sg As String
    Dim us As String
    Dim ms As String
    Dim ss As String
    Dim u As Long
    Dim m As Long
    Dim fmt
    

    sg = "+"                                ' Assume positive
    If n < 0 Then                           ' Check neg.
        n = -n                              ' Make pos.
        sg = "-"                            ' Remember sign
    End If

    m = Fix(n)                              ' Units (deg or hr)
    us = Format$(m, "00")

    n = (n - m) * 60#
    m = n                              ' Minutes
    ms = Format$(m, "00")

    FmtSexa2 = us & ":" & ms
    If ShowPlus Or (sg = "-") Then FmtSexa2 = sg & FmtSexa2
    
End Function
Private Sub RefreshList()
Dim i As Integer
Dim size As Integer
Dim tmpstr As String

    List1.Clear
    List1.AddItem (oLangDll.GetLangString(2111) & vbTab & oLangDll.GetLangString(2110) & vbTab & oLangDll.GetLangString(2112) & vbTab & oLangDll.GetLangString(2113))
    
    On Error GoTo endsub
    
    size = UBound(LimitArray)

    For i = 0 To size - 1
        tmpstr = CStr(CInt(LimitArray(i).Az)) & vbTab & FmtSexa2(LimitArray(i).Alt, True) & vbTab & FmtSexa2(LimitArray(i).ha, True) & vbTab & FmtSexa2(LimitArray(i).DEC, True)
        List1.AddItem (tmpstr)
    Next i
    
endsub:
    PlotLimits
End Sub

Private Sub PlotLimits()
Dim xoffset, yoffset, midy As Double
Dim xscale, yscale As Double
Dim x, Y, lastx, lasty, firstx, firsty As Double
Dim i As Integer
Dim str As String

    On Error Resume Next
    
    plot.Cls
    PicX.Cls
    PicY.Cls
    plot.DrawWidth = 1
    xscale = (plot.ScaleWidth) / 359
    yscale = 0.5 * plot.ScaleHeight / VScroll1.Value
    
    midy = plot.ScaleHeight / 2
    
    ' plot gridlines
    For i = 20 To 340 Step 20
        x = i * xscale
        plot.Line (x, 0)-(x, plot.ScaleHeight), &H80&
        If (i Mod 60) = 0 Then
            str = CStr(i)
            PicX.CurrentX = x - TextWidth(str) / 2
            PicX.CurrentY = 0
            PicX.Print str
        End If
    Next i
    For i = -90 To 90 Step 5
        Y = midy - i * yscale
        plot.Line (0, Y)-(plot.ScaleWidth, Y), &H80&
        If (i Mod 10) = 0 Then
            str = CStr(i)
            PicY.CurrentY = Y - TextHeight("0") / 2
            PicY.CurrentX = 0
            PicY.Print str
        End If
    Next i
    
    ' plot axis
    plot.Line (0, midy)-(plot.ScaleWidth, midy), vbRed
    
    ' plot graph
    plot.DrawWidth = 2
    
    lastx = 0
    lasty = midy - LimitArray2(0).Alt * yscale
    
    For x = 1 To plot.ScaleWidth
        Y = LimitArray2(x / xscale).Alt
        Y = midy - Y * yscale
        plot.Line (lastx, lasty)-(x, Y), &HC000C0
        lastx = x
        lasty = Y
    Next x
    
    
    x = gAz * xscale
    Y = midy - gAlt * yscale
   
    plot.DrawWidth = 1
    plot.Line (x - 4, Y)-(x + 4, Y), &HFF00FF
    plot.Line (x, Y - 4)-(x, Y + 4), &HFF00FF
    

End Sub


Public Sub UpdateDisplay()
Dim t, x, Y As Double
Dim i As Double
Dim lim1 As Double
Dim lim2 As Double
    
    t = Limits_TimeToHorizon()
    If t = -1 Then
        Label3.Caption = "--:--:--"
    Else
        Label3.Caption = FmtSexa(t, False)
    End If
    
    t = Limits_TimeToMeridian()
    If t = -1 Then
        Label6.Caption = "--:--:--"
    Else
        Label6.Caption = FmtSexa(t, False)
    End If
    
    PlotLimits

    i = (Range24(gRA_Hours - 6) / 24) * 100
    
    If gRA_Limit_East <> 0 And gRA_Limit_West <> 0 Then
        lim1 = 100 * (gRAEncoder_Zero_pos - gRA_Limit_East) / (gTot_RA)
        If lim1 < 0 Then lim1 = 100 + lim1
        lim2 = 100 * (gRAEncoder_Zero_pos - gRA_Limit_West) / (gTot_RA)
        If lim2 < 0 Then lim2 = 100 + lim2
    Else
        lim1 = -1
        lim2 = -1
    End If
    
    If gHemisphere <> 0 Then
        i = 100 - i
    End If
    Call DrawAxis(PictureRA, 0, i, lim1, lim2)



End Sub


Private Sub Form_Activate()
    UpdateDisplay
End Sub

Private Sub Form_Load()

    If HC.HCOnTop.Value = 1 Then Call PutWindowOnTop(LimitEditForm)

    LimitEditForm.Caption = oLangDll.GetLangString(2107)
    Command19.ToolTipText = oLangDll.GetLangString(157)
    Command20.ToolTipText = oLangDll.GetLangString(158)
    Command38.ToolTipText = oLangDll.GetLangString(183)
    CmdDel.ToolTipText = oLangDll.GetLangString(183)
    CmdLoad.ToolTipText = oLangDll.GetLangString(184)
    CmdSave.ToolTipText = oLangDll.GetLangString(185)
    CmdAddAuto.ToolTipText = oLangDll.GetLangString(2115)
    CmdUserAdd.ToolTipText = oLangDll.GetLangString(2116)
    CmdClear.ToolTipText = oLangDll.GetLangString(1129)
    Label37.Caption = oLangDll.GetLangString(159)
    Label38.Caption = oLangDll.GetLangString(160)
    Label4.Caption = oLangDll.GetLangString(2124)
    Label5.Caption = oLangDll.GetLangString(2125)
    Frame1.Caption = oLangDll.GetLangString(2123)
    Frame3.Caption = oLangDll.GetLangString(2108)
    Frame4.Caption = oLangDll.GetLangString(2109)
    ChkPark.Caption = oLangDll.GetLangString(2121)
    chkSlew.Caption = oLangDll.GetLangString(2122)
    chkAutoFLip.Caption = oLangDll.GetLangString(2204)
    Combo1.Clear
    Combo1.AddItem (oLangDll.GetLangString(2117))
    Combo1.AddItem (oLangDll.GetLangString(2118))
    Combo2.Clear
    Combo2.AddItem (oLangDll.GetLangString(2119))
    Combo2.AddItem (oLangDll.GetLangString(2120))

    If gHemisphere = 0 Then
        Label35.Caption = printhex(gRA_Limit_West)
        Label36.Caption = printhex(gRA_Limit_East)
    Else
        Label36.Caption = printhex(gRA_Limit_West)
        Label35.Caption = printhex(gRA_Limit_East)
    End If
    
    Tabulator(1) = 30
    Tabulator(2) = 60
    Tabulator(3) = 90
    Tabulator(4) = 120
    SendMessage List1.hwnd, LB_SETTABSTOPS, OBorder, Tabulator(1)
    
    Combo1.ListIndex = 0
    Combo1.Text = Combo1.List(0)
    Combo2.ListIndex = 0
    Combo2.Text = Combo2.List(0)
    On Error Resume Next
    Check1.SetFocus
    RefreshList
    
    If gLimitPark Then ChkPark.Value = 1 Else ChkPark.Value = 0
    If gLimitSlews Then chkSlew.Value = 1 Else chkSlew.Value = 0
    If gAutoFlipEnabled Then chkAutoFLip.Value = 1 Else chkAutoFLip.Value = 0
    If gAutoFlipAllowed Then
        chkAutoFLip.Enabled = True
    Else
        ' show state but don't allow it to be changed.
        chkAutoFLip.Enabled = False
    End If
    If gAutoFlipEnabled = True Then chkAutoFLip.Value = 1 Else chkAutoFLip.Value = 0
    
    Timer1.Enabled = True
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
    UpdateDisplay
End Sub

Private Sub VScroll1_Change()
    PlotLimits
End Sub

Private Sub VScroll1_Scroll()
    VScroll1_Change
End Sub



Private Sub Command19_Click()
    'Routine to add slew limits

    If (GetEmulRA() < gRAEncoder_Zero_pos) Then
        gRA_Limit_East = GetEmulRA()
    Else
        gRA_Limit_West = GetEmulRA()
    End If
    
    If gHemisphere = 0 Then
        Label35.Caption = printhex(gRA_Limit_West)
        Label36.Caption = printhex(gRA_Limit_East)
    Else
        Label36.Caption = printhex(gRA_Limit_West)
        Label35.Caption = printhex(gRA_Limit_East)
    End If
    
    Call writeRAlimit
    Call UpdateDisplay

End Sub

Private Sub Command38_Click()

    gRA_Limit_East = 0
    gRA_Limit_West = 0

    Label35.Caption = printhex(gRA_Limit_West)
    Label36.Caption = printhex(gRA_Limit_East)
    
    Call writeRAlimit
    Call UpdateDisplay
    
 End Sub
 
 Private Sub Command20_Click()
    
    Call SetRaLimitDefaults
    
    If gHemisphere = 0 Then
        Label35.Caption = printhex(gRA_Limit_West)
        Label36.Caption = printhex(gRA_Limit_East)
    Else
        Label36.Caption = printhex(gRA_Limit_West)
        Label35.Caption = printhex(gRA_Limit_East)
    End If
    
    ' store limits
    Call writeRAlimit
    Call UpdateDisplay
End Sub
