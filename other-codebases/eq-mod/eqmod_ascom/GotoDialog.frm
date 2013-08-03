VERSION 5.00
Begin VB.Form GotoDialog 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Goto"
   ClientHeight    =   3360
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   3345
   Icon            =   "GotoDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1020
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   5
      Left            =   2400
      MousePointer    =   3  'I-Beam
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   4
      Left            =   1680
      MousePointer    =   3  'I-Beam
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   3
      Left            =   960
      MousePointer    =   3  'I-Beam
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   2
      Left            =   2400
      MousePointer    =   3  'I-Beam
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   1
      Left            =   1680
      MousePointer    =   3  'I-Beam
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   960
      MousePointer    =   3  'I-Beam
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   120
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "J2000"
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H0095C1CB&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H0095C1CB&
      Caption         =   "GOTO"
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton CmdStore 
      BackColor       =   &H0095C1CB&
      Height          =   375
      Left            =   120
      Picture         =   "GotoDialog.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton ClearCommand 
      BackColor       =   &H0095C1CB&
      Height          =   375
      Left            =   2880
      Picture         =   "GotoDialog.frx":124C
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2280
      TabIndex        =   14
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "DEC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "RA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "GotoDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Type GOTO_BOOKMARK
    RA As Double
    DEC As Double
End Type

Private bookmarks() As GOTO_BOOKMARK
Private StoreRA As Double
Private StoreDEC As Double

Private Sub CancelButton_Click()
    Call Unload(Me)
End Sub

Private Sub ClearCommand_Click()
    ReDim bookmarks(0)
    List1.Clear
End Sub



Private Sub CmdStore_Click()
Dim RA As Double
Dim DEC As Double

    If Not GetCoords(RA, DEC) Then Exit Sub
    StoreRA = RA
    StoreDEC = DEC
    
    List1.AddItem CStr(List1.ListCount + 1) & ")  " & FmtSexa(RA, False) & ", " & FmtSexa(DEC, False)
    ReDim Preserve bookmarks(List1.ListCount)
    bookmarks(List1.ListCount).RA = RA
    bookmarks(List1.ListCount).DEC = DEC

End Sub

Private Sub Form_Load()
Dim i As Integer

    Call SetText

    Text1(0).Text = Left$(HC.ralbl.Caption, 2)
    Text1(1).Text = mid$(HC.ralbl.Caption, 4, 2)
    Text1(2).Text = Right$(HC.ralbl.Caption, 2)
    Text1(3).Text = Left$(HC.declbl.Caption, 3)
    Text1(4).Text = mid$(HC.declbl.Caption, 5, 2)
    Text1(5).Text = Right$(HC.declbl.Caption, 2)
    
    On Error GoTo endload
    For i = 1 To UBound(bookmarks)
        List1.AddItem CStr(List1.ListCount + 1) & ")  " & FmtSexa(bookmarks(i).RA, False) & ", " & FmtSexa(bookmarks(i).DEC, False)
    Next i
endload:
    Call PutWindowOnTop(GotoDialog)

End Sub

Private Function GetCoords(ByRef RA As Double, ByRef DEC As Double) As Boolean
    Dim rh As Integer
    Dim rm As Integer
    Dim rs As Double
    Dim dd As Integer
    Dim dm As Integer
    Dim ds As Double
    
    On Error GoTo errhandler
    
    rh = CInt(Text1(0).Text)
    rm = CInt(Text1(1).Text)
    rs = CDbl(Text1(2).Text)
    dd = CInt(Text1(3).Text)
    dm = CInt(Text1(4).Text)
    ds = CDbl(Text1(5).Text)

    If rh < 0 Or rh > 23 Then GoTo errhandler
    If dd < -90 Or dd > 90 Then GoTo errhandler
    
    If rm < 0 Or rm > 59 Then GoTo errhandler
    If dm < 0 Or dm > 59 Then GoTo errhandler

    If rs < 0 Or rs >= 60 Then GoTo errhandler
    If ds < 0 Or ds >= 60 Then GoTo errhandler

    RA = CDbl(rh) + CDbl(rm) / 60 + rs / 3600
    DEC = CDbl(dd) + CDbl(dm) / 60 + ds / 3600
    GetCoords = True
    Exit Function
    
errhandler:
    ' bad data!
    GetCoords = False

End Function


Private Sub List1_Click()
Dim strRa As String
Dim StrDec As String

    If List1.ListIndex <> -1 Then
        strRa = FmtSexa(bookmarks(List1.ListIndex + 1).RA, False)
        StrDec = FmtSexa(bookmarks(List1.ListIndex + 1).DEC, True)
        
        Text1(0).Text = Left$(strRa, 2)
        Text1(1).Text = mid$(strRa, 4, 2)
        Text1(2).Text = Right$(strRa, 2)
        Text1(3).Text = Left$(StrDec, 3)
        Text1(4).Text = mid$(StrDec, 5, 2)
        Text1(5).Text = Right$(StrDec, 2)
    End If

End Sub

Private Sub List1_DblClick()
    List1_Click
    OKButton_Click
End Sub

Private Sub OKButton_Click()
    Dim rh As Integer
    Dim rm As Integer
    Dim rs As Double
    Dim dd As Integer
    Dim dm As Integer
    Dim ds As Double
    
    Dim RA As Double
    Dim DEC As Double
    Dim epochnow As Double
    
    On Error GoTo errhandler
    
    If Not GetCoords(RA, DEC) Then GoTo errhandler

    If Check1(0).Value = 1 Then
        ' coordinates are J2000 so precess
        epochnow = 2000 + (now_mjd() - J2000) / 365.25
        Call Precess(RA, DEC, 2000, epochnow)
    End If

    If gEQparkstatus = 0 Then
        gTargetRA = RA
        gTargetDec = DEC
        HC.Add_Message ("Goto: " & oLangDll.GetLangString(105) & "[ " & FmtSexa(gTargetRA, False) & " ] " & oLangDll.GetLangString(106) & "[ " & FmtSexa(gTargetDec, True) & " ]")
        gSlewCount = gMaxSlewCount   'NUM_SLEW_RETRIES               'Set initial iterative slew count
        Call radecAsyncSlew(gGotoRate)
        EQ_Beep (20)
    Else
        HC.Add_Message (oLangDll.GetLangString(5000))
    End If

    Call Unload(GotoDialog)
    Exit Sub
    
errhandler:
    ' bad data!

End Sub


Private Sub SetText()
    Label1.Caption = oLangDll.GetLangString(105)
    Label2.Caption = oLangDll.GetLangString(106)
    CancelButton.Caption = oLangDll.GetLangString(1131)
    OKButton.Caption = oLangDll.GetLangString(2712)
    CmdStore.ToolTipText = oLangDll.GetLangString(2710)
    ClearCommand.ToolTipText = oLangDll.GetLangString(1003)
End Sub

