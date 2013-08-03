VERSION 5.00
Begin VB.Form CustomMountDlg 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Customise"
   ClientHeight    =   3000
   ClientLeft      =   2760
   ClientTop       =   3705
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox CheckCustom 
      BackColor       =   &H00000000&
      Caption         =   "Custom Mount Enabled"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   4935
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "DEC"
      ForeColor       =   &H000080FF&
      Height          =   1935
      Left            =   2640
      TabIndex        =   2
      Top             =   480
      Width           =   2415
      Begin VB.TextBox TextDecOffset 
         BackColor       =   &H00000080&
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox TextDecWormSteps 
         BackColor       =   &H00000080&
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox TextDecTotalSteps 
         BackColor       =   &H00000080&
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         ItemData        =   "CustomMountDlg.frx":0000
         Left            =   120
         List            =   "CustomMountDlg.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "Tracking Offset"
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "Worm Steps"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "Total Steps"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "RA"
      ForeColor       =   &H000080FF&
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2415
      Begin VB.TextBox TextRAOffset 
         BackColor       =   &H00000080&
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox TextRAWormSteps 
         BackColor       =   &H00000080&
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox TextRATotalSteps 
         BackColor       =   &H00000080&
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00000080&
         ForeColor       =   &H000080FF&
         Height          =   315
         ItemData        =   "CustomMountDlg.frx":002D
         Left            =   120
         List            =   "CustomMountDlg.frx":003A
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Tracking Offset"
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Worm Steps"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Total Steps"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   4935
   End
End
Attribute VB_Name = "CustomMountDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Dim RADefns() As MountDefn
Dim DECDefns() As MountDefn

Private Sub Combo1_Click()
    With RADefns(Combo1.ListIndex)
        TextRATotalSteps.Text = .TotalSteps
        TextRAWormSteps.Text = .wormsteps
        TextRAOffset.Text = .offset
    End With
End Sub

Private Sub Combo2_Click()
    With RADefns(Combo2.ListIndex)
        TextDecTotalSteps.Text = .TotalSteps
        TextDecWormSteps.Text = .wormsteps
        TextDecOffset.Text = .offset
    End With
End Sub

Private Sub Form_Load()
    Call SetText
    Call LoadDefinitions
    Call readCustomMount
    TextRATotalSteps.Text = CStr(gCustomRA360)
    TextDecTotalSteps.Text = CStr(gCustomDEC360)
    TextRAWormSteps.Text = CStr(gCustomRAWormSteps)
    TextDecWormSteps.Text = CStr(gCustomDECWormSteps)
    TextRAOffset.Text = CStr(gCustomTrackingOffsetRA)
    TextDecOffset.Text = CStr(gCustomTrackingOffsetDEC)
    CheckCustom.Value = gCustomMount
    
    Call PutWindowOnTop(CustomMountDlg)
    
    
End Sub

Private Sub OKButton_Click()
Dim RaTotalSteps As Long
Dim DecTotalSteps As Long
Dim RaWormSteps As Long
Dim DecWormSteps As Long
Dim RAOffset As Long
Dim DecOffset As Long

    On Error GoTo endok
    RaTotalSteps = val(TextRATotalSteps.Text)
    DecTotalSteps = val(TextDecTotalSteps.Text)
    RaWormSteps = val(TextRAWormSteps.Text)
    DecWormSteps = val(TextDecWormSteps.Text)
    RAOffset = val(TextRAOffset.Text)
    DecOffset = val(TextDecOffset.Text)

    ' some limited error checking
    If RaTotalSteps <= 0 Then GoTo endok
    If DecTotalSteps <= 0 Then GoTo endok
    If DecWormSteps <= 0 Then GoTo endok
    If RaWormSteps <= 0 Then GoTo endok
    If DecWormSteps > DecTotalSteps Then GoTo endok
    If RaWormSteps > RaTotalSteps Then GoTo endok

    gCustomRA360 = RaTotalSteps
    gCustomDEC360 = DecTotalSteps
    gCustomRAWormSteps = RaWormSteps
    gCustomDECWormSteps = DecWormSteps
    gCustomTrackingOffsetRA = RAOffset
    gCustomTrackingOffsetDEC = DecOffset

    gCustomMount = CheckCustom.Value

    Call writeCustomMount
    
    Unload Me
    Exit Sub
    
endok:

End Sub

Public Sub LoadDefinitions()
    Dim temp1 As String
    Dim data() As String
    Dim RACount As Integer
    Dim DECCount As Integer
    
    On Error GoTo Loaderror
    
    RACount = 0
    DECCount = 0
    
    Combo1.Clear
    Combo2.Clear
    
    ReDim RADefns(0)
    ReDim DECDefns(0)
    
    temp1 = App.Path & "\eqmod_custom.txt"
    
    Close #1
    Open temp1 For Input As #1
    
    While Not EOF(1)
        Line Input #1, temp1
        If Left$(temp1, 1) <> "#" Then
            ' parse parameters
            data = Split(temp1, ";")
            If UBound(data) = 4 Then
                If data(0) = "R" Then
                    Combo1.AddItem (data(1))
                    With RADefns(RACount)
                        .TotalSteps = data(2)
                        .wormsteps = data(3)
                        .offset = data(4)
                    End With
                    RACount = RACount + 1
                    ReDim Preserve RADefns(RACount)
                End If
                If data(0) = "D" Then
                    Combo2.AddItem (data(1))
                    With RADefns(DECCount)
                        .TotalSteps = data(2)
                        .wormsteps = data(3)
                        .offset = data(4)
                    End With
                    DECCount = DECCount + 1
                    ReDim Preserve DECDefns(DECCount)
                End If
            End If
        End If
    Wend
    
Loaderror:
    Close #1
End Sub

Private Sub SetText()

    CustomMountDlg.Caption = oLangDll.GetLangString(2900)
    CheckCustom.Caption = oLangDll.GetLangString(2901)
    Label1.Caption = oLangDll.GetLangString(2902)
    Label4.Caption = oLangDll.GetLangString(2902)
    Label2.Caption = oLangDll.GetLangString(2903)
    Label5.Caption = oLangDll.GetLangString(2903)
    Label3.Caption = oLangDll.GetLangString(2904)
    Label6.Caption = oLangDll.GetLangString(2904)

End Sub
