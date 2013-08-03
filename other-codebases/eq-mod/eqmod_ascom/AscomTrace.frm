VERSION 5.00
Begin VB.Form AscomTrace 
   BackColor       =   &H00000000&
   Caption         =   "ASCOM Trace"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10110
   Icon            =   "AscomTrace.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check9 
      BackColor       =   &H00000000&
      Caption         =   "Log To File"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5880
      TabIndex        =   12
      Top             =   240
      Width           =   1335
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00000000&
      Caption         =   "Capabilities (Cans)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Syncs, Park, Unpark, Set Park"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0095C1CB&
      Caption         =   "Pause"
      Height          =   375
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Filters"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   5535
      Begin VB.CheckBox Check8 
         BackColor       =   &H00000000&
         Caption         =   "Other"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3600
         TabIndex        =   11
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00000000&
         Caption         =   "Slews, Moves"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3600
         TabIndex        =   10
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00000000&
         Caption         =   "Tracking"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3600
         TabIndex        =   9
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00000000&
         Caption         =   "Guiding"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3600
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00000000&
         Caption         =   "Get Position RA/DEC,ALT/AZ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   3135
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00000000&
         Caption         =   "Connection"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0095C1CB&
      Caption         =   "Clear"
      Height          =   375
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1920
      Width           =   9855
   End
End
Attribute VB_Name = "AscomTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public AscomTraceEnabled As Boolean
Private debugfile As String
Private logcount As Integer
Private logfileindex As Integer



Private Sub Check9_Click()
    On Error Resume Next
    If Check9.Value Then
        debugfile = HC.oPersist.GetIniPath() + "\ascom_debug1.txt"
        On Error Resume Next
        Close #1
        Open debugfile For Output As #1
        logfileindex = 1
        logcount = 0
    Else
        Close #1
        debugfile = ""
    End If

End Sub

Private Sub Command1_Click()
    Text1.Text = ""
End Sub

Public Sub Add_log(logclass As Integer, dtaLog As String)
Dim showlog As Boolean

    If AscomTraceEnabled = False Then Exit Sub

    On Error Resume Next
   showlog = False
   ' apply filters
   Select Case logclass
        Case 0:
            If Check8.Value Then showlog = True
        Case 1:
            If Check2.Value Then showlog = True
        Case 2:
            If Check3.Value Then showlog = True
        Case 3:
            If Check4.Value Then showlog = True
        Case 4:
            If Check5.Value Then showlog = True
        Case 5:
            If Check6.Value Then showlog = True
        Case 6:
            If Check7.Value Then showlog = True
        Case 7:
            If Check1.Value Then showlog = True
            
   End Select
    
    If showlog Then
        'Add the message, limit to 2000 characters
        Text1.Text = Right(Text1.Text & vbCrLf & time$ & " " & dtaLog, 2000)
        Text1.SelStart = Len(Text1.Text)
        If Check9.Value Then
            Print #1, time$ & " " & dtaLog
            logcount = logcount + 1
            If logcount > 2000 Then
                Close #1
                If logfileindex = 1 Then
                    Open HC.oPersist.GetIniPath() + "\ascom_debug2.txt" For Output As #1
                    logfileindex = 2
                Else
                    Open HC.oPersist.GetIniPath() + "\ascom_debug1.txt" For Output As #1
                    logfileindex = 1
                End If
                logcount = 0
            End If
        End If
        
    End If
End Sub

Private Sub Command2_Click()
    If Command2.Caption = "Pause" Then
        AscomTraceEnabled = False
        Command2.Caption = "Go"
    Else
        AscomTraceEnabled = True
        Command2.Caption = "Pause"
    End If

End Sub

Private Sub Form_Load()
    AscomTraceEnabled = True
    Check1.Value = 1
    Check2.Value = 1
    Check3.Value = 1
    Check4.Value = 1
    Check5.Value = 1
    Check6.Value = 1
    Check7.Value = 1
    Check8.Value = 1
    Command2.Caption = "Pause"
End Sub

Private Sub Form_Resize()
    Text1.width = AscomTrace.width - 350
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    AscomTraceEnabled = False
    Close #1
End Sub
