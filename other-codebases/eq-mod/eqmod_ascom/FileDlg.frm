VERSION 5.00
Begin VB.Form FileDlg 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Load / Save PE File"
   ClientHeight    =   3840
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   6030
   Icon            =   "FileDlg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Show Hidden Folders"
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
      Left            =   3240
      TabIndex        =   6
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000080&
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2880
      Width           =   5775
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00000080&
      ForeColor       =   &H000080FF&
      Height          =   2235
      Left            =   3240
      Pattern         =   "*.txt*"
      TabIndex        =   4
      Top             =   120
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00000080&
      ForeColor       =   &H000080FF&
      Height          =   2115
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2895
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      ForeColor       =   &H000080FF&
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H0095C1CB&
      Caption         =   "Cancel"
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
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
End
Attribute VB_Name = "FileDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FileName As String
Public filename2 As String
Public lastdrive As String
Public lastdir As String
Public notfirst As Boolean
Public filter As String

Option Explicit

Private Sub CancelButton_Click()
    FileName = ""
    filename2 = ""
    Unload FileDlg
End Sub

Private Sub Check1_Click()
    Call HC.oPersist.WriteIniValue("FILE_HIDDEN_DIR", CStr(Check1.Value))
    Call ShowHiddenFolders
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    Text1.Text = Dir1.Path & "\"
    Call ShowHiddenFolders

End Sub

Private Sub Dir1_Click()
    Dir1.Path = Dir1.List(Dir1.ListIndex)
End Sub

Private Sub Drive1_Change()
    On Error Resume Next
    Dir1.Path = Drive1.Drive
    Text1.Text = Dir1.Path
End Sub

Private Sub File1_Click()
    Text1.Text = Dir1.Path & "\" & File1.FileName
    FileName = Text1.Text
End Sub

Private Sub File1_DblClick()
    Text1.Text = Dir1.Path & "\" & File1.FileName
    OKButton_Click
End Sub

Private Sub Form_Activate()
    FileName = ""
    If filter = "" Then filter = "*.*"
    File1.Pattern = filter
    Call PutWindowOnTop(FileDlg)
    Text1.SetFocus
    Text1.Text = Dir1.Path & "\"
    ShowHiddenFolders
End Sub

Private Sub Form_Load()
    Dim tmptxt As String
    On Error GoTo errhandle:
    If filter = "" Then filter = "*.*"
       
    FileDlg.Caption = oLangDll.GetLangString(400) & " (" & filter & ")"
    OKButton.Caption = oLangDll.GetLangString(401)
    CancelButton.Caption = oLangDll.GetLangString(402)

    If Not notfirst Then
        Dir1.Path = App.Path
        notfirst = True
    Else
        Dir1.Path = lastdir
        Drive1.Drive = lastdrive
    End If
    
    tmptxt = HC.oPersist.ReadIniValue("FILE_HIDDEN_DIR")
    If tmptxt = "" Then
        Call HC.oPersist.WriteIniValue("FILE_HIDDEN_DIR", "0")
        Check1.Value = 0
    Else
        If tmptxt = "1" Then
            Check1.Value = 1
        Else
            Check1.Value = 0
        End If
    End If

    
    
    Exit Sub
errhandle:
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    lastdrive = Drive1.Drive
    lastdir = Dir1.Path
End Sub

Private Sub OKButton_Click()
    FileName = Text1.Text
    filename2 = File1.FileName
    Unload FileDlg
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        OKButton_Click
    End If
End Sub



Private Sub ShowHiddenFolders()
Dim Folders(1000) As String
Dim FolderAttr(1000) As VbFileAttribute
Dim i As Integer
Dim F As String
Dim F1 As String
Dim k As Integer
  
    On Error Resume Next
    
    If Check1.Value = 1 Then
        i = -1
        F1 = Dir1.Path
        If Right(F1, 1) <> "\" Then F1 = F1 & "\"
        F = dir(F1, vbDirectory + vbHidden)
        
        ' Search out hidden directories
        While (F <> "")
            If F <> "." And F <> ".." Then
                i = i + 1
                Folders(i) = F
                FolderAttr(i) = GetAttr(F1 & F)
            End If
            F = dir()
        Wend
    
        'Change hidden directories attributs to normal
        k = i
        While (k >= 0)
            Call SetAttr(F1 & Folders(k), vbNormal)
        k = k - 1
        Wend
        
        'Show hidden folders in "dirlistbox" control
        Dir1.Refresh
        
        'Put back hidden attributes
        While (i >= 0)
           If FolderAttr(i) And vbHidden Then
               Call SetAttr(F1 & Folders(i), vbHidden)
           End If
    '       If FolderAttr(I) And vbSystem Then
    '           Call SetAttr(F1 & Folders(I), vbSystem)
    '       End If
           i = i - 1
        Wend
    Else
        Dir1.Refresh
    End If

End Sub



