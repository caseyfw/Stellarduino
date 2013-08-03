VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "EQASCOM_Run"
   ClientHeight    =   30
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   2820
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   30
   ScaleWidth      =   2820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   0
      Top             =   0
   End
   Begin VB.Menu mPopupsys 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu setup 
         Caption         =   "Setup EQASCOM"
         Visible         =   0   'False
      End
      Begin VB.Menu setup2 
         Caption         =   "Setup EQASCOM2"
         Visible         =   0   'False
      End
      Begin VB.Menu sep 
         Caption         =   "--------"
      End
      Begin VB.Menu mPopExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private m_scope As Object
Private m_scope2 As Object
Private m_Util As DriverHelper.Util
Private ascomID
Private timerflag As Boolean


Private Sub Form_Load()
    Me.Show
    Me.Refresh
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "EQASCOM_RUN" & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nid


    timeflag = False
    Timer1.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim kill As Boolean
    
    On Error Resume Next
    
    Timer1.Enabled = False
    If UnloadMode = 0 Then
        If MsgBox("Stop EQASCOM?", vbCritical Or vbYesNo) = vbYes Then
            Cancel = False
        Else
            Cancel = True
        End If
    End If

    If Cancel = False Then
        ' kill eqascom irrespective of other clients
        m_scope.StopClientCount
        m_scope2.StopClientCount
    Else
        Timer1.Enabled = True
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Shell_NotifyIcon NIM_DELETE, nid
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'this procedure receives the callbacks from the System Tray icon.
Dim Result As Long
Dim msg As Long
    'the value of X will vary depending upon the scalemode setting
    If Me.ScaleMode = vbPixels Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If
    Select Case msg
        Case WM_LBUTTONUP        '514 restore form window
'           Me.WindowState = vbNormal
'           Result = SetForegroundWindow(Me.hwnd)
'           Me.Show
        Case WM_LBUTTONDBLCLK    '515 restore form window
'            Me.WindowState = vbNormal
'            Result = SetForegroundWindow(Me.hwnd)
'            Me.Show
        Case WM_RBUTTONUP        '517 display popup menu
            Result = SetForegroundWindow(Me.hwnd)
            Me.PopupMenu Me.mPopupsys
    End Select
End Sub

Private Sub Form_Resize()
    'this is necessary to assure that the minimized window is hidden
    If Me.WindowState = vbMinimized Then Me.Hide
End Sub

Private Sub Hide_Click()
    Me.Hide
End Sub

Private Sub mPopExit_Click()
    'called when user clicks the popup menu Exit command
    If MsgBox("Stop EQASCOM?", vbCritical Or vbYesNo) = vbYes Then
        Unload Me
    End If
End Sub

Private Sub mPopRestore_Click()
    'called when the user clicks the popup menu Restore command
    Dim Result As Long
    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hwnd)
    Me.Show
End Sub

Private Sub Setup_Click()
    On Error Resume Next
    If Not m_scope Is Nothing Then
        m_scope.SetupDialog
    End If
End Sub

Private Sub Setup2_Click()
    On Error Resume Next
    If Not m_scope2 Is Nothing Then
        m_scope2.SetupDialog
    End If
End Sub


Private Sub Timer1_Timer()
 On Error GoTo endtimer
    If Not timerflag Then
        timerflag = True
        If m_scope Is Nothing Then
            ' get an ascom scope object
            Set m_scope = CreateObject("EQMOD.Telescope")
            If Not m_scope Is Nothing Then
                setup.Visible = True
                ' if sucessfull then connect
                m_scope.Connected = True
                If Not m_scope.Connected Then
                    Set m_scope = Nothing
                End If
            End If
        Else
            setup.Visible = True
            'reconnect if disconnected
            If m_scope.Connected = False Then
                m_scope.Connected = True
            End If
        End If
        
        If m_scope2 Is Nothing Then
            ' get an ascom scope object
            Set m_scope2 = CreateObject("EQMOD_2.Telescope")
            If Not m_scope2 Is Nothing Then
                setup2.Visible = True
                ' if sucessfull then connect
                m_scope2.Connected = True
                If Not m_scope2.Connected Then
                    Set m_scope2 = Nothing
                End If
            End If
        Else
            setup2.Visible = True
            'reconnect if disconnected
            If m_scope2.Connected = False Then
                m_scope2.Connected = True
            End If
        End If

    End If
endtimer:
    timerflag = False
    
End Sub
