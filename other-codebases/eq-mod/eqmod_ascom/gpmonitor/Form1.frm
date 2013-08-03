VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Gamepad Monitor"
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3165
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   3165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00000080&
      ForeColor       =   &H000080FF&
      Height          =   315
      ItemData        =   "Form1.frx":0CCA
      Left            =   0
      List            =   "Form1.frx":0CCC
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2640
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   1095
      Left            =   0
      Picture         =   "Form1.frx":0CCE
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   205
      TabIndex        =   1
      Top             =   0
      Width           =   3135
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3600
      Top             =   120
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const HWND_TOPMOST = -1
Const HWND_NOTTOPMOST = -2
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Private Declare Function SetWindowPos Lib "user32" _
         (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
          ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long



Private Sub Combo1_Click()
    If Combo1.ListIndex >= 0 Then
        gpID = Combo1.ListIndex
    End If

End Sub

Private Sub Form_Load()
Dim i As Integer
Dim j As Integer

    On Error Resume Next

    JoystickDat.dwSize = Len(JoystickDat)
    JoystickDat.dwFlags = JOY_RETURNALL

    j = 0
    While i = JOYERR_NOERROR
        i = joyGetDevCaps(j, joyinfo, Len(joyinfo))
        If i = JOYERR_NOERROR Then
            Combo1.AddItem (CStr(j) & ": " & joyinfo.szPname)
        End If
        j = j + 1
    Wend
    
    gpID = JOYSTICKID1
    
    Combo1.ListIndex = 0

    lngWindowPosition = SetWindowPos(Form1.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Form1.Height = 3360

    Timer1.Enabled = True
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
       Form1.Height = 3360
    Else
        If Button = 2 Then
            Form1.Height = 1425
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    Dim x As Double
    Dim y As Double

    List1.Clear
        
    i = joyGetPosEx(gpID, JoystickDat)
    If i = JOYERR_NOERROR Then
            
        List1.AddItem ("Buttons=" & CStr(JoystickDat.dwButtons))
        List1.AddItem ("Number=" & CStr(JoystickDat.dwButtonNumber))
        List1.AddItem ("Flags=" & CStr(JoystickDat.dwFlags))
        List1.AddItem ("POV=" & CStr(JoystickDat.dwPOV))
        List1.AddItem ("R=" & CStr(JoystickDat.dwRpos))
        List1.AddItem ("Z=" & CStr(JoystickDat.dwZpos))
        List1.AddItem ("X=" & CStr(JoystickDat.dwXpos))
        List1.AddItem ("Y=" & CStr(JoystickDat.dwYpos))
        Picture1.Cls
        Picture1.DrawWidth = 5
        
        x = 33 + 5 * (JoystickDat.dwXpos - 32767) / 32767
        y = 43 + 5 * (JoystickDat.dwYpos - 32767) / 32767
        Picture1.Circle (x, y), 0, vbGreen
        
        x = 62 + 5 * (JoystickDat.dwRpos - 32767) / 32767
        y = 43 + 5 * (JoystickDat.dwZpos - 32767) / 32767
        Picture1.Circle (x, y), 0, vbGreen
        
        If JoystickDat.dwButtons And 1 Then
            Picture1.Circle (69, 27), 0, vbRed
        End If
        If JoystickDat.dwButtons And 2 Then
            Picture1.Circle (76, 20), 0, vbRed
        End If
        If JoystickDat.dwButtons And 4 Then
            Picture1.Circle (76, 34), 0, vbRed
        End If
        If JoystickDat.dwButtons And 8 Then
            Picture1.Circle (84, 27), 0, vbRed
        End If
        If JoystickDat.dwButtons And 16 Then
            Picture1.Circle (194, 31), 0, vbRed
        End If
        If JoystickDat.dwButtons And 32 Then
            Picture1.Circle (194, 42), 0, vbRed
        End If
        If JoystickDat.dwButtons And 64 Then
            Picture1.Circle (151, 37), 0, vbRed
        End If
        If JoystickDat.dwButtons And 128 Then
            Picture1.Circle (151, 48), 0, vbRed
        End If
        If JoystickDat.dwButtons And 256 Then
            Picture1.Circle (37, 20), 0, vbRed
        End If
        If JoystickDat.dwButtons And 512 Then
            Picture1.Circle (58, 20), 0, vbRed
        End If
        If JoystickDat.dwButtons And 1024 Then
            Picture1.Circle (33, 43), 0, vbRed
        End If
        If JoystickDat.dwButtons And 2048 Then
            Picture1.Circle (62, 43), 0, vbRed
        End If
        
        Select Case JoystickDat.dwPOV
            Case 0
                Picture1.Circle (19, 19), 0, vbRed
            Case 4500
                Picture1.Circle (23, 22), 0, vbRed
            Case 9000
                Picture1.Circle (26, 26), 0, vbRed
            Case 13500
                Picture1.Circle (23, 31), 0, vbRed
            Case 18000
                Picture1.Circle (19, 33), 0, vbRed
            Case 22500
                Picture1.Circle (15, 31), 0, vbRed
            Case 27000
                Picture1.Circle (13, 26), 0, vbRed
            Case 31500
                Picture1.Circle (15, 22), 0, vbRed
        End Select
    
    End If
    
End Sub
