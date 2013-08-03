Attribute VB_Name = "Monitor"
Option Explicit

'Turning the Computer Screen-Monitor on and off

'To switch a monitor on/off or to standby, use the following code. Note, this may
'not work on some NT machines.

Public gMonitorState As Integer
Public gMonitorMode As Integer

Private Const SC_MONITORPOWER = &HF170&
Private Const SPI_GETSCREENSAVEACTIVE As Long = &H10&
Private Const SPI_GETSCREENSAVERRUNNING As Long = &H72&

Const WM_CLOSE = &H10&
Const WM_SYSCOMMAND = &H112
Const SC_SCREENSAVE = &HF140&
Const HWND_BROADCAST = 65535
Const MONITOR_ON = -1&
Const MONITOR_OFF = 2&

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Boolean
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Sub ToggleMonitorPower()
    Select Case gMonitorMode
    
        Case 0
            ' power on/off
            If gMonitorState = 0 Then
                PostMessage HWND_BROADCAST, WM_SYSCOMMAND, SC_MONITORPOWER, MONITOR_ON
                gMonitorState = 1
                Call EQ_Beep(35)
            Else
                PostMessage HWND_BROADCAST, WM_SYSCOMMAND, SC_MONITORPOWER, MONITOR_OFF
                gMonitorState = 0
                Call EQ_Beep(36)
            End If
       
        Case 1
            ' screenscaver on/off
            If IsScreenSaverActivated Then
                ' close screensaver
                SendMessage GetForegroundWindow(), WM_CLOSE, 0&, 0&
                Call EQ_Beep(36)
            Else
                ' start screensaver
                SendMessage HC.hwnd, WM_SYSCOMMAND, SC_SCREENSAVE, 1&
                Call EQ_Beep(35)
            End If
            
    End Select
End Sub

Public Function IsScreenSaverActivated() As Boolean
  Dim a As Boolean
  SystemParametersInfo SPI_GETSCREENSAVERRUNNING, 0, a, False
  IsScreenSaverActivated = a
End Function

