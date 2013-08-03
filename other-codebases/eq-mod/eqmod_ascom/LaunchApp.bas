Attribute VB_Name = "LaunchApp"
Option Explicit

Public Type RECT
  Left    As Long
  Top     As Long
  Right   As Long
  Bottom  As Long
End Type

Public Type POINTAPI
  x       As Long
  Y       As Long
End Type

Public Type WINDOWPLACEMENT
  Length            As Long
  flags             As Long
  showCmd           As Long
  ptMinPosition     As POINTAPI
  ptMaxPosition     As POINTAPI
  rcNormalPosition  As RECT
End Type

Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_RESTORE = 9
Private Const GW_HWNDNEXT As Integer = 2
Private Const SW_NORMAL As Integer = 1

Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Integer) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Integer
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Integer) As Integer

' Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
' Private Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long


Public Function GetFirstWindowHandle(ByVal sStartingWith As String) As Long

    Dim hwnd As Long
    Dim sWindowName As String
    Dim iHandle As Long
   
    hwnd = GetTopWindow(GetDesktopWindow())
    
    Do While hwnd <> 0
       sWindowName = zGetWindowName(hwnd)
       If InStr(1, sWindowName, sStartingWith) = 1 Then
          iHandle = hwnd
          Exit Do
       End If
       hwnd = GetWindow(hwnd, GW_HWNDNEXT)
    Loop
    
    GetFirstWindowHandle = iHandle

End Function

Private Function zGetWindowName(ByVal hwnd As Long) As String

    Dim nBufferLength As Integer
    Dim nTextLength As Integer
    Dim sName As String
    
    sName = String(100, Chr$(0))
'    nBufferLength = GetWindowTextLength(hwnd) + 4
'    sName = Space(nBufferLength + 1)
    nTextLength = GetWindowText(hwnd, sName, 100)
    sName = Left(sName, nTextLength)
    zGetWindowName = sName

End Function


Public Sub LaunchUtilityApp(Index As Integer)
Dim tmptxt As String
Dim nofile As Boolean
Dim THandle As Long
Dim processId As Long
Dim iret As Long
Dim inistr As String
Dim strwnd As String
Dim clientini As String
Dim found As Boolean
Dim pos As Integer
Dim lState As Long
Dim lpwndpl As WINDOWPLACEMENT

On Error GoTo launcherr:

    Select Case Index
        Case 0
            inistr = "TOUR_EXE"
            clientini = Environ("APPDATA") & "\EQMOD\EQTOUR.ini"
        Case 1
            inistr = "MOSAIC_EXE"
            clientini = Environ("APPDATA") & "\EQMOD\EQMOSAIC.ini"
    End Select

    nofile = False
    ' get application path
    tmptxt = HC.oPersist.ReadIniValue(inistr)
    If tmptxt = "" Then
        nofile = True
    Else
        If dir(tmptxt) = "" Then
            nofile = True
        End If
    End If
    
    ' no application path set so let user assign one
    If nofile Then
        Call SetUtilityApp(Index)
        tmptxt = HC.oPersist.ReadIniValue(inistr)
        If tmptxt = "" Then
            Exit Sub
        Else
            If dir(tmptxt) = "" Then
                Exit Sub
            End If
        End If
    End If
    
    pos = InStrRev(tmptxt, "\")
    strwnd = Right(tmptxt, Len(tmptxt) - pos)
    strwnd = Left(strwnd, Len(strwnd) - 4)
    
    Select Case strwnd
        Case "EQTOUR"
            strwnd = "EQTOUR V"
        Case "EQMOSAIC"
            strwnd = "EQMOSAIC V"
        Case "TonightSky"
            strwnd = "Tonight Sky"
    End Select

    ' set tour or mosaic to connect to this driver.
    Call HC.oPersist.WriteIniValueEx("ASCOM_ID", ASCOM_id, "[default]", clientini)
    
    found = False
    THandle = GetFirstWindowHandle(strwnd)
'    THandle = FindWindow(vbNullString, strwnd)
    If THandle <> 0 Then
        ' found a window bit is a the main window?
'        If GetParent(THandle) = 0 Then
            RestoreWindow THandle
            found = True
'        End If
    End If
    If Not found Then
        ' open application
        processId = Shell(tmptxt, vbNormalFocus)
        If processId <> 0 Then
            ' bring it to the top
'            THandle = FindWindow(vbEmpty, strwnd)
            THandle = GetFirstWindowHandle(strwnd)
            If THandle <> 0 Then
               iret = BringWindowToTop(THandle)
            End If
        End If
    End If
    Exit Sub
    
    
launcherr:
    ' been an error of some sort! - remove infile entry
    Call HC.oPersist.WriteIniValue(inistr, "")

endlaunch:
End Sub

Public Sub SetUtilityApp(Index As Integer)
Dim tmptxt As String
Dim nofile As Boolean
Dim THandle As Long
Dim processId As Long
Dim iret As Long
Dim inistr As String
Dim strfilter As String

On Error GoTo Seterr:

    Select Case Index
        Case 0
            inistr = "TOUR_EXE"
            strfilter = "eqtour.exe;tonightsky.exe"
        Case 1
            inistr = "MOSAIC_EXE"
            strfilter = "eqmosaic.exe"
    End Select

    FileDlg.filter = strfilter
    FileDlg.Show (1)
    tmptxt = FileDlg.FileName
    If tmptxt <> "" Then
        If dir(tmptxt) <> "" Then
            Call HC.oPersist.WriteIniValue(inistr, tmptxt)
        End If
    End If
    GoTo endset
    
Seterr:
    ' been an error of some sort! - remove infile entry
    Call HC.oPersist.WriteIniValue(inistr, "")

endset:
End Sub


Private Sub RestoreWindow(hWndCtlApp As Long)

Dim currWinP As WINDOWPLACEMENT
    
    'prepare the WINDOWPLACEMENT type
    currWinP.Length = Len(currWinP)
    If GetWindowPlacement(hWndCtlApp, currWinP) > 0 Then
        'determine the window state
        If currWinP.showCmd = SW_SHOWMINIMIZED Then
            'minimized, so restore
            currWinP.Length = Len(currWinP)
            currWinP.flags = 0&
            currWinP.showCmd = SW_SHOWNORMAL
            Call SetWindowPlacement(hWndCtlApp, currWinP)
        Else
            'on screen, so assure visible
            Call SetForegroundWindow(hWndCtlApp)
            Call BringWindowToTop(hWndCtlApp)
        End If
    End If

End Sub

