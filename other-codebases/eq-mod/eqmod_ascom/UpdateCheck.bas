Attribute VB_Name = "UpdateCheck"
Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Const INTERNET_FLAG_EXISTING_CONNECT = &H20000000
Const INTERNET_FLAG_DONT_CACHE = &H4000000
Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
Const INTERNET_FLAG_RELOAD = &H80000000

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
    (ByVal lpszAgent As String, ByVal dwAccessType As Long, _
    ByVal lpszProxyName As String, ByVal lpszProxyBypass As String, _
    ByVal dwFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias _
    "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, _
    ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, _
    ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As _
    Long) As Integer
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As _
    Long, ByVal lpBuffer As String, ByVal dwNumberOfBytesToRead As Long, _
    lNumberOfBytesRead As Long) As Integer

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long


Public gUpdateFileUrl As String
Public gUpdateFullUrl As String
Public gUpdateTestUrl As String
Public gUpdateMode As Integer
Public gUpdateAvailable As Boolean
Public gStrUpdateVersion As String

Type VersionData
    Major As Integer
    Minor As Integer
    alpha As Integer
End Type

Public Sub CheckForUpdate()
    gUpdateAvailable = False
    Call ReadUpdateParams
    Select Case gUpdateMode
        Case 0
        Case 1, 2
            If CopyURLToFile(gUpdateFileUrl, "versions.txt") = True Then
                Call CheckUpdateFile
            End If
    End Select
End Sub


Public Sub ReadUpdateParams()
    Dim tmptxt As String
    
    gUpdateFileUrl = HC.oPersist.ReadIniValue("UpdateFileUrl")
    If gUpdateFileUrl = "" Then
        gUpdateFileUrl = "http://eq-mod.sourceforge.net/versions/versions.txt"
        Call HC.oPersist.WriteIniValue("UpdateFileUrl", gUpdateFileUrl)
    End If
    
    gUpdateFullUrl = HC.oPersist.ReadIniValue("UpdateReleaseUrl")
    If gUpdateFullUrl = "" Then
        gUpdateFullUrl = "http://sourceforge.net/projects/eq-mod/files/"
        Call HC.oPersist.WriteIniValue("UpdateReleaseUrl", gUpdateFullUrl)
    End If
    
    gUpdateTestUrl = HC.oPersist.ReadIniValue("UpdateTestUrl")
    If gUpdateTestUrl = "" Then
        gUpdateTestUrl = "http://tech.groups.yahoo.com/group/EQMOD/"
        Call HC.oPersist.WriteIniValue("UpdateTestUrl", gUpdateTestUrl)
    End If
    
    
    tmptxt = HC.oPersist.ReadIniValue("UpdateMode")
    Select Case tmptxt
        Case "0", "1", "2"
            gUpdateMode = val(tmptxt)
        Case Else
            gUpdateMode = 0
            Call HC.oPersist.WriteIniValue("UpdateMode", "0")
    End Select

End Sub

' Download a file from Internet and save it to a local file
'
' it works with HTTP and FTP, but you must explicitly include
' the protocol name in the URL, as in
'    CopyURLToFile "http://www.vb2themax.com/default.asp", "C:\vb2themax.htm"
Public Function CopyURLToFile(ByVal URL As String, ByVal FileName As String) As Boolean
    Dim hInternetSession As Long
    Dim hUrl As Long
    Dim FileNum As Integer
    Dim ok As Boolean
    Dim NumberOfBytesRead As Long
    Dim Buffer As String
    Dim fileIsOpen As Boolean
    Dim error As Boolean

    error = False
    On Error GoTo errorhandler

    ' check obvious syntax errors
    If Len(URL) = 0 Or Len(FileName) = 0 Then
        error = True
    Else
        ' open an Internet session, and retrieve its handle
        hInternetSession = InternetOpen(App.EXEName, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
        If hInternetSession = 0 Then
            error = True
        Else
            ' open the file and retrieve its handle
            hUrl = InternetOpenUrl(hInternetSession, URL, vbNullString, 0, INTERNET_FLAG_EXISTING_CONNECT + INTERNET_FLAG_RELOAD, 0)
            If hUrl = 0 Then
                error = True
            Else
                ' ensure that there is no local file
                On Error Resume Next
                Kill FileName
                On Error GoTo errorhandler
                
                ' open the local file
                FileNum = FreeFile
                Open FileName For Binary As FileNum
                fileIsOpen = True
            
                ' prepare the receiving buffer
                Buffer = Space(4096)
                
                Do
                    ' read a chunk of the file - returns True if no error
                    ok = InternetReadFile(hUrl, Buffer, Len(Buffer), NumberOfBytesRead)
                    
                    ' makes sure our lines have CRLF terminators
                    Buffer = Replace(Buffer, vbCr, "")
                    Buffer = Replace(Buffer, vbLf, vbCrLf)
        
                    ' exit if error or no more data
                    If NumberOfBytesRead = 0 Or Not ok Then Exit Do
                    
                    ' save the data to the local file
                    Put #FileNum, , Left$(Buffer, NumberOfBytesRead + Len(Buffer) - 4096)
                Loop
            End If
        End If
        GoTo endfunc
    End If

errorhandler:
    On Error Resume Next
    error = True
endfunc:
    ' close the local file, if necessary
    Close #FileNum
    ' close internet handles, if necessary
    If hUrl Then InternetCloseHandle hUrl
    If hInternetSession Then InternetCloseHandle hInternetSession
    CopyURLToFile = Not error
End Function

Private Sub CheckUpdateFile()
    Dim FileNum As Integer
    Dim tmp1 As String
    Dim tmp2() As String
    Dim tmp3 As String
    Dim ver1 As VersionData
    Dim ver2 As VersionData
    
    On Error GoTo errhandler
    
    ver1 = GetVersionData(gVersion)
    
    ' open the local file
    FileNum = FreeFile
    Close FileNum
    Open "versions.txt" For Input As FileNum

    While Not EOF(1)
        Line Input #1, tmp1
        If Left(tmp1, 1) <> "#" Then
            tmp2 = Split(tmp1, " ")
            If tmp2(0) = "EQASCOM" Then
                Select Case gUpdateMode
                    Case 1
                        ver2 = GetVersionData(tmp2(1))
                        gUpdateAvailable = False
                        If ver2.Major > ver1.Major Then
                            gUpdateAvailable = True
                        Else
                            If ver2.Major = ver1.Major Then
                                If ver2.Minor > ver1.Minor Then
                                    gUpdateAvailable = True
                                Else
                                    If ver2.Minor = ver1.Minor Then
                                        If ver2.alpha > ver1.alpha Then
                                            gUpdateAvailable = True
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        If gUpdateAvailable Then
                            gStrUpdateVersion = tmp2(1)
                            
                        End If
                    Case 2
                        ver2 = GetVersionData(tmp2(2))
                        gUpdateAvailable = False
                        If ver2.Major > ver1.Major Then
                            gUpdateAvailable = True
                        Else
                            If ver2.Major = ver1.Major Then
                                If ver2.Minor > ver1.Minor Then
                                    gUpdateAvailable = True
                                Else
                                    If ver2.Minor = ver1.Minor Then
                                        If ver2.alpha > ver1.alpha Then
                                            gUpdateAvailable = True
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        If gUpdateAvailable Then
                            gStrUpdateVersion = Replace(tmp2(2), vbLf, "")
                        End If
                End Select
            End If
        End If
    Wend

    GoTo closefile

errhandler:
    gUpdateAvailable = False
closefile:
    Close FileNum


End Sub

Private Function GetVersionData(ByVal str As String) As VersionData
    Dim ver As VersionData
    Dim tmp() As String
    
    str = Replace(str, vbLf, "")
    ver.alpha = Asc(Right(str, 1))
    ' strip of the V and alpha
    str = mid(str, 2, Len(str) - 2)
    tmp = Split(str, ".")
    ver.Major = val(tmp(0))
    ver.Minor = val(tmp(1))
    
    GetVersionData = ver
    
End Function

' Open the default browser on a given URL
' Returns True if successful, False otherwise

Public Function OpenBrowser(ByVal URL As String) As Boolean
    Dim res As Long
    res = ShellExecute(0&, "open", URL, vbNullString, vbNullString, vbNormalFocus)
    OpenBrowser = (res > 32)
End Function


