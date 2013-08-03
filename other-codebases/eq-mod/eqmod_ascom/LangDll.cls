VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LangDll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private DllName As String
Private DllHandle  As Long

Private Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hLib As Long, ByVal ResourceID As Long, ByVal lpBuffer As String, ByVal nBufferSize As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLib As Long) As Long 'BOOL
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal strFilePath As String) As Long


Private Declare Function GetThreadLocale Lib "kernel32" () As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" _
   (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long


Private Const LOCALE_SISO639LANGNAME = &H59      '  ISO abbreviated language name


Public Sub LoadLangDll(ByVal strDll As String)
  
  ' Close any dll already open
  Call CloseLibrary
  
  ' Load the new library
  DllHandle = LoadLibrary(strDll & Chr(0))
  If DllHandle <> 0 Then
    DllName = strDll
  Else
    DllName = ""
  End If
  
End Sub

Private Sub Class_Initialize()
    
    DllName = ""
    DllHandle = 0

End Sub
Private Sub Class_Terminate()
    CloseLibrary
End Sub

Public Sub CloseLibrary()
    If DllHandle <> 0 Then FreeLibrary DllHandle
    DllName = ""
    DllHandle = 0
End Sub

Public Function GetDefaultDllName() As String
    Dim LCID As Long
    Dim x As Long
    Dim lang As String
    
    lang = String$(16, " ")
    
    LCID = GetThreadLocale() 'Get Current locale
    
    x = GetLocaleInfo(LCID, &H59, lang, Len(lang))

    If lang <> "" Then
        GetDefaultDllName = "EQMOD" & Left$(lang, 2)
    Else
        GetDefaultDllName = ""
    End If
    
End Function



Public Function GetLangString(ByVal ResourceID As Integer) As String
  
    Dim hRes    As Long
    Dim strLen  As Long
    Dim tmpstr As String
    Dim error As Boolean
    
    tmpstr = "?"
    error = False
'    GoTo check_err
    
    If DllHandle = 0 Then
        error = True
        GoTo check_err
    End If
    
    If ResourceID < 1 Or ResourceID > 32767 Then
        error = True
        GoTo check_err
    Else
        ' Get the string from the dll
        tmpstr = String(256, Chr(0))
        strLen = LoadString(DllHandle, ResourceID, tmpstr, 256)
        If strLen <> 0 Then
            tmpstr = Left(tmpstr, strLen)
        Else
            error = True
        End If
    End If

check_err:
    If error Then
        ' get it from project resource instead
        tmpstr = LoadResString(ResourceID)
        If tmpstr = "" Then
            tmpstr = "?"
        End If
    End If
  
    GetLangString = tmpstr
    
End Function
