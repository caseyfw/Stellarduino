VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Persist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private iniPath As String
Private INIfilename As String
Private key As String

Private Sub Class_Initialize()
    iniPath = Environ("APPDATA") + "\" + AppName
    On Error Resume Next
    If GetAttr(iniPath) <> vbDirectory Then
        MkDir (iniPath)
    End If
    INIfilename = iniPath + "\EQMOD.ini"
    key = "[default]"
End Sub

Public Sub Store()
Dim Index As Integer
End Sub

Public Sub Retrieve()

End Sub


' write using the default key and path
Public Sub WriteIniValue(PutVariable As String, PutValue As String)
    Call WriteIniValueEx(PutVariable, PutValue, key, INIfilename)
End Sub

' read using the default key and path
Public Function ReadIniValue(Variable As String) As String
ReadIniValue = ReadIniValueEx(Variable, key, INIfilename)
End Function

' write using any key and path
Public Sub DeleteSection(IniKey As String, filepath As String)

Dim NF1 As Integer
Dim NF2 As Integer
Dim filepath2 As String
Dim keyStr As String
Dim skip As Boolean
Dim temp1 As String

    On Error GoTo handleErr

    
    filepath2 = Left$(filepath, Len(filepath) - 4) & "_tmp.ini"
    
    NF1 = FreeFile
    NF2 = FreeFile
    
    Close NF1
    Close NF2
    Open filepath For Input As NF1
    Open filepath2 For Output As NF2
    
    skip = False
    While Not EOF(NF1)
        Line Input #NF1, temp1
        
        If Left$(temp1, 1) = "[" Then
            If temp1 = IniKey Then
                ' found secion to skip
                skip = True
            Else
                skip = False
            End If
        End If
        If Not skip Then
            ' copy to tempfile
            Print #NF2, temp1
        End If
    Wend
    
    Close #NF1
    Close #NF2
    
    On Error Resume Next
    Kill filepath
    Name filepath2 As filepath
    GoTo endsub
    
handleErr:
    Close #NF1
    Close #NF2
    Exit Sub

endsub:
    Close #NF1
    Close #NF2
End Sub




' write using any key and path
Public Sub WriteIniValueEx(PutVariable As String, PutValue As String, IniKey As String, filepath As String)
Dim temp As String
Dim LcaseTemp As String
Dim ReadKey As String
Dim ReadVariable As String
Dim LOKEY As Long
Dim HIKEY As Long
Dim KEYLEN As Long
Dim VAR As Integer
Dim VARENDOFLINE As Long
Dim NF As Integer
Dim x As Long

AssignVariables:
    NF = FreeFile
    ReadKey = vbCrLf & IniKey & Chr$(13)
    KEYLEN = Len(ReadKey)
    ReadVariable = Chr$(10) & LCase$(PutVariable) & "="
        
EnsureFileExists:
    Close NF
    Open filepath For Binary As NF
    Close NF
    SetAttr filepath, vbArchive
    
LoadFile:
    Open filepath For Input As NF
    temp = Input$(LOF(NF), NF)
    temp = vbCrLf & temp & "[]"
    Close NF
    LcaseTemp = LCase$(temp)
    
LogicMenu:
    LOKEY = InStr(LcaseTemp, ReadKey)
    If LOKEY = 0 Then GoTo AddKey:
    HIKEY = InStr(LOKEY + KEYLEN, LcaseTemp, "[")
    VAR = InStr(LOKEY, LcaseTemp, ReadVariable)
    If VAR > HIKEY Or VAR < LOKEY Then GoTo AddVariable:
    GoTo RenewVariable:
    
AddKey:
        temp = Left$(temp, Len(temp) - 2)
        temp = temp & vbCrLf & vbCrLf & IniKey & vbCrLf & PutVariable & "=" & PutValue
        GoTo TrimFinalString:
        
AddVariable:
        temp = Left$(temp, Len(temp) - 2)
        temp = Left$(temp, LOKEY + KEYLEN) & PutVariable & "=" & PutValue & vbCrLf & mid$(temp, LOKEY + KEYLEN + 1)
        GoTo TrimFinalString:
        
RenewVariable:
        temp = Left$(temp, Len(temp) - 2)
        VARENDOFLINE = InStr(VAR, temp, Chr$(13))
        temp = Left$(temp, VAR) & PutVariable & "=" & PutValue & mid$(temp, VARENDOFLINE)
        GoTo TrimFinalString:

TrimFinalString:
        temp = mid$(temp, 2)
        Do Until InStr(temp, vbCrLf & vbCrLf & vbCrLf) = 0
        temp = Replace(temp, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
        Loop
    
        Do Until Right$(temp, 1) > Chr$(13)
        temp = Left$(temp, Len(temp) - 1)
        Loop
    
        Do Until Left$(temp, 1) > Chr$(13)
        temp = mid$(temp, 2)
        Loop
    
OutputAmendedINIFile:
        Open filepath For Output As NF
        Print #NF, temp
        Close NF
    
End Sub

' read using any key and path
Public Function ReadIniValueEx(Variable As String, IniKey As String, filepath As String) As String
Dim NF As Integer
Dim temp As String
Dim LcaseTemp As String
Dim ReadyToRead As Boolean
    
AssignVariables:
        NF = FreeFile
        ReadIniValueEx = ""
        Variable = LCase$(Variable)
    
EnsureFileExists:
    Close NF
    Open filepath For Binary As NF
    Close NF
    SetAttr filepath, vbArchive
    
LoadFile:
    Open filepath For Input As NF
    While Not EOF(NF)
    Line Input #NF, temp
    LcaseTemp = LCase$(temp)
    If InStr(LcaseTemp, "[") <> 0 Then ReadyToRead = False
    If LcaseTemp = IniKey Then ReadyToRead = True
    If InStr(LcaseTemp, "[") = 0 And ReadyToRead = True Then
        If InStr(LcaseTemp, Variable & "=") = 1 Then
            ReadIniValueEx = mid$(temp, 1 + Len(Variable & "="))
            Close NF: Exit Function
            End If
        End If
    Wend
    Close NF
    ' didn't find a value, look in the registry instead
    ReadIniValueEx = oProfile.GetValue(oID, Variable)
End Function

Public Function GetIniPath() As String
    GetIniPath = iniPath
End Function
