Attribute VB_Name = "Language"
Option Explicit

Public LanguageDll As String
Public oLangDll As New LangDll

Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Declare Function GetUserDefaultLCID% Lib "kernel32" ()

' Constants used

Public Const LOCALE_SDECIMAL = &HE


' Function to fix the Numerical string to match the  Regional Settings

Public Function EQFixNum(numstring As String) As String

Dim Symbol As String
Dim iRet1 As Long
Dim iRet2 As Long
Dim lpLCDataVar As String
Dim pos As Integer
Dim Locale As Long
Dim tmpstring As String


    If numstring = "" Then
        EQFixNum = 0
        Exit Function
    End If

    ' Get the Local ID
     Locale = GetUserDefaultLCID()

    ' Get the Local info
   iRet1 = GetLocaleInfo(Locale, LOCALE_SDECIMAL, lpLCDataVar, 0)

   
   ' Get the symbol from the Local info
   Symbol = String$(iRet1, 0)
   iRet2 = GetLocaleInfo(Locale, LOCALE_SDECIMAL, Symbol, iRet1)
   pos = InStr(Symbol, Chr$(0))
   If pos > 0 Then Symbol = Left$(Symbol, pos - 1)

   tmpstring = Replace(numstring, ".", Symbol)

   If IsNumeric(tmpstring) = False Then
        tmpstring = Replace(numstring, ",", Symbol)
        If IsNumeric(tmpstring) = False Then
            MsgBox ("EQMOD: Invalid String Format " & tmpstring)
            tmpstring = "-1"
        End If
   End If
   
   EQFixNum = tmpstring

End Function

Public Function EQFixNum2(ByRef numstring As String) As Boolean

Dim Symbol As String
Dim iRet1 As Long
Dim iRet2 As Long
Dim lpLCDataVar As String
Dim pos As Integer
Dim Locale As Long
Dim tmpstring As String


    If numstring = "" Then
        EQFixNum2 = False
        Exit Function
    End If

    ' Get the Local ID
     Locale = GetUserDefaultLCID()

    ' Get the Local info
   iRet1 = GetLocaleInfo(Locale, LOCALE_SDECIMAL, lpLCDataVar, 0)

   
   ' Get the symbol from the Local info
   Symbol = String$(iRet1, 0)
   iRet2 = GetLocaleInfo(Locale, LOCALE_SDECIMAL, Symbol, iRet1)
   pos = InStr(Symbol, Chr$(0))
   If pos > 0 Then Symbol = Left$(Symbol, pos - 1)

   tmpstring = Replace(numstring, ".", Symbol)

   If IsNumeric(tmpstring) = False Then
        tmpstring = Replace(numstring, ",", Symbol)
        If IsNumeric(tmpstring) = False Then
            EQFixNum2 = False
            Exit Function
        End If
   End If
   
   EQFixNum2 = True

End Function

Public Function EQFixLocale2(numstring As String) As String

Dim Symbol As String
Dim iRet1 As Long
Dim iRet2 As Long
Dim lpLCDataVar As String
Dim pos As Integer
Dim Locale As Long
Dim tmpstring As String

    ' Get the Local ID
     Locale = GetUserDefaultLCID()

    ' Get the Local info
   iRet1 = GetLocaleInfo(Locale, LOCALE_SDECIMAL, lpLCDataVar, 0)

   
   ' Get the symbol from the Local info
   Symbol = String$(iRet1, 0)
   iRet2 = GetLocaleInfo(Locale, LOCALE_SDECIMAL, Symbol, iRet1)
   pos = InStr(Symbol, Chr$(0))
   If pos > 0 Then Symbol = Left$(Symbol, pos - 1)

   tmpstring = Replace(numstring, Symbol, ".")
   
   EQFixLocale2 = tmpstring

End Function


Public Function EQFixLocale(numstring As String) As String

Dim Symbol As String
Dim iRet1 As Long
Dim iRet2 As Long
Dim lpLCDataVar As String
Dim pos As Integer
Dim Locale As Long
Dim tmpstring As String

    ' Get the Local ID
     Locale = GetUserDefaultLCID()

    ' Get the Local info
   iRet1 = GetLocaleInfo(Locale, LOCALE_SDECIMAL, lpLCDataVar, 0)

   
   ' Get the symbol from the Local info
   Symbol = String$(iRet1, 0)
   iRet2 = GetLocaleInfo(Locale, LOCALE_SDECIMAL, Symbol, iRet1)
   pos = InStr(Symbol, Chr$(0))
   If pos > 0 Then Symbol = Left$(Symbol, pos - 1)

   tmpstring = Replace(numstring, ".", Symbol)
   tmpstring = Replace(tmpstring, ",", Symbol)

   EQFixLocale = tmpstring

End Function

Public Sub LoadLanguageDll()
     LanguageDll = HC.oPersist.ReadIniValue("LANG_DLL")
     If LanguageDll = "" Then
        ' no value exists - create a default from locale
        LanguageDll = oLangDll.GetDefaultDllName
        ' create an ini file entry but don't assign a value
        ' this way the loacale is always used unless manually edited.
        Call HC.oPersist.WriteIniValue("LANG_DLL", "")
     End If
End Sub
