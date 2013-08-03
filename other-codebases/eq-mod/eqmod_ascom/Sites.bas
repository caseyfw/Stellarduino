Attribute VB_Name = "Sites"
Option Explicit

Public Sub ReadSiteValues()
    
    Dim tmptxt As String
    Dim mins As Double
    Dim secs As Double
    
    HC.cbNS.ListIndex = 0
    HC.cbEW.ListIndex = 0
    HC.cbhem.ListIndex = 0
       
    tmptxt = HC.oPersist.ReadIniValue("LongitudeDeg")
    If tmptxt <> "" Then HC.txtLongDeg.Text = tmptxt
    
    tmptxt = HC.oPersist.ReadIniValue("LongitudeMin")
    If tmptxt <> "" Then
        HC.txtLongMin.Text = tmptxt
        mins = CDbl(HC.txtLongMin.Text)
        secs = 60 * (mins - Int(mins))
        HC.txtLongMin.Text = CStr(Int(mins))
        If secs <> 0 Then
            Call HC.oPersist.WriteIniValue("LongitudeSec", CStr(secs))
            Call HC.oPersist.WriteIniValue("LongitudeMin", HC.txtLongMin.Text)
        End If
    End If
    
    tmptxt = HC.oPersist.ReadIniValue("LongitudeSec")
    If tmptxt <> "" Then HC.txtLongSec.Text = tmptxt
     
    tmptxt = HC.oPersist.ReadIniValue("LongitudeEW")
    If tmptxt <> "" Then HC.cbEW.ListIndex = val(tmptxt)
    
    tmptxt = HC.oPersist.ReadIniValue("LatitudeDeg")
    If tmptxt <> "" Then HC.txtLatDeg.Text = tmptxt
    
    tmptxt = HC.oPersist.ReadIniValue("LatitudeMin")
    If tmptxt <> "" Then
        HC.txtLatMin.Text = tmptxt
        mins = CDbl(HC.txtLatMin.Text)
        secs = 60 * (mins - Int(mins))
        HC.txtLatMin.Text = CStr(Int(mins))
        If secs <> 0 Then
            Call HC.oPersist.WriteIniValue("LatitudeMin", HC.txtLatMin.Text)
            Call HC.oPersist.WriteIniValue("LatitudeSec", CStr(secs))
        End If
    End If
    
    tmptxt = HC.oPersist.ReadIniValue("LatitudeSec")
    If tmptxt <> "" Then HC.txtLatSec.Text = tmptxt
    
    tmptxt = HC.oPersist.ReadIniValue("LatitudeNS")
    If tmptxt <> "" Then HC.cbNS.ListIndex = val(tmptxt)
    
    tmptxt = HC.oPersist.ReadIniValue("Elevation")
    If tmptxt <> "" Then HC.txtElevation = tmptxt
    
    tmptxt = HC.oPersist.ReadIniValue("TimeDelta")
    If tmptxt <> "" Then gEQTimeDelta = val(EQFixNum(tmptxt))
    
    HC.cbhem.ListIndex = HC.cbNS.ListIndex
'     tmptxt = HC.oPersist.ReadIniValue("HemisphereNS")
'     If tmptxt <> "" Then HC.cbhem.ListIndex = val(tmptxt)
     
    gLongitude = CDbl(EQFixNum(HC.txtLongDeg)) + (CDbl(EQFixNum(HC.txtLongMin)) / 60#) + (CDbl(EQFixNum(HC.txtLongSec)) / 3600#)
    If HC.cbEW.Text = oLangDll.GetLangString(115) Then gLongitude = -gLongitude  ' W is neg
    
    gLatitude = CDbl(EQFixNum(HC.txtLatDeg)) + (CDbl(EQFixNum(HC.txtLatMin)) / 60#) + (CDbl(EQFixNum(HC.txtLatSec)) / 3600#)
    If HC.cbNS.Text = oLangDll.GetLangString(116) Then gLatitude = -gLatitude
    gElevation = CDbl(EQFixNum(HC.txtElevation))
    
    If HC.cbhem.Text = oLangDll.GetLangString(1110) Then
       gHemisphere = 0
    Else
       gHemisphere = 1
    End If
    
    tmptxt = HC.oPersist.ReadIniValue("SiteName")
    HC.SitesCombo.Text = HC.oPersist.ReadIniValue("SiteName")
          
End Sub

Public Sub WriteSiteValues()

        HC.oPersist.WriteIniValue "LatitudeDeg", CStr(HC.txtLatDeg.Text)
        HC.oPersist.WriteIniValue "LatitudeMin", CStr(HC.txtLatMin.Text)
        HC.oPersist.WriteIniValue "LatitudeSec", CStr(HC.txtLatSec.Text)
        HC.oPersist.WriteIniValue "LatitudeNS", CStr(HC.cbNS.ListIndex)
        HC.oPersist.WriteIniValue "LongitudeDeg", CStr(HC.txtLongDeg.Text)
        HC.oPersist.WriteIniValue "LongitudeMin", CStr(HC.txtLongMin.Text)
        HC.oPersist.WriteIniValue "LongitudeSec", CStr(HC.txtLongSec.Text)
        HC.oPersist.WriteIniValue "LongitudeEW", CStr(HC.cbEW.ListIndex)
        HC.oPersist.WriteIniValue "HemisphereNS", CStr(HC.cbhem.ListIndex)
        HC.oPersist.WriteIniValue "Elevation", CStr(HC.txtElevation.Text)
        HC.oPersist.WriteIniValue "TimeDelta", CStr(EQFixNum(str(gEQTimeDelta)))
        HC.oPersist.WriteIniValue "SiteName", CStr(HC.SitesCombo.Text)
        
End Sub

Public Sub LoadSites(combo As ComboBox)
Dim tmptxt As String
Dim key As String
Dim valstr As String
Dim Ini As String
Dim Index As Integer

    ' set up a file path for the align.ini file
    Ini = HC.oPersist.GetIniPath & "\EQMOD.ini"
        
    combo.Clear
    For Index = 1 To 10
        key = "[site" & CStr(Index) & "]"
        tmptxt = HC.oPersist.ReadIniValueEx("Name", key, Ini)
        If tmptxt <> "" Then
            combo.AddItem (tmptxt)
        Else
            combo.AddItem (oLangDll.GetLangString(187) & CStr(Index))
        End If
    Next Index

    ' set text
    combo.Text = HC.oPersist.ReadIniValue("SiteName")


End Sub

Public Sub LoadSite(ByVal Index As Integer)
Dim tmptxt As String
Dim key As String
Dim Ini As String
Dim Count As Integer
Dim secs As Double
Dim mins As Double

    ' set up a file path for the align.ini file
    Ini = HC.oPersist.GetIniPath & "\EQMOD.ini"
        
    key = "[site" & CStr(Index + 1) & "]"
    
    tmptxt = HC.oPersist.ReadIniValueEx("LongitudeDeg", key, Ini)
    If tmptxt <> "" Then HC.txtLongDeg.Text = tmptxt
    
    tmptxt = HC.oPersist.ReadIniValueEx("LongitudeMin", key, Ini)
    If tmptxt <> "" Then
        HC.txtLongMin.Text = tmptxt
        mins = CDbl(HC.txtLongMin.Text)
        secs = 60 * (mins - Int(mins))
        HC.txtLongMin.Text = CStr(Int(mins))
        If secs <> 0 Then
            Call HC.oPersist.WriteIniValueEx("LongitudeMin", HC.txtLongMin.Text, key, Ini)
            Call HC.oPersist.WriteIniValueEx("LongitudeSec", CStr(secs), key, Ini)
        End If
    End If
    
    tmptxt = HC.oPersist.ReadIniValueEx("LongitudeSec", key, Ini)
    If tmptxt <> "" Then HC.txtLongSec.Text = tmptxt
    
    tmptxt = HC.oPersist.ReadIniValueEx("LongitudeEW", key, Ini)
    If tmptxt <> "" Then HC.cbEW.ListIndex = val(tmptxt)
    
    tmptxt = HC.oPersist.ReadIniValueEx("LatitudeDeg", key, Ini)
    If tmptxt <> "" Then HC.txtLatDeg.Text = tmptxt
    
    tmptxt = HC.oPersist.ReadIniValueEx("LatitudeMin", key, Ini)
    If tmptxt <> "" Then
        HC.txtLatMin.Text = tmptxt
        mins = CDbl(HC.txtLatMin.Text)
        secs = 60 * (mins - Int(mins))
        HC.txtLatMin.Text = CStr(Int(mins))
        If secs <> 0 Then
            Call HC.oPersist.WriteIniValueEx("LatitudeMin", HC.txtLatMin.Text, key, Ini)
            Call HC.oPersist.WriteIniValueEx("LatitudeSec", CStr(secs), key, Ini)
        End If
    End If
    
    tmptxt = HC.oPersist.ReadIniValueEx("LatitudeSec", key, Ini)
    If tmptxt <> "" Then HC.txtLatSec.Text = tmptxt
    
    tmptxt = HC.oPersist.ReadIniValueEx("LatitudeNS", key, Ini)
    If tmptxt <> "" Then HC.cbNS.ListIndex = val(tmptxt)
    
    tmptxt = HC.oPersist.ReadIniValueEx("Elevation", key, Ini)
    If tmptxt <> "" Then HC.txtElevation = tmptxt
    
    tmptxt = HC.oPersist.ReadIniValueEx("TimeDelta", key, Ini)
    If tmptxt <> "" Then gEQTimeDelta = val(EQFixNum(tmptxt))
    
    HC.cbhem.ListIndex = HC.cbNS.ListIndex
'    tmptxt = HC.oPersist.ReadIniValueEx("HemisphereNS", key, Ini)
'    If tmptxt <> "" Then HC.cbhem.ListIndex = val(tmptxt)
  
End Sub

Public Sub SaveSite(ByVal Index As Integer, ByVal name As String)
    Dim key As String
    Dim Ini As String

    ' set up a file path for the align.ini file
    Ini = HC.oPersist.GetIniPath & "\EQMOD.ini"
        
    key = "[site" & CStr(Index + 1) & "]"
    
    Call HC.oPersist.WriteIniValueEx("Name", name, key, Ini)
    Call HC.oPersist.WriteIniValueEx("LatitudeDeg", CStr(HC.txtLatDeg.Text), key, Ini)
    Call HC.oPersist.WriteIniValueEx("LatitudeMin", CStr(HC.txtLatMin.Text), key, Ini)
    Call HC.oPersist.WriteIniValueEx("LatitudeSec", CStr(HC.txtLatSec.Text), key, Ini)
    Call HC.oPersist.WriteIniValueEx("LatitudeNS", CStr(HC.cbNS.ListIndex), key, Ini)
    Call HC.oPersist.WriteIniValueEx("LongitudeDeg", CStr(HC.txtLongDeg.Text), key, Ini)
    Call HC.oPersist.WriteIniValueEx("LongitudeMin", CStr(HC.txtLongMin.Text), key, Ini)
    Call HC.oPersist.WriteIniValueEx("LongitudeSec", CStr(HC.txtLongSec.Text), key, Ini)
    Call HC.oPersist.WriteIniValueEx("LongitudeEW", CStr(HC.cbEW.ListIndex), key, Ini)
    Call HC.oPersist.WriteIniValueEx("HemisphereNS", CStr(HC.cbhem.ListIndex), key, Ini)
    Call HC.oPersist.WriteIniValueEx("Elevation", CStr(HC.txtElevation.Text), key, Ini)
    Call HC.oPersist.WriteIniValueEx("TimeDelta", CStr(EQFixNum(str(gEQTimeDelta))), key, Ini)

End Sub

Public Sub UpdateSiteControls()
Dim tmp As Double
Dim h As Integer
Dim m As Integer
Dim s As Double

    tmp = Abs(gLongitude)
    If gLongitude < 0 Then
        HC.cbEW.ListIndex = 1
    Else
        HC.cbEW.ListIndex = 0
    End If
    h = Int(tmp)
    tmp = Int((tmp - h) * 60)
    m = Int(tmp)
    s = (tmp - m) * 60
    HC.txtLongDeg.Text = CStr(h)
    HC.txtLongMin.Text = CStr(m)
    HC.txtLongSec.Text = CStr(s)
    
    tmp = Abs(gLatitude)
    If gLatitude < 0 Then
        HC.cbNS.ListIndex = 1
        HC.cbhem.ListIndex = 1
    Else
        HC.cbNS.ListIndex = 0
        HC.cbhem.ListIndex = 0
    End If
    h = Int(tmp)
    tmp = (tmp - h) * 60
    m = Int(tmp)
    s = (tmp - m) * 60
    HC.txtLatDeg.Text = CStr(h)
    HC.txtLatMin.Text = CStr(m)
    HC.txtLatSec.Text = CStr(s)
    
    HC.txtElevation.Text = CStr(gElevation)
    
    Call WriteSiteValues
    
End Sub
