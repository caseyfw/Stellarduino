Attribute VB_Name = "Common"
'---------------------------------------------------------------------
' Copyright © 2006 Raymund Sarmiento
'
' Permission is hereby granted to use this Software for any purpose
' including combining with commercial products, creating derivative
' works, and redistribution of source or binary code, without
' limitation or consideration. Any redistributed copies of this
' Software must include the above Copyright Notice.
'
' THIS SOFTWARE IS PROVIDED "AS IS". THE AUTHOR OF THIS CODE MAKES NO
' WARRANTIES REGARDING THIS SOFTWARE, EXPRESS OR IMPLIED, AS TO ITS
' SUITABILITY OR FITNESS FOR A PARTICULAR PURPOSE.
'---------------------------------------------------------------------
'
' Common.bas - Common functions for EQMOD ASCOM driver
'
'
' Written:  07-Oct-06   Raymund Sarmiento
'
'

'  DISCLAIMER:

'  You can use the information on this site COMPLETELY AT YOUR OWN RISK.
'  The modification steps and other information on this site is provided
'  to you "AS IS" and WITHOUT WARRANTY OF ANY KIND, express, statutory,
'  implied or otherwise, including without limitation any warranty of
'  merchantability or fitness for any particular or intended purpose.
'  In no event the author will  be liable for any direct, indirect,
'  punitive, special, incidental or consequential damages or loss of any
'  kind whether or not the author  has been advised of the possibility
'  of such loss.

'  WARNING:

'  Circuit modifications implemented on your setup could invalidate
'  any warranty that you may have with your product. Use this
'  information at your own risk. The modifications involve direct
'  access to the stepper motor controls of your mount. Any "mis-control"
'  or "mis-command"  / "invalid parameter" or "garbage" data sent to the
'  mount could accidentally activate the stepper motors and allow it to
'  rotate "freely" damaging any equipment connected to your mount.
'  It is also possible that any garbage or invalid data sent to the mount
'  could cause its firmware to generate mis-steps pulse sequences to the
'  motors causing it to overheat. Make sure that you perform the
'  modifications and testing while there is no physical "load" or
'  dangling wires on your mount. Be sure to disconnect the power once
'  this event happens or if you notice any unusual sound coming from
'  the motor assembly.
'
'  CREDITS:
'
'  Portions of the information on this code should be attributed
'  to Mr. John Archbold from his initial observations and analysis
'  of the interface circuits and of the ASCII data stream between
'  the Hand Controller (HC) and the Go To Controller.
'
Option Explicit


'------------------- EQCONTRL.DLL Constants -----------------------
Public Const EQ_OK As Double = &H0
Public Const EQ_COMNOTOPEN As Double = &H1
Public Const EQ_COMTIMEOUT As Double = &H3
Public Const EQ_MOTORBUSY As Double = &H10
Public Const EQ_NOTINITIALIZED As Double = &HC8
Public Const EQ_INVALIDCOORDINATE As Double = &H1000000
Public Const EQ_INVALID As Double = &H3000000

' Protocol types
Public Const CURMOUNT As Long = 0 'Detected Current Mount
Public Const EQMOUNT As Long = 1 'EQG Protocol
Public Const NXMOUNT As Long = 2 'NexStar Protocol
Public Const LXMOUNT As Long = 3 'LX200 Protocol
Public Const TKMOUNT As Long = 4 'Takahashi Protocol
Public Const HBXMOUNT As Long = 5 'Meade HBX

' coordinate types
Public Const CT_STEP As Long = 0
Public Const CT_RADEC As Long = 1
Public Const CT_AZALT As Long = 2
'------------------------------------------------------------------

'Virtual Desktop sizes
Const SM_XVIRTUALSCREEN = 76    'Virtual Left
Const SM_YVIRTUALSCREEN = 77    'Virtual Top
Const SM_CXVIRTUALSCREEN = 78   'Virtual Width
Const SM_CYVIRTUALSCREEN = 79   'Virtual Height
Const SM_CMONITORS = 80         'Get number of monitors
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long


Const HWND_TOPMOST = -1
Const HWND_NOTTOPMOST = -2
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1

Type ASCOM_COMPLIANCE
    SlewWithTrackingOff As Boolean
    AllowPulseGuide As Boolean
    AllowExceptions As Boolean
    AllowPulseGuideExceptions As Boolean
    BlockPark As Boolean
    AllowSiteWrites As Boolean
    Epoch As Integer
End Type


Public gAscomCompatibility As ASCOM_COMPLIANCE
Public oProfile As DriverHelper.Profile
Public Const oID As String = "EQMOD.Telescope"
Public m_telescope As Telescope
Public gPresetSlewRates(1 To 10) As Double
Public gRateButtons(1 To 4) As Integer
Public gPresetSlewRatesCount As Integer
Public gCurrentRatePreset As Integer
Public gPoleStarRa As Double
Public gPoleStarDec As Double
Public gPoleStarRaJ2000 As Double
Public gPoleStarDecJ2000 As Double
Public gPoleStarReticuleDec As Double
Public gPolarReticuleEpoch As Double
Public gPolHa As Double
Public gVersion As String
Public gShowPolarAlign As Integer
Public gAlignmentMode As Integer
Public gCoordType As Long
Public gDllVer As Double
Public g3PointAlgorithm As Integer
Public gAdvanced As Integer
Public gPointFilter As Integer
Public gBacklashDec As Integer
Public gDriftComp As Integer
Public gPoleStarIdx As Integer

Public gPulseguideRateRa As Double
Public gPulseguideRateDec As Double

Public gCommErrorStop As Integer

Public ClientCount As Integer
Public gInitResult As Double

Private Const oDESC As String = "EQMOD ASCOM Scope Driver"


Private Const SC_CLOSE As Long = &HF060&
Private Const MIIM_STATE As Long = &H1&
Private Const MIIM_ID As Long = &H2&
Private Const MFS_GRAYED As Long = &H3&
Private Const WM_NCACTIVATE As Long = &H86

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

    

Private Declare Function GetSystemMenu Lib "user32" ( _
    ByVal hwnd As Long, ByVal bRevert As Long) As Long

Private Declare Function GetMenuItemInfo Lib "user32" Alias _
    "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, _
    ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long

Private Declare Function SetMenuItemInfo Lib "user32" Alias _
    "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, _
    ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long

Private Declare Function SendMessage Lib "user32" Alias _
    "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Any) As Long

Private Declare Function IsWindow Lib "user32" _
    (ByVal hwnd As Long) As Long


' Locale Info for Regional Settings processing
    


Private Declare Function SetWindowPos Lib "user32" _
         (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, _
          ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long



Public Sub PutWindowOnTop(pFrm As Form)
  Dim lngWindowPosition As Long
  
  lngWindowPosition = SetWindowPos(pFrm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)

End Sub
Public Sub PutWindowNormal(pFrm As Form)
  Dim lngWindowPosition As Long
  
  lngWindowPosition = SetWindowPos(pFrm.hwnd, HWND_NOTTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)

End Sub




Public Function EnableCloseButton(ByVal hwnd As Long, Enable As Boolean) _
                                                                As Integer
    Const xSC_CLOSE As Long = -10

    ' Check that the window handle passed is valid
    
    EnableCloseButton = -1
    If IsWindow(hwnd) = 0 Then Exit Function
    
    ' Retrieve a handle to the window's system menu
    
    Dim hMenu As Long
    hMenu = GetSystemMenu(hwnd, 0)
    
    ' Retrieve the menu item information for the close menu item/button
    
    Dim MII As MENUITEMINFO
    MII.cbSize = Len(MII)
    MII.dwTypeData = String(80, 0)
    MII.cch = Len(MII.dwTypeData)
    MII.fMask = MIIM_STATE
    
    If Enable Then
        MII.wID = xSC_CLOSE
    Else
        MII.wID = SC_CLOSE
    End If
    
    EnableCloseButton = -0
    If GetMenuItemInfo(hMenu, MII.wID, False, MII) = 0 Then Exit Function
    
    ' Switch the ID of the menu item so that VB can not undo the action itself
    
    Dim lngMenuID As Long
    lngMenuID = MII.wID
    
    If Enable Then
        MII.wID = SC_CLOSE
    Else
        MII.wID = xSC_CLOSE
    End If
    
    MII.fMask = MIIM_ID
    EnableCloseButton = -2
    If SetMenuItemInfo(hMenu, lngMenuID, False, MII) = 0 Then Exit Function
    
    ' Set the enabled / disabled state of the menu item
    
    If Enable Then
        MII.fState = (MII.fState Or MFS_GRAYED)
        MII.fState = MII.fState - MFS_GRAYED
    Else
        MII.fState = (MII.fState Or MFS_GRAYED)
    End If
    
    MII.fMask = MIIM_STATE
    EnableCloseButton = -3
    If SetMenuItemInfo(hMenu, MII.wID, False, MII) = 0 Then Exit Function
    
    ' Activate the non-client area of the window to update the titlebar, and
    ' draw the close button in its new state.
    
    SendMessage hwnd, WM_NCACTIVATE, True, 0
    
    EnableCloseButton = 0
    
End Function

Public Sub Main()

    ClientCount = 0
    
    Set oProfile = New DriverHelper.Profile
    'Dim m_telescope As Telescope
    Set m_telescope = New Telescope
    
    oProfile.DeviceType = "Telescope"
    
    Set g_TrackingRates = New TrackingRates
    g_TrackingRates.Add driveSidereal
'    g_TrackingRates.Add driveLunar
'    g_TrackingRates.Add driveSolar
    
    
    If App.StartMode = vbSModeStandalone Then
        MsgBox AppName & " is an ASCOM driver. It cannot be run stand-alone", _
                (vbOKOnly + vbCritical + vbMsgBoxSetForeground), App.FileDescription
        Exit Sub
   End If

End Sub



Public Sub SyncToRADEC(ByVal RightAscension As Double, ByVal Declination As Double, ByVal pLongitude As Double, ByVal pHemisphere As Long)

                                    
Dim targetRAEncoder As Double
Dim targetDECEncoder As Double
Dim currentRAEncoder As Double
Dim currentDECEncoder As Double
Dim SaveRaSync As Double
Dim SaveDecSync As Double


Dim tRA As Double
Dim tha As Double
Dim tPier As Double

Dim tmpcoord As Coordt

    If HC.ListSyncMode.ListIndex = 1 Then
        ' Append via sync mode!
        Call EQ_NPointAppend(RightAscension, Declination, pLongitude, pHemisphere)
        Exit Sub
    Else
        ' its an ascom sync - shift whole model
        SaveDecSync = gDECSync01
        SaveRaSync = gRASync01
        gRASync01 = 0
        gDECSync01 = 0
    
        HC.EncoderTimer.Enabled = False
        
        If gThreeStarEnable = False Then
            currentRAEncoder = EQGetMotorValues(0) + gRA1Star
            currentDECEncoder = EQGetMotorValues(1) + gDEC1Star
        Else
    
            Select Case gAlignmentMode
    
                Case 2
                    ' nearest
                   tmpcoord = DeltaSync_Matrix_Map(EQGetMotorValues(0), EQGetMotorValues(1))
                   currentRAEncoder = tmpcoord.x
                   currentDECEncoder = tmpcoord.Y
                
                Case 1
                    ' n-star
                    tmpcoord = Delta_Matrix_Reverse_Map(EQGetMotorValues(0), EQGetMotorValues(1))
                    currentRAEncoder = tmpcoord.x
                    currentDECEncoder = tmpcoord.Y
                    
                Case Else
                    'n-star+nearest
                    tmpcoord = Delta_Matrix_Reverse_Map(EQGetMotorValues(0), EQGetMotorValues(1))
                    currentRAEncoder = tmpcoord.x
                    currentDECEncoder = tmpcoord.Y
                
                    If tmpcoord.F = 0 Then
                        tmpcoord = DeltaSync_Matrix_Map(EQGetMotorValues(0), EQGetMotorValues(1))
                        currentRAEncoder = tmpcoord.x
                        currentDECEncoder = tmpcoord.Y
                    End If
            
            End Select
    
        End If
        
        HC.EncoderTimer.Enabled = True
        tha = RangeHA(RightAscension - EQnow_lst(pLongitude * DEG_RAD))
      
        If tha < 0 Then
            If pHemisphere = 0 Then
                tPier = 1
            Else
                tPier = 0
            End If
            tRA = Range24(RightAscension - 12)
        Else
            If pHemisphere = 0 Then
                tPier = 0
            Else
                tPier = 1
            End If
            tRA = RightAscension
        End If
    
        'Compute for Sync RA/DEC Encoder Values
    
         targetRAEncoder = Get_RAEncoderfromRA(tRA, 0, pLongitude, gRAEncoder_Zero_pos, gTot_RA, pHemisphere)
         targetDECEncoder = Get_DECEncoderfromDEC(Declination, tPier, gDECEncoder_Zero_pos, gTot_DEC, pHemisphere)
         
        If (Abs(targetRAEncoder - currentRAEncoder) > gEQ_MAXSYNC) Or (Abs(targetDECEncoder - currentDECEncoder) > gEQ_MAXSYNC) Then
            Call HC.Add_Message(oLangDll.GetLangString(6004))
'            gDECSync01 = SaveDecSync
'            gRASync01 = SaveRaSync
            HC.Add_Message ("RA=" & FmtSexa(gRA, False) & " " & CStr(currentRAEncoder))
            HC.Add_Message ("SyncRA=" & FmtSexa(RightAscension, False) & " " & CStr(targetRAEncoder))
            HC.Add_Message ("DEC=" & FmtSexa(gDec, True) & " " & CStr(currentDECEncoder))
            HC.Add_Message ("Sync   DEC=" & FmtSexa(Declination, True) & " " & CStr(targetDECEncoder))
        Else
            gRASync01 = targetRAEncoder - currentRAEncoder
            gDECSync01 = targetDECEncoder - currentDECEncoder
        End If
        
        Call WriteSyncMap
        gEmulOneShot = True    ' Re Sync Display
        HC.DxSalbl.Caption = Format$(str(gRASync01), "000000000")
        HC.DxSblbl.Caption = Format$(str(gDECSync01), "000000000")
    End If
End Sub

Public Sub readlastpos()

     Dim tmptxt As String

     tmptxt = HC.oPersist.ReadIniValue("LASTPOS_RA")
     If tmptxt <> "" Then
        gRAEncoderlastpos = val(tmptxt)
     Else
        gRAEncoderlastpos = RAEncoder_Home_pos
     End If
     
     tmptxt = HC.oPersist.ReadIniValue("LASTPOS_DEC")
     If tmptxt <> "" Then
        gDECEncoderlastpos = val(tmptxt)
     Else
        gDECEncoderlastpos = gDECEncoder_Home_pos
     End If
 
End Sub

Public Sub writelastpos()

    HC.oPersist.WriteIniValue "LASTPOS_RA", CStr(gRAEncoderlastpos)
    HC.oPersist.WriteIniValue "LASTPOS_DEC", CStr(gDECEncoderlastpos)

End Sub
Public Sub WriteSyncMap()

    HC.oPersist.WriteIniValue "RSYNC01", CStr(gRASync01)
    HC.oPersist.WriteIniValue "DSYNC01", CStr(gDECSync01)

End Sub

Public Sub WriteAlignMap()

    HC.oPersist.WriteIniValue "RALIGN01", CStr(gRA1Star)
    HC.oPersist.WriteIniValue "DALIGN01", CStr(gDEC1Star)

End Sub

Public Sub readPolarHomeGoto()

     Dim tmptxt As String

     tmptxt = HC.oPersist.ReadIniValue("POLARHOME_GOTO_RA")
     If tmptxt <> "" Then
        gRAEncoderPolarHomeGoto = val(tmptxt)
     Else
        gRAEncoderPolarHomeGoto = 0
     End If
     
     tmptxt = HC.oPersist.ReadIniValue("POLARHOME_GOTO_DEC")
     If tmptxt <> "" Then
        gDECEncoderPolarHomeGoto = val(tmptxt)
     Else
        gDECEncoderPolarHomeGoto = 0
     End If

End Sub

Public Sub writePolarHomeGoto(ByVal StartPos As Integer)

    HC.oPersist.WriteIniValue "POLARHOME_GOTO_RA", CStr(gRAEncoderPolarHomeGoto)
    HC.oPersist.WriteIniValue "POLARHOME_GOTO_DEC", CStr(gDECEncoderPolarHomeGoto)
    Call HC.oPersist.WriteIniValue("POLARHOME_RETICULE_START", CStr(StartPos))
    
End Sub



Public Sub resetsync()

    gRASync01 = 0
    gDECSync01 = 0
    
    WriteSyncMap
    
    HC.DxSalbl.Caption = Format$(str(gRASync01), "000000000")
    HC.DxSblbl.Caption = Format$(str(gDECSync01), "000000000")
    
End Sub

Public Sub writeratebarstateHC()

    HC.oPersist.WriteIniValue "BAR01_1", CStr(HC.VScrollRASlewRate.Value)
    HC.oPersist.WriteIniValue "BAR01_2", CStr(HC.VScrollDecSlewRate.Value)
    HC.oPersist.WriteIniValue "BAR01_3", CStr(HC.HScrollRARate.Value)
    HC.oPersist.WriteIniValue "BAR01_4", CStr(HC.HScrollDecRate.Value)
    HC.oPersist.WriteIniValue "BAR01_5", CStr(HC.HScrollRAOride.Value)
    HC.oPersist.WriteIniValue "BAR01_6", CStr(HC.HScrollDecOride.Value)

End Sub

Public Sub readratebarstateHC()

     Dim tmptxt As String

     tmptxt = HC.oPersist.ReadIniValue("BAR01_1")
     If tmptxt <> "" Then
        HC.VScrollRASlewRate.Value = val(tmptxt)
     End If
     
     tmptxt = HC.oPersist.ReadIniValue("BAR01_2")
     If tmptxt <> "" Then
        HC.VScrollDecSlewRate.Value = val(tmptxt)
     End If
 
     tmptxt = HC.oPersist.ReadIniValue("BAR01_3")
     If tmptxt <> "" Then
        HC.HScrollRARate.Value = val(tmptxt)
     End If
     
     tmptxt = HC.oPersist.ReadIniValue("BAR01_4")
     If tmptxt <> "" Then
        HC.HScrollDecRate.Value = val(tmptxt)
     End If

     tmptxt = HC.oPersist.ReadIniValue("BAR01_5")
     If tmptxt <> "" Then
        HC.HScrollRAOride.Value = val(tmptxt)
     End If
     
     tmptxt = HC.oPersist.ReadIniValue("BAR01_6")
     If tmptxt <> "" Then
        HC.HScrollDecOride.Value = val(tmptxt)
     End If
End Sub

Public Sub writeratebarstateAlign()

    HC.oPersist.WriteIniValue "BAR02_1", CStr(Align.HScroll1.Value)
    HC.oPersist.WriteIniValue "BAR02_2", CStr(Align.HScroll2.Value)


End Sub

Public Sub readratebarstateAlign()

     Dim tmptxt As String

     tmptxt = HC.oPersist.ReadIniValue("BAR02_1")
     If tmptxt <> "" Then
        Align.HScroll1.Value = val(tmptxt)
     End If
     
     tmptxt = HC.oPersist.ReadIniValue("BAR02_2")
     If tmptxt <> "" Then
        Align.HScroll2.Value = val(tmptxt)
     End If
 
  
End Sub

Public Sub writeratebarstatePad()

    HC.oPersist.WriteIniValue "BAR03_1", CStr(Slewpad.VScroll1.Value)
    HC.oPersist.WriteIniValue "BAR03_2", CStr(Slewpad.VScroll2.Value)


End Sub

Public Sub writeOnTop()

    HC.oPersist.WriteIniValue "ON_TOP1", CStr(HC.HCOnTop.Value)

End Sub

Public Sub readOnTop()

     Dim tmptxt As String

     tmptxt = HC.oPersist.ReadIniValue("ON_TOP1")
     If tmptxt <> "" Then
        HC.HCOnTop.Value = val(tmptxt)
     End If

End Sub

Public Sub writeAlignCheck1()
    Select Case gAlignmentMode
        Case 0
            ' n-star+nearset
            HC.oPersist.WriteIniValue "SYNCNSTAR", "0"
        Case 1
            ' n-star - no longer used so force to n-star_nearest
            HC.oPersist.WriteIniValue "SYNCNSTAR", "0"
        Case 2
            ' nearest
            HC.oPersist.WriteIniValue "SYNCNSTAR", "2"
    End Select
End Sub

Public Sub writeAlignCheck2()
    HC.oPersist.WriteIniValue "APPENDSYNCNSTAR", CStr(HC.ListSyncMode.ListIndex)
    Select Case HC.ListSyncMode.ListIndex
        Case 0
            ' ascom standard
            HC.CommandAddPoint.Visible = True
        Case 1
            ' append syncs
            HC.CommandAddPoint.Visible = False
     End Select
End Sub

Public Sub readAlignCheck()

     Dim tmptxt As String

     tmptxt = HC.oPersist.ReadIniValue("SYNCNSTAR")
     If tmptxt <> "" Then
        Select Case tmptxt
            Case "0"
                ' nstar+nearest
                gAlignmentMode = 0
                HC.ListAlignMode.ListIndex = 0 ' N-Star+nearest
            Case "1"
                ' nstar - no longer provided!
                gAlignmentMode = 0 ' use nstar+nearest insted
                HC.ListAlignMode.ListIndex = 0 ' N-Star+nearest
            Case "2"
                ' nearest
                gAlignmentMode = 2
                HC.ListAlignMode.ListIndex = 1 ' nearest
        End Select
     Else
        gAlignmentMode = 0
        HC.ListAlignMode.ListIndex = 0
     End If

     tmptxt = HC.oPersist.ReadIniValue("APPENDSYNCNSTAR")
     Select Case tmptxt
        Case "0"
            ' ascom standard
            HC.ListSyncMode.ListIndex = 0
        Case "1"
            ' append syncs
            HC.ListSyncMode.ListIndex = 1
        Case Else
            ' default = append syncs
            HC.ListSyncMode.ListIndex = 1
            ' write default to ini file
            Call writeAlignCheck2
     End Select
     
     tmptxt = HC.oPersist.ReadIniValue("NSTAR_MAXCOMBINATION")
     If tmptxt <> "" Then
        gMaxCombinationCount = val(tmptxt)
     Else
        gMaxCombinationCount = MAX_COMBINATION_COUNT
        HC.oPersist.WriteIniValue "NSTAR_MAXCOMBINATION", CStr(gMaxCombinationCount)
     End If
     
     tmptxt = HC.oPersist.ReadIniValue("ALIGN_PROXIMITY")
     If tmptxt <> "" Then
       HC.HScrollProximity.Value = val(tmptxt)
     Else
        HC.HScrollProximity.Value = 0
        Call writeAlignProximity
     End If
     
     tmptxt = HC.oPersist.ReadIniValue("ALIGN_SELECTION")
     If tmptxt <> "" Then
        
        gPointFilter = val(tmptxt)
     Else
        gPointFilter = 0
        HC.oPersist.WriteIniValue "ALIGN_SELECTION", "0"
     End If
    HC.ComboActivePoints.ListIndex = gPointFilter
     
     tmptxt = HC.oPersist.ReadIniValue("ALIGN_LOCALTOPIER")
     If tmptxt <> "" Then
       HC.CheckLocalPier.Value = val(tmptxt)
     Else
        HC.CheckLocalPier.Value = 1
        HC.oPersist.WriteIniValue "ALIGN_LOCALTOPIER", "1"
     End If
     
End Sub

Public Sub writeAlignProximity()
    HC.oPersist.WriteIniValue "ALIGN_PROXIMITY", CStr(HC.HScrollProximity.Value)
End Sub

Public Sub readAlignProximity()

     Dim tmptxt As String

     tmptxt = HC.oPersist.ReadIniValue("ALIGN_PROXIMITY")
     If tmptxt <> "" Then
       HC.HScrollProximity.Value = val(tmptxt)
     Else
        HC.HScrollProximity.Value = 0
        Call writeAlignProximity
     End If
     CalcPromximityLimits (HC.HScrollProximity.Value)
End Sub

Public Sub writeColorDat(a1 As Long, a2 As Long, a3 As Long, b1 As Long, b2 As Long, b3 As Long, F1 As Long)

    HC.oPersist.WriteIniValue "FOR_R", CStr(a1)
    HC.oPersist.WriteIniValue "FOR_G", CStr(a2)
    HC.oPersist.WriteIniValue "FOR_B", CStr(a3)
    HC.oPersist.WriteIniValue "BAK_R", CStr(b1)
    HC.oPersist.WriteIniValue "BAK_G", CStr(b2)
    HC.oPersist.WriteIniValue "BAK_B", CStr(b3)
    HC.oPersist.WriteIniValue "FONT_S", CStr(F1)

End Sub


Public Sub readColorDat()

     Dim i As Long
     Dim j As Long
     Dim k As Long
     
     Dim tmptxt As String

     tmptxt = HC.oPersist.ReadIniValue("FOR_R")
     If tmptxt <> "" Then
        i = val(tmptxt)
     Else
        i = &HFF
     End If

     tmptxt = HC.oPersist.ReadIniValue("FOR_G")
     If tmptxt <> "" Then
        j = val(tmptxt)
     Else
        j = &H80
     End If

     tmptxt = HC.oPersist.ReadIniValue("FOR_B")
     If tmptxt <> "" Then
        k = val(tmptxt)
     Else
        k = &H0
     End If

    i = i And &HFF
    j = (j * 256) And &HFF00
    k = (k * 65536) And &HFF0000
    
    HC.HCMessage.ForeColor = i + j + k
'    HC.HCTextAlign.ForeColor = i + j + k
    
     tmptxt = HC.oPersist.ReadIniValue("BAK_R")
     If tmptxt <> "" Then
        i = val(tmptxt)
     Else
        i = &H80
     End If

     tmptxt = HC.oPersist.ReadIniValue("BAK_G")
     If tmptxt <> "" Then
        j = val(tmptxt)
     Else
        j = &H0
     End If

     tmptxt = HC.oPersist.ReadIniValue("BAK_B")
     If tmptxt <> "" Then
        k = val(tmptxt)
     Else
        k = &H0
     End If

    i = i And &HFF
    j = (j * 256) And &HFF00
    k = (k * 65536) And &HFF0000
    
    HC.HCMessage.BackColor = i + j + k
'    HC.HCTextAlign.BackColor = i + j + k
    
     tmptxt = HC.oPersist.ReadIniValue("FONT_S")
     If tmptxt <> "" Then
        i = val(tmptxt)
     Else
        i = 7
     End If
     
    HC.HCMessage.FontSize = i
'    HC.HCTextAlign.FontSize = i
     
 
End Sub


Public Sub readratebarstatePad()

     Dim tmptxt As String

     tmptxt = HC.oPersist.ReadIniValue("BAR03_1")
     If tmptxt <> "" Then
        Slewpad.VScroll1.Value = val(tmptxt)
     End If
     
     tmptxt = HC.oPersist.ReadIniValue("BAR03_2")
     If tmptxt <> "" Then
        Slewpad.VScroll2.Value = val(tmptxt)
     End If
   
End Sub

Public Sub readportrate()

     Dim raval As String
     Dim decval As String

     raval = HC.oPersist.ReadIniValue("AUTOGUIDER_RA")
     Select Case raval
        Case "x1.00":
            eqres = EQ_SetAutoguiderPortRate(0, 3)
            HC.RAGuideRateList.ListIndex = 3
        Case "x0.75":
            eqres = EQ_SetAutoguiderPortRate(0, 2)
            HC.RAGuideRateList.ListIndex = 2
        Case "x0.50":
            eqres = EQ_SetAutoguiderPortRate(0, 1)
            HC.RAGuideRateList.ListIndex = 1
        Case "x0.25"
            eqres = EQ_SetAutoguiderPortRate(0, 0)
            HC.RAGuideRateList.ListIndex = 0
        Case Else
            HC.RAGuideRateList.ListIndex = 4
    End Select
        
    decval = HC.oPersist.ReadIniValue("AUTOGUIDER_DEC")
    Select Case decval
        Case "x1.00":
            eqres = EQ_SetAutoguiderPortRate(1, 3)
             HC.DECGuideRateList.ListIndex = 3
        Case "x0.75":
            eqres = EQ_SetAutoguiderPortRate(1, 2)
            HC.DECGuideRateList.ListIndex = 2
        Case "x0.50":
            eqres = EQ_SetAutoguiderPortRate(1, 1)
            HC.DECGuideRateList.ListIndex = 1
        Case "x0.25"
            eqres = EQ_SetAutoguiderPortRate(1, 0)
            HC.DECGuideRateList.ListIndex = 0
        Case Else
            HC.DECGuideRateList.ListIndex = 4
    End Select
   
   
End Sub


Public Sub writeportrateRa(strRate As String)

    HC.oPersist.WriteIniValue "AUTOGUIDER_RA", strRate

End Sub
Public Sub writeportrateDec(strRate As String)

    HC.oPersist.WriteIniValue "AUTOGUIDER_DEC", strRate

End Sub

Public Sub writePulseguidepwidth()
    HC.oPersist.WriteIniValue "PULSEGUIDE_TIMER_INTERVAL", CStr(HC.PltimerHscroll.Value)
End Sub

Public Sub readPulseguidepwidth()

     Dim tmptxt As String

     tmptxt = HC.oPersist.ReadIniValue("PULSEGUIDE_TIMER_INTERVAL")
 
     If tmptxt = "" Then
        HC.PltimerHscroll.Value = 20
        HC.Label40.Caption = " 20"
        HC.Pulseguide_Timer.Interval = 20
        gpl_interval = 20
     Else
        gpl_interval = val(tmptxt)
        If gpl_interval < HC.PltimerHscroll.min Then
            gpl_interval = HC.PltimerHscroll.min
        Else
            If gpl_interval > HC.PltimerHscroll.max Then
                gpl_interval = HC.PltimerHscroll.max
            End If
        End If
        HC.PltimerHscroll.Value = gpl_interval
        HC.Label40.Caption = tmptxt
        HC.Pulseguide_Timer.Interval = gpl_interval
     End If
     
     tmptxt = HC.oPersist.ReadIniValue("DEC_BACKLASH")
     If tmptxt = "" Then
        gBacklashDec = 0
     Else
        gBacklashDec = val(tmptxt)
        If gBacklashDec > 200 Or gBacklashDec < 0 Then
            gBacklashDec = 0
        End If
     End If
     HC.HScrollBacklashDec.Value = gBacklashDec
     
End Sub
Public Sub writeRASyncCheckVal()
    HC.oPersist.WriteIniValue "AUTOSYNCRA", CStr(HC.CheckRASync.Value)
End Sub
Public Sub readRASyncCheckVal()
    
     Dim tmptxt As String
     
     tmptxt = HC.oPersist.ReadIniValue("AUTOSYNCRA")
     If tmptxt = "" Then
        HC.CheckRASync.Value = 1
        Call writeRASyncCheckVal
     Else
        HC.CheckRASync.Value = val(tmptxt)
     End If

End Sub

Public Sub writeDriftVal()
    HC.oPersist.WriteIniValue "RA_DRIFT_VAL", CStr(gDriftComp)
End Sub

Public Sub readDriftVal()
     Dim tmptxt As String
     tmptxt = HC.oPersist.ReadIniValue("RA_DRIFT_VAL")

     If tmptxt = "" Then
        gDriftComp = 0
        HC.DriftScroll.Value = 0
        HC.Driftlbl.Caption = "0"
     Else
        gDriftComp = val(tmptxt)
        HC.DriftScroll.Value = gDriftComp
        HC.Driftlbl.Caption = tmptxt
     End If
    Call EQSetOffsets

End Sub


Public Sub writeAxisRevRA()

    HC.oPersist.WriteIniValue "RA_REVERSE", CStr(HC.RA_inv.Value)

End Sub


Public Sub writeAxisRevDEC()

    HC.oPersist.WriteIniValue "DEC_REVERSE", CStr(HC.DEC_Inv.Value)

End Sub


Public Sub readDevelopmentOptions()
Dim tmp As String
Dim ver As Double

    If HC.oPersist.ReadIniValue("Advanced") = "1" Then
        gAdvanced = 1
        HC.Combo3PointAlgorithm.Visible = True
        HC.CheckRASync.Visible = True
        HC.Label35.Visible = True
        HC.CheckLocalPier.Visible = True
        HC.FrameAdvanced.Visible = True
        HC.FramePGAvanced.Visible = True
        HC.LabelSlewLimit.Visible = True
        HC.Label31.Visible = True
        HC.HScrollSlewLimit.Visible = True
    Else
        HC.oPersist.WriteIniValue "Advanced", "0"
        gAdvanced = 0
        HC.Combo3PointAlgorithm.Visible = False
        HC.CheckRASync.Visible = False
        HC.Label35.Visible = False
        HC.CheckLocalPier.Visible = False
        HC.FrameAdvanced.Visible = False
        HC.FramePGAvanced.Visible = False
        HC.LabelSlewLimit.Visible = False
        HC.Label31.Visible = False
        HC.HScrollSlewLimit.Visible = False
    End If
    
    If HC.oPersist.ReadIniValue("POLAR_ALIGNMENT") = "1" Then
        gShowPolarAlign = 1
        HC.puPolar.Visible = True
    Else
        HC.oPersist.WriteIniValue "POLAR_ALIGNMENT", "0"
        gShowPolarAlign = 0
        HC.puPolar.Visible = False
    End If
    
    If HC.oPersist.ReadIniValue("3POINT_ALGORITHM") = "1" Then
        g3PointAlgorithm = 1
    Else
        HC.oPersist.WriteIniValue "3POINT_ALGORITHM", "0"
        g3PointAlgorithm = 0
    End If
    HC.Combo3PointAlgorithm.ListIndex = g3PointAlgorithm

    tmp = HC.oPersist.ReadIniValue("MAX_GOTO_INTERATIONS")
    If tmp <> "" Then
        gMaxSlewCount = val(tmp)
    Else
        HC.oPersist.WriteIniValue "MAX_GOTO_INTERATIONS", "5"
        gMaxSlewCount = 5
    End If
    HC.HScrollSlewRetries.Value = gMaxSlewCount

    tmp = HC.oPersist.ReadIniValue("GOTO_RESOLUTION")
    If tmp <> "" Then
        gGotoResolution = val(tmp)
    Else
        HC.oPersist.WriteIniValue "GOTO_RESOLUTION", "20"
        gGotoResolution = 20
    End If
    HC.HScrollGotoRes.Value = gGotoResolution

    tmp = HC.oPersist.ReadIniValue("GOTO_RA_COMPENSATE")
    If tmp <> "" Then
        gRA_Compensate = val(tmp)
    Else
        HC.oPersist.WriteIniValue "GOTO_RA_COMPENSATE", "40"
        gRA_Compensate = 40
    End If
    HC.HScrollSlewAdjust.Value = gRA_Compensate
    
    tmp = HC.oPersist.ReadIniValue("COMMS_ERROR_STOP")
    If tmp <> "" Then
        gCommErrorStop = val(tmp)
    Else
        HC.oPersist.WriteIniValue "COMMS_ERROR_STOP", "0"
        gCommErrorStop = 0
    End If
    
End Sub


Public Sub readAscomCompatibiity()
Dim tmptxt As String

    On Error GoTo readerr1
    tmptxt = HC.oPersist.ReadIniValue("ASCOM_COMPAT_SLEWTRACKOFF")
    If tmptxt <> "" Then
        gAscomCompatibility.SlewWithTrackingOff = CBool(tmptxt)
    Else
        gAscomCompatibility.SlewWithTrackingOff = True
        HC.oPersist.WriteIniValue "ASCOM_COMPAT_SLEWTRACKOFF", CStr(gAscomCompatibility.SlewWithTrackingOff)
    End If

    tmptxt = HC.oPersist.ReadIniValue("ASCOM_COMPAT_PULSEGUIDE")
    If tmptxt <> "" Then
        gAscomCompatibility.AllowPulseGuide = CBool(tmptxt)
    Else
        gAscomCompatibility.AllowPulseGuide = True
        HC.oPersist.WriteIniValue "ASCOM_COMPAT_PULSEGUIDE", CStr(gAscomCompatibility.AllowPulseGuide)
    End If
    
    If gAscomCompatibility.AllowPulseGuide Then
        HC.Frame5.Visible = True
        HC.Frame6.Visible = False
    Else
        HC.Frame6.Visible = True
        HC.Frame5.Visible = False
    End If

    tmptxt = HC.oPersist.ReadIniValue("ASCOM_COMPAT_EXCEPTIONS")
    If tmptxt <> "" Then
        gAscomCompatibility.AllowExceptions = CBool(tmptxt)
    Else
        gAscomCompatibility.AllowExceptions = True
        HC.oPersist.WriteIniValue "ASCOM_COMPAT_EXCEPTIONS", CStr(gAscomCompatibility.AllowExceptions)
    End If
    
    tmptxt = HC.oPersist.ReadIniValue("ASCOM_COMPAT_PG_EXCEPTIONS")
    If tmptxt <> "" Then
        gAscomCompatibility.AllowPulseGuideExceptions = CBool(tmptxt)
    Else
        gAscomCompatibility.AllowPulseGuideExceptions = True
        HC.oPersist.WriteIniValue "ASCOM_COMPAT_PG_EXCEPTIONS", CStr(gAscomCompatibility.AllowPulseGuideExceptions)
    End If
    
    tmptxt = HC.oPersist.ReadIniValue("ASCOM_COMPAT_BLOCK_PARK")
    If tmptxt <> "" Then
        gAscomCompatibility.BlockPark = CBool(tmptxt)
    Else
        gAscomCompatibility.BlockPark = False
        HC.oPersist.WriteIniValue "ASCOM_COMPAT_BLOCK_PARK", CStr(gAscomCompatibility.BlockPark)
    End If
    
    tmptxt = HC.oPersist.ReadIniValue("ASCOM_COMPAT_SITEWRITES")
    If tmptxt <> "" Then
        gAscomCompatibility.AllowSiteWrites = CBool(tmptxt)
    Else
        gAscomCompatibility.AllowSiteWrites = False
        HC.oPersist.WriteIniValue "ASCOM_COMPAT_SITEWRITES", CStr(gAscomCompatibility.AllowSiteWrites)
    End If
    
    tmptxt = HC.oPersist.ReadIniValue("ASCOM_COMPAT_EPOCH")
    If tmptxt <> "" Then
        gAscomCompatibility.Epoch = val(tmptxt)
    Else
        gAscomCompatibility.Epoch = 0
        HC.oPersist.WriteIniValue "ASCOM_COMPAT_EPOCH", CStr(gAscomCompatibility.Epoch)
    End If
    
    Exit Sub


readerr1:
        gAscomCompatibility.SlewWithTrackingOff = True
        gAscomCompatibility.AllowPulseGuide = True
        gAscomCompatibility.AllowExceptions = True
        gAscomCompatibility.AllowPulseGuideExceptions = True
        gAscomCompatibility.Epoch = 0
        gAscomCompatibility.BlockPark = False
        WriteAscomCompatibiity
End Sub
Public Sub WriteAscomCompatibiity()
    HC.oPersist.WriteIniValue "ASCOM_COMPAT_SLEWTRACKOFF", CStr(gAscomCompatibility.SlewWithTrackingOff)
    HC.oPersist.WriteIniValue "ASCOM_COMPAT_PULSEGUIDE", CStr(gAscomCompatibility.AllowPulseGuide)
    HC.oPersist.WriteIniValue "ASCOM_COMPAT_EXCEPTIONS", CStr(gAscomCompatibility.AllowExceptions)
    HC.oPersist.WriteIniValue "ASCOM_COMPAT_PG_EXCEPTIONS", CStr(gAscomCompatibility.AllowPulseGuideExceptions)
    HC.oPersist.WriteIniValue "ASCOM_COMPAT_BLOCK_PARK", CStr(gAscomCompatibility.BlockPark)
    HC.oPersist.WriteIniValue "ASCOM_COMPAT_SITEWRITES", CStr(gAscomCompatibility.AllowSiteWrites)
    HC.oPersist.WriteIniValue "ASCOM_COMPAT_EPOCH", CStr(gAscomCompatibility.Epoch)
    
End Sub

Public Sub readAutoFlipData()
Dim tmptxt As String

    On Error GoTo readerr1
    tmptxt = HC.oPersist.ReadIniValue("FLIP_AUTO_ALLOWED")
    If tmptxt <> "" Then
        gAutoFlipAllowed = CBool(tmptxt)
    Else
        ' default to allow slews when not tracking - not ASCOM compliant but is CDC compliant!

        gAutoFlipAllowed = False
    End If
    
    tmptxt = HC.oPersist.ReadIniValue("FLIP_AUTO_ENABLED")
    If tmptxt <> "" Then
        gAutoFlipEnabled = CBool(tmptxt)
    Else
        ' default to allow slews when not tracking - not ASCOM compliant but is CDC compliant!
        gAutoFlipEnabled = False
    End If
    
    Call WriteAutoFlipData
    Exit Sub
    
readerr1:
    gAutoFlipEnabled = False
    gAutoFlipAllowed = False
    Call WriteAutoFlipData
    
End Sub

Public Sub WriteAutoFlipData()
    HC.oPersist.WriteIniValue "FLIP_AUTO_ALLOWED", CStr(gAutoFlipAllowed)
    HC.oPersist.WriteIniValue "FLIP_AUTO_ENABLED", CStr(gAutoFlipEnabled)
End Sub

Public Sub readAxisRev()

     Dim tmptxt As String

     tmptxt = HC.oPersist.ReadIniValue("RA_REVERSE")
     If tmptxt <> "" Then
        HC.RA_inv.Value = val(tmptxt)
     End If
     
     tmptxt = HC.oPersist.ReadIniValue("DEC_REVERSE")
     If tmptxt <> "" Then
        HC.DEC_Inv.Value = val(tmptxt)
     End If

End Sub

Public Sub writePresetSlewRates()
Dim tmptxt As String
Dim key As String
Dim valstr As String
Dim Ini As String
Dim Count As Integer

    ' set up a file path for the align.ini file
    Ini = HC.oPersist.GetIniPath & "\EQMOD.ini"

    key = "[slewrates]"

    Call HC.oPersist.WriteIniValueEx("COUNT", CStr(gPresetSlewRatesCount), key, Ini)
    
    For Count = 1 To gPresetSlewRatesCount
        valstr = "RATE_" & CStr(Count)
        Call HC.oPersist.WriteIniValueEx(valstr, CStr(gPresetSlewRates(Count)), key, Ini)
    Next Count

    For Count = 1 To 4
        valstr = "RATEBTN_" & CStr(Count)
        Call HC.oPersist.WriteIniValueEx(valstr, CStr(gRateButtons(Count)), key, Ini)
    Next Count

End Sub
Public Sub readPresetSlewRates()

Dim tmptxt As String
Dim key As String
Dim valstr As String
Dim Ini As String
Dim Count As Integer
Dim DefaultRates(1 To 10) As Integer

    DefaultRates(1) = 1
    DefaultRates(2) = 8
    DefaultRates(3) = 64
    DefaultRates(4) = 800
    DefaultRates(5) = 0
    DefaultRates(6) = 0
    DefaultRates(7) = 0
    DefaultRates(8) = 0
    DefaultRates(9) = 0
    DefaultRates(10) = 0
    
    HC.PresetRateCombo.Clear
    HC.PresetRate2Combo.Clear
    
    ' set up a file path for the align.ini file
    Ini = HC.oPersist.GetIniPath & "\EQMOD.ini"

    key = "[slewrates]"

    ' read preset count
    tmptxt = HC.oPersist.ReadIniValueEx("COUNT", key, Ini)
    If tmptxt <> "" Then
        gPresetSlewRatesCount = val(tmptxt)
        If gPresetSlewRatesCount > 10 Then
            gPresetSlewRatesCount = 10
            Call HC.oPersist.WriteIniValueEx("COUNT", "10", key, Ini)
        End If
    Else
        gPresetSlewRatesCount = 4
        Call HC.oPersist.WriteIniValueEx("COUNT", "4", key, Ini)
    End If
    
    For Count = 1 To gPresetSlewRatesCount
        valstr = "RATE_" & CStr(Count)
        tmptxt = HC.oPersist.ReadIniValueEx(valstr, key, Ini)
        If tmptxt <> "" Then
            gPresetSlewRates(Count) = val(tmptxt)
        Else
            gPresetSlewRates(Count) = DefaultRates(Count)
            Call HC.oPersist.WriteIniValueEx(valstr, CStr(gPresetSlewRates(Count)), key, Ini)
        End If
        ' add preset to combo
        HC.PresetRateCombo.AddItem (CStr(Count))
        HC.PresetRate2Combo.AddItem (CStr(Count))
    Next Count
    
    tmptxt = HC.oPersist.ReadIniValueEx("InitalPreset", key, Ini)
    If tmptxt <> "" Then
        gCurrentRatePreset = val(tmptxt)
        If gCurrentRatePreset > 10 Then
            gCurrentRatePreset = 1
            Call HC.oPersist.WriteIniValueEx("InitalPreset", "1", key, Ini)
        End If
    Else
        gCurrentRatePreset = 1
        Call HC.oPersist.WriteIniValueEx("InitalPreset", "1", key, Ini)
    End If

    HC.PresetRateCombo.ListIndex = gCurrentRatePreset - 1
    HC.PresetRate2Combo.ListIndex = gCurrentRatePreset - 1
    
    
    For Count = 1 To 4
        valstr = "RATEBTN_" & CStr(Count)
        tmptxt = HC.oPersist.ReadIniValueEx(valstr, key, Ini)
        If tmptxt <> "" Then
            gRateButtons(Count) = val(tmptxt)
        Else
            gRateButtons(Count) = Count
            Call HC.oPersist.WriteIniValueEx(valstr, CStr(Count), key, Ini)
        End If
    Next Count
    
    
End Sub

Public Sub readPoleStar()
     Dim tmptxt As String
     Dim RA As Double
     Dim DEC As Double
     Dim RA2 As Double
     Dim DEC2 As Double
     Dim epochnow As Double

    'J2000 = RA: 02h31m50.209s DE:+89°15'50.86"
    
    tmptxt = HC.oPersist.ReadIniValue("PoleStarId")
    If tmptxt <> "" Then
        gPoleStarIdx = val(tmptxt)
        If gPoleStarIdx >= HC.ComboPoleStar.ListCount Then
            gPoleStarIdx = 0
        End If
    Else
        gPoleStarIdx = 0
    End If
    
    Select Case gPoleStarIdx
        Case 0
            ' polaris
            RA = 2.53019444444444
            DEC = 89.2641666666667
        Case 1
            ' Chi Oct
            RA = 18.91286139
            DEC = -87.60628056
        Case 2
            'Tau Oct
            RA = 23.46775278
            DEC = -87.48219167
        Case 3
            ' Sigma Oct
            RA = 21.146498333333
            DEC = -88.9547972222
            
           
        Case Else
            tmptxt = HC.oPersist.ReadIniValue("POLE_STAR_J2000RA")
            If tmptxt <> "" Then
               RA = CDbl(tmptxt)
            Else
               RA = 2.53061361
               HC.oPersist.WriteIniValue "POLE_STAR_J2000RA", CStr(RA)
            End If
            
            tmptxt = HC.oPersist.ReadIniValue("POLE_STAR_J2000DEC")
            If tmptxt <> "" Then
               DEC = CDbl(tmptxt)
            Else
               DEC = 89.2641278
               HC.oPersist.WriteIniValue "POLE_STAR_J2000DEC", CStr(DEC)
            End If
    End Select
    
'    HC.oPersist.WriteIniValue "POLE_STAR_J2000RA", CStr(RA)
'    HC.oPersist.WriteIniValue "POLE_STAR_J2000DEC", CStr(DEC)
    
    tmptxt = HC.oPersist.ReadIniValue("PolarReticuleEpoch")
    If tmptxt <> "" Then
        gPolarReticuleEpoch = val(tmptxt)
    Else
        gPolarReticuleEpoch = 2000
        HC.oPersist.WriteIniValue "PolarReticuleEpoch", "2000"
    End If
    
    HC.ComboPoleStar.ListIndex = gPoleStarIdx
    
    gPoleStarRaJ2000 = RA
    gPoleStarDecJ2000 = DEC
    RA2 = RA
    DEC2 = DEC
    epochnow = 2000 + (now_mjd() - J2000) / 365.25
    Call Precess(RA2, DEC2, 2000, epochnow)
    gPoleStarRa = RA2
    gPoleStarDec = DEC2
    Call Precess(RA, DEC, 2000, gPolarReticuleEpoch)
    gPoleStarReticuleDec = DEC
    
    
End Sub
Public Sub writePoleStar()
    Dim tmptxt As String
    Dim RA As Double
    Dim DEC As Double
    Dim epochnow As Double
    
    HC.oPersist.WriteIniValue "PoleStarId", CStr(gPoleStarIdx)
    Select Case gPoleStarIdx
        Case 0
            ' polaris
            RA = 2.53019444444444 ' 2.53061361 ' 2.53019444444444
            DEC = 89.2641666666667 ' 89.2641278  ' 89.2641666666667
        
        Case 1
            ' Chi Oct
            RA = 18.91286139
            DEC = -87.60628056
        
        Case 2
            'Tau Oct
            RA = 23.46775278
            DEC = -87.48219167
        
        Case Else
            tmptxt = HC.oPersist.ReadIniValue("POLE_STAR_J2000RA")
            If tmptxt <> "" Then
               RA = CDbl(tmptxt)
            Else
               RA = 2.53061361
               HC.oPersist.WriteIniValue "POLE_STAR_J2000RA", CStr(RA)
            End If
            
            tmptxt = HC.oPersist.ReadIniValue("POLE_STAR_J2000DEC")
            If tmptxt <> "" Then
               DEC = CDbl(tmptxt)
            Else
               DEC = 89.2641278
               HC.oPersist.WriteIniValue "POLE_STAR_J2000DEC", CStr(DEC)
            End If
    End Select
    
'    HC.oPersist.WriteIniValue "POLE_STAR_J2000RA", CStr(RA)
'    HC.oPersist.WriteIniValue "POLE_STAR_J2000DEC", CStr(DEC)
    
    epochnow = 2000 + (now_mjd() - J2000) / 365.25
    Call Precess(RA, DEC, 2000, epochnow)
    gPoleStarRa = RA
    gPoleStarDec = DEC
End Sub


Public Function GetEmulRA() As Double

Dim emulinc As Double

    ' Compute for elapsed Time
                    
    If gTrackingStatus = 1 Then

        gCurrent_time = EQnow_lst_norange()
        If gLast_time = 0 Then gCurrent_time = 0.000002
        If gEmulRA_Init = 0 Then gEmulRA_Init = gEmulRA
        
        If gLast_time > gCurrent_time Then      ' Counter wrap around ?
            gLast_time = EQnow_lst_norange()
            gCurrent_time = gLast_time
            gEmulRA_Init = gEmulRA
        End If
                    
        ' Compute Elapste stepper count based on Elapsed Local Sidreal time (PC time)
        
        emulinc = gEMUL_RATE2 * (gCurrent_time - gLast_time)
'       If gRA_LastRate = 0 Then
'          emulinc = gEMUL_RATE2 * (gCurrent_time - gLast_time)
'       Else  ' PEC tracking
'          emulinc = (gRightAscensionRate / gARCSECSTEP) * (gCurrent_time - gLast_time)
'          emulinc = (gCurrent_time - gLast_time) * gTot_RA / (1296000 / gRightAscensionRate)
'       End If
        
        If gHemisphere = 0 Then
            GetEmulRA = gEmulRA_Init + emulinc
        Else
            GetEmulRA = gEmulRA_Init - emulinc
        End If

    Else
        GetEmulRA = gEmulRA
    End If

End Function


Public Function GetEmulRA_EQ() As Double

Dim emulinc As Double
Dim tmpEmulRA As Double
Dim tmpgRA_Hours As Double

Dim tmpgRA_Encoder As Double
Dim tmpgDec_Encoder As Double

Dim tRA As Double
Dim tmpgDec_DegNoAdjust As Double

Dim tmpcoord As Coordt

        'Compute for elapsed Time
            
    If gTrackingStatus = 1 Then
        
         gCurrent_time = EQnow_lst_norange()
         If gLast_time = 0 Then gCurrent_time = 0.000002
         If gEmulRA_Init = 0 Then gEmulRA_Init = gEmulRA
         
         If gLast_time > gCurrent_time Then      ' Counter wrap around ?
             gLast_time = EQnow_lst_norange()
             gCurrent_time = gLast_time
             gEmulRA_Init = gEmulRA
         End If
         
         ' Compute Elapste stepper count based on Elapsed Local Sidreal time (PC time)
         
         emulinc = gEMUL_RATE2 * (gCurrent_time - gLast_time)
'        If gRA_LastRate = 0 Then
'           emulinc = gEMUL_RATE2 * (gCurrent_time - gLast_time)
'        Else  ' PEC tracking
'           emulinc = (gRightAscensionRate / gARCSECSTEP) * (gCurrent_time - gLast_time)
'           emulinc = (gCurrent_time - gLast_time) * gTot_RA / (1296000 / gRightAscensionRate)
'        End If
         
         If gHemisphere = 0 Then
             tmpEmulRA = gEmulRA_Init + emulinc
         Else
             tmpEmulRA = gEmulRA_Init - emulinc
         End If
    Else
        tmpEmulRA = gEmulRA
    End If

    If gThreeStarEnable = False Then
        tmpgRA_Encoder = Delta_RA_Map(tmpEmulRA)
        tmpgDec_Encoder = Delta_DEC_Map(gEmulDEC)
    Else
    
        Select Case gAlignmentMode
        
            Case 2
                ' nearest
                tmpcoord = DeltaSync_Matrix_Map(tmpEmulRA, gEmulDEC)
                tmpgRA_Encoder = tmpcoord.x
                tmpgDec_Encoder = tmpcoord.Y
    
            Case 1
                ' n-star+nearest
                tmpcoord = Delta_Matrix_Reverse_Map(tmpEmulRA, gEmulDEC)
                tmpgRA_Encoder = tmpcoord.x
                tmpgDec_Encoder = tmpcoord.Y
                
            Case Else
                tmpcoord = Delta_Matrix_Reverse_Map(tmpEmulRA, gEmulDEC)
                tmpgRA_Encoder = tmpcoord.x
                tmpgDec_Encoder = tmpcoord.Y
                
                If tmpcoord.F = 0 Then
                
                    tmpcoord = DeltaSync_Matrix_Map(tmpEmulRA, gEmulDEC)
                    tmpgRA_Encoder = tmpcoord.x
                    tmpgDec_Encoder = tmpcoord.Y
    
                End If
                        
        End Select
       
    End If
            
    If (tmpgRA_Encoder < &H1000000) Then tmpgRA_Hours = Get_EncoderHours(gRAEncoder_Zero_pos, tmpgRA_Encoder, gTot_RA, gHemisphere)

    tRA = EQnow_lst(gLongitude * DEG_RAD) + tmpgRA_Hours
    
    tmpgDec_DegNoAdjust = Get_EncoderDegrees(gDECEncoder_Zero_pos, tmpgDec_Encoder, gTot_DEC, gHemisphere)
          
    If gHemisphere = 0 Then
        If (tmpgDec_DegNoAdjust > 90) And (tmpgDec_DegNoAdjust <= 270) Then tRA = tRA - 12
    Else
        If (tmpgDec_DegNoAdjust <= 90) Or (tmpgDec_DegNoAdjust > 270) Then tRA = tRA + 12
    End If
    
    GetEmulRA_EQ = Range24(tRA)

End Function

' checks that a guiding rate commanded is in range
Public Function RateIsInRange(ByVal rate As Double, ByVal Rates As Rates) As Boolean
    Dim i As Integer
    Dim r As rate
    
    For i = 1 To Rates.Count
        Set r = Rates.Item(i)
        If Abs(rate) > r.Maximum Or Abs(rate) < r.Minimum Then
            RateIsInRange = False
            Exit Function
        End If
    Next i
    RateIsInRange = True
End Function

Public Function EQGP(ByVal motor_id As Long, ByVal p_id As Long) As Long
Dim ret As Long

    Select Case p_id
        Case 10006
            ' get worm steps from ini file this way we can easilly simulate heq5
            If gCustomMount = 1 Then
                Select Case motor_id
                    Case 0
                        ret = gCustomRAWormSteps
                    Case 1
                        ret = gCustomDECWormSteps
                    Case Else
                        ret = EQ_GP(motor_id, p_id)
                End Select
            Else
                ret = EQ_GP(motor_id, p_id)
            End If
        Case Else
            ret = EQ_GP(motor_id, p_id)
    End Select
    EQGP = ret
End Function

Public Sub ReadFormPosition()
Dim tmptxt As String
Dim tmp As Single
Dim DesktopLeft As Long
Dim DesktopTop As Long
Dim DesktopWidth As Long
Dim DesktopHeight As Long
Dim DesktopRight As Long
Dim DesktopBottom As Long

    If GetSystemMetrics(SM_CMONITORS) = 0 Then
        'No multi monitor
        DesktopLeft = 0
        DesktopRight = Screen.width
        DesktopTop = 0
        DesktopBottom = Screen.Height
    Else
        DesktopLeft = GetSystemMetrics(SM_XVIRTUALSCREEN)
        DesktopLeft = DesktopLeft * Screen.TwipsPerPixelX
        DesktopTop = GetSystemMetrics(SM_YVIRTUALSCREEN)
        DesktopTop = DesktopTop * Screen.TwipsPerPixelY
        DesktopWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN) * Screen.TwipsPerPixelX
        DesktopHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN) * Screen.TwipsPerPixelY
        DesktopRight = DesktopLeft + DesktopWidth
        DesktopBottom = DesktopTop + DesktopHeight
    End If
    
    tmptxt = HC.oPersist.ReadIniValue("form_height")
    If tmptxt = "" Then
        Call HC.oPersist.WriteIniValue("form_height", HC.Height)
    Else
        HC.Height = val(tmptxt)
    End If
    
    tmptxt = HC.oPersist.ReadIniValue("form_top")
    If tmptxt = "" Then
        Call HC.oPersist.WriteIniValue("form_top", 0)
        HC.Top = 0
    Else
        tmp = val(tmptxt)
        If tmp < DesktopTop Then tmp = DesktopTop
        If tmp > DesktopBottom - HC.Height Then tmp = DesktopBottom - HC.Height
        HC.Top = tmp
    End If
    
    tmptxt = HC.oPersist.ReadIniValue("form_left")
    If tmptxt = "" Then
        Call HC.oPersist.WriteIniValue("form_left", 0)
        HC.Left = 0
    Else
        tmp = val(tmptxt)
        If tmp < DesktopLeft Then tmp = DesktopLeft
        If tmp > DesktopRight - HC.width Then tmp = DesktopRight - HC.width
        HC.Left = tmp
    End If

End Sub

Public Sub WriteFormPosition()
    Call HC.oPersist.WriteIniValue("form_height", HC.Height)
    Call HC.oPersist.WriteIniValue("form_top", HC.Top)
    Call HC.oPersist.WriteIniValue("form_left", HC.Left)
End Sub

Public Sub ResetFormPosition()
Dim tmptxt As String
    tmptxt = HC.oPersist.ReadIniValue("form_dleft")
    HC.Left = val(tmptxt)
    HC.Top = val(HC.oPersist.ReadIniValue("form_dtop"))
    HC.Height = val(HC.oPersist.ReadIniValue("form_dheight"))
    Call WriteFormPosition
End Sub

Public Sub GetDllVer()
Dim dllver As Long
Dim tmpstr As String
    dllver = EQ_DriverVersion()
    
    tmpstr = Hex$((dllver And &HF000) / 4096 And &HF) + Hex$((dllver And &HF00) / 256 And &HF)
    tmpstr = tmpstr & "." & Hex$((dllver And &HF0) / 16 And &HF) + Hex$(dllver And &HF)
'    tmpstr = Hex$((dllver And &HF00000) / 1048576 And &HF) + Hex$((dllver And &HF0000) / 65536 And &HF)
    gDllVer = val(tmpstr)

End Sub

' Interceptor functions for different mount types

Public Function EQGetMotorValues(ByVal motor_id As Long) As Long
Dim ret As Long

    On Error GoTo errhandle

    If gDllVer < 3.04 Then
        ret = EQ_GetMotorValues(motor_id)
    Else
        Select Case gCoordType
    
            Case CT_RADEC
                Select Case motor_id
                    Case 0
                        ' RA
                    Case 1
                        ' DEC
                End Select
        
            Case CT_AZALT
                Select Case motor_id
                    Case 0
                        ' Az
                    Case 1
                        ' Alt
                End Select
        
            Case Else
                ' microstep
                    ret = EQ_GetMotorValues(motor_id)
'                    ret = EQ_GetMotorValues2(motor_id, "", "")
        
        End Select
    End If
    
    EQGetMotorValues = ret
    Exit Function

errhandle:
    EQGetMotorValues = EQ_INVALID

End Function


Public Function EQSetMotorValues(ByVal motor_id As Long, motor_val As Long) As Long

Dim ret As Long

    On Error GoTo errhandle
    
    If gDllVer < 3.04 Then
        ret = EQ_SetMotorValues(motor_id, motor_val)
    Else
        Select Case gCoordType

            Case CT_RADEC
                Select Case motor_id
                    Case 0
                        ' RA
                    Case 1
                        ' DEC
                    End Select

            Case CT_AZALT
                Select Case motor_id
                    Case 0
                        ' Az
                    Case 1
                        ' Alt
                    End Select

                Case Else
            ' microstep
                ret = EQ_SetMotorValues(motor_id, motor_val)
'                ret = EQ_SetMotorValues2(motor_id, motor_val, "", "")
    
        End Select

    End If
    EQSetMotorValues = ret
    Exit Function

errhandle:
    EQSetMotorValues = EQ_INVALID


End Function

Public Function EQStartMoveMotor(ByVal motor_id As Long, ByVal hemisphere As Long, ByVal direction As Long, ByVal Steps As Long, ByVal stepslowdown As Long) As Long

Dim ret As Long

    Select Case gCoordType

        Case CT_RADEC
            Select Case motor_id
                Case 0
                    ' RA
                Case 1
                    ' DEC
            End Select
    
        Case CT_AZALT
            Select Case motor_id
                Case 0
                    ' Az
                Case 1
                    ' Alt
            End Select
    
        Case Else
            ' microstep
            ret = EQ_StartMoveMotor(motor_id, hemisphere, direction, Steps, stepslowdown)
    
    End Select

    EQStartMoveMotor = ret

End Function

Public Sub EQSetOffsets()

     If gCustomMount = 0 Then
        ' apply drift compenstation only to standard mounts
        eqres = EQ_SetOffset(0, gDriftComp * -1)
        eqres = EQ_SetOffset(1, 0)
     Else
        ' for customised mounts apply tracking offsets
        eqres = EQ_SetOffset(0, (gCustomTrackingOffsetRA + gDriftComp) * -1)
        eqres = EQ_SetOffset(1, (gCustomTrackingOffsetDEC) * -1)
     End If
     
End Sub

Public Function StripPath(str As String) As String
Dim i As Integer
    i = InStrRev(str, "\")
    StripPath = Right$(str, Len(str) - i)
End Function

Public Function GetPath(str As String) As String
Dim i As Integer
    i = InStrRev(str, "\")
    GetPath = Left$(str, i)
End Function
