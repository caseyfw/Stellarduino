Attribute VB_Name = "Sounds"
Option Explicit
Public Type EQMOD_SOUNDS
    mode As Integer
    PositionBeep As Boolean
    ButtonClick As Boolean
    RateClick As Boolean
    GotoClick As Boolean
    GotoStartClick As Boolean
    ParkClick As Boolean
    ParkedClick As Boolean
    Stopclick As Boolean
    Unparkclick As Boolean
    FlipWarning As Boolean
    TrackClick As Boolean
    AlignClick As Boolean
    PolarClick As Boolean
    DMSClick As Boolean
    GPLClick As Boolean
    MonitorClick As Boolean
    ReverseClick As Boolean
    BeepWav As String
    ClickWav As String
    AlarmWav As String
    RateWav(1 To 10) As String
    SyncWav As String
    GotoWav As String
    GotoStartWav As String
    ParkWav As String
    ParkedWav As String
    StopWav As String
    Unparkwav As String
    SiderealWav As String
    LunarWav As String
    SolarWav As String
    CustomWav As String
    AcceptWav As String
    CancelWav As String
    EndWav As String
    PHomeWav As String
    PAlignwav As String
    PAlignedwav As String
    DMSwav As String
    DMS2wav As String
    GPLOnwav As String
    GPLOffwav As String
    MonitorOnwav As String
    MonitorOffwav As String
    RAReverseOnwav As String
    RaReverseOffwav As String
    DecReverseOnwav As String
    DecReverseOffwav As String
End Type

Public EQSounds As EQMOD_SOUNDS

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Const SND_ASYNC = &H1         '  play asynchronously
Private Const SND_SYNC = &H0

Public Declare Function BeepAPI Lib "kernel32" Alias "Beep" (ByVal dwFrequency _
    As Long, ByVal dwMilliseconds As Long) As Long

Public Sub EQ_Beep(BeepType As Integer)
    
    On Error Resume Next
    
    Select Case BeepType
        ' Beep
        Case 0
            If EQSounds.PositionBeep = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 600, 100
                    Case 1
                    ' play asynchronously
                    Call sndPlaySound(EQSounds.BeepWav, SND_ASYNC)
                End Select
            End If
            
        ' Click
        Case 1
            If EQSounds.ButtonClick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 100, 1
                    Case 1
                    ' play asynchronously
                    Call sndPlaySound(EQSounds.ClickWav, SND_ASYNC)
                End Select
            End If
            
        ' Alarm
        Case 2
            If EQSounds.FlipWarning = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 200, 500
                        BeepAPI 100, 500
                        BeepAPI 200, 500
                        BeepAPI 100, 500
                        BeepAPI 200, 500
                        BeepAPI 100, 500
                        BeepAPI 200, 500
                        BeepAPI 100, 500
                        BeepAPI 200, 500
                        BeepAPI 100, 500
                    Case 1
                        ' play asynchronously
                        Call sndPlaySound(EQSounds.AlarmWav, SND_ASYNC)
                End Select
            End If
        
        ' Beep - always sounds
        Case 3
            Select Case EQSounds.mode
                Case 0
                    BeepAPI 600, 100
                Case 1
                ' play asynchronously
                Call sndPlaySound(EQSounds.BeepWav, SND_ASYNC)
            End Select
            
        ' Sync
        Case 4
            If EQSounds.AlignClick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 600, 100
                    Case 1
                    ' play asynchronously
                    Call sndPlaySound(EQSounds.SyncWav, SND_ASYNC)
                End Select
            End If
            
        ' Park
        Case 5
            If EQSounds.ParkClick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 600, 100
                    Case 1
                    ' play asynchronously
                    Call sndPlaySound(EQSounds.ParkWav, SND_ASYNC)
                End Select
            End If
            
        ' Goto
        Case 6
            If EQSounds.GotoClick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 600, 100
                    Case 1
                    ' play asynchronously
                    Call sndPlaySound(EQSounds.GotoWav, SND_ASYNC)
                End Select
            End If
            
        ' Stop
        Case 7
            If EQSounds.Stopclick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 600, 100
                    Case 1
                    ' play asynchronously
                    Call sndPlaySound(EQSounds.StopWav, SND_ASYNC)
                End Select
            End If
        
        ' Parked
        Case 8
            If EQSounds.ParkedClick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 600, 100
                    Case 1
                    ' play asynchronously
                    Call sndPlaySound(EQSounds.ParkedWav, SND_ASYNC)
                End Select
            End If
            
        ' Unpark
        Case 9
            If EQSounds.Unparkclick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 600, 100
                    Case 1
                    ' play asynchronously
                    Call sndPlaySound(EQSounds.Unparkwav, SND_ASYNC)
                End Select
            End If
            
        ' Sidereal
        Case 10
            If EQSounds.TrackClick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 600, 100
                    Case 1
                    ' play asynchronously
                    Call sndPlaySound(EQSounds.SiderealWav, SND_ASYNC)
                End Select
            End If
            
        ' lunar
        Case 11
            If EQSounds.TrackClick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 600, 100
                    Case 1
                    ' play asynchronously
                    Call sndPlaySound(EQSounds.LunarWav, SND_ASYNC)
                End Select
            End If
            
        ' solar
        Case 12
            If EQSounds.TrackClick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 600, 100
                    Case 1
                    ' play asynchronously
                    Call sndPlaySound(EQSounds.SolarWav, SND_ASYNC)
                End Select
            End If
            
        ' custom
        Case 13
            If EQSounds.TrackClick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 600, 100
                    Case 1
                    ' play asynchronously
                    Call sndPlaySound(EQSounds.CustomWav, SND_ASYNC)
                End Select
            End If
            
        ' Goto
        Case 20
            If EQSounds.GotoStartClick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 600, 100
                    Case 1
                    ' play asynchronously
                    Call sndPlaySound(EQSounds.GotoStartWav, SND_ASYNC)
                End Select
            End If
        
        ' Accept
        Case 21
            If EQSounds.AlignClick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 600, 100
                    Case 1
                    ' play asynchronously
                    Call sndPlaySound(EQSounds.AcceptWav, SND_ASYNC)
                End Select
            End If
            
        ' Cancel
        Case 22
            If EQSounds.AlignClick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 600, 100
                    Case 1
                    ' play asynchronously
                    Call sndPlaySound(EQSounds.CancelWav, SND_ASYNC)
                End Select
            End If
            
        ' End
        Case 23
            If EQSounds.AlignClick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 600, 100
                    Case 1
                    ' play asynchronously
                    Call sndPlaySound(EQSounds.EndWav, SND_ASYNC)
                End Select
            End If
            
        ' Polar Home
        Case 24
            If EQSounds.PolarClick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 600, 100
                    Case 1
                    ' play asynchronously
                    Call sndPlaySound(EQSounds.PHomeWav, SND_ASYNC)
                End Select
            End If
            
        ' Polar Aligning
        Case 25
            If EQSounds.PolarClick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 600, 100
                    Case 1
                    ' play asynchronously
                    Call sndPlaySound(EQSounds.PAlignwav, SND_ASYNC)
                End Select
            End If
            
        ' Polar Aligned
        Case 26
            If EQSounds.PolarClick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 600, 100
                    Case 1
                    ' play asynchronously
                    Call sndPlaySound(EQSounds.PAlignedwav, SND_ASYNC)
                End Select
            End If
            
        Case 30
            ' end beep
            If EQSounds.ButtonClick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 600, 100
                    Case 1
                    ' play asynchronously
                End Select
            End If
            
        Case 31
            ' Dead mans switch armed
            If EQSounds.DMSClick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 600, 100
                    Case 1
                        Call sndPlaySound(EQSounds.DMSwav, SND_ASYNC)
                End Select
            End If
            
        Case 32
            ' Dead mans switch armed
            If EQSounds.DMSClick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 600, 100
                    Case 1
                        Call sndPlaySound(EQSounds.DMS2wav, SND_ASYNC)
                End Select
            End If
            
        Case 33
            ' GamePad Lock on
            If EQSounds.GPLClick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 600, 100
                    Case 1
                        Call sndPlaySound(EQSounds.GPLOnwav, SND_ASYNC)
                End Select
            End If
            
        Case 34
            ' GamePad Lock off
            If EQSounds.GPLClick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 600, 100
                    Case 1
                        Call sndPlaySound(EQSounds.GPLOffwav, SND_ASYNC)
                End Select
            End If
            
        Case 35
            ' Monitor On
            If EQSounds.MonitorClick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 600, 100
                    Case 1
                        Call sndPlaySound(EQSounds.MonitorOnwav, SND_ASYNC)
                End Select
            End If
            
        Case 36
            ' Monitor Off
            If EQSounds.MonitorClick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 600, 100
                    Case 1
                        Call sndPlaySound(EQSounds.MonitorOffwav, SND_ASYNC)
                End Select
            End If
            
        Case 40
            ' RaReverseOn
            If EQSounds.ReverseClick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 600, 100
                    Case 1
                        Call sndPlaySound(EQSounds.RAReverseOnwav, SND_ASYNC)
                End Select
            End If
            
        Case 41
            ' RaReverseOff
            If EQSounds.ReverseClick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 600, 100
                    Case 1
                        Call sndPlaySound(EQSounds.RaReverseOffwav, SND_ASYNC)
                End Select
            End If
            
        Case 42
            ' DecReverseOn
            If EQSounds.ReverseClick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 600, 100
                    Case 1
                        Call sndPlaySound(EQSounds.DecReverseOnwav, SND_ASYNC)
                End Select
            End If
            
        Case 43
            ' DecReverseOff
            If EQSounds.ReverseClick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 600, 100
                    Case 1
                        Call sndPlaySound(EQSounds.DecReverseOffwav, SND_ASYNC)
                End Select
            End If
            
            
        ' rate sounds
        Case 101, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110
            If EQSounds.RateClick = True Then
                Select Case EQSounds.mode
                    Case 0
                        BeepAPI 100, 1
                    Case 1
                    ' play asynchronously
                    Call sndPlaySound(EQSounds.RateWav(BeepType - 100), SND_ASYNC)
                End Select
            End If
        
    End Select
End Sub

Public Sub writeBeep()
Dim key As String
Dim i As Integer

    With EQSounds
        HC.oPersist.WriteIniValue "SND_WAV_ALARM", .AlarmWav
        HC.oPersist.WriteIniValue "SND_WAV_CLICK", .ClickWav
        HC.oPersist.WriteIniValue "SND_WAV_BEEP", .BeepWav
        HC.oPersist.WriteIniValue "SND_WAV_SYNC", .SyncWav
        HC.oPersist.WriteIniValue "SND_WAV_PARK", .ParkWav
        HC.oPersist.WriteIniValue "SND_WAV_PARKED", .ParkedWav
        HC.oPersist.WriteIniValue "SND_WAV_GOTO", .GotoWav
        HC.oPersist.WriteIniValue "SND_WAV_GOTOSTART", .GotoStartWav
        HC.oPersist.WriteIniValue "SND_WAV_STOP", .StopWav
        HC.oPersist.WriteIniValue "SND_WAV_UNPARK", .Unparkwav
        HC.oPersist.WriteIniValue "SND_WAV_SIDEREAL", .SiderealWav
        HC.oPersist.WriteIniValue "SND_WAV_LUNAR", .LunarWav
        HC.oPersist.WriteIniValue "SND_WAV_SOLAR", .SolarWav
        HC.oPersist.WriteIniValue "SND_WAV_ACCEPT", .AcceptWav
        HC.oPersist.WriteIniValue "SND_WAV_CANCEL", .CancelWav
        HC.oPersist.WriteIniValue "SND_WAV_END", .EndWav
        HC.oPersist.WriteIniValue "SND_WAV_PHOME", .PHomeWav
        HC.oPersist.WriteIniValue "SND_WAV_PALIGN", .PAlignwav
        HC.oPersist.WriteIniValue "SND_WAV_PALIGNED", .PAlignedwav
        HC.oPersist.WriteIniValue "SND_WAV_DMS", .DMSwav
        HC.oPersist.WriteIniValue "SND_WAV_DMS2", .DMS2wav
        HC.oPersist.WriteIniValue "SND_WAV_GPLON", .GPLOnwav
        HC.oPersist.WriteIniValue "SND_WAV_GPLOFF", .GPLOffwav
        HC.oPersist.WriteIniValue "SND_WAV_MONITORON", .MonitorOnwav
        HC.oPersist.WriteIniValue "SND_WAV_MONITOROFF", .MonitorOffwav
        HC.oPersist.WriteIniValue "SND_WAV_CUSTOM", .CustomWav
        HC.oPersist.WriteIniValue "SND_WAV_RAREVERSEON", .RAReverseOnwav
        HC.oPersist.WriteIniValue "SND_WAV_RAREVERSEOFF", .RaReverseOffwav
        HC.oPersist.WriteIniValue "SND_WAV_DECREVERSEON", .DecReverseOnwav
        HC.oPersist.WriteIniValue "SND_WAV_DECREVERSEOFF", .DecReverseOffwav
        

        For i = 1 To 10
            key = "SND_WAV_RATE" & CStr(i)
            HC.oPersist.WriteIniValue key, .RateWav(i)
        Next i
        HC.oPersist.WriteIniValue "SND_MODE", CStr(.mode)
        If .PositionBeep Then
            HC.oPersist.WriteIniValue "SND_ENABLE_BEEP", "1"
        Else
            HC.oPersist.WriteIniValue "SND_ENABLE_BEEP", "0"
        End If
        
        If .ButtonClick Then
            HC.oPersist.WriteIniValue "SND_ENABLE_CLICK", "1"
        Else
            HC.oPersist.WriteIniValue "SND_ENABLE_CLICK", "0"
        End If
        If .FlipWarning Then
            HC.oPersist.WriteIniValue "SND_ENABLE_ALARM", "1"
        Else
            HC.oPersist.WriteIniValue "SND_ENABLE_ALARM", "0"
        End If
        
        If .RateClick Then
            HC.oPersist.WriteIniValue "SND_ENABLE_RATE", "1"
        Else
            HC.oPersist.WriteIniValue "SND_ENABLE_RATE", "0"
        End If
        If .ParkClick Then
            HC.oPersist.WriteIniValue "SND_ENABLE_PARK", "1"
        Else
            HC.oPersist.WriteIniValue "SND_ENABLE_PARK", "0"
        End If
        If .ParkedClick Then
            HC.oPersist.WriteIniValue "SND_ENABLE_PARKED", "1"
        Else
            HC.oPersist.WriteIniValue "SND_ENABLE_PARKED", "0"
        End If
        If .GotoClick Then
            HC.oPersist.WriteIniValue "SND_ENABLE_GOTO", "1"
        Else
            HC.oPersist.WriteIniValue "SND_ENABLE_GOTO", "0"
        End If
        If .GotoStartClick Then
            HC.oPersist.WriteIniValue "SND_ENABLE_GOTOSTART", "1"
        Else
            HC.oPersist.WriteIniValue "SND_ENABLE_GOTOSTART", "0"
        End If
        If .Stopclick Then
            HC.oPersist.WriteIniValue "SND_ENABLE_STOP", "1"
        Else
            HC.oPersist.WriteIniValue "SND_ENABLE_STOP", "0"
        End If
        If .Unparkclick Then
            HC.oPersist.WriteIniValue "SND_ENABLE_UNPARK", "1"
        Else
            HC.oPersist.WriteIniValue "SND_ENABLE_UNPARK", "0"
        End If
        If .TrackClick Then
            HC.oPersist.WriteIniValue "SND_ENABLE_TRACKING", "1"
        Else
            HC.oPersist.WriteIniValue "SND_ENABLE_TRACKING", "0"
        End If
        If .AlignClick Then
            HC.oPersist.WriteIniValue "SND_ENABLE_ALIGN", "1"
        Else
            HC.oPersist.WriteIniValue "SND_ENABLE_ALIGN", "0"
        End If
        If .PolarClick Then
            HC.oPersist.WriteIniValue "SND_ENABLE_POLAR", "1"
        Else
            HC.oPersist.WriteIniValue "SND_ENABLE_POLAR", "0"
        End If
        If .DMSClick Then
            HC.oPersist.WriteIniValue "SND_ENABLE_DMS", "1"
        Else
            HC.oPersist.WriteIniValue "SND_ENABLE_DMS", "0"
        End If
        If .GPLClick Then
            HC.oPersist.WriteIniValue "SND_ENABLE_GPL", "1"
        Else
            HC.oPersist.WriteIniValue "SND_ENABLE_GPL", "0"
        End If
        If .MonitorClick Then
            HC.oPersist.WriteIniValue "SND_ENABLE_MONITOR", "1"
        Else
            HC.oPersist.WriteIniValue "SND_ENABLE_MONITOR", "0"
        End If
        If .ReverseClick Then
            HC.oPersist.WriteIniValue "SND_ENABLE_REVERSE", "1"
        Else
            HC.oPersist.WriteIniValue "SND_ENABLE_REVERSE", "0"
        End If
   
    
    End With
End Sub
Public Sub readBeep()

Dim tmptxt As String
Dim key As String
Dim i As Integer

    With EQSounds
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_ALARM")
        If tmptxt <> "" Then
            .AlarmWav = tmptxt
        Else
            .AlarmWav = "EQMOD_klaxton.wav"
            HC.oPersist.WriteIniValue "SND_WAV_ALARM", .AlarmWav
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_CLICK")
        If tmptxt <> "" Then
            .ClickWav = tmptxt
        Else
            .ClickWav = "EQMOD_click.wav"
            HC.oPersist.WriteIniValue "SND_WAV_CLICK", .ClickWav
        End If
    
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_BEEP")
        If tmptxt <> "" Then
            .BeepWav = tmptxt
        Else
            .BeepWav = "EQMOD_beep.wav"
            HC.oPersist.WriteIniValue "SND_WAV_BEEP", .BeepWav
        End If
         
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_SYNC")
        If tmptxt <> "" Then
            .SyncWav = tmptxt
        Else
            .SyncWav = "EQMOD_beep.wav"
            HC.oPersist.WriteIniValue "SND_WAV_SYNC", .SyncWav
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_PARK")
        If tmptxt <> "" Then
            .ParkWav = tmptxt
        Else
            .ParkWav = "EQMOD_beep.wav"
            HC.oPersist.WriteIniValue "SND_WAV_PARK", .ParkWav
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_UNPARK")
        If tmptxt <> "" Then
            .Unparkwav = tmptxt
        Else
            .Unparkwav = "EQMOD_beep.wav"
            HC.oPersist.WriteIniValue "SND_WAV_UNPARK", .Unparkwav
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_PARKED")
        If tmptxt <> "" Then
            .ParkedWav = tmptxt
        Else
            .ParkedWav = "EQMOD_beep.wav"
            HC.oPersist.WriteIniValue "SND_WAV_PARKED", .ParkedWav
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_GOTO")
        If tmptxt <> "" Then
            .GotoWav = tmptxt
        Else
            .GotoWav = "EQMOD_beep.wav"
            HC.oPersist.WriteIniValue "SND_WAV_GOTO", .GotoWav
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_GOTOSTART")
        If tmptxt <> "" Then
            .GotoStartWav = tmptxt
        Else
            .GotoStartWav = "EQMOD_beep.wav"
            HC.oPersist.WriteIniValue "SND_WAV_GOTOSTART", .GotoStartWav
        End If
         
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_STOP")
        If tmptxt <> "" Then
            .StopWav = tmptxt
        Else
            .StopWav = "EQMOD_beep.wav"
            HC.oPersist.WriteIniValue "SND_WAV_STOP", .StopWav
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_SIDEREAL")
        If tmptxt <> "" Then
            .SiderealWav = tmptxt
        Else
            .SiderealWav = "EQMOD_beep.wav"
            HC.oPersist.WriteIniValue "SND_WAV_SIDEREAL", .SiderealWav
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_LUNAR")
        If tmptxt <> "" Then
            .LunarWav = tmptxt
        Else
            .LunarWav = "EQMOD_beep.wav"
            HC.oPersist.WriteIniValue "SND_WAV_LUNAR", .LunarWav
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_SOLAR")
        If tmptxt <> "" Then
            .SolarWav = tmptxt
        Else
            .SolarWav = "EQMOD_beep.wav"
            HC.oPersist.WriteIniValue "SND_WAV_SOLAR", .SolarWav
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_CUSTOM")
        If tmptxt <> "" Then
            .CustomWav = tmptxt
        Else
            .CustomWav = "EQMOD_beep.wav"
            HC.oPersist.WriteIniValue "SND_WAV_CUSTOM", .CustomWav
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_ACCEPT")
        If tmptxt <> "" Then
            .AcceptWav = tmptxt
        Else
            .AcceptWav = "EQMOD_beep.wav"
            HC.oPersist.WriteIniValue "SND_WAV_ACCEPT", .AcceptWav
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_CANCEL")
        If tmptxt <> "" Then
            .CancelWav = tmptxt
        Else
            .CancelWav = "EQMOD_beep.wav"
            HC.oPersist.WriteIniValue "SND_WAV_CANCEL", .CancelWav
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_END")
        If tmptxt <> "" Then
            .EndWav = tmptxt
        Else
            .EndWav = "EQMOD_beep.wav"
            HC.oPersist.WriteIniValue "SND_WAV_END", .EndWav
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_PHOME")
        If tmptxt <> "" Then
            .PHomeWav = tmptxt
        Else
            .PHomeWav = "EQMOD_beep.wav"
            HC.oPersist.WriteIniValue "SND_WAV_PHOME", .PHomeWav
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_PALIGN")
        If tmptxt <> "" Then
            .PAlignwav = tmptxt
        Else
            .PAlignwav = "EQMOD_beep.wav"
            HC.oPersist.WriteIniValue "SND_WAV_PALIGN", .PAlignwav
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_PALIGNED")
        If tmptxt <> "" Then
            .PAlignedwav = tmptxt
        Else
            .PAlignedwav = "EQMOD_beep.wav"
            HC.oPersist.WriteIniValue "SND_WAV_PALIGNED", .PAlignedwav
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_DMS")
        If tmptxt <> "" Then
            .DMSwav = tmptxt
        Else
            .DMSwav = "EQMOD_beep.wav"
            HC.oPersist.WriteIniValue "SND_WAV_DMS", .DMSwav
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_DMS2")
        If tmptxt <> "" Then
            .DMS2wav = tmptxt
        Else
            .DMS2wav = "EQMOD_beep.wav"
            HC.oPersist.WriteIniValue "SND_WAV_DMS2", .DMS2wav
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_GPLON")
        If tmptxt <> "" Then
            .GPLOnwav = tmptxt
        Else
            .GPLOnwav = "EQMOD_beep.wav"
            HC.oPersist.WriteIniValue "SND_WAV_GPLON", .GPLOnwav
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_GPLOFF")
        If tmptxt <> "" Then
            .GPLOffwav = tmptxt
        Else
            .GPLOffwav = "EQMOD_beep.wav"
            HC.oPersist.WriteIniValue "SND_WAV_GPLOFF", .GPLOffwav
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_MONITORON")
        If tmptxt <> "" Then
            .MonitorOnwav = tmptxt
        Else
            .MonitorOnwav = "EQMOD_beep.wav"
            HC.oPersist.WriteIniValue "SND_WAV_MONITORON", .MonitorOnwav
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_MONITOROFF")
        If tmptxt <> "" Then
            .MonitorOffwav = tmptxt
        Else
            .MonitorOffwav = "EQMOD_beep.wav"
            HC.oPersist.WriteIniValue "SND_WAV_MONITOROFF", .MonitorOffwav
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_RAREVERSEOFF")
        If tmptxt <> "" Then
            .RaReverseOffwav = tmptxt
        Else
            .RaReverseOffwav = "EQMOD_beep.wav"
            HC.oPersist.WriteIniValue "SND_WAV_RAREVERSEOFF", .RaReverseOffwav
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_RAREVERSEON")
        If tmptxt <> "" Then
            .RAReverseOnwav = tmptxt
        Else
            .RAReverseOnwav = "EQMOD_beep.wav"
            HC.oPersist.WriteIniValue "SND_WAV_RAREVERSEON", .RAReverseOnwav
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_DECREVERSEOFF")
        If tmptxt <> "" Then
            .DecReverseOffwav = tmptxt
        Else
            .DecReverseOffwav = "EQMOD_beep.wav"
            HC.oPersist.WriteIniValue "SND_WAV_DECREVERSEOFF", .DecReverseOffwav
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_WAV_DECREVERSEON")
        If tmptxt <> "" Then
            .DecReverseOnwav = tmptxt
        Else
            .DecReverseOnwav = "EQMOD_beep.wav"
            HC.oPersist.WriteIniValue "SND_WAV_DECREVERSEON", .DecReverseOnwav
        End If
        
        
        
        For i = 1 To 10
            key = "SND_WAV_RATE" & CStr(i)
            tmptxt = HC.oPersist.ReadIniValue(key)
            If tmptxt <> "" Then
                .RateWav(i) = tmptxt
            Else
                .RateWav(i) = "EQMOD_click.wav"
                HC.oPersist.WriteIniValue key, .ClickWav
            End If
        Next i
        
        tmptxt = HC.oPersist.ReadIniValue("SND_MODE")
        If tmptxt <> "" Then
            .mode = val(tmptxt)
        Else
            .mode = 0
            HC.oPersist.WriteIniValue "SND_MODE", CStr(.mode)
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_ENABLE_BEEP")
        If tmptxt <> "" Then
            If tmptxt = "1" Then
                .PositionBeep = True
            Else
                .PositionBeep = False
            End If
        Else
            .PositionBeep = False
            HC.oPersist.WriteIniValue "SND_ENABLE_BEEP", "0"
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_ENABLE_CLICK")
        If tmptxt <> "" Then
            If tmptxt = "1" Then
                .ButtonClick = True
            Else
                .ButtonClick = False
            End If
        Else
            .ButtonClick = False
            HC.oPersist.WriteIniValue "SND_ENABLE_CLICK", "0"
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_ENABLE_ALARM")
        If tmptxt <> "" Then
            If tmptxt = "1" Then
                .FlipWarning = True
            Else
                .FlipWarning = False
            End If
        Else
            .FlipWarning = False
            HC.oPersist.WriteIniValue "SND_ENABLE_ALARM", "0"
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_ENABLE_RATE")
        If tmptxt <> "" Then
            If tmptxt = "1" Then
                .RateClick = True
            Else
                .RateClick = False
            End If
        Else
            .RateClick = False
            HC.oPersist.WriteIniValue "SND_ENABLE_RATE", "0"
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_ENABLE_PARK")
        If tmptxt <> "" Then
            If tmptxt = "1" Then
                .ParkClick = True
            Else
                .ParkClick = False
            End If
        Else
            .ParkClick = False
            HC.oPersist.WriteIniValue "SND_ENABLE_PARK", "0"
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_ENABLE_UNPARK")
        If tmptxt <> "" Then
            If tmptxt = "1" Then
                .Unparkclick = True
            Else
                .Unparkclick = False
            End If
        Else
            .Unparkclick = False
            HC.oPersist.WriteIniValue "SND_ENABLE_UNPARK", "0"
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_ENABLE_PARKED")
        If tmptxt <> "" Then
            If tmptxt = "1" Then
                .ParkedClick = True
            Else
                .ParkedClick = False
            End If
        Else
            .ParkedClick = False
            HC.oPersist.WriteIniValue "SND_ENABLE_PARKED", "0"
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_ENABLE_GOTO")
        If tmptxt <> "" Then
            If tmptxt = "1" Then
                .GotoClick = True
            Else
                .GotoClick = False
            End If
        Else
            .GotoClick = False
            HC.oPersist.WriteIniValue "SND_ENABLE_GOTO", "0"
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_ENABLE_GOTOSTART")
        If tmptxt <> "" Then
            If tmptxt = "1" Then
                .GotoStartClick = True
            Else
                .GotoStartClick = False
            End If
        Else
            .GotoStartClick = False
            HC.oPersist.WriteIniValue "SND_ENABLE_GOTOSTART", "0"
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_ENABLE_STOP")
        If tmptxt <> "" Then
            If tmptxt = "1" Then
                .Stopclick = True
            Else
                .Stopclick = False
            End If
        Else
            .Stopclick = False
            HC.oPersist.WriteIniValue "SND_ENABLE_STOP", "0"
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_ENABLE_TRACKING")
        If tmptxt <> "" Then
            If tmptxt = "1" Then
                .TrackClick = True
            Else
                .TrackClick = False
            End If
        Else
            .TrackClick = False
            HC.oPersist.WriteIniValue "SND_ENABLE_TRACKING", "0"
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_ENABLE_ALIGN")
        If tmptxt <> "" Then
            If tmptxt = "1" Then
                .AlignClick = True
            Else
                .AlignClick = False
            End If
        Else
            .AlignClick = False
            HC.oPersist.WriteIniValue "SND_ENABLE_ALIGN", "0"
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_ENABLE_POLAR")
        If tmptxt <> "" Then
            If tmptxt = "1" Then
                .PolarClick = True
            Else
                .PolarClick = False
            End If
        Else
            .PolarClick = False
            HC.oPersist.WriteIniValue "SND_ENABLE_POLAR", "0"
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_ENABLE_DMS")
        If tmptxt <> "" Then
            If tmptxt = "1" Then
                .DMSClick = True
            Else
                .DMSClick = False
            End If
        Else
            .DMSClick = False
            HC.oPersist.WriteIniValue "SND_ENABLE_DMS", "0"
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_ENABLE_GPL")
        If tmptxt <> "" Then
            If tmptxt = "1" Then
                .GPLClick = True
            Else
                .GPLClick = False
            End If
        Else
            .GPLClick = False
            HC.oPersist.WriteIniValue "SND_ENABLE_GPL", "0"
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_ENABLE_MONITOR")
        If tmptxt <> "" Then
            If tmptxt = "1" Then
                .MonitorClick = True
            Else
                .MonitorClick = False
            End If
        Else
            .MonitorClick = False
            HC.oPersist.WriteIniValue "SND_ENABLE_MONITOR", "0"
        End If
        
        tmptxt = HC.oPersist.ReadIniValue("SND_ENABLE_REVERSE")
        If tmptxt <> "" Then
            If tmptxt = "1" Then
                .ReverseClick = True
            Else
                .ReverseClick = False
            End If
        Else
            .ReverseClick = False
            HC.oPersist.WriteIniValue "SND_ENABLE_REVERSE", "0"
        End If
        
    End With
End Sub

