Attribute VB_Name = "FFT"
'***************************************************************
' Copyright © 2006 Chris Shillito
'
' Fast Fourier Transform
'
'***************************************************************

Option Explicit

Private ReX() As Double
Private ImX() As Double
Private FFTSampleRate As Double
Private SampleCount As Integer
Private MaxMag As Double
Private N As Integer

Const pi = 3.14159265  'Set constants


Public Function FFT_Free()
    ReDim ReX(0)
    ReDim ImX(0)
    N = 0
End Function

Public Sub FFT_Initialise(ByVal size As Integer, ByVal rate As Double)
Dim i As Integer
    ReDim ReX(size)
    ReDim ImX(size)
    N = size
    For i = 0 To N
        ReX(i) = 0
        ImX(i) = 0
    Next i
    FFTSampleRate = rate
    MaxMag = 0
    SampleCount = 0
End Sub


'Upon entry, N% contains the number of points in the DFT, fftReal[ ] and
'fft[].Img contain the real and imaginary parts of the input.
' Upon return, fft[].Real and fft[].Img contain the DFT output. All signals run from 0 to N-1.

Public Sub FFT_ForwardFFTComplex()

Dim TI, TR As Double
Dim i, j, k, l, m As Integer
Dim IP, LE, LE2 As Integer
Dim ur, ui, sr, si As Double
Dim NDiv2 As Integer
Dim NSub1 As Integer
Dim NSub2 As Integer
    
    If N <> 0 Then
        ImX(0) = 0
        ReX(0) = 0

        NSub1 = N - 1
        NSub2 = N - 2
        NDiv2 = N / 2
        m = CInt(Log(N) / Log(2))
        j = NDiv2
    
        For i = 1 To (NSub2) 'Bit reversal sorting
            
            If i < j Then
                TR = ReX(j)
                TI = ImX(j)
                ReX(j) = ReX(i)
                ImX(j) = ImX(i)
                ReX(i) = TR
                ImX(i) = TI
            End If
            
            k = NDiv2
            
            While k <= j
                j = j - k
                k = k / 2
            Wend
            
            j = j + k
        
        Next i
    
        For l = 1 To m 'Loop for each stage
            LE = CInt(2 ^ l)
            LE2 = LE / 2
            ur = 1
            ui = 0
            sr = Cos(pi / LE2) 'Calculate sine & cosine values
            si = -Sin(pi / LE2)
            
            For j = 1 To LE2 'Loop for each sub DFT
                
                For i = (j - 1) To (NSub1) Step LE 'Loop for each butterfly
                    IP = i + LE2
                    TR = ReX(IP) * ur - ImX(IP) * ui 'Butterfly calculation
                    TI = ReX(IP) * ui + ImX(IP) * ur
                    ReX(IP) = ReX(i) - TR
                    ImX(IP) = ImX(i) - TI
                    ReX(i) = ReX(i) + TR
                    ImX(i) = ImX(i) + TI
                    
                Next i
                
                TR = ur
                ur = TR * sr - ui * si
                ui = TR * si + ui * sr
            
            Next j
        
        Next l
    End If
End Sub
Public Sub FFT_NormaliseMag()
Dim idx As Integer
Dim mag As Double
Dim max As Double
    MaxMag = 0
    max = 0
    For idx = 1 To N / 2 - 1
        mag = FFT_GetMagnitude(idx)
        If mag > max Then
            max = mag
        End If
    Next idx
    MaxMag = max

End Sub

Public Sub FFT_InverseFFTComplex()
' upon entry N is the numbr of real and imaginary points.
' real[] and img[] contain the real and imaginary parts of the frequency domain
' running for index 0 to n/2
' On return real[] containds the real time domain, img[] contains zeros.
Dim k As Integer
Dim NSub1 As Integer
    
    NSub1 = N - 1

    If N <> 0 Then
        For k = 0 To NSub1
            ImX(k) = -ImX(k)
        Next k
        
        FFT_ForwardFFTComplex
        
        For k = 0 To NSub1
            ReX(k) = ReX(k) / N
' don't really need the imaginary part of the time domain
'            ImX(K) = -ImX(K) / N
        Next k
    End If

End Sub

Public Sub FFT_InverseFFTReal()
' upon entry N is the numbr of real and imaginary points.
' real[] and img[] contain the real and imaginary parts of the frequency domain
' running for index 0 to n/2
' On return real[] contains the real time domain, img[] contains zeros.
Dim k As Integer

    If N <> 0 Then
        For k = N / 2 + 1 To N - 1
            ReX(k) = ReX(N - k)
            ImX(k) = -ImX(N - k)
        Next k
        
        For k = 0 To N - 1
            ReX(k) = ReX(k) + ImX(k)
        Next k
        
        FFT_ForwardFFTComplex
        
        For k = 0 To N - 1
            ReX(k) = (ReX(k) + ImX(k)) / N
            ImX(k) = 0
        Next k
    End If

End Sub

Public Sub FFT_ApplyFilter(lofilter As Double, hifilter As Double, MagLimit As Double)
Dim lo As Double
Dim hi As Double

    If lofilter = 0 Then
        lo = N
    Else
        lo = FFT_Freq2Bin(lofilter)
    End If
  
    hi = FFT_Freq2Bin(hifilter)
    
    If lo <> -1 And hi <> -1 Then Call FFT_Filter(lo, hi, MagLimit)
End Sub


Public Sub FFT_Filter(lofilter As Double, hifilter As Double, MagLimit As Double)
Dim k As Integer
Dim NDiv2Sub1 As Integer
Dim NSub1 As Integer

    NDiv2Sub1 = N / 2 - 1
    NSub1 = N - 1
    If N <> 0 Then
        If Not (lofilter = 0 And hifilter = 0) Then
            For k = 0 To NDiv2Sub1
                If (k >= lofilter) Or (k <= hifilter) Then
                    ImX(k) = 0
                    ReX(k) = 0
                Else
                    If FFT_GetMagnitude(k) < MagLimit Then
                        ImX(k) = 0
                        ReX(k) = 0
                    End If
                    
                End If
            Next k
        End If
    
        For k = N / 2 + 1 To NSub1
            ReX(k) = ReX(N - k)
            ImX(k) = -ImX(N - k)
        Next k
    
    End If

End Sub

Public Function FFT_Bin2Freq(bin As Double) As Double
Dim NDiv2 As Integer

    NDiv2 = N / 2
    
    If bin > NDiv2 Then
        FFT_Bin2Freq = -1
    Else
        FFT_Bin2Freq = bin * FFTSampleRate / N ' Div2
    End If

End Function

Public Function FFT_Freq2Bin(freq As Double) As Double
Dim NDiv2 As Integer

    NDiv2 = N / 2
    
    If freq > FFTSampleRate / 2 Then
        FFT_Freq2Bin = -1
    Else
        FFT_Freq2Bin = freq * N / FFTSampleRate
    End If

End Function


Public Function FFT_GetPhase(idx As Integer) As Double
    If idx <= N Then
        If ReX(idx) = 0 Then
            ' divide by 0 error
            If ImX(idx) < 0 Then
                FFT_GetPhase = -pi / 2
            Else
                FFT_GetPhase = pi / 2
            End If
        Else
            ' calculate phase
            FFT_GetPhase = Atn(ImX(idx) / ReX(idx))
        
            ' fix incorrect arctan
            If ReX(idx) < 0 Then
                If ImX(idx) < 0 Then
                    ' Rex < 0 and imx < 0
                    FFT_GetPhase = FFT_GetPhase - pi
                Else
                    ' rex < 0 and imx > 0
                    FFT_GetPhase = FFT_GetPhase + pi
                End If
            End If
        End If
    Else
        FFT_GetPhase = 0
    End If
End Function

Public Function FFT_GetMagnitude(idx As Integer) As Double
    If idx <= N Then
        If MaxMag <> 0 Then
            FFT_GetMagnitude = Sqr((ImX(idx) * ImX(idx)) + (ReX(idx) * ReX(idx)))
            ' normalise magnitude
            FFT_GetMagnitude = 100 * FFT_GetMagnitude / MaxMag
        Else
            FFT_GetMagnitude = Sqr((ImX(idx) * ImX(idx)) + (ReX(idx) * ReX(idx)))
        End If
    Else
        FFT_GetMagnitude = 0
    End If
End Function

Public Function FFT_GetReX(idx As Integer) As Double
    If idx <= N Then
        FFT_GetReX = ReX(idx)
    Else
        FFT_GetReX = 0
    End If
End Function

Public Function FFT_GetImX(idx As Integer) As Double
    If idx <= N Then
        FFT_GetImX = ImX(idx)
    Else
        FFT_GetImX = 0
    End If
End Function

Public Sub FFT_SetSample(idx As Integer, sample As Double)
    If idx <= N Then
        ReX(idx) = sample
        ImX(idx) = 0
        If idx >= SampleCount Then SampleCount = idx + 1
    End If
End Sub

Public Sub FFT_SetFSample(idx As Integer, r As Double, i As Double)
    If idx <= N Then
        ReX(idx) = r
        ImX(idx) = i
    End If
End Sub

Public Sub FFT_MoveFSample(idx1 As Integer, idx2 As Integer)
    If idx1 <= N And idx2 < N Then
        ReX(idx2) = ReX(idx1)
        ImX(idx2) = ImX(idx1)
        ReX(idx1) = 0
        ImX(idx1) = 0
    End If
End Sub

Public Function FFT_GetSample(idx As Integer) As Double
    If idx <= N Then
        FFT_GetSample = ReX(idx)
    End If
End Function

Public Function FFT_GetSampleCount() As Integer
    FFT_GetSampleCount = SampleCount
End Function

Public Function FFT_GetSize() As Integer
    FFT_GetSize = N
End Function

Public Function FFT_GetSampleRate() As Double
    FFT_GetSampleRate = FFTSampleRate
End Function

