Attribute VB_Name = "Graphics"
Option Explicit

Private Declare Function Arc Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function AngleArc Lib "GDI32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal dwRadius As Long, ByVal eStartAngle As Single, ByVal eSweepAngle As Single) As Long
Private Declare Function MoveToEx Lib "GDI32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByRef lpPoint As Any) As Long

Public Sub PlotInit()

Dim i As Integer

    HC.Plot_RA.DrawMode = 13
    HC.Plot_DEC.DrawMode = 13
    
    HC.Plot_RA.Cls
    HC.Plot_DEC.Cls
    
    HC.Plot_RA.Line (0, HC.Plot_RA.ScaleHeight / 2)-(HC.Plot_RA.ScaleWidth, HC.Plot_RA.ScaleHeight / 2), vbBlue
    HC.Plot_DEC.Line (0, HC.Plot_DEC.ScaleHeight / 2)-(HC.Plot_DEC.ScaleWidth, HC.Plot_DEC.ScaleHeight / 2), vbBlue

    gPlot_ra_pos = 0
    gPlot_dec_pos = 0
    
    gRAHeight = (HC.Plot_RA.ScaleHeight / 2)
    gDecHeight = (HC.Plot_DEC.ScaleHeight / 2)

    gplot_ra_cur = gRAHeight
    gPlot_dec_cur = gDecHeight
    
End Sub

Public Sub Plot_PG(id As Integer, side As Integer, ppvalue As Long)
Dim pheight As Double
Dim pscale As Double
Dim pvalue As Double
Dim nextpos As Double

'    If HC.Frame16.Visible = True Then
        pvalue = ppvalue
        gMAX_RAlevel = HC.RAdisplay_gain.Value + 100
        gMAX_DEClevel = HC.DECdisplay_gain.Value + 100
    
        If id = 0 Then
    
            nextpos = gPlot_ra_pos + 3
            If nextpos > HC.Plot_RA.ScaleWidth Then
                gPlot_ra_pos = 0
                nextpos = 3
                HC.Plot_RA.Line (0, 0)-(0, HC.Plot_RA.ScaleHeight), &H40&
            End If
        
            If pvalue > gMAX_RAlevel Then pvalue = gMAX_RAlevel
        
            pscale = (pvalue / gMAX_RAlevel) * gRAHeight
            HC.Plot_RA.Line (gPlot_ra_pos + 1, 0)-(gPlot_ra_pos + 1, HC.Plot_RA.ScaleHeight), &H40&
            HC.Plot_RA.Line (gPlot_ra_pos + 2, 0)-(gPlot_ra_pos + 2, HC.Plot_RA.ScaleHeight), &H40&
            HC.Plot_RA.Line (gPlot_ra_pos + 3, 0)-(gPlot_ra_pos + 3, HC.Plot_RA.ScaleHeight), &H40&
            HC.Plot_RA.Line (gPlot_ra_pos + 4, 0)-(gPlot_ra_pos + 4, HC.Plot_RA.ScaleHeight), vbBlue
            HC.Plot_RA.Line (gPlot_ra_pos + 5, 0)-(gPlot_ra_pos + 5, HC.Plot_RA.ScaleHeight), &H40&
            HC.Plot_RA.Line (gPlot_ra_pos, gRAHeight)-(nextpos + 1, gRAHeight), vbBlue
            
            If side = 0 Then
                HC.Plot_RA.Line (gPlot_ra_pos, gplot_ra_cur)-(nextpos, gRAHeight - pscale), vbRed
                gplot_ra_cur = gRAHeight - pscale
            Else
                HC.Plot_RA.Line (gPlot_ra_pos, gplot_ra_cur)-(nextpos, gRAHeight + pscale), vbRed
                gplot_ra_cur = gRAHeight + pscale
            End If
            
            gPlot_ra_pos = nextpos
        
        Else
        
            nextpos = gPlot_dec_pos + 3
            If nextpos > HC.Plot_DEC.ScaleWidth Then
                gPlot_dec_pos = 0
                nextpos = 3
                HC.Plot_DEC.Line (0, 0)-(0, HC.Plot_DEC.ScaleHeight), &H40&
            End If
        
            If pvalue > gMAX_DEClevel Then pvalue = gMAX_DEClevel
        
            pscale = (pvalue / gMAX_DEClevel) * gDecHeight
            HC.Plot_DEC.Line (gPlot_dec_pos + 1, 0)-(gPlot_dec_pos + 1, HC.Plot_DEC.ScaleHeight), &H40&
            HC.Plot_DEC.Line (gPlot_dec_pos + 2, 0)-(gPlot_dec_pos + 2, HC.Plot_DEC.ScaleHeight), &H40&
            HC.Plot_DEC.Line (gPlot_dec_pos + 3, 0)-(gPlot_dec_pos + 3, HC.Plot_DEC.ScaleHeight), &H40&
            HC.Plot_DEC.Line (gPlot_dec_pos + 4, 0)-(gPlot_dec_pos + 4, HC.Plot_DEC.ScaleHeight), vbBlue
            HC.Plot_DEC.Line (gPlot_dec_pos + 5, 0)-(gPlot_dec_pos + 5, HC.Plot_DEC.ScaleHeight), &H40&
            HC.Plot_DEC.Line (gPlot_dec_pos, gDecHeight)-(nextpos + 1, gRAHeight), vbBlue
            
            If side = 0 Then
                HC.Plot_DEC.Line (gPlot_dec_pos, gPlot_dec_cur)-(nextpos, gDecHeight - pscale), vbRed
                gPlot_dec_cur = gDecHeight - pscale
            Else
                HC.Plot_DEC.Line (gPlot_dec_pos, gPlot_dec_cur)-(nextpos, gDecHeight + pscale), vbRed
                gPlot_dec_cur = gDecHeight + pscale
            End If
            
            gPlot_dec_pos = nextpos
        
        End If
'    End If
End Sub

Public Sub DrawAxis(pic As PictureBox, mode As Integer, val As Double, lowlimit As Double, highlimit As Double)

    Dim i As Integer
    Dim x1, y1, x2, y2, tmp As Double
    
    pic.Cls
    val = val * 3.6
    pic.DrawWidth = 10
    If mode = -1 Then
        pic.Circle (40, 40), 35, &H808080
    Else
        pic.Circle (40, 40), 35, &H8000&
        pic.DrawWidth = 2
        For i = 1 To 9
            pic.ForeColor = vbRed
           Call MoveToEx(pic.hDC, 40, i, ByVal 0&)
           AngleArc pic.hDC, 40, 40, 40 - i, 90, -val
        Next i
    End If
    pic.DrawWidth = 1
    pic.Circle (40, 40), 30, vbBlack
    pic.Circle (40, 40), 40, vbBlack
    
    For i = 0 To 345 Step 15
       tmp = i * PI / 180
       x1 = 30 * Cos(tmp) + 40
       y1 = 30 * Sin(tmp) + 40
       x2 = (40) * Cos(tmp) + 40
       y2 = (40) * Sin(tmp) + 40
       If i = 0 Or i = 180 Then
          pic.Line (x1, y1)-(x2, y2), vbCyan
       Else
          pic.Line (x1, y1)-(x2, y2), vbBlack
       End If
    Next i
    
    pic.DrawWidth = 2
    
    If lowlimit > 0 Then
        tmp = lowlimit * 3.6 - 90
        tmp = tmp * PI / 180
        x1 = 30 * Cos(tmp) + 40
        y1 = 30 * Sin(tmp) + 40
        x2 = (40) * Cos(tmp) + 40
        y2 = (40) * Sin(tmp) + 40
        pic.Line (x1, y1)-(x2, y2), vbYellow
    End If
    
    If highlimit > 0 Then
        tmp = highlimit * 3.6 - 90
        tmp = tmp * PI / 180
        x1 = 30 * Cos(tmp) + 40
        y1 = 30 * Sin(tmp) + 40
        x2 = (40) * Cos(tmp) + 40
        y2 = (40) * Sin(tmp) + 40
        pic.Line (x1, y1)-(x2, y2), vbYellow
        
    End If
    
    pic.DrawWidth = 1
    
    pic.ForeColor = &H80FF&
    i = pic.TextHeight("0") / 2
    Select Case mode
        Case 0
            pic.CurrentX = 40 - pic.TextWidth("6") / 2
            pic.CurrentY = 20 - i
            pic.Print "6"
            pic.CurrentX = 40 - pic.TextWidth("18") / 2
            pic.CurrentY = 60 - i
            pic.Print "18"
            pic.CurrentX = 20 - pic.TextWidth("0") / 2
            pic.CurrentY = 40 - i
            pic.Print "0"
            pic.CurrentX = 60 - pic.TextWidth("12") / 2
            pic.CurrentY = 40 - i
            pic.Print "12"
            pic.CurrentX = 40 - pic.TextWidth(oLangDll.GetLangString(105)) / 2
            pic.CurrentY = 40 - i
            pic.Print oLangDll.GetLangString(105)
        
        Case 1
            pic.CurrentX = 40 - pic.TextWidth("90") / 2
            pic.CurrentY = 20 - i
            pic.Print "90"
            pic.CurrentX = 40 - pic.TextWidth("-90") / 2
            pic.CurrentY = 60 - i
            pic.Print "-90"
            pic.CurrentX = 20 - pic.TextWidth("0") / 2
            pic.CurrentY = 40 - i
            pic.Print "0"
            pic.CurrentX = 60 - pic.TextWidth("0") / 2
            pic.CurrentY = 40 - i
            pic.Print "0"
            pic.CurrentX = 40 - pic.TextWidth(oLangDll.GetLangString(106)) / 2
            pic.CurrentY = 40 - i
            pic.Print oLangDll.GetLangString(106)
    End Select

End Sub


