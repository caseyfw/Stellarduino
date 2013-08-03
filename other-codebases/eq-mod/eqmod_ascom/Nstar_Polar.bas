Attribute VB_Name = "Nstar_Polar"
'---------------------------------------------------------------------
' Copyright © 2007 Raymund Sarmiento
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
' Nstar_polar.bas - Polar Alignment using the N-star table
'
'
' Written:  07-Oct-06   Raymund Sarmiento
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' 21-Dec-07 rcs     Initial edit for EQ Mount Driver Function Prototype
'---------------------------------------------------------------------
'
'
'  SYNOPSIS:
'
'  This is a demonstration of a EQ6/ATLAS/EQG direct stepper motor control access
'  using the EQCONTRL.DLL driver code.
'
'  File EQCONTROL.bas contains all the function prototypes of all subroutines
'  encoded in the EQCONTRL.dll
'
'  The EQ6CONTRL.DLL simplifies execution of the Mount controller board stepper
'  commands.
'
'  The mount circuitry needs to be modified for this test program to work.
'  Circuit details can be found at http://www.freewebs.com/eq6mod/
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


Public Function EQGet_Polar_Offset(RA As Double, DEC As Double, radius As Double, raprobe As Double, pscale As Double) As Double

Dim i As Integer

Dim tmpcoord1 As Coord
Dim tmpcoord2 As Coord

Dim dy1 As Double
Dim dy2 As Double
Dim dx As Double


    ' Must perform the usual Update Affine here

    ' Transform using the Negative RA boundary
    
    i = EQ_UpdateAffine_PolarDrift(RA - raprobe, DEC)
    
    tmpcoord1.x = RA - raprobe
    tmpcoord1.Y = DEC
    
    
    tmpcoord1 = EQ_plAffine2(tmpcoord1)

    ' Transform using the Positive RA Boundary

    i = EQ_UpdateAffine_PolarDrift(RA + raprobe, DEC)

    tmpcoord2.x = RA + raprobe
    tmpcoord2.Y = DEC
    
        
    tmpcoord2 = EQ_plAffine2(tmpcoord2)
    
    
    HC.polarplot.Line (pxlate_x(EQ_pl2Cs(tmpcoord1).x, pscale), pxlate_y(EQ_pl2Cs(tmpcoord1).Y, pscale))-(pxlate_x(EQ_pl2Cs(tmpcoord2).x, pscale), pxlate_y(EQ_pl2Cs(tmpcoord2).Y, pscale)), vbRed
    

    ' Get the drift points

    dy1 = DEC - tmpcoord1.Y
    dy2 = DEC - tmpcoord2.Y
    
    ' Coompute for the run data for slope computations
    
    dx = raprobe * Sin(360 * (radius / gTot_step) * DEG_RAD) * 2
    
    If dx = 0 Then dx = 0.00000001

    ' Get the Perpendicular offset error

    EQGet_Polar_Offset = Tan(Atn((dy2 - dy1) / dx)) * Abs(radius)

End Function

' Function to Normalize the Virtual Horizon Measurement data

Public Function EQNormalize_Polar(Alt As Double, Az As Double, vhoriz As Double) As Coord

Dim crt As CartesCoord
Dim crt2 As CartesCoord

    ' Transform Alt/Az data based on the horiz value
    
    ' 90 degrees from the virtual horizon
    crt = EQ_Polar2Cartes(vhoriz + (gTot_step / 4), Alt, gTot_step, 0, 0)
    
    ' 180 degrees from the virtual horizon
    crt2 = EQ_Polar2Cartes(vhoriz + (gTot_step / 2), Az, gTot_step, 0, 0)

    ' Return the normalized data
    EQNormalize_Polar.x = (crt.x + crt2.x) * -1
    EQNormalize_Polar.Y = crt.Y + crt2.Y
    

End Function


'Function to convert polar coordinates to Cartesian using the Coord structure (for Polar Alignment function)

Public Function EQ_pl2Cs_Polar(ByRef obj As Coord, poffset As Double) As Coord

Dim tmpobj As CartesCoord

        tmpobj = EQ_Polar2Cartes(obj.x, obj.Y - poffset, gTot_step, RAEncoder_Home_pos, gDECEncoder_Home_pos)
    
        EQ_pl2Cs_Polar.x = tmpobj.x
        EQ_pl2Cs_Polar.Y = tmpobj.Y
        EQ_pl2Cs_Polar.z = 1
    
End Function

Public Function EQ_UpdateAffine_PolarDrift(x As Double, Y As Double) As Integer

Dim tmpcoord As Coord

Dim i As Long
Dim j As Long
Dim k As Long

Dim datholder(1 To MAX_STARS) As Double
Dim dotidholder(1 To MAX_STARS) As Double

    ' Adjust only if there are four alignment stars
    If gAlignmentStars_count < 3 Then Exit Function

    tmpcoord.x = x
    tmpcoord.Y = Y
    tmpcoord = EQ_sp2Cs(tmpcoord)

    For i = 1 To gAlignmentStars_count
         ' Compute for total X-Y distance.
         datholder(i) = Abs(my_PointsC(i).x - tmpcoord.x) + Abs(my_PointsC(i).Y - tmpcoord.Y)
         ' Also save the reference star id for this particular reference star
         dotidholder(i) = i
    Next i
    
    Call EQ_Quicksort(datholder(), dotidholder(), 1, gAlignmentStars_count)
    ' Get the nearest Star (lowest at the head of the sorted list)
    i = dotidholder(1)
    j = dotidholder(2)
    k = dotidholder(3)

    EQ_UpdateAffine_PolarDrift = EQ_AssembleMatrix_Affine(tmpcoord.x, tmpcoord.Y, ct_PointsC(i), ct_PointsC(j), ct_PointsC(k), my_PointsC(i), my_PointsC(j), my_PointsC(k))
    
End Function


'Implement an Affine transformation on a Polar coordinate system
'This is done by converting the Polar Data to Cartesian, Apply affine transformation
'then return the transformed coordinates

Public Function EQ_plAffineCartes(ByRef obj As Coord) As Coord

Dim tmpobj1 As CartesCoord
Dim tmpobj2 As Coord
Dim tmpobj3 As Coord

        tmpobj1 = EQ_Polar2Cartes(obj.x, obj.Y, gTot_step, RAEncoder_Home_pos, gDECEncoder_Home_pos)
    
        tmpobj2.x = tmpobj1.x
        tmpobj2.Y = tmpobj1.Y
        tmpobj2.z = 1
    
        tmpobj3 = EQ_Transform_Affine(tmpobj2)
        
        EQ_plAffineCartes.x = tmpobj3.x
        EQ_plAffineCartes.Y = tmpobj3.Y
        EQ_plAffineCartes.z = 1


End Function


Public Sub PolarAlign_init(stepcount As Integer)
Dim i As Integer

    HC.polarplot.DrawMode = 13
    HC.polarplot.Cls
    
    If stepcount > 50 Then
        For i = 0 To HC.polarplot.width Step stepcount
    
            HC.polarplot.Circle (HC.polarplot.width / 2, HC.polarplot.Height / 2), i, vbBlue
    
        Next i
    End If
    
    HC.polarplot.Line (0, HC.polarplot.Height / 2)-(HC.polarplot.width, HC.polarplot.Height / 2), vbRed
    HC.polarplot.Line (HC.polarplot.width / 2, 0)-(HC.polarplot.width / 2, HC.polarplot.Height), vbRed

End Sub


Public Sub Plot_PolarAlign(RA As Double, DEC As Double, pscale As Double)

    ' 0.0024 = 0.144 / 60  that is .0024 arcminute / microsteps
    
    HC.polarplot.Circle ((HC.polarplot.width / 2) + ((RA * 0.0024) * pscale), (HC.polarplot.Height / 2) - ((DEC * 0.0024) * pscale)), 50, vbYellow
    HC.polarplot.Line (HC.polarplot.width / 2, HC.polarplot.Height / 2)-((HC.polarplot.width / 2) + ((RA * 0.0024) * pscale), (HC.polarplot.Height / 2) - ((DEC * 0.0024) * pscale)), vbYellow

End Sub

Public Sub NStar_Polar_plot_init(stepcount As Integer)

Dim i As Integer

  HC.polarplot.DrawMode = 13
  HC.polarplot.Cls
    
   
    HC.polarplot.Line (gXshift, gYshift + (HC.polarplot.Height * 3 / 4))-(gXshift + HC.polarplot.width, gYshift + (HC.polarplot.Height * 3 / 4)), vbRed
    HC.polarplot.Line (gXshift + (HC.polarplot.width / 2), gYshift)-(gXshift + (HC.polarplot.width / 2), gYshift + HC.polarplot.Height), vbRed


End Sub

Function pxlate_x(inpx As Double, pscale As Double) As Double

    pxlate_x = (HC.polarplot.width / 2) - (inpx * pscale / (gTot_step / 2)) + gXshift
    
End Function
Function pxlate_y(inpy As Double, pscale As Double) As Double


    pxlate_y = (HC.polarplot.Height * 3 / 4) + (inpy * pscale / (gTot_step / 2)) + gYshift
    

End Function



Public Sub NStar_Polar_plot(RA1 As Double, DEC1 As Double, RA2 As Double, DEC2 As Double, pscale As Double)
Dim i As Double
Dim j As Double
Dim k As Double
Dim raprobe As Double


Dim tmpobj As Coord
Dim tmpobj2 As Coord
Dim datholder(1 To MAX_STARS) As Double
Dim dotidholder(1 To MAX_STARS) As Double

    If (gThreeStarEnable = False) Then
    
        Exit Sub
        
    End If

    raprobe = HC.HScroll4.Value
    raprobe = raprobe * 100

    tmpobj.x = RA1
    tmpobj.Y = DEC1

    HC.polarplot.Circle (pxlate_x(EQ_pl2Cs(tmpobj).x, pscale), pxlate_y(EQ_pl2Cs(tmpobj).Y, pscale)), 30, vbYellow
    HC.polarplot.Line (pxlate_x(0, pscale), pxlate_y(0, pscale))-(pxlate_x(EQ_pl2Cs(tmpobj).x, pscale), pxlate_y(EQ_pl2Cs(tmpobj).Y, pscale)), vbYellow
    
    tmpobj.x = RA1 - raprobe
    tmpobj2.x = RA1 + raprobe
    tmpobj2.Y = DEC1
    
    HC.polarplot.Line (pxlate_x(EQ_pl2Cs(tmpobj).x, pscale), pxlate_y(EQ_pl2Cs(tmpobj).Y, pscale))-(pxlate_x(EQ_pl2Cs(tmpobj2).x, pscale), pxlate_y(EQ_pl2Cs(tmpobj2).Y, pscale)), vbBlue
    
    
    tmpobj.x = RA2
    tmpobj.Y = DEC2
    
    HC.polarplot.Circle (pxlate_x(EQ_pl2Cs(tmpobj).x, pscale), pxlate_y(EQ_pl2Cs(tmpobj).Y, pscale)), 30, vbGreen
    HC.polarplot.Line (pxlate_x(0, pscale), pxlate_y(0, pscale))-(pxlate_x(EQ_pl2Cs(tmpobj).x, pscale), pxlate_y(EQ_pl2Cs(tmpobj).Y, pscale)), vbGreen
   
    tmpobj.x = RA2 - raprobe
    tmpobj2.x = RA2 + raprobe
    tmpobj2.Y = DEC2
    
    HC.polarplot.Line (pxlate_x(EQ_pl2Cs(tmpobj).x, pscale), pxlate_y(EQ_pl2Cs(tmpobj).Y, pscale))-(pxlate_x(EQ_pl2Cs(tmpobj2).x, pscale), pxlate_y(EQ_pl2Cs(tmpobj2).Y, pscale)), vbBlue
    
    tmpobj.x = RA1
    tmpobj.Y = DEC1

    For i = 1 To gAlignmentStars_count

          HC.polarplot.Circle (pxlate_x(EQ_pl2Cs(ct_Points(i)).x, pscale), pxlate_y(EQ_pl2Cs(ct_Points(i)).Y, pscale)), 30, vbCyan
         ' Compute for total X-Y distance.
         datholder(i) = Abs(my_PointsC(i).x - EQ_sp2Cs(tmpobj).x) + Abs(my_PointsC(i).Y - EQ_sp2Cs(tmpobj).Y)
         ' Also save the reference star id for this particular reference star
         dotidholder(i) = i
    
    Next i
    
    Call EQ_Quicksort(datholder(), dotidholder(), 1, gAlignmentStars_count)
    ' Get the nearest Star (lowest at the head of the sorted list)
    i = dotidholder(1)
    j = dotidholder(2)
    k = dotidholder(3)

    HC.polarplot.Line (pxlate_x(EQ_pl2Cs(my_Points(i)).x, pscale), pxlate_y(EQ_pl2Cs(my_Points(i)).Y, pscale))-(pxlate_x(EQ_pl2Cs(my_Points(j)).x, pscale), pxlate_y(EQ_pl2Cs(my_Points(j)).Y, pscale)), vbYellow
    HC.polarplot.Line (pxlate_x(EQ_pl2Cs(my_Points(j)).x, pscale), pxlate_y(EQ_pl2Cs(my_Points(j)).Y, pscale))-(pxlate_x(EQ_pl2Cs(my_Points(k)).x, pscale), pxlate_y(EQ_pl2Cs(my_Points(k)).Y, pscale)), vbYellow
    HC.polarplot.Line (pxlate_x(EQ_pl2Cs(my_Points(i)).x, pscale), pxlate_y(EQ_pl2Cs(my_Points(i)).Y, pscale))-(pxlate_x(EQ_pl2Cs(my_Points(k)).x, pscale), pxlate_y(EQ_pl2Cs(my_Points(k)).Y, pscale)), vbYellow
    
    HC.polarplot.Line (pxlate_x(EQ_pl2Cs(ct_Points(i)).x, pscale), pxlate_y(EQ_pl2Cs(ct_Points(i)).Y, pscale))-(pxlate_x(EQ_pl2Cs(ct_Points(j)).x, pscale), pxlate_y(EQ_pl2Cs(ct_Points(j)).Y, pscale)), vbBlue
    HC.polarplot.Line (pxlate_x(EQ_pl2Cs(ct_Points(j)).x, pscale), pxlate_y(EQ_pl2Cs(ct_Points(j)).Y, pscale))-(pxlate_x(EQ_pl2Cs(ct_Points(k)).x, pscale), pxlate_y(EQ_pl2Cs(ct_Points(k)).Y, pscale)), vbBlue
    HC.polarplot.Line (pxlate_x(EQ_pl2Cs(ct_Points(i)).x, pscale), pxlate_y(EQ_pl2Cs(ct_Points(i)).Y, pscale))-(pxlate_x(EQ_pl2Cs(ct_Points(k)).x, pscale), pxlate_y(EQ_pl2Cs(ct_Points(k)).Y, pscale)), vbBlue
    
    
    tmpobj.x = RA2
    tmpobj.Y = DEC2
    
    For i = 1 To gAlignmentStars_count
         ' Compute for total X-Y distance.
         datholder(i) = Abs(my_PointsC(i).x - EQ_sp2Cs(tmpobj).x) + Abs(my_PointsC(i).Y - EQ_sp2Cs(tmpobj).Y)
         
         ' Also save the reference star id for this particular reference star
         dotidholder(i) = i
    Next i
    Call EQ_Quicksort(datholder(), dotidholder(), 1, gAlignmentStars_count)
    ' Get the nearest Star (lowest at the head of the sorted list)
    i = dotidholder(1)
    j = dotidholder(2)
    k = dotidholder(3)
 
 
    HC.polarplot.Line (pxlate_x(EQ_pl2Cs(my_Points(i)).x, pscale), pxlate_y(EQ_pl2Cs(my_Points(i)).Y, pscale))-(pxlate_x(EQ_pl2Cs(my_Points(j)).x, pscale), pxlate_y(EQ_pl2Cs(my_Points(j)).Y, pscale)), vbGreen
    HC.polarplot.Line (pxlate_x(EQ_pl2Cs(my_Points(j)).x, pscale), pxlate_y(EQ_pl2Cs(my_Points(j)).Y, pscale))-(pxlate_x(EQ_pl2Cs(my_Points(k)).x, pscale), pxlate_y(EQ_pl2Cs(my_Points(k)).Y, pscale)), vbGreen
    HC.polarplot.Line (pxlate_x(EQ_pl2Cs(my_Points(i)).x, pscale), pxlate_y(EQ_pl2Cs(my_Points(i)).Y, pscale))-(pxlate_x(EQ_pl2Cs(my_Points(k)).x, pscale), pxlate_y(EQ_pl2Cs(my_Points(k)).Y, pscale)), vbGreen
    
 
    HC.polarplot.Line (pxlate_x(EQ_pl2Cs(ct_Points(i)).x, pscale), pxlate_y(EQ_pl2Cs(ct_Points(i)).Y, pscale))-(pxlate_x(EQ_pl2Cs(ct_Points(j)).x, pscale), pxlate_y(EQ_pl2Cs(ct_Points(j)).Y, pscale)), vbBlue
    HC.polarplot.Line (pxlate_x(EQ_pl2Cs(ct_Points(j)).x, pscale), pxlate_y(EQ_pl2Cs(ct_Points(j)).Y, pscale))-(pxlate_x(EQ_pl2Cs(ct_Points(k)).x, pscale), pxlate_y(EQ_pl2Cs(ct_Points(k)).Y, pscale)), vbBlue
    HC.polarplot.Line (pxlate_x(EQ_pl2Cs(ct_Points(i)).x, pscale), pxlate_y(EQ_pl2Cs(ct_Points(i)).Y, pscale))-(pxlate_x(EQ_pl2Cs(ct_Points(k)).x, pscale), pxlate_y(EQ_pl2Cs(ct_Points(k)).Y, pscale)), vbBlue
    
End Sub


Public Function PolarAlignDrift_Map(ByVal RA1 As Double, ByVal DEC1 As Double, ByVal RA2 As Double, ByVal DEC2 As Double, ByVal raprobe As Double, ByVal pscale As Double) As Coord

Dim obtmp2 As Coord
Dim dy1 As Double
Dim dy2 As Double

    If (RA1 >= &H1000000) Or (DEC1 >= &H1000000) Or (gThreeStarEnable = False) Then
    
        PolarAlignDrift_Map.x = 0
        PolarAlignDrift_Map.Y = 0
        PolarAlignDrift_Map.z = 0
        
        Exit Function
        
    End If
    

    
    ' re transform using the 3 nearest stars
    
    HC.EncoderTimer.Enabled = False
    
    dy1 = EQGet_Polar_Offset(RA1, DEC1, gDECEncoder_Home_pos - DEC1, raprobe, pscale)
    dy2 = EQGet_Polar_Offset(RA2, DEC2, gDECEncoder_Home_pos - DEC1, raprobe, pscale)

    
    HC.EncoderTimer.Enabled = True

    obtmp2 = EQNormalize_Polar(dy1, dy2, RAEncoder_Home_pos - RA1)
    
    PolarAlignDrift_Map.x = obtmp2.x
    PolarAlignDrift_Map.Y = obtmp2.Y
    PolarAlignDrift_Map.z = 1
    
    
End Function

Public Sub Position_polar(pscale As Double)
Dim vh As Double
Dim vy As Double
Dim RA1 As Double
Dim DEC1 As Double
Dim RA2 As Double
Dim DEC2 As Double
Dim obtmp As Coord
Dim raprobe As Double

    If (gThreeStarEnable = False) Then
    
        Exit Sub
        
    End If

    raprobe = HC.HScroll4.Value
    raprobe = raprobe * 100

    vh = HC.HScroll2.Value
    vh = (vh / 360) * gTot_step
    
    vy = 90 + HC.HScroll3.Value
    vy = (vy / 360) * gTot_step

    RA1 = RAEncoder_Home_pos + vh
    DEC1 = gDECEncoder_Home_pos - vy
    RA2 = RAEncoder_Home_pos + vh - (gTot_step / 4)
    DEC2 = gDECEncoder_Home_pos + vy
    

    NStar_Polar_plot_init (pscale)
    Call NStar_Polar_plot(RA1, DEC1, RA2, DEC2, pscale)
    obtmp = PolarAlignDrift_Map(RA1, DEC1, RA2, DEC2, raprobe, pscale)
'    HC.Label62.Caption = Format(obtmp.x * 0.0024, "####0.0000000000")   '.144 * 60
'    HC.Label64.Caption = Format(obtmp.y * 0.0024, "####0.0000000000")
    

End Sub
