Attribute VB_Name = "eqmodvector"
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
' EQMODVECTOR.BAS - Matrix Transformation Routines for 3-Star Alignment
'
' Written:  10-Dec-06   Raymund Sarmiento
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' 10-Dec-06 rcs     Initial edit for EQ Mount 3-Star Matrix Transformation
' 14-Dec-06 rcs     Added Taki Method on top of Affine Mapping Method for Comparison
'                   Taki Routines based on John Archbold's Excel computation
' 08-Apr-07 rcs     N-star implementation
'---------------------------------------------------------------------
'
'
'  SYNOPSIS:
'
'  This is a demonstration of the 3-Star Alignment Algorithm for the EQContrl.DLL
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


Option Explicit

Public Type Coord
 x As Double           'x = X Coordinate
 Y As Double           'y = Y Coordinate
 z As Double
End Type

Type Tdatholder
    dat As Double
    idx As Integer
    cc  As Coord    ' cartesian coordinate
End Type

Type THolder
    a As Double
    b As Double
    c As Double
End Type


Public Type Matrix
    Element(1 To 3, 1 To 3) As Double    '2D array of elements
End Type

Public Type Matrix2
    Element(1 To 4, 1 To 4) As Double    '2D array of elements
End Type







Public Type Coordt
 x As Double           'x = X Coordinate
 Y As Double           'y = Y Coordinate
 z As Double
 f As Integer
End Type

Public Type CartesCoord
 x As Double           'x = X Coordinate
 Y As Double           'y = Y Coordinate
 r As Double           ' Radius Sign
 RA As Double          ' Radius Alpha
End Type

Public Type SphereCoord
 x As Double           'x = X Coordinate
 Y As Double           'y = Y Coordinate
 r As Double           'r = RA Range Flag
End Type


Public Type TriangleCoord
 i As Double           ' Offset 1
 j As Double           ' Offset 2
 k As Double           ' offset 3

End Type

'Define Affine Matrix

Public EQMP As Matrix
Public EQMQ As Matrix

Public EQMI As Matrix
Public EQMM As Matrix
Public EQCO As Coord


'Define Taki Matrix

Public EQLMN1 As Matrix
Public EQLMN2 As Matrix

Public EQMI_T As Matrix
Public EQMT As Matrix
Public EQCT As Coord




'Function to put coordinate values into a LMN/lmn matrix array

Public Function GETLMN(ByRef p1 As Coord, ByRef p2 As Coord, ByRef p3 As Coord) As Matrix

Dim temp As Matrix
Dim UnitVect As Matrix

    With temp
    
        .Element(1, 1) = p2.x - p1.x
        .Element(2, 1) = p3.x - p1.x
        
        .Element(1, 2) = p2.Y - p1.Y
        .Element(2, 2) = p3.Y - p1.Y
        
        .Element(1, 3) = p2.z - p1.z
        .Element(2, 3) = p3.z - p1.z
            
    End With


    With UnitVect
    
        .Element(1, 1) = (temp.Element(1, 2) * temp.Element(2, 3)) - (temp.Element(1, 3) * temp.Element(2, 2))
        .Element(1, 2) = (temp.Element(1, 3) * temp.Element(2, 1)) - (temp.Element(1, 1) * temp.Element(2, 3))
        .Element(1, 3) = (temp.Element(1, 1) * temp.Element(2, 2)) - (temp.Element(1, 2) * temp.Element(2, 1))
        .Element(2, 1) = .Element(1, 1) ^ 2 + .Element(1, 2) ^ 2 + .Element(1, 3) ^ 2
        .Element(2, 2) = Sqr(.Element(2, 1))
        If .Element(2, 2) <> 0 Then .Element(2, 3) = 1 / .Element(2, 2)
        
    End With

    With temp
    
        .Element(3, 1) = UnitVect.Element(2, 3) * UnitVect.Element(1, 1)
        .Element(3, 2) = UnitVect.Element(2, 3) * UnitVect.Element(1, 2)
        .Element(3, 3) = UnitVect.Element(2, 3) * UnitVect.Element(1, 3)
   
    End With



    GETLMN = temp
    
End Function

'Function to put coordinate values into a P/Q Affine matrix array

Public Function GETPQ(ByRef p1 As Coord, ByRef p2 As Coord, ByRef p3 As Coord) As Matrix

Dim temp As Matrix

    With temp
        .Element(1, 1) = p2.x - p1.x
        .Element(2, 1) = p3.x - p1.x
        .Element(1, 2) = p2.Y - p1.Y
        .Element(2, 2) = p3.Y - p1.Y
    End With

    GETPQ = temp
    
End Function

' Subroutine to draw the Transformation Matrix (Taki Method)

Public Function EQ_AssembleMatrix_Taki(x As Double, Y As Double, ByRef a1 As Coord, ByRef a2 As Coord, ByRef a3 As Coord, ByRef m1 As Coord, ByRef m2 As Coord, ByRef m3 As Coord) As Integer


Dim Det As Double
    
     
    ' Get the LMN Matrix

    EQLMN1 = GETLMN(a1, a2, a3)
    
    ' Get the lmn Matrix
    
    EQLMN2 = GETLMN(m1, m2, m3)

   
    
    With EQLMN1
    
        ' Get the Determinant
        
        Det = .Element(1, 1) * ((.Element(2, 2) * .Element(3, 3)) - (.Element(3, 2) * .Element(2, 3)))
        Det = Det - (.Element(1, 2) * ((.Element(2, 1) * .Element(3, 3)) - (.Element(3, 1) * .Element(2, 3))))
        Det = Det + (.Element(1, 3) * ((.Element(2, 1) * .Element(3, 2)) - (.Element(3, 1) * .Element(2, 2))))
        
        
        ' Compute for the Matrix Inverse of EQLMN1
        
        If Det = 0 Then
            Err.Raise 999, "AssembleMatrix", "Cannot invert matrix with Determinant = 0"
        Else
    
            EQMI_T.Element(1, 1) = ((.Element(2, 2) * .Element(3, 3)) - (.Element(3, 2) * .Element(2, 3))) / Det
            EQMI_T.Element(1, 2) = ((.Element(1, 3) * .Element(3, 2)) - (.Element(1, 2) * .Element(3, 3))) / Det
            EQMI_T.Element(1, 3) = ((.Element(1, 2) * .Element(2, 3)) - (.Element(2, 2) * .Element(1, 3))) / Det
            EQMI_T.Element(2, 1) = ((.Element(2, 3) * .Element(3, 1)) - (.Element(3, 3) * .Element(2, 1))) / Det
            EQMI_T.Element(2, 2) = ((.Element(1, 1) * .Element(3, 3)) - (.Element(3, 1) * .Element(1, 3))) / Det
            EQMI_T.Element(2, 3) = ((.Element(1, 3) * .Element(2, 1)) - (.Element(2, 3) * .Element(1, 1))) / Det
            EQMI_T.Element(3, 1) = ((.Element(2, 1) * .Element(3, 2)) - (.Element(3, 1) * .Element(2, 2))) / Det
            EQMI_T.Element(3, 2) = ((.Element(1, 2) * .Element(3, 1)) - (.Element(3, 2) * .Element(1, 1))) / Det
            EQMI_T.Element(3, 3) = ((.Element(1, 1) * .Element(2, 2)) - (.Element(2, 1) * .Element(1, 2))) / Det
        End If
    
    End With
   
   
    ' Get the M Matrix by Multiplying EQMI and EQLMN2
    ' EQMI_T - Matrix A
    ' EQLMN2 - Matrix B
        
    
    EQMT.Element(1, 1) = (EQMI_T.Element(1, 1) * EQLMN2.Element(1, 1)) + (EQMI_T.Element(1, 2) * EQLMN2.Element(2, 1)) + (EQMI_T.Element(1, 3) * EQLMN2.Element(3, 1))
    EQMT.Element(1, 2) = (EQMI_T.Element(1, 1) * EQLMN2.Element(1, 2)) + (EQMI_T.Element(1, 2) * EQLMN2.Element(2, 2)) + (EQMI_T.Element(1, 3) * EQLMN2.Element(3, 2))
    EQMT.Element(1, 3) = (EQMI_T.Element(1, 1) * EQLMN2.Element(1, 3)) + (EQMI_T.Element(1, 2) * EQLMN2.Element(2, 3)) + (EQMI_T.Element(1, 3) * EQLMN2.Element(3, 3))
    
    EQMT.Element(2, 1) = (EQMI_T.Element(2, 1) * EQLMN2.Element(1, 1)) + (EQMI_T.Element(2, 2) * EQLMN2.Element(2, 1)) + (EQMI_T.Element(2, 3) * EQLMN2.Element(3, 1))
    EQMT.Element(2, 2) = (EQMI_T.Element(2, 1) * EQLMN2.Element(1, 2)) + (EQMI_T.Element(2, 2) * EQLMN2.Element(2, 2)) + (EQMI_T.Element(2, 3) * EQLMN2.Element(3, 2))
    EQMT.Element(2, 3) = (EQMI_T.Element(2, 1) * EQLMN2.Element(1, 3)) + (EQMI_T.Element(2, 2) * EQLMN2.Element(2, 3)) + (EQMI_T.Element(2, 3) * EQLMN2.Element(3, 3))
    
    EQMT.Element(3, 1) = (EQMI_T.Element(3, 1) * EQLMN2.Element(1, 1)) + (EQMI_T.Element(3, 2) * EQLMN2.Element(2, 1)) + (EQMI_T.Element(3, 3) * EQLMN2.Element(3, 1))
    EQMT.Element(3, 2) = (EQMI_T.Element(3, 1) * EQLMN2.Element(1, 2)) + (EQMI_T.Element(3, 2) * EQLMN2.Element(2, 2)) + (EQMI_T.Element(3, 3) * EQLMN2.Element(3, 2))
    EQMT.Element(3, 3) = (EQMI_T.Element(3, 1) * EQLMN2.Element(1, 3)) + (EQMI_T.Element(3, 2) * EQLMN2.Element(2, 3)) + (EQMI_T.Element(3, 3) * EQLMN2.Element(3, 3))
    
        
    ' Get the Coordinate Offset Vector and store it at EQCO Matrix

    EQCT.x = m1.x - ((a1.x * EQMT.Element(1, 1)) + (a1.Y * EQMT.Element(2, 1)) + (a1.z * EQMT.Element(3, 1)))
    EQCT.Y = m1.Y - ((a1.x * EQMT.Element(1, 2)) + (a1.Y * EQMT.Element(2, 2)) + (a1.z * EQMT.Element(3, 2)))
    EQCT.z = m1.z - ((a1.x * EQMT.Element(1, 3)) + (a1.Y * EQMT.Element(2, 3)) + (a1.z * EQMT.Element(3, 3)))
    
    
     If (x + Y) = 0 Then
        EQ_AssembleMatrix_Taki = 0
     Else
        EQ_AssembleMatrix_Taki = EQ_CheckPoint_in_Triangle(x, Y, a1.x, a1.Y, a2.x, a2.Y, a3.x, a3.Y)
     End If
    

End Function


'Function to transform the Coordinates (Taki Method)  using the MT Matrix and Offset Vector

Public Function EQ_Transform_Taki(ByRef ob As Coord) As Coord

    ' CoordTransform = Offset + CoordObject * Matrix MT

    EQ_Transform_Taki.x = EQCT.x + ((ob.x * EQMT.Element(1, 1)) + (ob.Y * EQMT.Element(2, 1)) + (ob.z * EQMT.Element(3, 1)))
    EQ_Transform_Taki.Y = EQCT.Y + ((ob.x * EQMT.Element(1, 2)) + (ob.Y * EQMT.Element(2, 2)) + (ob.z * EQMT.Element(3, 2)))
    EQ_Transform_Taki.z = EQCT.z + ((ob.x * EQMT.Element(1, 3)) + (ob.Y * EQMT.Element(2, 3)) + (ob.z * EQMT.Element(3, 3)))


End Function

' Subroutine to draw the Transformation Matrix (Affine Mapping Method)

Public Function EQ_AssembleMatrix_Affine(x As Double, Y As Double, ByRef a1 As Coord, ByRef a2 As Coord, ByRef a3 As Coord, ByRef m1 As Coord, ByRef m2 As Coord, ByRef m3 As Coord) As Integer

Dim Det As Double
    
    ' Get the P Matrix
    EQMP = GETPQ(a1, a2, a3)
    
    ' Get the Q Matrix
    EQMQ = GETPQ(m1, m2, m3)

    ' Get the Inverse of P
    With EQMP
        ' Get the EQMP Determinant for Inverse Computation
        Det = (.Element(1, 1) * .Element(2, 2)) - (.Element(1, 2) * .Element(2, 1))

        ' Make sure Determinant is NON ZERO
        If Det = 0 Then
            Err.Raise 999, "AssembleMatrix", "Cannot invert matrix with Determinant = 0"
        Else
            'Perform the Matrix Inversion, put result to EQMI matrix
            EQMI.Element(1, 1) = (.Element(2, 2)) / Det
            EQMI.Element(1, 2) = (-.Element(1, 2)) / Det
            EQMI.Element(2, 1) = (-.Element(2, 1)) / Det
            EQMI.Element(2, 2) = (.Element(1, 1)) / Det
        End If
    End With
   
    ' Get the M Matrix by Multiplying EQMI and EQMQ
    ' EQMI - Matrix A
    ' EQMQ - Matrix B
    EQMM.Element(1, 1) = (EQMI.Element(1, 1) * EQMQ.Element(1, 1)) + (EQMI.Element(1, 2) * EQMQ.Element(2, 1))
    EQMM.Element(1, 2) = (EQMI.Element(1, 1) * EQMQ.Element(1, 2)) + (EQMI.Element(1, 2) * EQMQ.Element(2, 2))
    EQMM.Element(2, 1) = (EQMI.Element(2, 1) * EQMQ.Element(1, 1)) + (EQMI.Element(2, 2) * EQMQ.Element(2, 1))
    EQMM.Element(2, 2) = (EQMI.Element(2, 1) * EQMQ.Element(1, 2)) + (EQMI.Element(2, 2) * EQMQ.Element(2, 2))
    
    ' Get the Coordinate Offset Vector and store it at EQCO Matrix
    EQCO.x = m1.x - ((a1.x * EQMM.Element(1, 1)) + (a1.Y * EQMM.Element(2, 1)))
    EQCO.Y = m1.Y - ((a1.x * EQMM.Element(1, 2)) + (a1.Y * EQMM.Element(2, 2)))

    If (x + Y) = 0 Then
       EQ_AssembleMatrix_Affine = 0
    Else
       EQ_AssembleMatrix_Affine = EQ_CheckPoint_in_Triangle(x, Y, m1.x, m1.Y, m2.x, m2.Y, m3.x, m3.Y)
    End If

End Function


'Function to transform the Coordinates (Affine Mapping) using the M Matrix and Offset Vector

Public Function EQ_Transform_Affine(ByRef ob As Coord) As Coord

    ' CoordTransform = Offset + CoordObject * Matrix M

    EQ_Transform_Affine.x = EQCO.x + ((ob.x * EQMM.Element(1, 1)) + (ob.Y * EQMM.Element(2, 1)))
    EQ_Transform_Affine.Y = EQCO.Y + ((ob.x * EQMM.Element(1, 2)) + (ob.Y * EQMM.Element(2, 2)))

End Function

'Function to convert spherical coordinates to Cartesian using the Coord structure

Public Function EQ_sp2Cs(ByRef obj As Coord) As Coord

Dim tmpobj As CartesCoord
Dim tmpobj4 As SphereCoord

    If HC.PolarEnable.Value = 1 Then
        tmpobj4 = EQ_SphericalPolar(obj.x, obj.Y, gTot_step, RAEncoder_Home_pos, gDECEncoder_Home_pos, gLatitude)
        tmpobj = EQ_Polar2Cartes(tmpobj4.x, tmpobj4.Y, gTot_step, RAEncoder_Home_pos, gDECEncoder_Home_pos)
        EQ_sp2Cs.x = tmpobj.x
        EQ_sp2Cs.Y = tmpobj.Y
        EQ_sp2Cs.z = 1
    Else
        EQ_sp2Cs.x = obj.x
        EQ_sp2Cs.Y = obj.Y
        EQ_sp2Cs.z = 1
    End If
    
End Function


'Function to convert spherical coordinates to Cartesian using the Coord structure

Public Function EQ_sp2Cs2(ByRef obj As Coord) As Coord

Dim tmpobj As CartesCoord
Dim tmpobj4 As SphereCoord
Dim lat As Double

    If HC.PolarEnable.Value = 1 Then
        tmpobj4 = EQ_SphericalPolar(obj.x, obj.Y, gTot_step, RAEncoder_Home_pos, gDECEncoder_Home_pos, Abs(gLatitude))
        tmpobj = EQ_Polar2Cartes(tmpobj4.x, tmpobj4.Y, gTot_step, RAEncoder_Home_pos, gDECEncoder_Home_pos)
        EQ_sp2Cs2.x = tmpobj.x
        EQ_sp2Cs2.Y = tmpobj.Y
        EQ_sp2Cs2.z = 1
    Else
        EQ_sp2Cs2.x = obj.x
        EQ_sp2Cs2.Y = obj.Y
        EQ_sp2Cs2.z = 1
    End If
    
End Function


'Function to convert polar coordinates to Cartesian using the Coord structure


Public Function EQ_pl2Cs(ByRef obj As Coord) As Coord

Dim tmpobj As CartesCoord

    If HC.PolarEnable.Value = 1 Then
        tmpobj = EQ_Polar2Cartes(obj.x, obj.Y, gTot_step, RAEncoder_Home_pos, gDECEncoder_Home_pos)
    
        EQ_pl2Cs.x = tmpobj.x
        EQ_pl2Cs.Y = tmpobj.Y
        EQ_pl2Cs.z = 1
    Else
        EQ_pl2Cs.x = obj.x
        EQ_pl2Cs.Y = obj.Y
        EQ_pl2Cs.z = 1
    End If
    
End Function

'Implement an Affine transformation on a Polar coordinate system
'This is done by converting the Polar Data to Cartesian, Apply affine transformation
'Then restore the transformed Cartesian Coordinates back to polar


Public Function EQ_plAffine(ByRef obj As Coord) As Coord

Dim tmpobj1 As CartesCoord
Dim tmpobj2 As Coord
Dim tmpobj3 As Coord
Dim tmpobj4 As SphereCoord

       If HC.PolarEnable.Value = 1 Then
            tmpobj4 = EQ_SphericalPolar(obj.x, obj.Y, gTot_step, RAEncoder_Home_pos, gDECEncoder_Home_pos, gLatitude)

            tmpobj1 = EQ_Polar2Cartes(tmpobj4.x, tmpobj4.Y, gTot_step, RAEncoder_Home_pos, gDECEncoder_Home_pos)
    
            tmpobj2.x = tmpobj1.x
            tmpobj2.Y = tmpobj1.Y
            tmpobj2.z = 1
    
            tmpobj3 = EQ_Transform_Affine(tmpobj2)
    
            tmpobj2 = EQ_Cartes2Polar(tmpobj3.x, tmpobj3.Y, tmpobj1.r, tmpobj1.RA, gTot_step, RAEncoder_Home_pos, gDECEncoder_Home_pos)

            EQ_plAffine = EQ_PolarSpherical(tmpobj2.x, tmpobj2.Y, gTot_step, RAEncoder_Home_pos, gDECEncoder_Home_pos, gLatitude, tmpobj4.r)

        Else
            tmpobj3 = EQ_Transform_Affine(obj)
            EQ_plAffine.x = tmpobj3.x
            EQ_plAffine.Y = tmpobj3.Y
            EQ_plAffine.z = 1
        End If

End Function


Public Function EQ_plAffine2(ByRef obj As Coord) As Coord

Dim tmpobj1 As CartesCoord
Dim tmpobj2 As Coord
Dim tmpobj3 As Coord
Dim tmpobj4 As SphereCoord

       If HC.PolarEnable.Value = 1 Then
            tmpobj4 = EQ_SphericalPolar(obj.x, obj.Y, gTot_step, RAEncoder_Home_pos, gDECEncoder_Home_pos, gLatitude)

            tmpobj1 = EQ_Polar2Cartes(tmpobj4.x, tmpobj4.Y, gTot_step, RAEncoder_Home_pos, gDECEncoder_Home_pos)
    
            tmpobj2.x = tmpobj1.x
            tmpobj2.Y = tmpobj1.Y
            tmpobj2.z = 1
    
            tmpobj3 = EQ_Transform_Affine(tmpobj2)
    
            tmpobj2 = EQ_Cartes2Polar(tmpobj3.x, tmpobj3.Y, tmpobj1.r, tmpobj1.RA, gTot_step, RAEncoder_Home_pos, gDECEncoder_Home_pos)


            EQ_plAffine2 = EQ_PolarSpherical(tmpobj2.x, tmpobj2.Y, gTot_step, RAEncoder_Home_pos, gDECEncoder_Home_pos, gLatitude, tmpobj4.r)

        Else
            tmpobj3 = EQ_Transform_Affine(obj)
            EQ_plAffine2.x = tmpobj3.x
            EQ_plAffine2.Y = tmpobj3.Y
            EQ_plAffine2.z = 1
        End If

End Function
'Implement a TAKI transformation on a Polar coordinate system
'This is done by converting the Polar Data to Cartesian, Apply TAKI transformation
'Then restore the transformed Cartesian Coordinates back to polar

Public Function EQ_plTaki(ByRef obj As Coord) As Coord

Dim tmpobj1 As CartesCoord
Dim tmpobj2 As Coord
Dim tmpobj3 As Coord
Dim tmpobj4 As SphereCoord

    If HC.PolarEnable.Value = 1 Then
        tmpobj4 = EQ_SphericalPolar(obj.x, obj.Y, gTot_step, RAEncoder_Home_pos, gDECEncoder_Home_pos, gLatitude)
        tmpobj1 = EQ_Polar2Cartes(tmpobj4.x, tmpobj4.Y, gTot_step, RAEncoder_Home_pos, gDECEncoder_Home_pos)
    
        tmpobj2.x = tmpobj1.x
        tmpobj2.Y = tmpobj1.Y
        tmpobj2.z = 1
   
        tmpobj3 = EQ_Transform_Taki(tmpobj2)
        
        tmpobj2 = EQ_Cartes2Polar(tmpobj3.x, tmpobj3.Y, tmpobj1.r, tmpobj1.RA, gTot_step, RAEncoder_Home_pos, gDECEncoder_Home_pos)
    
        EQ_plTaki = EQ_PolarSpherical(tmpobj2.x, tmpobj2.Y, gTot_step, RAEncoder_Home_pos, gDECEncoder_Home_pos, gLatitude, tmpobj4.r)

    Else
        tmpobj3 = EQ_Transform_Taki(obj)
        EQ_plTaki.x = tmpobj3.x
        EQ_plTaki.Y = tmpobj3.Y
        EQ_plTaki.z = 1
    End If

End Function

' Function to Convert Polar RA/DEC Stepper coordinates to Cartesian Coordinates

Public Function EQ_Polar2Cartes(RA As Double, DEC As Double, TOT As Double, RACENTER As Double, DECCENTER As Double) As CartesCoord


Dim x2 As Double
Dim y2 As Double

Dim theta As Double
Dim radius As Double

Dim angle As Double
Dim radiusder As Double
Dim i As Double

Dim radpeak As Double


    ' make angle stays within the 360 bound

    If RA > RACENTER Then
        i = ((RA - RACENTER) / TOT) * 360
    Else
        i = ((RACENTER - RA) / TOT) * 360
        i = 360 - i
    End If
    
    theta = Range360(i) * DEG_RAD

    'treat y as the radius of the polar coordinate

    radius = DEC - DECCENTER

    radpeak = 0
    
  '  Removed
    
  '  If Abs(radius) > DECPEAK Then
  '      radpeak = radius
  '      If radius > 0 Then
  '          radius = (2 * DECPEAK) - radius
  '      Else
  '          radius = ((2 * DECPEAK) + radius) * -1
  '      End If
  '      radpeak = radpeak - radius
  '  End If


    ' Avoid division 0 errors
    
    If radius = 0 Then radius = 1
    
  ' Get the cartesian coordinates
    
    EQ_Polar2Cartes.x = Cos(theta) * radius
    EQ_Polar2Cartes.Y = Sin(theta) * radius
    EQ_Polar2Cartes.RA = radpeak
    
  ' if radius is a negative number, pass this info on the next conversion routine
    
    If radius > 0 Then
        EQ_Polar2Cartes.r = 1
    Else
        EQ_Polar2Cartes.r = -1
    End If

End Function

'Function to convert the Cartesian Coordinate data back to RA/DEC polar

Public Function EQ_Cartes2Polar(x As Double, Y As Double, r As Double, RA As Double, TOT As Double, RACENTER As Double, DECCENTER As Double) As Coord

Dim radiusder As Double
Dim angle As Double

    ' Ah the famous radius formula

    radiusder = Sqr((x * x) + (Y * Y)) * r
    
    
    ' And the nasty angle compute routine (any simpler way to impelent this ?)
    
    angle = 0
    If x > 0 Then angle = Atn(Y / x)
    If x < 0 Then
        If Y >= 0 Then
          angle = Atn(Y / x) + PI
          
        Else
          angle = Atn(Y / x) - PI
        End If
    End If
    If x = 0 Then
        If Y > 0 Then
            angle = PI / 2
        Else
            angle = -1 * (PI / 2)
        End If
    End If
    
    ' Convert angle to degrees
    
    angle = angle * RAD_DEG
   
    If angle < 0 Then angle = 360 + angle
    
    If r < 0 Then angle = Range360(angle + 180)
    
    If (angle > 180) Then
            EQ_Cartes2Polar.x = RACENTER - (((360 - angle) / 360) * TOT)
        Else
            EQ_Cartes2Polar.x = ((angle / 360) * TOT) + RACENTER
    End If
    
    'treat y as the polar coordinate radius (ra var not used - always 0)
    
    EQ_Cartes2Polar.Y = radiusder + DECCENTER + RA
    
End Function

Public Function EQ_UpdateTaki(x As Double, Y As Double) As Integer

Dim tr As TriangleCoord
Dim tmpcoord As Coord

    ' Adjust only if there are four alignment stars
    If gAlignmentStars_count < 3 Then Exit Function


    Select Case g3PointAlgorithm
        Case 1
            ' find the 50 nearest points - then find the nearest enclosing triangle
            tr = EQ_ChooseNearest3Points(x, Y)
        Case Else
            ' find the 50 nearest points - then find the enclosing triangle with the nearest centre point
            tr = EQ_Choose_3Points(x, Y)
    End Select
    
    gTaki1 = tr.i
    gTaki2 = tr.j
    gTaki3 = tr.k
    
    If gTaki1 = 0 Or gTaki2 = 0 Or gTaki3 = 0 Then
        EQ_UpdateTaki = 0
        Exit Function
    End If
    
    tmpcoord.x = x
    tmpcoord.Y = Y
    tmpcoord = EQ_sp2Cs(tmpcoord)
    EQ_UpdateTaki = EQ_AssembleMatrix_Taki(tmpcoord.x, tmpcoord.Y, ct_PointsC(gTaki1), ct_PointsC(gTaki2), ct_PointsC(gTaki3), my_PointsC(gTaki1), my_PointsC(gTaki2), my_PointsC(gTaki3))
    
End Function

Public Function EQ_UpdateAffine(x As Double, Y As Double) As Integer

Dim tmpcoord As Coord
Dim tr As TriangleCoord

    If gAlignmentStars_count < 3 Then Exit Function
    
    Select Case g3PointAlgorithm
        Case 1
            ' find the 50 nearest points - then find the nearest enclosing triangle
            tr = EQ_ChooseNearest3Points(x, Y)
        Case Else
            ' find the 50 nearest points - then find the enclosing triangle with the nearest centre point
            tr = EQ_Choose_3Points(x, Y)
    End Select
    
    gAffine1 = tr.i
    gAffine2 = tr.j
    gAffine3 = tr.k
    
    If gAffine1 = 0 Or gAffine1 = 0 Or gAffine1 = 0 Then
        EQ_UpdateAffine = 0
        Exit Function
    End If
 
    tmpcoord.x = x
    tmpcoord.Y = Y
    tmpcoord = EQ_sp2Cs(tmpcoord)

    EQ_UpdateAffine = EQ_AssembleMatrix_Affine(tmpcoord.x, tmpcoord.Y, my_PointsC(gAffine1), my_PointsC(gAffine2), my_PointsC(gAffine3), ct_PointsC(gAffine1), ct_PointsC(gAffine2), ct_PointsC(gAffine3))
    
    If EQ_UpdateAffine = 0 Then
        gAffine1 = 0
        gAffine2 = 0
        gAffine3 = 0
    End If
    
End Function

' Subroutine to implement find Array index with the lowest value
Public Function EQ_FindLowest(List() As Double, min As Integer, max As Integer) As Integer
Dim val As Double
Dim newval As Double
Dim i As Integer
Dim idx As Integer
    
    idx = -1
    If min >= max Or max > UBound(List) Then GoTo endfn
        
    val = List(min)
    For i = min To max Step 1
        newval = List(i)
        If newval <= val Then
            val = newval
            idx = i
        End If
    Next i

endfn:
    EQ_FindLowest = idx
End Function

Public Sub EQ_FindLowest3(List() As Double, Sublist() As Integer, min As Integer, max As Integer)
Dim val As Double
Dim min1 As Double
Dim min2 As Double
Dim min3 As Double
Dim i As Integer
    
    If min >= max Or max > UBound(List) Then GoTo endfn
        
    If List(1) <= List(2) And List(1) <= List(3) Then
        'List 1 is first
        min1 = List(1)
        If List(2) <= List(3) Then
            'List2 is second
            'List3 is third
            min2 = List(2)
            min3 = List(3)
        Else
            'List3 is second
            'List2 is third
            min2 = List(3)
            min3 = List(2)
        End If
    Else
        If List(2) <= List(1) And List(2) <= List(3) Then
            'List 2 is first
            min1 = List(2)
            If List(1) <= List(3) Then
                'List1 is second
                'List3 is third
                min2 = List(1)
                min3 = List(3)
            Else
                'List3 is second
                'List1 is third
                min2 = List(3)
                min3 = List(1)
            End If
        Else
            If List(3) <= List(1) And List(3) <= List(2) Then
                'List 3 is first
                min1 = List(3)
                If List(1) <= List(2) Then
                    'List1 is second
                    'List2 is third
                    min2 = List(1)
                    min3 = List(2)
                Else
                    'List2 is second
                    'List1 is third
                    min2 = List(2)
                    min3 = List(1)
                End If
            End If
        End If
    End If
        
    val = List(min)
    
    For i = min To max Step 1
        val = List(i)
        If val < min1 Then
            min1 = val
            Sublist(3) = Sublist(2)
            Sublist(2) = Sublist(1)
            Sublist(1) = i
        Else
            If val < min2 Then
                min2 = val
                Sublist(3) = Sublist(2)
                Sublist(2) = i
            Else
                If val < min3 Then
                    Sublist(3) = i
                End If
            End If
        End If
    Next i
    
endfn:
    
End Sub




' Subroutine to implement an Array sort
Public Sub EQ_Quicksort(List() As Double, Sublist() As Double, min As Integer, max As Integer)

Dim med_value As Double
Dim submed As Double

Dim hi As Integer
Dim lo As Integer
Dim i As Integer

    If min >= max Then Exit Sub

    i = Int((max - min + 1) * Rnd + min)
    med_value = List(i)
    submed = Sublist(i)

    List(i) = List(min)
    Sublist(i) = Sublist(min)

    lo = min
    hi = max
    Do
        Do While List(hi) >= med_value
            hi = hi - 1
            If hi <= lo Then Exit Do
        Loop
        If hi <= lo Then
            List(lo) = med_value
            Sublist(lo) = submed
            Exit Do
        End If
        
        List(lo) = List(hi)
        Sublist(lo) = Sublist(hi)

        lo = lo + 1
        Do While List(lo) < med_value
            lo = lo + 1
            If lo >= hi Then Exit Do
        Loop
        
        If lo >= hi Then
            lo = hi
            List(hi) = med_value
            Sublist(hi) = submed
            Exit Do
        End If

        List(hi) = List(lo)
        Sublist(hi) = Sublist(lo)
    
    Loop

    EQ_Quicksort List(), Sublist(), min, lo - 1
    EQ_Quicksort List(), Sublist(), lo + 1, max
    
End Sub


' Subroutine to implement an Array sort

Public Sub EQ_Quicksort2(List() As Tdatholder, min As Integer, max As Integer)

Dim med_value As Tdatholder

Dim hi As Integer
Dim lo As Integer
Dim i As Integer

    If min >= max Then Exit Sub

    i = Int((max - min + 1) * Rnd + min)
    med_value = List(i)
    
    List(i) = List(min)
    
    lo = min
    hi = max
    
    Do
        Do While List(hi).dat >= med_value.dat
            hi = hi - 1
            If hi <= lo Then Exit Do
        Loop
        If hi <= lo Then
            List(lo) = med_value
            Exit Do
        End If
    
        List(lo) = List(hi)
    
        lo = lo + 1
        Do While List(lo).dat < med_value.dat
            lo = lo + 1
            If lo >= hi Then Exit Do
        Loop
        If lo >= hi Then
            lo = hi
            List(hi) = med_value
            Exit Do
        End If
    
        List(hi) = List(lo)
    Loop
    
    EQ_Quicksort2 List(), min, lo - 1
    EQ_Quicksort2 List(), lo + 1, max

End Sub

' Subroutine to implement an Array sort with three sublists

Public Sub EQ_Quicksort3(List() As Double, Sublist1() As Double, Sublist2() As Double, Sublist3() As Double, min As Integer, max As Integer)

Dim med_value As Double
Dim submed1 As Double
Dim submed2 As Double
Dim submed3 As Double

Dim hi As Integer
Dim lo As Integer
Dim i As Integer


    If min >= max Then Exit Sub

    i = Int((max - min + 1) * Rnd + min)
    med_value = List(i)
    submed1 = Sublist1(i)
    submed2 = Sublist2(i)
    submed3 = Sublist3(i)

    List(i) = List(min)
    Sublist1(i) = Sublist1(min)
    Sublist2(i) = Sublist2(min)
    Sublist3(i) = Sublist3(min)

    lo = min
    hi = max
    Do

        Do While List(hi) >= med_value
            hi = hi - 1
            If hi <= lo Then Exit Do
        Loop
        If hi <= lo Then
            List(lo) = med_value
            Sublist1(lo) = submed1
            Sublist2(lo) = submed2
            Sublist3(lo) = submed3
            Exit Do
        End If


        List(lo) = List(hi)
        Sublist1(lo) = Sublist1(hi)
        Sublist2(lo) = Sublist2(hi)
        Sublist3(lo) = Sublist3(hi)
        
        lo = lo + 1
        Do While List(lo) < med_value
            lo = lo + 1
            If lo >= hi Then Exit Do
        Loop
        If lo >= hi Then
            lo = hi
            List(hi) = med_value
            Sublist1(hi) = submed1
            Sublist2(hi) = submed2
            Sublist3(hi) = submed3
            Exit Do
        End If

        List(hi) = List(lo)
        Sublist1(hi) = Sublist1(lo)
        Sublist2(hi) = Sublist2(lo)
        Sublist3(hi) = Sublist3(lo)
    Loop

    EQ_Quicksort3 List(), Sublist1(), Sublist2(), Sublist3(), min, lo - 1
    EQ_Quicksort3 List(), Sublist1(), Sublist2(), Sublist3(), lo + 1, max
    
End Sub

' Function to compute for an area of a triangle

Public Function EQ_Triangle_Area(px1 As Double, py1 As Double, px2 As Double, py2 As Double, px3 As Double, py3 As Double) As Double

Dim ta As Double

'True formula is this
'    EQ_Triangle_Area = Abs(((px2 * py1) - (px1 * py2)) + ((px3 * py2) - (px2 * py3)) + ((px1 * py3) - (px3 * py1))) / 2

' Make LARGE  numerical value safe for Windows by adding a scaling factor

    ta = (((px2 * py1) - (px1 * py2)) / 10000) + (((px3 * py2) - (px2 * py3)) / 10000) + (((px1 * py3) - (px3 * py1)) / 10000)

    EQ_Triangle_Area = Abs(ta) / 2
    
End Function

' Function to check if a point is inside the triangle. Computed based sum of areas method

Public Function EQ_CheckPoint_in_Triangle(px As Double, py As Double, px1 As Double, py1 As Double, px2 As Double, py2 As Double, px3 As Double, py3 As Double) As Integer

Dim ta As Double
Dim t1 As Double
Dim t2 As Double
Dim t3 As Double

    ta = EQ_Triangle_Area(px1, py1, px2, py2, px3, py3)
    t1 = EQ_Triangle_Area(px, py, px2, py2, px3, py3)
    t2 = EQ_Triangle_Area(px1, py1, px, py, px3, py3)
    t3 = EQ_Triangle_Area(px1, py1, px2, py2, px, py)


    If Abs(ta - t1 - t2 - t3) < 2 Then
        EQ_CheckPoint_in_Triangle = 1
    Else
        EQ_CheckPoint_in_Triangle = 0
    End If

End Function




Public Function EQ_GetCenterPoint(p1 As Coord, p2 As Coord, p3 As Coord) As Coord

Dim p1x As Double
Dim p1y As Double
Dim p2x As Double
Dim p2y As Double
Dim p3x As Double
Dim p3y As Double
Dim p4x As Double
Dim p4y As Double

Dim XD1 As Double
Dim YD1 As Double
Dim XD2 As Double
Dim YD2 As Double
Dim XD3 As Double
Dim YD3 As Double

Dim ua As Double
Dim ub As Double
Dim dv As Double



' Get the two line 4 point data

    p1x = p1.x
    p1y = p1.Y
    
    
    If p3.x > p2.x Then
        p2x = ((p3.x - p2.x) / 2) + p2.x
    Else
        p2x = ((p2.x - p3.x) / 2) + p3.x
    End If
    
    If p3.Y > p2.Y Then
        p2y = ((p3.Y - p2.Y) / 2) + p2.Y
    Else
        p2y = ((p2.Y - p3.Y) / 2) + p3.Y
    End If
    
    p3x = p2.x
    p3y = p2.Y
    
    
    If p1.x > p3.x Then
        p4x = ((p1.x - p3.x) / 2) + p3.x
    Else
        p4x = ((p3.x - p1.x) / 2) + p1.x
    End If
    
    If p1.Y > p3.Y Then
        p4y = ((p1.Y - p3.Y) / 2) + p3.Y
    Else
        p4y = ((p3.Y - p1.Y) / 2) + p1.Y
    End If
    
    
    XD1 = p2x - p1x
    XD2 = p4x - p3x
    YD1 = p2y - p1y
    YD2 = p4y - p3y
    XD3 = p1x - p3x
    YD3 = p1y - p3y
    
   
    dv = (YD2 * XD1) - (XD2 * YD1)
    
    If dv = 0 Then dv = 0.00000001   'avoid div 0 errors
    
    
    ua = ((XD2 * YD3) - (YD2 * XD3)) / dv
    ub = ((XD1 * YD3) - (YD1 * XD3)) / dv
    
    EQ_GetCenterPoint.x = p1x + (ua * XD1)
    EQ_GetCenterPoint.Y = p1y + (ub * YD1)
    
End Function


Public Function EQ_SphericalPolar(RA As Double, DEC As Double, TOT As Double, RACENTER As Double, DECCENTER As Double, Latitude As Double) As SphereCoord
Dim i As Double
Dim j As Double
Dim x As Double
Dim Y As Double

    i = Get_EncoderHours(RACENTER, RA, TOT, 0)
    j = Get_EncoderDegrees(DECCENTER, DEC, TOT, 0) + 270
    j = Range360(j)
  
    Call hadec_aa(Latitude * DEG_RAD, i * HRS_RAD, j * DEG_RAD, Y, x)
   
    EQ_SphericalPolar.x = ((((x * RAD_DEG) - 180) / 360) * TOT) + RACENTER
    EQ_SphericalPolar.Y = ((((Y * RAD_DEG) + 90) / 180) * TOT) + DECCENTER
    
    ' Check if RA value is within allowed visible range
    i = TOT / 4
    If (RA <= (RACENTER + i)) And (RA >= (RACENTER - i)) Then
        EQ_SphericalPolar.r = 1
    Else
        EQ_SphericalPolar.r = 0
    End If

End Function

Public Function EQ_PolarSpherical(RA As Double, DEC As Double, TOT As Double, RACENTER As Double, DECCENTER As Double, Latitude As Double, range As Double) As Coord
Dim i As Double
Dim j As Double
Dim x As Double
Dim Y As Double
Dim pr As Double


    i = (((RA - RACENTER) / TOT) * 360) + 180
    j = (((DEC - DECCENTER) / TOT) * 180) - 90

    Call aa_hadec(Latitude * DEG_RAD, j * DEG_RAD, i * DEG_RAD, x, Y)
    
    If i > 180 Then
         If range = 0 Then
            Y = Range360(180 - (Y * RAD_DEG))
        Else
            Y = Range360(Y * RAD_DEG)
        End If
    Else
        If range = 0 Then
            Y = Range360(Y * RAD_DEG)
        Else
            Y = Range360(180 - (Y * RAD_DEG))
        End If
    End If
    
    j = Range360(Y + 90)
      
    If j < 180 Then
        If range = 1 Then
            x = Range24(x * RAD_HRS)
        Else
            x = Range24(24 + (x * RAD_HRS))
        End If
    Else
        x = Range24(12 + (x * RAD_HRS))
    End If
    
     
    EQ_PolarSpherical.x = Get_EncoderfromHours(RACENTER, x, TOT, 0)
    EQ_PolarSpherical.Y = Get_EncoderfromDegrees(DECCENTER, Y + 90, TOT, 0, 0)
    
End Function


Public Function EQ_Spherical2Cartes(RA As Double, DEC As Double, TOT As Double, RACENTER As Double, DECCENTER As Double) As CartesCoord

Dim tmpobj1 As CartesCoord
Dim tmpobj4 As SphereCoord

        tmpobj4 = EQ_SphericalPolar(RA, DEC, TOT, RACENTER, DECCENTER, gLatitude)

        tmpobj1 = EQ_Polar2Cartes(tmpobj4.x, tmpobj4.Y, TOT, RACENTER, DECCENTER)
    
        EQ_Spherical2Cartes.x = tmpobj1.x
        EQ_Spherical2Cartes.Y = tmpobj1.Y
        EQ_Spherical2Cartes.RA = tmpobj1.RA
        EQ_Spherical2Cartes.r = tmpobj1.r
    
End Function

Public Function EQ_Cartes2Spherical(x As Double, Y As Double, r As Double, RA As Double, range As Double, TOT As Double, RACENTER As Double, DECCENTER As Double) As Coord
Dim tmpobj2 As Coord

    tmpobj2 = EQ_Cartes2Polar(x, Y, r, RA, TOT, RACENTER, DECCENTER)
    EQ_Cartes2Spherical = EQ_PolarSpherical(tmpobj2.x, tmpobj2.Y, gTot_step, RAEncoder_Home_pos, gDECEncoder_Home_pos, gLatitude, range)

End Function


Public Function EQ_Choose_3Points(x As Double, Y As Double) As TriangleCoord

Dim i, j, k, l, m, n As Integer
Dim tmpcoords As Coord
Dim tmpcoord As Coord
Dim p1 As Coord
Dim p2 As Coord
Dim p3 As Coord
Dim pc As Coord
Dim Count As Integer

Dim datholder(1 To MAX_STARS) As Tdatholder
Dim combi_cnt, tmp1, tmp2 As Integer
Dim first As Boolean
Dim last_dist, new_dist As Double

' Adjust only if there are three alignment stars

    If gAlignmentStars_count <= 3 Then
        EQ_Choose_3Points.i = 1
        EQ_Choose_3Points.j = 2
        EQ_Choose_3Points.k = 3
        Exit Function
    End If
    
    tmpcoords.x = x
    tmpcoords.Y = Y
    tmpcoord = EQ_sp2Cs(tmpcoords)
    
    Count = 0
    ' first find out the distances to the alignment stars
    For i = 1 To gAlignmentStars_count
        
        With datholder(Count + 1)
            .cc = my_PointsC(i)
            Select Case gPointFilter
                Case 0
                    ' all points
            
                Case 1
                    ' only consider points on this side of the meridian
                    If .cc.Y * tmpcoord.Y < 0 Then
                        GoTo NextPoint
                    End If
                    
                Case 2
                    ' local quadrant
                    If GetQuadrant(tmpcoord) <> GetQuadrant(.cc) Then
                        GoTo NextPoint
                    End If
                    
            End Select
            
            If HC.CheckLocalPier.Value = 1 Then
                ' calculate polar distance
                .dat = (my_Points(i).x - x) ^ 2 + (my_Points(i).Y - Y) ^ 2
            Else
                ' calculate cartesian disatnce
                .dat = (CDbl(.cc.x - tmpcoord.x)) ^ 2 + (CDbl(.cc.Y - tmpcoord.Y)) ^ 2
            End If
            
            ' Also save the reference star id for this particular reference star
            .idx = i
        
        End With
        Count = Count + 1
NextPoint:
    Next i
        
    If Count < 3 Then
        ' not enough points to do 3-point
        EQ_Choose_3Points.i = 0
        EQ_Choose_3Points.j = 0
        EQ_Choose_3Points.k = 0
        Exit Function
    End If
    
    ' now sort the disatnces so the closest stars are at the top
    Call EQ_Quicksort2(datholder(), 1, Count)
    
    'Just use the nearest 50 stars (max) - saves processing time
    If Count > gMaxCombinationCount - 1 Then
        combi_cnt = gMaxCombinationCount
    Else
        combi_cnt = Count
    End If
    
'    combi_offset = 1
    tmp1 = combi_cnt - 1
    tmp2 = combi_cnt - 2
    first = True
    ' iterate through all the triangles posible using the nearest alignment points
    l = 1
    m = 2
    n = 3
    For i = 1 To (tmp2)
        p1 = datholder(i).cc
        For j = i + 1 To (tmp1)
            p2 = datholder(j).cc
            For k = (j + 1) To combi_cnt
                p3 = datholder(k).cc
                
                If EQ_CheckPoint_in_Triangle(tmpcoord.x, tmpcoord.Y, p1.x, p1.Y, p2.x, p2.Y, p3.x, p3.Y) = 1 Then
                    ' Compute for the center point
                    pc = EQ_GetCenterPoint(p1, p2, p3)
                    ' don't need full pythagoras - sum of squares is good enough
                    new_dist = (pc.x - tmpcoord.x) ^ 2 + (pc.Y - tmpcoord.Y) ^ 2
                    
                    If first Then
                       ' first time through
                        last_dist = new_dist
                        first = False
                        l = i
                        m = j
                        n = k
                    Else
                        If new_dist < last_dist Then
                            l = i
                            m = j
                            n = k
                            last_dist = new_dist
                        End If
                    End If
                End If
            Next k
        Next j
    Next i
    
    If first = True Then
        EQ_Choose_3Points.i = 0
        EQ_Choose_3Points.j = 0
        EQ_Choose_3Points.k = 0
    Else
        EQ_Choose_3Points.i = datholder(l).idx
        EQ_Choose_3Points.j = datholder(m).idx
        EQ_Choose_3Points.k = datholder(n).idx
    End If

End Function

Public Function EQ_ChooseNearest3Points(x As Double, Y As Double) As TriangleCoord

Dim i, j, k, l, m, n As Integer
Dim tmpcoords As Coord
Dim tmpcoord As Coord
Dim p1 As Coord
Dim p2 As Coord
Dim p3 As Coord
Dim pc As Coord
Dim Count As Integer

Dim datholder(1 To MAX_STARS) As Tdatholder
Dim combi_cnt, tmp1, tmp2 As Integer
Dim first As Boolean
Dim last_dist, new_dist As Double

' Adjust only if there are three alignment stars

    If gAlignmentStars_count <= 3 Then
        EQ_ChooseNearest3Points.i = 1
        EQ_ChooseNearest3Points.j = 2
        EQ_ChooseNearest3Points.k = 3
        Exit Function
    End If
    
    tmpcoords.x = x
    tmpcoords.Y = Y
    tmpcoord = EQ_sp2Cs(tmpcoords)
    
    Count = 0
    ' first find out the distances to the alignment stars
    For i = 1 To gAlignmentStars_count
        
        With datholder(Count + 1)
            .cc = my_PointsC(i)
            
            Select Case gPointFilter
                Case 0
                    ' all points
            
                Case 1
                    ' only consider points on this side of the meridian
                    If .cc.Y * tmpcoord.Y < 0 Then
                        GoTo NextPoint
                    End If
                    
                Case 2
                    ' local quadrant
                    If GetQuadrant(tmpcoord) <> GetQuadrant(.cc) Then
                        GoTo NextPoint
                    End If
                    
            End Select
            
            If HC.CheckLocalPier.Value = 1 Then
                ' calculate polar distance
                .dat = (my_Points(i).x - x) ^ 2 + (my_Points(i).Y - Y) ^ 2
            Else
                ' calculate cartesian disatnce
                .dat = (CDbl(.cc.x - tmpcoord.x)) ^ 2 + (CDbl(.cc.Y - tmpcoord.Y)) ^ 2
            End If
           
           ' Also save the reference star id for this particular reference star
            .idx = i
        
        End With
        Count = Count + 1
NextPoint:
    Next i
        
    If Count < 3 Then
        ' not enough points to do 3-point
        EQ_ChooseNearest3Points.i = 0
        EQ_ChooseNearest3Points.j = 0
        EQ_ChooseNearest3Points.k = 0
        Exit Function
    End If

    ' now sort the disatnces so the closest stars are at the top
    Call EQ_Quicksort2(datholder(), 1, Count)
    
    'Just use the nearest 50 stars (max) - saves processing time
    If Count > gMaxCombinationCount - 1 Then
        combi_cnt = gMaxCombinationCount
    Else
        combi_cnt = Count
    End If
    
    tmp1 = combi_cnt - 1
    tmp2 = combi_cnt - 2
    first = True
    
    ' iterate through all the triangles posible using the nearest alignment points
    l = 1
    m = 2
    n = 3
    For i = 1 To (tmp2)
        p1 = datholder(i).cc
        For j = i + 1 To (tmp1)
            p2 = datholder(j).cc
            For k = (j + 1) To combi_cnt
                p3 = datholder(k).cc
                
                If EQ_CheckPoint_in_Triangle(tmpcoord.x, tmpcoord.Y, p1.x, p1.Y, p2.x, p2.Y, p3.x, p3.Y) = 1 Then
                    l = i
                    m = j
                    n = k
                    GoTo alldone
                End If
            Next k
        Next j
    Next i
    
    EQ_ChooseNearest3Points.i = 0
    EQ_ChooseNearest3Points.j = 0
    EQ_ChooseNearest3Points.k = 0
    Exit Function
    
alldone:
    EQ_ChooseNearest3Points.i = datholder(l).idx
    EQ_ChooseNearest3Points.j = datholder(m).idx
    EQ_ChooseNearest3Points.k = datholder(n).idx

End Function



