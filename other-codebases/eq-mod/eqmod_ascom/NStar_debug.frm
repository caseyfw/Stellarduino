VERSION 5.00
Begin VB.Form NStar_debug 
   BackColor       =   &H00000000&
   Caption         =   "N-Star Mapper"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8760
   Icon            =   "NStar_debug.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox SkyPlot 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   240
      ScaleHeight     =   6345
      ScaleWidth      =   6945
      TabIndex        =   5
      Top             =   960
      Width           =   6975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Redraw"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Transformation Map"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   $"NStar_debug.frx":0CCA
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   8535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000080&
      Caption         =   "Y - 00000"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5640
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
      Caption         =   "X - 00000"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   4560
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "NStar_debug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public RA_Zpos As Double
Public DEC_Zpos As Double
Public MotorTotPos As Double
Public TotSize As Double
Public tmpx1 As Double
Public tmpy1 As Double

Private Sub Command1_Click()

Dim RA As Double
Dim DEC As Double
Dim tmpcoord As Coordt
Dim tmpc1 As Coord
Dim tmpc2 As Coord
Dim tmpsp As SphereCoord
Dim Onestar As Integer

Dim ra_step As Double
Dim dec_step As Double

    If gAlignmentStars_count < 1 Then Exit Sub
    
    SkyPlot.Cls

    ra_step = 150000
    dec_step = 150000
    
    For RA = (gRAEncoder_Zero_pos - (gTot_step / 4)) To (gRAEncoder_Zero_pos + (gTot_step / 4) - ra_step) Step ra_step
       For DEC = (gDECEncoder_Zero_pos + (gTot_step / 4)) To (gDECEncoder_Zero_pos - (gTot_step / 4)) Step (dec_step * -1)

        tmpsp = EQ_SphericalPolar(RA, DEC, gTot_step, gRAEncoder_Zero_pos, gDECEncoder_Home_pos, gLatitude)

        ' Check if sky visible
        
        If tmpsp.Y < ((gTot_step / 2) + gDECEncoder_Home_pos) Then
        
            Onestar = 0
            
            Select Case gAlignmentMode
                
                Case 0
                    ' nstar+nearest
                    tmpcoord = Delta_Matrix_Map(RA - gRASync01, DEC - gDECSync01)
                    If tmpcoord.f = 0 Then
                        Onestar = 1
                        tmpcoord = DeltaSyncReverse_Matrix_Map(RA - gRASync01, DEC - gDECSync01)
                    End If
                
                Case 1
                    ' nstar
                    tmpcoord = Delta_Matrix_Map(RA - gRASync01, DEC - gDECSync01)
                
                Case 2
                    ' nearest
                    Onestar = 1
                    tmpcoord = DeltaSyncReverse_Matrix_Map(RA - gRASync01, DEC - gDECSync01)
            
            End Select
            
            tmpc1.x = tmpcoord.x
            tmpc1.Y = tmpcoord.Y
            tmpc2.x = RA - gRASync01
            tmpc2.Y = DEC - gDECSync01
        
            Call NStarPlotCross(SkyPlot, EQ_sp2Cs(tmpc2).x, EQ_sp2Cs(tmpc2).Y, 0, 0, 50, vbYellow)

            Call NStarPlotCross(SkyPlot, EQ_sp2Cs(tmpc1).x, EQ_sp2Cs(tmpc1).Y, 0, 0, 100, vbWhite)

            If Onestar = 1 Then
                Call NStarPlotLine(SkyPlot, EQ_sp2Cs(tmpc1).x, EQ_sp2Cs(tmpc1).Y, EQ_sp2Cs(tmpc2).x, EQ_sp2Cs(tmpc2).Y, 0, 0, vbGreen)
            Else
                Call NStarPlotLine(SkyPlot, EQ_sp2Cs(tmpc1).x, EQ_sp2Cs(tmpc1).Y, EQ_sp2Cs(tmpc2).x, EQ_sp2Cs(tmpc2).Y, 0, 0, vbRed)
            End If
        End If

       Next DEC
    Next RA

    For RA = (gRAEncoder_Zero_pos + (gTot_step / 4)) To (gRAEncoder_Zero_pos - (gTot_step / 4)) Step (ra_step * -1)
       For DEC = (gDECEncoder_Zero_pos + (gTot_step / 4)) To (gDECEncoder_Zero_pos + (gTot_step / 2) + (gTot_step / 4)) Step dec_step

        tmpsp = EQ_SphericalPolar(RA, DEC, gTot_step, gRAEncoder_Zero_pos, gDECEncoder_Home_pos, gLatitude)

        ' Check if sky visible
        
        If tmpsp.Y < ((gTot_step / 2) + gDECEncoder_Home_pos) Then
            
            Onestar = 0
 
            Select Case gAlignmentMode
                
                Case 0
                    ' nstar+nearest
                    tmpcoord = Delta_Matrix_Map(RA - gRASync01, DEC - gDECSync01)
                    If tmpcoord.f = 0 Then
                        Onestar = 1
                        tmpcoord = DeltaSyncReverse_Matrix_Map(RA - gRASync01, DEC - gDECSync01)
                    End If
                 
                Case 1
                    ' nstar
                    tmpcoord = Delta_Matrix_Map(RA - gRASync01, DEC - gDECSync01)
                
                Case Else
                    ' nearest
                    Onestar = 1
                    tmpcoord = DeltaSyncReverse_Matrix_Map(RA - gRASync01, DEC - gDECSync01)
            
            End Select
            
            tmpc1.x = tmpcoord.x
            tmpc1.Y = tmpcoord.Y
            tmpc2.x = RA - gRASync01
            tmpc2.Y = DEC - gDECSync01
        
            Call NStarPlotCross(SkyPlot, EQ_sp2Cs(tmpc2).x, EQ_sp2Cs(tmpc2).Y, 0, 0, 50, vbYellow)
            Call NStarPlotCross(SkyPlot, EQ_sp2Cs(tmpc1).x, EQ_sp2Cs(tmpc1).Y, 0, 0, 100, vbWhite)

            If Onestar = 1 Then
                Call NStarPlotLine(SkyPlot, EQ_sp2Cs(tmpc1).x, EQ_sp2Cs(tmpc1).Y, EQ_sp2Cs(tmpc2).x, EQ_sp2Cs(tmpc2).Y, 0, 0, vbGreen)
            Else
                Call NStarPlotLine(SkyPlot, EQ_sp2Cs(tmpc1).x, EQ_sp2Cs(tmpc1).Y, EQ_sp2Cs(tmpc2).x, EQ_sp2Cs(tmpc2).Y, 0, 0, vbRed)
            End If
        End If

       Next DEC
    Next RA

End Sub

Private Sub Command2_Click()
    SkyPlot.Cls
End Sub

Private Sub Form_Load()

Dim i As Integer

    Call SetText
    RA_Zpos = RAEncoder_Home_pos
    DEC_Zpos = gDECEncoder_Home_pos
    MotorTotPos = gTot_step
    TotSize = 7000
    tmpx1 = 0
    tmpy1 = 0

End Sub



Public Sub drawdots(ByVal x1 As Double, ByVal y1 As Double)

Dim i As Integer
Dim tmpobj As Coord
Dim tmpobj2 As Coordt

    Call draw_lines(SkyPlot)

    ' delete the old data from the screen
    tmpobj.x = tmpx1
    tmpobj.Y = tmpy1
    
    Call NStarPlotCircle(SkyPlot, EQ_sp2Cs(tmpobj).x, EQ_sp2Cs(tmpobj).Y, 0, 0, 30, vbBlack)
'    tmpobj = EQ_Transform_Taki(EQ_sp2Cs(tmpobj))
'    Call NStarPlotCircle(tmpobj.x, tmpobj.y, 0, 0, 30, vbBlack)
    
    Select Case gAlignmentMode
    
        Case 0
            tmpobj2 = Delta_Matrix_Map(tmpobj.x, tmpobj.Y)
            If tmpobj2.f = 0 Then
              tmpobj2 = DeltaSyncReverse_Matrix_Map(tmpobj.x, tmpobj.Y)
            End If
            
        Case 1
            tmpobj2 = Delta_Matrix_Map(tmpobj.x, tmpobj.Y)
        
        Case Else
            tmpobj2 = DeltaSyncReverse_Matrix_Map(tmpobj.x, tmpobj.Y)
     
     End Select
             
     tmpobj.x = tmpobj2.x
     tmpobj.Y = tmpobj2.Y
     
     Call NStarPlotCircle(SkyPlot, EQ_sp2Cs(tmpobj).x, EQ_sp2Cs(tmpobj).Y, 0, 0, 30, vbBlack)
    
    ' Draw the center DOT (NCP/SCP)
    
    tmpobj.x = RAEncoder_Home_pos
    tmpobj.Y = DECEncoder_Home_pos_const
    
    Call NStarPlotCircle(SkyPlot, EQ_sp2Cs(tmpobj).x, EQ_sp2Cs(tmpobj).Y, 0, 0, 100, vbCyan)
    Call NStarPlotCircle(SkyPlot, EQ_sp2Cs(tmpobj).x, EQ_sp2Cs(tmpobj).Y, 0, 0, 30, vbCyan)
    

    ' Draw the Reference Points (catalog and measured stars)
    
    For i = 1 To gAlignmentStars_count
      
      Call NStarPlotCircle(SkyPlot, ct_PointsC(i).x, ct_PointsC(i).Y, 0, 0, 30, vbGreen)
      
        If gAlignmentStars_count <= 3 Then
            If gSelectStar <> 0 Then
                If i = gSelectStar Then
                    Call NStarPlotCircle(SkyPlot, my_PointsC(i).x, my_PointsC(i).Y, 0, 0, 30, vbYellow)
                Else
                    Call NStarPlotCircle(SkyPlot, my_PointsC(i).x, my_PointsC(i).Y, 0, 0, 30, vbBlue)
                End If
            Else
                Call NStarPlotCircle(SkyPlot, my_PointsC(i).x, my_PointsC(i).Y, 0, 0, 30, vbBlue)
            End If
        Else
            If gSelectStar <> 0 Then
                If i = gSelectStar Then
                   Call NStarPlotCircle(SkyPlot, my_PointsC(i).x, my_PointsC(i).Y, 0, 0, 30, vbYellow)
                Else
                   Call NStarPlotCircle(SkyPlot, my_PointsC(i).x, my_PointsC(i).Y, 0, 0, 30, vbBlue)
                End If
            Else
                If i = gAffine1 Or i = gAffine2 Or i = gAffine3 Then
                    Call NStarPlotCircle(SkyPlot, my_PointsC(i).x, my_PointsC(i).Y, 0, 0, 30, vbYellow)
                Else
                    Call NStarPlotCircle(SkyPlot, my_PointsC(i).x, my_PointsC(i).Y, 0, 0, 30, vbBlue)
                End If
            End If
        End If
    Next i
    
    tmpobj.x = x1 - gRASync01
    tmpobj.Y = y1 - gDECSync01
    
    ' Draw the mount's current position
    Call NStarPlotCircle(SkyPlot, EQ_sp2Cs(tmpobj).x, EQ_sp2Cs(tmpobj).Y, 0, 0, 30, vbRed)
            
'     tmpobj = EQ_Transform_Taki(EQ_sp2Cs(tmpobj))
'     Call NStarPlotCircle(tmpobj.x, tmpobj.y, 0, 0, 30, vbWhite)
             
     Select Case gAlignmentMode
     
        Case 0
            ' nstar+nearest
            tmpobj2 = Delta_Matrix_Map(tmpobj.x, tmpobj.Y)
            If tmpobj2.f = 0 Then
              tmpobj2 = DeltaSyncReverse_Matrix_Map(tmpobj.x, tmpobj.Y)
            End If
            
        Case 1
            ' nstar
            tmpobj2 = Delta_Matrix_Map(tmpobj.x, tmpobj.Y)
            
        Case Else
            ' nearest
            tmpobj2 = DeltaSyncReverse_Matrix_Map(tmpobj.x, tmpobj.Y)
     
     End Select
             
     tmpobj.x = tmpobj2.x
     tmpobj.Y = tmpobj2.Y
     
     Call NStarPlotCircle(SkyPlot, EQ_sp2Cs(tmpobj).x, EQ_sp2Cs(tmpobj).Y, 0, 0, 30, vbWhite)
             
     tmpx1 = x1 - gRASync01
     tmpy1 = y1 - gDECSync01

End Sub

Public Sub draw_gotodot(ByVal x1 As Double, ByVal y1 As Double)

Dim tmpobj As Coord
Dim tmpobj2 As Coordt

    ' Draw the target goto dots, both the actual (white) and transformed (red)

    tmpobj.x = x1
    tmpobj.Y = y1

    Call NStarPlotCircle(SkyPlot, EQ_sp2Cs(tmpobj).x, EQ_sp2Cs(tmpobj).Y, 0, 0, 30, vbRed)
    
'    tmpobj = EQ_Transform_Taki(EQ_sp2Cs(tmpobj))
'    Call NStarPlotCircle(tmpobj.x, tmpobj.y, 0, 0, 30, vbWhite)
    
    Select Case gAlignmentMode
    
        Case 0
            ' nstar+nearest
            tmpobj2 = Delta_Matrix_Map(tmpobj.x, tmpobj.Y)
            If tmpobj2.f = 0 Then
              tmpobj2 = DeltaSyncReverse_Matrix_Map(tmpobj.x, tmpobj.Y)
            End If
            
        Case 1
            ' nstar
            tmpobj2 = Delta_Matrix_Map(tmpobj.x, tmpobj.Y)
        
        Case Else
            ' nearest
            tmpobj2 = DeltaSyncReverse_Matrix_Map(tmpobj.x, tmpobj.Y)
            
    End Select
             
    tmpobj.x = tmpobj2.x
    tmpobj.Y = tmpobj2.Y
     
    Call NStarPlotCircle(SkyPlot, EQ_sp2Cs(tmpobj).x, EQ_sp2Cs(tmpobj).Y, 0, 0, 30, vbWhite)
    
    If gSelectStar <> 0 Then
        Call NStarPlotCircle(SkyPlot, my_PointsC(gSelectStar).x, my_PointsC(gSelectStar).Y, 0, 0, 30, vbYellow)
    Else
        If gAlignmentStars_count > 3 Then
            If gTaki1 <> 0 And gTaki2 <> 0 And gTaki3 <> 0 Then
                Call NStarPlotCircle(SkyPlot, my_PointsC(gTaki1).x, my_PointsC(gTaki1).Y, 0, 0, 30, vbYellow)
                Call NStarPlotCircle(SkyPlot, my_PointsC(gTaki2).x, my_PointsC(gTaki2).Y, 0, 0, 30, vbYellow)
                Call NStarPlotCircle(SkyPlot, my_PointsC(gTaki3).x, my_PointsC(gTaki3).Y, 0, 0, 30, vbYellow)
            End If
        End If
    End If

End Sub


Public Function xlate_X(ByVal x As Double) As Double

    If HC.PolarEnable.Value = 1 Then
        x = x * -1   ' Reverse the image to match the planetarium screen
        xlate_X = ((x - RA_Zpos) * (TotSize / MotorTotPos)) + (TotSize / 2) + 7000
    Else
        xlate_X = ((x - RA_Zpos) * (TotSize / MotorTotPos)) + (TotSize / 2)
    End If

End Function

Public Function xlate_y(ByVal Y As Double) As Double

    If HC.PolarEnable.Value = 1 Then
        xlate_y = ((Y - DEC_Zpos) * (TotSize / MotorTotPos)) + (TotSize / 2) + 9000
    Else
        xlate_y = ((Y - DEC_Zpos) * (TotSize / MotorTotPos)) + (TotSize / 2)
    End If

End Function

Public Sub draw_lines(ByRef plot As PictureBox)

Dim i As Long

    If plot.Height < plot.width Then
    
        i = plot.Height
    Else
        i = plot.width
    End If

    plot.Circle (i / 2, i / 2), i / 2, vbBlue
    plot.Line (i / 2, 0)-(i / 2, i), vbBlue
    plot.Line (0, i / 2)-(i, i / 2), vbBlue

End Sub

Public Sub NStarPlotCircle(ByRef plot As PictureBox, y1 As Double, x1 As Double, xofst As Double, yofst As Double, size As Long, dcolor As Long)

Dim xs As Double
Dim ys As Double
Dim xr As Double
Dim yr As Double
Dim i As Long

    If plot.Height < plot.width Then
        i = plot.Height
    Else
        i = plot.width
    End If

    xs = 0
    ys = 0
    
    x1 = x1 * -1
    
   
    xr = i / 9024000
    yr = i / 9024000
    
    plot.DrawMode = 13
    
    plot.Circle ((xs + (i / 2) + (x1 * xr) + xofst), ys + (i / 2) + (y1 * yr) + yofst), size, dcolor

End Sub

Public Sub NStarPlotCross(ByRef plot As PictureBox, y1 As Double, x1 As Double, xofst As Double, yofst As Double, size As Long, dcolor As Long)

Dim xs As Double
Dim ys As Double
Dim xr As Double
Dim yr As Double
Dim i As Long

    If plot.Height < plot.width Then
        i = plot.Height
    Else
        i = plot.width
    End If

    xs = 0
    ys = 0
    
    x1 = x1 * -1
    
   
    xr = i / 9024000
    yr = i / 9024000
    
    plot.DrawMode = 13
    
    plot.Line ((xs + (i / 2) + (x1 * xr) + xofst) - (size / 2), ys + (i / 2) + (y1 * yr) + yofst)-((xs + (i / 2) + (x1 * xr) + xofst) + (size / 2), ys + (i / 2) + (y1 * yr) + yofst), dcolor
    plot.Line ((xs + (i / 2) + (x1 * xr) + xofst), ys + (i / 2) + (y1 * yr) + yofst - (size / 2))-((xs + (i / 2) + (x1 * xr) + xofst), ys + (i / 2) + (y1 * yr) + yofst + (size / 2)), dcolor
    

End Sub

Public Sub NStarPlotLine(ByRef plot As PictureBox, y1 As Double, x1 As Double, y2 As Double, x2 As Double, xofst As Double, yofst As Double, dcolor As Long)

Dim xs As Double
Dim ys As Double
Dim xr As Double
Dim yr As Double
Dim i As Long

    If plot.Height < plot.width Then
        i = plot.Height
    Else
        i = plot.width
    End If

    xs = 0
    ys = 0
    
    x1 = x1 * -1
    x2 = x2 * -1
   
    xr = i / 9024000
    yr = i / 9024000
    
    plot.DrawMode = 13
    
    plot.Line ((xs + (i / 2) + (x1 * xr) + xofst), ys + (i / 2) + (y1 * yr) + yofst)-((xs + (i / 2) + (x2 * xr) + xofst), ys + (i / 2) + (y2 * yr) + yofst), dcolor
    

End Sub

Private Sub SetText()
    NStar_debug.Caption = HC.oLangDll.GetLangString(1000)
    Label5.Caption = HC.oLangDll.GetLangString(1001)
End Sub

Private Sub Form_Resize()
    SkyPlot.Cls
End Sub



Private Sub SkyPlot_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Label1.Caption = "X - " & str(x)
    Label2.Caption = "Y - " & str(Y)

End Sub
