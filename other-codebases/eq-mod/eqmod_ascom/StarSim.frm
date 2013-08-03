VERSION 5.00
Begin VB.Form StarSim 
   BackColor       =   &H00000000&
   Caption         =   "EyepieceSim"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3030
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   3030
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   2520
      Top             =   2520
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   0
      ScaleHeight     =   199
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   199
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "StarSim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private simx As Double
Private simy As Double


Private Sub Form_Load()

simx = Picture1.ScaleWidth / 2
simx = Picture1.ScaleHeight / 2


End Sub


Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    simx = X
    simy = Y
End Sub

Private Sub Timer1_Timer()

    Dim rarate As Double
    Dim decrate As Double

    Picture1.Cls
    Picture1.DrawWidth = 1
    Picture1.Line (0, Picture1.ScaleHeight / 2 - 2)-(Picture1.ScaleWidth, Picture1.ScaleHeight / 2 - 2), vbRed
    Picture1.Line (0, Picture1.ScaleHeight / 2 + 2)-(Picture1.ScaleWidth, Picture1.ScaleHeight / 2 + 2), vbRed
    Picture1.Line (Picture1.ScaleWidth / 2 - 2, 0)-(Picture1.ScaleWidth / 2 - 2, Picture1.ScaleHeight), vbRed
    Picture1.Line (Picture1.ScaleWidth / 2 + 2, 0)-(Picture1.ScaleWidth / 2 + 2, Picture1.ScaleHeight), vbRed

    If gTrackingStatus <> 1 Then
        simx = simx - 0.2
    End If
    
    rarate = HC.VScrollRASlewRate.Value / 50
    decrate = HC.VScrollDecSlewRate.Value / 50

    Select Case SlewActive
        Case 0
            'none
        Case 1
            ' north
            simy = simy - decrate
        Case 2
            ' northeast
            simx = simx + rarate
            simy = simy - decrate
        Case 3
            ' east
            simx = simx + rarate
        Case 4
            ' southeast
            simx = simx + rarate
            simy = simy + decrate
        Case 5
            'south
            simy = simy + decrate
        Case 6
            'southwest
            simx = simx - rarate
            simy = simy + decrate
        Case 7
            ' west
            simx = simx - rarate
        Case 8
            ' northwest
            simx = simx - rarate
            simy = simy - decrate
    End Select
    
    Picture1.DrawWidth = 4
    Picture1.Circle (simx, simy), 2, vbWhite
    Picture1.DrawWidth = 35
    Picture1.Circle (Picture1.ScaleWidth / 2, Picture1.ScaleHeight / 2), 120, &H40&
End Sub
