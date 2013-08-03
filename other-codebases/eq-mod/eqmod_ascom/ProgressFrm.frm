VERSION 5.00
Begin VB.Form ProgressFrm 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Progress"
   ClientHeight    =   735
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   49
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   120
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   295
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "ProgressFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private max As Double
Private min As Double

Private Sub Form_Load()
    Call PutWindowOnTop(Me)
    Me.AutoRedraw = True
    min = 0
    max = 100
    Picture1.Cls
    SetProgress (0)
End Sub

Public Sub SetProgress(val As Double)
    Dim x As Double
    val = (val - min) / (max - min)
    x = Picture1.width * val
    Picture1.DrawWidth = Picture1.ScaleHeight
    Picture1.Line (0, 0)-(x, 0), vbRed
    Label1.Caption = FormatNumber(val * 100, 0) & "%"
End Sub

Public Sub SetMax(val As Double)
    max = val
    Picture1.Cls
End Sub
Public Sub SetMin(val As Double)
    min = val
    Picture1.Cls
End Sub

