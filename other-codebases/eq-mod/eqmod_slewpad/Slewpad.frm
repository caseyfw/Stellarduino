VERSION 5.00
Begin VB.Form Slewpad 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EQMOD Mouse Slew Pad"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Top             =   6120
      Width           =   6495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Slew Region"
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
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6495
      Begin VB.Label Label7 
         BackColor       =   &H00000080&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   735
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000040&
      Caption         =   "Put the mouse cursor on the Slew Region below and click the mouse buttons to slew the mount"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   6735
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "EQMOD MOUSE BUTTON SLEW PAD"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Slewpad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Command1_Click()
    Unload Slewpad
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   Label7.Caption = Button
   
End Sub

Private Sub Frame1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
   Label7.Caption = 0
  
End Sub

