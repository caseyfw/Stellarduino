VERSION 5.00
Begin VB.Form Slewpad 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EQMOD Mouse Slew Pad"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   7140
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
      Left            =   5280
      TabIndex        =   6
      Top             =   6120
      Width           =   1335
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   2175
      Left            =   6720
      Max             =   800
      Min             =   1
      TabIndex        =   2
      Top             =   3840
      Value           =   400
      Width           =   375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2055
      Left            =   6720
      Max             =   800
      Min             =   1
      TabIndex        =   1
      Top             =   1080
      Value           =   400
      Width           =   375
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
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "DEC"
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
      Left            =   6720
      TabIndex        =   11
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "RA"
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
      Left            =   6720
      TabIndex        =   10
      Top             =   840
      Width           =   495
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
      TabIndex        =   9
      Top             =   480
      Width           =   6735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000040&
      Caption         =   "Slew: NONE"
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
      Left            =   4080
      TabIndex        =   8
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000040&
      Caption         =   "BUTTONS: Left:West, Right: East, Middle: Switch Axis"
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
      TabIndex        =   7
      Top             =   6120
      Width           =   3855
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
      TabIndex        =   5
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000040&
      Caption         =   "400"
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
      Left            =   6720
      TabIndex        =   4
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
      Caption         =   "400"
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
      Left            =   6720
      TabIndex        =   3
      Top             =   3240
      Width           =   375
   End
End
Attribute VB_Name = "Slewpad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sPadMode As Boolean
Public sPadLastState As Long


Private Sub Command1_Click()
    writeratebarstatePad
    Unload Slewpad
End Sub

Private Sub Form_Activate()
    WheelHook (Me.hWnd)
End Sub


Private Sub Form_Deactivate()
    WheelUnHook (Me.hWnd)
End Sub

Private Sub Form_Load()
    
    EnableCloseButton Me.hWnd, False
    sPadMode = False
    sPadLastState = 0
    readratebarstatePad

End Sub



Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Select Case (Button)
    
    Case 1 ' Left Button
    
        If sPadMode = False Then
        
            eqres = EQ_MotorStop(0)          ' Stop RA Motor
            If eqres <> EQ_OK Then
                GoTo spadEND01
            End If

            'Wait until RA motor is stable

            Do
                eqres = EQ_GetMotorStatus(0)
                    If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo spadEND01
       
            Loop While (eqres And EQ_MOTORBUSY) <> 0
            eqres = EQ_Slew(0, 0, 0, val(VScroll1.Value))
            Label4.Caption = "Slew: WEST"
        
        Else
        
            eqres = EQ_MotorStop(1)          ' Stop DEC Motor
            If eqres <> EQ_OK Then
               GoTo spadEND01
            End If

            ' Wait for motor stop

            Do
        
                eqres = EQ_GetMotorStatus(1)
                If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo spadEND01
    
            Loop While (eqres And EQ_MOTORBUSY) <> 0
   
            eqres = EQ_Slew(1, 0, 0, val(VScroll2.Value))
            Label4.Caption = "Slew: NORTH"

        
        End If

    
    Case 2 ' Right Button
    
        If sPadMode = False Then
        
            eqres = EQ_MotorStop(0)          ' Stop RA Motor
    
            If eqres <> EQ_OK Then
                GoTo spadEND01
            End If

            'Wait until RA motor is stable

            Do
                 eqres = EQ_GetMotorStatus(0)
                If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo spadEND01

            Loop While (eqres And EQ_MOTORBUSY) <> 0

            eqres = EQ_Slew(0, 0, 1, val(VScroll1.Value))
            Label4.Caption = "Slew: EAST"
            
        Else
            eqres = EQ_MotorStop(1)          ' Stop DEC Motor
            If eqres <> EQ_OK Then
               GoTo spadEND01
            End If

            ' Wait for motor stop

            Do
        
                eqres = EQ_GetMotorStatus(1)
                If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo spadEND01
    
            Loop While (eqres And EQ_MOTORBUSY) <> 0
   
            eqres = EQ_Slew(1, 0, 1, val(VScroll2.Value))
            Label4.Caption = "Slew: SOUTH"
            
        End If

    
    Case Else ' Assume Middle Button
    
        If sPadMode = False Then
            sPadMode = True
            Slewpad.Label3.Caption = "BUTTONS: Left: North, Right: South, Middle: Switch Axis"
        Else
            sPadMode = False
            Slewpad.Label3.Caption = "BUTTONS: Left: West, Right: East, Middle: Switch Axis"
        End If
            
    End Select

    sPadLastState = Button
    
spadEND01:
End Sub

Private Sub Frame1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case (sPadLastState)
        Case 1
            If sPadMode = False Then
            
                eqres = EQ_MotorStop(0)          ' Stop RA Motor
                If eqres <> EQ_OK Then
                    GoTo spadEND02
                End If

                'Wait until RA motor is stable

                Do
                    eqres = EQ_GetMotorStatus(0)
                    If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo spadEND02
       
                Loop While (eqres And EQ_MOTORBUSY) <> 0

                If gTrackingStatus <> 0 Then eqres = EQ_StartRATrack(gTrackingStatus - 1, gHemisphere, gHemisphere)
        
            Else
            
                eqres = EQ_MotorStop(1)
            
            End If
        Case 2
            If sPadMode = False Then
            
                eqres = EQ_MotorStop(0)          ' Stop RA Motor
                If eqres <> EQ_OK Then
                    GoTo spadEND02
                End If

                'Wait until RA motor is stable

                Do
                    eqres = EQ_GetMotorStatus(0)
                    If (eqres = EQ_NOTINITIALIZED) Or (eqres = EQ_COMNOTOPEN) Or (eqres = EQ_COMTIMEOUT) Then GoTo spadEND02
       
                Loop While (eqres And EQ_MOTORBUSY) <> 0

                If gTrackingStatus <> 0 Then eqres = EQ_StartRATrack(gTrackingStatus - 1, gHemisphere, gHemisphere)
        
            Else
                eqres = EQ_MotorStop(1)
            End If
        Case Else
                eqres = 0       ' Do Nothing
    End Select
    
    Label4.Caption = "Slew: NONE"
    
spadEND02:
End Sub

Private Sub VScroll1_Change()
    Slewpad.Label1.Caption = Str(VScroll1.Value)
End Sub

Private Sub VScroll1_Scroll()
    Slewpad.Label1.Caption = Str(VScroll1.Value)
End Sub

Private Sub VScroll2_Change()
    Slewpad.Label2.Caption = Str(VScroll2.Value)
End Sub

Private Sub VScroll2_Scroll()
    Slewpad.Label2.Caption = Str(VScroll2.Value)
End Sub

Public Sub Mousewheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
   
   Dim i As Double
      
   If sPadMode = False Then
        i = Slewpad.VScroll1.Value
        If Rotation > 0 Then
            i = i + 20
            If i >= 800 Then i = 800
        Else
            i = i - 20
            If i <= 0 Then i = 1
        End If
        Slewpad.VScroll1.Value = i
   Else
        i = Slewpad.VScroll2.Value
        If Rotation > 0 Then
            i = i + 20
            If i >= 800 Then i = 800
        Else
            i = i - 20
            If i <= 0 Then i = 1
        End If
        Slewpad.VScroll2.Value = i
   End If
   
End Sub

