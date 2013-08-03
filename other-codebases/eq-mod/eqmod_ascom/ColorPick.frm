VERSION 5.00
Begin VB.Form ColorPick 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Color Picker"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H0095C1CB&
      Caption         =   "Set to Default"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3960
      Width           =   2415
   End
   Begin VB.HScrollBar HScroll7 
      Height          =   135
      Left            =   480
      Max             =   24
      Min             =   7
      TabIndex        =   8
      Top             =   3600
      Value           =   7
      Width           =   2295
   End
   Begin VB.HScrollBar HScroll6 
      Height          =   135
      Left            =   480
      Max             =   255
      TabIndex        =   7
      Top             =   3000
      Width           =   2295
   End
   Begin VB.HScrollBar HScroll5 
      Height          =   135
      Left            =   480
      Max             =   255
      TabIndex        =   6
      Top             =   2760
      Width           =   2295
   End
   Begin VB.HScrollBar HScroll4 
      Height          =   135
      Left            =   480
      Max             =   255
      TabIndex        =   5
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0095C1CB&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   2415
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   135
      Left            =   480
      Max             =   255
      TabIndex        =   3
      Top             =   1920
      Width           =   2295
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   135
      Left            =   480
      Max             =   255
      TabIndex        =   2
      Top             =   1680
      Width           =   2295
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   135
      Left            =   480
      Max             =   255
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox HCMessage 
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "ColorPick.frx":0000
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "B"
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   0
      TabIndex        =   19
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "G"
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   0
      TabIndex        =   18
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "R"
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   0
      TabIndex        =   17
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "B"
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   0
      TabIndex        =   16
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "G"
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   0
      TabIndex        =   15
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "R"
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   0
      TabIndex        =   14
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Font Size:"
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   480
      TabIndex        =   13
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Background Text:"
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   480
      TabIndex        =   12
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Foreground Text:"
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   480
      TabIndex        =   11
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "EQMOD COLOR SETUP"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   10
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "ColorPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
' ColorPick.frm - Color Picker
'
'
' Written:  07-Oct-06   Raymund Sarmiento
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' 11-Apr-07 rcs     Initial Edit
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


Private Sub Command1_Click()

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long
Dim n As Long


    i = val(HScroll1.Value)
    j = val(HScroll2.Value) * 256
    k = val(HScroll3.Value) * 65536
    
    HC.HCMessage.ForeColor = i + j + k
'    HC.HCTextAlign.ForeColor = i + j + k
    
    i = val(HScroll4.Value)
    j = val(HScroll5.Value) * 256
    k = val(HScroll6.Value) * 65536
    
    HC.HCMessage.BackColor = i + j + k
'    HC.HCTextAlign.BackColor = i + j + k
    
    HC.HCMessage.FontSize = HScroll7.Value
'    HC.HCTextAlign.FontSize = HScroll7.Value
    
    Call writeColorDat(HScroll1.Value, HScroll2.Value, HScroll3.Value, HScroll4.Value, HScroll5.Value, HScroll6.Value, HScroll7.Value)
    
    Unload ColorPick
    
End Sub

Private Sub Command2_Click()

    HScroll1.Value = &HFF
    HScroll2.Value = &H80
    HScroll3.Value = &H0

    HScroll4.Value = &H80
    HScroll5.Value = &H0
    HScroll6.Value = &H0
    
    HScroll7.Value = 7
    
End Sub



Private Sub Form_Load()

Dim i As Integer
Dim j As Integer
Dim k As Integer

    Call SetText
    
    HScroll1.Value = 1
    HScroll2.Value = 1
    HScroll3.Value = 1
    HScroll4.Value = 1
    HScroll5.Value = 1
    HScroll6.Value = 1
'    HScroll7.value = 7
    HScroll1.Value = 0
    HScroll2.Value = 0
    HScroll3.Value = 0
    HScroll4.Value = 0
    HScroll5.Value = 0
    HScroll6.Value = 0
'    HScroll7.value = 8

    i = val(HC.HCMessage.ForeColor) And &HFF
    j = ((val(HC.HCMessage.ForeColor) And &HFF00) / 256) And &HFF
    k = ((val(HC.HCMessage.ForeColor) And &HFF0000) / 65536) And &HFF
    
    HScroll1.Value = i
    HScroll2.Value = j
    HScroll3.Value = k
    
    i = val(HC.HCMessage.BackColor) And &HFF
    j = ((val(HC.HCMessage.BackColor) And &HFF00) / 256) And &HFF
    k = ((val(HC.HCMessage.BackColor) And &HFF0000) / 65536) And &HFF
    
    HScroll4.Value = i
    HScroll5.Value = j
    HScroll6.Value = k
    
 '   HScroll7.value = val(HC.HCMessage.FontSize)

End Sub

Private Sub HScroll1_Change()
    
    Call Get_Foreground

End Sub

Public Sub Get_Foreground()

Dim i As Long
Dim j As Long
Dim k As Long
Dim m As Long

    i = val(HScroll1.Value)
    j = val(HScroll2.Value) * 256
    k = val(HScroll3.Value) * 65536

    m = i + j + k

    If (m >= 0) And (m <= &HFFFFFF) Then
        HCMessage.ForeColor = m
    End If
    
End Sub

Public Sub Get_Background()

Dim i As Long
Dim j As Long
Dim k As Long
Dim m As Long

    i = val(HScroll4.Value)
    j = val(HScroll5.Value) * 256
    k = val(HScroll6.Value) * 65536
    
    m = i + j + k

    If (m >= 0) And (m <= &HFFFFFF) Then
        HCMessage.BackColor = m
    End If

End Sub

Public Sub SetFont()
Dim i As Long
Dim m As Long

    m = val(HScroll7.Value)
    
    If (m >= 7) And (m <= 24) Then
        HCMessage.FontSize = m
    End If

End Sub

Private Sub HScroll1_Scroll()

    Call Get_Foreground

End Sub

Private Sub HScroll2_Change()

    Call Get_Foreground

End Sub

Private Sub HScroll2_Scroll()

    Call Get_Foreground

End Sub


Private Sub HScroll3_Change()

    Call Get_Foreground

End Sub

Private Sub HScroll3_Scroll()

    Call Get_Foreground

End Sub

Private Sub HScroll4_Change()

    Call Get_Background

End Sub

Private Sub HScroll4_Scroll()

    Call Get_Background
    
End Sub

Private Sub HScroll5_Change()

    Call Get_Background

End Sub

Private Sub HScroll5_Scroll()

    Call Get_Background
    
End Sub

Private Sub HScroll6_Change()
    Call Get_Background

End Sub

Private Sub HScroll6_Scroll()

    Call Get_Background
    
End Sub

Private Sub HScroll7_Change()
    Call SetFont
End Sub

Private Sub HScroll7_Scroll()
    Call SetFont
End Sub

Private Sub SetText()
    ColorPick.Caption = oLangDll.GetLangString(600)
    Label6.Caption = oLangDll.GetLangString(601)
    Label2.Caption = oLangDll.GetLangString(602)
    Label1.Caption = oLangDll.GetLangString(603)
    Label3.Caption = oLangDll.GetLangString(604)
    Label4.Caption = oLangDll.GetLangString(605)
    Label5.Caption = oLangDll.GetLangString(606)
    Label7.Caption = oLangDll.GetLangString(607)
    Label8.Caption = oLangDll.GetLangString(605)
    Label9.Caption = oLangDll.GetLangString(606)
    Label10.Caption = oLangDll.GetLangString(607)
    Command2.Caption = oLangDll.GetLangString(608)
    Command1.Caption = oLangDll.GetLangString(609)
End Sub
