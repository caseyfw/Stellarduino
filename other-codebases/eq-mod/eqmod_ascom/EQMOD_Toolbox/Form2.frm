VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6795
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6750
   LinkTopic       =   "Form2"
   ScaleHeight     =   6795
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   6735
   End
   Begin VB.Menu refreshini 
      Caption         =   "Refresh"
   End
   Begin VB.Menu saveini 
      Caption         =   "Save"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public filename As String


Private Sub Form_Load()
Dim temp1 As String
    On Error GoTo exitsub
    
    Form2.Caption = filename
    
    Open filename For Input As #1

    Text1.Text = ""
    While Not EOF(1)
        Line Input #1, temp1
        Text1.Text = Text1.Text & temp1 & vbCrLf
    Wend
            
exitsub:
    Close #1
            
End Sub

Private Sub Form_Resize()
    Text1.Width = Form2.ScaleWidth
    Text1.Height = Form2.ScaleHeight
End Sub

Private Sub refreshini_Click()
    Call Form_Load
End Sub

Private Sub saveini_Click()
    Dim txt() As String
    Dim i As Integer
    
    If MsgBox("Are you sure you wish to overwrite this file?", vbYesNo, "Save ini file") = vbYes Then
    
        txt = Split(Text1.Text, vbCrLf)
    
        On Error GoTo exitsub
        
        Open filename For Output As #1
    
        For i = 0 To UBound(txt)
            Print #1, txt(i)
        Next i
    End If
            
exitsub:
    Close #1

    Call Form_Load
End Sub

