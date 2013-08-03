VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "EQASCOM Toolbox"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6105
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "View/Edit"
      Height          =   375
      Index           =   3
      Left            =   4080
      TabIndex        =   11
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Setup"
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   5895
      Begin VB.CommandButton CommandConnect 
         Caption         =   "Test Connect"
         Height          =   375
         Left            =   3960
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Driver Setup"
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "Form1.frx":0CCA
         Left            =   120
         List            =   "Form1.frx":0CD4
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Configuration files"
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   5895
      Begin VB.CommandButton Command3 
         Caption         =   "Backup"
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   20
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         Height          =   375
         Index           =   2
         Left            =   3960
         TabIndex        =   19
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Form1.frx":0CEC
         Left            =   120
         List            =   "Form1.frx":0CF6
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   270
         Width           =   1695
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         ItemData        =   "Form1.frx":0D0E
         Left            =   2040
         List            =   "Form1.frx":0D18
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1440
         Width           =   1695
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         ItemData        =   "Form1.frx":0D30
         Left            =   120
         List            =   "Form1.frx":0D3A
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Copy Configuration"
         Height          =   375
         Left            =   3960
         TabIndex        =   10
         ToolTipText     =   "Copy all Simulator files to EQASCOM"
         Top             =   1410
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":0D52
         Left            =   120
         List            =   "Form1.frx":0D5F
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   750
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Restore"
         Height          =   375
         Index           =   1
         Left            =   2040
         TabIndex        =   4
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Windows and ASCOM Registration"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.CommandButton Command2 
         Caption         =   "Register"
         Height          =   375
         Left            =   3960
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "Form1.frx":0D8C
         Left            =   120
         List            =   "Form1.frx":0D96
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton CommandClone 
         Caption         =   "Clone"
         Height          =   375
         Left            =   4560
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Deregister"
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim path As String
Dim path2 As String
Dim gscope As Object

Private Sub Command1_Click()
    Dim str As String
    Dim m_Profile As DriverHelper.Profile
    On Error GoTo unregerror
    
    Set m_Profile = New DriverHelper.Profile
    str = UCase(Left(Combo4.Text, Len(Combo4.Text) - 4))
    
    m_Profile.DeviceType = "Telescope"          ' We're a Telescope driver
    m_Profile.Unregister str & ".Telescope"
'    m_Profile.Unregister "EQMOD.Telescope"
'    m_Profile.Unregister "EQMOD_SIM.Telescope"

    Shell path & Combo4.Text & " /UNREGSERVER", vbHide
'    Shell path & "EQMOD.EXE /UNREGSERVER", vbHide
'    Shell path & "EQMOD_SIM.EXE /UNREGSERVER", vbHide
    MsgBox ("Success!")
    GoTo endsub:

unregerror:
    MsgBox ("Error! Please check the following:" & vbCrLf & vbCrLf & "1. The ASCOM Platform has been installed." & vbCrLf & "2. This toolbox applicaitons has adminisrator rights")

endsub:
End Sub

Private Sub Command2_Click()
    Dim str As String
    On Error GoTo regerror
    Dim m_Profile As DriverHelper.Profile
    Set m_Profile = New DriverHelper.Profile
    
    Set m_Profile = New DriverHelper.Profile
    str = UCase(Left(Combo4.Text, Len(Combo4.Text) - 4))
    
    m_Profile.DeviceType = "Telescope"          ' We're a Telescope driver
    m_Profile.Register str & ".Telescope", str & " HEQ5/6"
    Shell path & Combo4.Text & " /REGSERVER", vbHide
    
'    m_Profile.Register "EQMOD_SIM.Telescope", "EQMOD ASCOM Simulator"
'    m_Profile.Register "EQMOD.Telescope", "EQMOD ASCOM EQ5/6"
'    Shell path & "EQMOD.EXE /REGSERVER", vbHide
'    Shell path & "EQMOD_SIM.EXE /REGSERVER", vbHide
    MsgBox ("Success!")
    GoTo endsub
    
regerror:
    MsgBox ("Error! Please check the following:" & vbCrLf & vbCrLf & "1. The ASCOM Platform has been installed." & vbCrLf & "2. This toolbox applicaitons has adminisrator rights")
    
endsub:
End Sub

Private Sub Command3_Click(Index As Integer)
    Dim str1 As String
    Dim str2 As String
    Dim str3 As String
    Dim file1 As String
    Dim file2 As String
    Dim fso As Object
    
    On Error GoTo errHandle
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    str2 = Combo2.List(Combo2.ListIndex)
    str1 = "\" & Left(str2, Len(str2) - 4)
    
'    Select Case Combo2.ListIndex
'        Case 0
'            str1 = "\EQMOD"
'        Case 1
'            str1 = "\EQMOD_SIM"
'    End Select
    
    Select Case Combo1.ListIndex
        Case 0
            str2 = "\EQMOD.ini"
            str3 = "\EQMOD_bak.ini"
        Case 1
            str2 = "\ALIGN.ini"
            str3 = "\ALIGN_bak.ini"
        Case 2
            str2 = "\JOYSTICK.ini"
            str3 = "\JOYSTICK_bak.ini"
                
    End Select
    
    file1 = path2 & str1 & str2
    file2 = path2 & str1 & str3
    
    If fso.FileExists(file1) Then
        Select Case Index
            Case 0
                If fso.FileExists(file2) Then
                    If MsgBox("Overwrite existing backup?", vbYesNo, "Warning!") = vbYes Then
                        fso.CopyFile file1, file2
                    End If
                Else
                    fso.CopyFile file1, file2
                End If
            Case 1
                fso.CopyFile file2, file1
            Case 2
                fso.Deletefile file1
            Case 3
                Form2.filename = file1
                Form2.Show (0)
        End Select
    Else
        MsgBox ("Backup file " & file1 & " doesn't exist!")
    End If
    
errHandle:
End Sub

Private Sub Command4_Click()
Dim scope As Object
Dim id As String

    id = UCase(Left(Combo3.Text, Len(Combo3.Text) - 4)) & ".Telescope"

'    Select Case Combo3.ListIndex
'        Case 1
'            id = "EQMOD_sim.Telescope"
'        Case 0
'            id = "EQMOD.Telescope"
'    End Select
    
    Set scope = CreateObject(id)
    
    If Not scope Is Nothing Then
        scope.SetupDialog
    End If

End Sub

Private Sub Command5_Click()
    Dim fso As Object
    Dim FromPath As String
    Dim ToPath As String
    
    On Error Resume Next
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Combo5.Text = "" Then Exit Sub
    If Combo6.Text = "" Then Exit Sub
    If Combo5.Text = Combo6.Text Then Exit Sub
    
    FromPath = path2 & "\" & UCase(Left(Combo5.Text, Len(Combo5.Text) - 4)) & "\"
    ToPath = path2 & "\" & UCase(Left(Combo6.Text, Len(Combo6.Text) - 4)) & "\"
    
    fso.CopyFile FromPath & "EQMOD.ini", ToPath & "EQMOD.ini"
    fso.CopyFile FromPath & "JOYSTICK.ini", ToPath & "JOYSTICK.ini"
    fso.CopyFile FromPath & "ALIGN.ini", ToPath & "ALIGN.ini"
    MsgBox ("Success!")

End Sub

Private Sub CommandClone_Click()
    Dim fso As Object
    Dim m_Profile As DriverHelper.Profile
    Set m_Profile = New DriverHelper.Profile
    
    On Error Resume Next
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile path & "eqmod.exe", path & "eqmod_2.exe"

    On Error GoTo regerror
    m_Profile.DeviceType = "Telescope"
    m_Profile.Register "EQMOD_2.Telescope", "EQMOD_2 HEQ5/6"
    Shell path & "eqmod_2.exe /REGSERVER", vbHide
    MsgBox ("Success!")
    Exit Sub
regerror:
    MsgBox ("Clone Failed!")


End Sub

Private Sub CommandConnect_Click()

Dim id As String
    On Error Resume Next

    If CommandConnect.Caption = "Test Connect" Then
    
    id = UCase(Left(Combo3.Text, Len(Combo3.Text) - 4)) & ".Telescope"
'        Select Case Combo3.ListIndex
'            Case 1
'                id = "EQMOD_sim.Telescope"
'            Case 0
'                id = "EQMOD.Telescope"
'        End Select
        Set gscope = CreateObject(id)
        
        If Not gscope Is Nothing Then
            CommandConnect.Caption = "Disconnect"
            gscope.Connected = True
        End If
    Else
        If Not gscope Is Nothing Then
            gscope.Connected = False
            If Not gscope.Connected Then
                Set gscope = Nothing
                CommandConnect.Caption = "Test Connect"
            End If
        End If
    End If

End Sub

Private Sub Form_Load()
    path = Environ("CommonProgramFiles") & "\ASCOM\Telescope\"
    path2 = Environ("Appdata")
    Call ListEQMODs
    Combo1.ListIndex = 0
    Combo2.ListIndex = 0
    Combo3.ListIndex = 0
    Combo4.ListIndex = 0
End Sub

Private Sub ListEQMODs()
    'Leave Extention blank for all files
    Dim File As String
    
    Combo3.Clear
    Combo2.Clear
    Combo4.Clear
    Combo5.Clear
    Combo6.Clear
    File = Dir$(path & "eqmod*.exe")
    Do While Len(File)
        If Left(File, 5) = "eqmod" Then
            Combo3.AddItem File
            Combo2.AddItem File
            Combo4.AddItem File
            Combo5.AddItem File
            Combo6.AddItem File
        End If
        File = Dir$
    Loop
End Sub

