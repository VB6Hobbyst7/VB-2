VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmConfig 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  '단일 고정
   Caption         =   "통신설정"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3915
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   3915
   StartUpPosition =   2  '화면 가운데
   Begin Threed.SSPanel SSPanel3 
      Height          =   4095
      Left            =   60
      TabIndex        =   22
      Top             =   1080
      Width           =   3795
      _Version        =   65536
      _ExtentX        =   6694
      _ExtentY        =   7223
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Begin VB.CommandButton Command4 
         Caption         =   "종 료"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2010
         TabIndex        =   44
         Top             =   3510
         Width           =   1125
      End
      Begin VB.CommandButton Command3 
         Caption         =   "확 인"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   780
         TabIndex        =   43
         Top             =   3510
         Width           =   1125
      End
      Begin VB.CheckBox chkACK2 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00E0E0E0&
         Caption         =   "ACK 사용"
         Height          =   315
         Left            =   660
         TabIndex        =   31
         Top             =   3120
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.ComboBox Combo_Port2 
         Height          =   315
         ItemData        =   "frmConfig.frx":0000
         Left            =   1785
         List            =   "frmConfig.frx":0002
         TabIndex        =   30
         Top             =   480
         Width           =   1695
      End
      Begin VB.ComboBox Combo_BPS2 
         Height          =   315
         ItemData        =   "frmConfig.frx":0004
         Left            =   1785
         List            =   "frmConfig.frx":0006
         TabIndex        =   29
         Top             =   930
         Width           =   1695
      End
      Begin VB.ComboBox Combo_Databit2 
         Height          =   315
         ItemData        =   "frmConfig.frx":0008
         Left            =   1785
         List            =   "frmConfig.frx":000A
         TabIndex        =   28
         Top             =   1380
         Width           =   1695
      End
      Begin VB.ComboBox Combo_Startbit2 
         Height          =   315
         Left            =   1785
         TabIndex        =   27
         Top             =   1830
         Width           =   1695
      End
      Begin VB.ComboBox Combo_Stopbit2 
         Height          =   315
         Left            =   1785
         TabIndex        =   26
         Top             =   2280
         Width           =   1695
      End
      Begin VB.ComboBox Combo_Parity2 
         Height          =   315
         ItemData        =   "frmConfig.frx":000C
         Left            =   1785
         List            =   "frmConfig.frx":000E
         TabIndex        =   25
         Top             =   2760
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         Height          =   225
         Left            =   4335
         TabIndex        =   24
         Top             =   4395
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CheckBox Check1 
         Height          =   225
         Left            =   4335
         TabIndex        =   23
         Top             =   4050
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "COM PORT"
         Height          =   195
         Index           =   15
         Left            =   420
         TabIndex        =   40
         Top             =   540
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "전송속도"
         Height          =   195
         Index           =   14
         Left            =   420
         TabIndex        =   39
         Top             =   990
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "데이터 비트"
         Height          =   195
         Index           =   13
         Left            =   405
         TabIndex        =   38
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "시작 비트"
         Height          =   195
         Index           =   12
         Left            =   405
         TabIndex        =   37
         Top             =   1890
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "정지 비트"
         Height          =   195
         Index           =   11
         Left            =   405
         TabIndex        =   36
         Top             =   2340
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "패리티"
         Height          =   195
         Index           =   10
         Left            =   405
         TabIndex        =   35
         Top             =   2820
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "DTR Enabled"
         Height          =   195
         Index           =   9
         Left            =   3030
         TabIndex        =   34
         Top             =   4410
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "RTS Enabled"
         Height          =   195
         Index           =   8
         Left            =   3030
         TabIndex        =   33
         Top             =   4065
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "Com2"
         Height          =   225
         Left            =   3180
         TabIndex        =   32
         Top             =   120
         Width           =   855
      End
   End
   Begin Threed.SSPanel spPort1 
      Height          =   345
      Left            =   60
      TabIndex        =   19
      Top             =   720
      Width           =   825
      _Version        =   65536
      _ExtentX        =   1455
      _ExtentY        =   609
      _StockProps     =   15
      Caption         =   "Com1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.76
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   4095
      Left            =   60
      TabIndex        =   0
      Top             =   1080
      Width           =   3795
      _Version        =   65536
      _ExtentX        =   6694
      _ExtentY        =   7223
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Begin VB.CommandButton Command2 
         Caption         =   "종 료"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2010
         TabIndex        =   42
         Top             =   3510
         Width           =   1125
      End
      Begin VB.CommandButton Command1 
         Caption         =   "확 인"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   780
         TabIndex        =   41
         Top             =   3510
         Width           =   1125
      End
      Begin VB.CheckBox chkRTS 
         Height          =   225
         Left            =   3735
         TabIndex        =   16
         Top             =   3960
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CheckBox chkDTR 
         Height          =   225
         Left            =   3735
         TabIndex        =   15
         Top             =   4305
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.ComboBox Combo_Parity 
         Height          =   315
         ItemData        =   "frmConfig.frx":0010
         Left            =   1785
         List            =   "frmConfig.frx":0012
         TabIndex        =   7
         Top             =   2760
         Width           =   1695
      End
      Begin VB.ComboBox Combo_Stopbit 
         Height          =   315
         Left            =   1785
         TabIndex        =   6
         Top             =   2280
         Width           =   1695
      End
      Begin VB.ComboBox Combo_Startbit 
         Height          =   315
         Left            =   1785
         TabIndex        =   5
         Top             =   1830
         Width           =   1695
      End
      Begin VB.ComboBox Combo_Databit 
         Height          =   315
         ItemData        =   "frmConfig.frx":0014
         Left            =   1785
         List            =   "frmConfig.frx":0016
         TabIndex        =   4
         Top             =   1380
         Width           =   1695
      End
      Begin VB.ComboBox Combo_BPS 
         Height          =   315
         ItemData        =   "frmConfig.frx":0018
         Left            =   1770
         List            =   "frmConfig.frx":001A
         TabIndex        =   3
         Top             =   930
         Width           =   1695
      End
      Begin VB.ComboBox Combo_Port 
         Height          =   315
         ItemData        =   "frmConfig.frx":001C
         Left            =   1785
         List            =   "frmConfig.frx":001E
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.CheckBox chkACK 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00E0E0E0&
         Caption         =   "ACK 사용"
         Height          =   315
         Left            =   660
         TabIndex        =   1
         Top             =   3120
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "Com1"
         Height          =   225
         Left            =   3180
         TabIndex        =   21
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "RTS Enabled"
         Height          =   195
         Index           =   6
         Left            =   2460
         TabIndex        =   18
         Top             =   3960
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "DTR Enabled"
         Height          =   195
         Index           =   7
         Left            =   2430
         TabIndex        =   17
         Top             =   4320
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "패리티"
         Height          =   195
         Index           =   5
         Left            =   405
         TabIndex        =   13
         Top             =   2820
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "정지 비트"
         Height          =   195
         Index           =   4
         Left            =   405
         TabIndex        =   12
         Top             =   2340
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "시작 비트"
         Height          =   195
         Index           =   3
         Left            =   405
         TabIndex        =   11
         Top             =   1890
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "데이터 비트"
         Height          =   195
         Index           =   2
         Left            =   405
         TabIndex        =   10
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "전송속도"
         Height          =   195
         Index           =   1
         Left            =   420
         TabIndex        =   9
         Top             =   990
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "COM PORT"
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   8
         Top             =   540
         Width           =   1155
      End
   End
   Begin Threed.SSPanel SSPanel_machine 
      Height          =   675
      Left            =   45
      TabIndex        =   14
      Top             =   0
      Width           =   3795
      _Version        =   65536
      _ExtentX        =   6694
      _ExtentY        =   1191
      _StockProps     =   15
      Caption         =   "Coagu Check"
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   345
      Left            =   900
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   825
      _Version        =   65536
      _ExtentX        =   1455
      _ExtentY        =   609
      _StockProps     =   15
      Caption         =   "Com2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.76
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Serial Setting

Private Sub Command1_Click()
    Dim Parity As String
    Dim sEquipNo As String
    
    On Error GoTo ErrorHandler
    
    If MsgBox("설정을 저장하시겠습니까?", vbCritical + vbOKCancel + vbDefaultButton2, "확인!") = vbCancel Then
        Unload Me
        Exit Sub
    Else
    
        gSetup.gPort = Combo_Port.Text
        gSetup.gSpeed = Combo_BPS.Text
        gSetup.gDataBit = Combo_Databit.Text
        gSetup.gStartBit = Combo_Startbit.Text
        gSetup.gStopBit = Combo_Stopbit.Text
        If Combo_Parity.ListIndex = 0 Then
           gSetup.gParity = "N"
        End If
        If Combo_Parity.ListIndex = 1 Then
           gSetup.gParity = "E"
        End If
        If Combo_Parity.ListIndex = 2 Then
           gSetup.gParity = "O"
        End If
            
        If chkACK.Value = 1 Then
            gSetup.ACKUse = "1"
        Else
            gSetup.ACKUse = "0"
        End If
        
        
        Call WritePrivateProfileString("config1", "gPort", gSetup.gPort, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config1", "gSpeed", gSetup.gSpeed, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config1", "gParity", gSetup.gParity, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config1", "gDataBit", gSetup.gDataBit, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config1", "gStopBit", gSetup.gStopBit, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config1", "gStartBit", gSetup.gStartBit, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config1", "ACKUse", gSetup.ACKUse, App.Path & "\Interface.ini")
        
        If frmInterface.MSComm1.PortOpen = True Then
            frmInterface.MSComm1.PortOpen = False
        End If
        frmInterface.MSComm1.CommPort = gSetup.gPort
        frmInterface.MSComm1.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit
        
        frmInterface.MSComm1.PortOpen = True
        
        Unload Me
    End If
        
    Exit Sub
 
ErrorHandler:
    Resume Next
'    If MsgBox("통신설정이 맞지 않습니다", vbCritical + vbOKCancel + vbDefaultButton2, "종료버튼") = vbCancel Then
'        Exit Sub
'    End If
    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    Dim Parity As String
    Dim sEquipNo As String
    
    On Error GoTo ErrorHandler
    
    If MsgBox("설정을 저장하시겠습니까?", vbCritical + vbOKCancel + vbDefaultButton2, "확인!") = vbCancel Then
        Unload Me
        Exit Sub
    Else
    
        gSetup2.gPort = Combo_Port2.Text
        gSetup2.gSpeed = Combo_BPS2.Text
        gSetup2.gDataBit = Combo_Databit2.Text
        gSetup2.gStartBit = Combo_Startbit2.Text
        gSetup2.gStopBit = Combo_Stopbit2.Text
        If Combo_Parity2.ListIndex = 0 Then
           gSetup2.gParity = "N"
        End If
        If Combo_Parity2.ListIndex = 1 Then
           gSetup2.gParity = "E"
        End If
        If Combo_Parity2.ListIndex = 2 Then
           gSetup2.gParity = "O"
        End If
            
        If chkACK2.Value = 1 Then
            gSetup2.ACKUse = "1"
        Else
            gSetup2.ACKUse = "0"
        End If
        
        
        Call WritePrivateProfileString("config2", "gPort", gSetup2.gPort, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config2", "gSpeed", gSetup2.gSpeed, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config2", "gParity", gSetup2.gParity, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config2", "gDataBit", gSetup2.gDataBit, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config2", "gStopBit", gSetup2.gStopBit, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config2", "gStartBit", gSetup2.gStartBit, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config2", "ACKUse", gSetup2.ACKUse, App.Path & "\Interface.ini")
        
        If frmInterface.MSComm2.PortOpen = True Then
            frmInterface.MSComm2.PortOpen = False
        End If
        frmInterface.MSComm2.CommPort = gSetup2.gPort
        frmInterface.MSComm2.Settings = gSetup2.gSpeed & "," & gSetup2.gParity & "," & gSetup2.gDataBit & "," & gSetup2.gStopBit
        frmInterface.MSComm2.PortOpen = True
        
        Unload Me
    End If
        
    Exit Sub
 
ErrorHandler:
    Resume Next
'    If MsgBox("통신설정이 맞지 않습니다", vbCritical + vbOKCancel + vbDefaultButton2, "종료버튼") = vbCancel Then
'        Exit Sub
'    End If
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim ret As Integer
    
    'SSPanel_machine.Caption = gEquip
    
    Combo_Port.AddItem ("1")
    Combo_Port.AddItem ("2")
    Combo_Port.AddItem ("3")
    Combo_Port.AddItem ("4")
    Combo_Port.AddItem ("5")
    Combo_Port.AddItem ("6")
    Combo_Port.AddItem ("7")
    Combo_Port.AddItem ("8")
    Combo_Port.AddItem ("9")
    Combo_Port.AddItem ("10")
    
    Combo_BPS.AddItem ("150")
    Combo_BPS.AddItem ("300")
    Combo_BPS.AddItem ("600")
    Combo_BPS.AddItem ("1200")
    Combo_BPS.AddItem ("2400")
    Combo_BPS.AddItem ("4800")
    Combo_BPS.AddItem ("9600")
    Combo_BPS.AddItem ("14400")
    Combo_BPS.AddItem ("19200")
    Combo_BPS.AddItem ("57600")
    
    Combo_Databit.AddItem ("7")
    Combo_Databit.AddItem ("8")
    
    Combo_Startbit.AddItem ("1")
    Combo_Startbit.AddItem ("2")
    
    Combo_Stopbit.AddItem ("1")
    Combo_Stopbit.AddItem ("1.5")
    Combo_Stopbit.AddItem ("2")
    
    Combo_Parity.AddItem ("N")
    Combo_Parity.AddItem ("E")
    Combo_Parity.AddItem ("O")
    
    Combo_Port2.AddItem ("1")
    Combo_Port2.AddItem ("2")
    Combo_Port2.AddItem ("3")
    Combo_Port2.AddItem ("4")
    Combo_Port2.AddItem ("5")
    Combo_Port2.AddItem ("6")
    Combo_Port2.AddItem ("7")
    Combo_Port2.AddItem ("8")
    Combo_Port2.AddItem ("9")
    Combo_Port2.AddItem ("10")
    
    Combo_BPS2.AddItem ("150")
    Combo_BPS2.AddItem ("300")
    Combo_BPS2.AddItem ("600")
    Combo_BPS2.AddItem ("1200")
    Combo_BPS2.AddItem ("2400")
    Combo_BPS2.AddItem ("4800")
    Combo_BPS2.AddItem ("9600")
    Combo_BPS2.AddItem ("14400")
    Combo_BPS2.AddItem ("19200")
    Combo_BPS2.AddItem ("57600")
    
    Combo_Databit2.AddItem ("7")
    Combo_Databit2.AddItem ("8")
    
    Combo_Startbit2.AddItem ("1")
    Combo_Startbit2.AddItem ("2")
    
    Combo_Stopbit2.AddItem ("1")
    Combo_Stopbit2.AddItem ("1.5")
    Combo_Stopbit2.AddItem ("2")
    
    Combo_Parity2.AddItem ("N")
    Combo_Parity2.AddItem ("E")
    Combo_Parity2.AddItem ("O")
    
   
   
'    Combo_Port.ListIndex = 0

    ret = -1
    For i = 0 To Combo_Port.ListCount - 1
        If gSetup.gPort = Trim(Combo_Port.List(i)) Then
            Combo_Port.ListIndex = i
            ret = 1
            Exit For
        End If
    Next i
    If ret = -1 Then
        Combo_Port.ListIndex = 1
    End If
    
    ret = -1
    For i = 0 To Combo_BPS.ListCount - 1
        If gSetup.gSpeed = Trim(Combo_BPS.List(i)) Then
            Combo_BPS.ListIndex = i
            ret = 1
            Exit For
        End If
    Next i
    If ret = -1 Then
        Combo_BPS.ListIndex = 4
    End If
    
    ret = -1
    For i = 0 To Combo_Databit.ListCount - 1
        If gSetup.gDataBit = Trim(Combo_Databit.List(i)) Then
            Combo_Databit.ListIndex = i
            ret = 1
            Exit For
        End If
    Next i
    If ret = -1 Then
        Combo_Databit.ListIndex = 1
    End If

    ret = -1
    For i = 0 To Combo_Startbit.ListCount - 1
        If gSetup.gStartBit = Trim(Combo_Startbit.List(i)) Then
            Combo_Startbit.ListIndex = i
            ret = 1
            Exit For
        End If
    Next i
    If ret = -1 Then
        Combo_Startbit.ListIndex = 0
    End If
    
    ret = -1
    For i = 0 To Combo_Stopbit.ListCount - 1
        If gSetup.gStopBit = Trim(Combo_Stopbit.List(i)) Then
            Combo_Stopbit.ListIndex = i
            ret = 1
            Exit For
        End If
    Next i
    If ret = -1 Then
        Combo_Stopbit.ListIndex = 0
    End If
    
    ret = -1
    For i = 0 To Combo_Parity.ListCount - 1
        If gSetup.gParity = Trim(Combo_Parity.List(i)) Then
            Combo_Parity.ListIndex = i
            ret = 1
            Exit For
        End If
    Next i
    If ret = -1 Then
        Combo_Parity.ListIndex = 0
    End If
    
    If Trim(gSetup.ACKUse) = "1" Then
        chkACK.Value = 1
    Else
        chkACK.Value = 0
    End If
    
    
    
    '''''''''''''''''''-----------------------
    ret = -1
    For i = 0 To Combo_Port2.ListCount - 1
        If gSetup2.gPort = Trim(Combo_Port2.List(i)) Then
            Combo_Port2.ListIndex = i
            ret = 1
            Exit For
        End If
    Next i
    If ret = -1 Then
        Combo_Port2.ListIndex = 1
    End If
    
    ret = -1
    For i = 0 To Combo_BPS2.ListCount - 1
        If gSetup2.gSpeed = Trim(Combo_BPS2.List(i)) Then
            Combo_BPS2.ListIndex = i
            ret = 1
            Exit For
        End If
    Next i
    If ret = -1 Then
        Combo_BPS2.ListIndex = 4
    End If
    
    ret = -1
    For i = 0 To Combo_Databit2.ListCount - 1
        If gSetup2.gDataBit = Trim(Combo_Databit2.List(i)) Then
            Combo_Databit2.ListIndex = i
            ret = 1
            Exit For
        End If
    Next i
    If ret = -1 Then
        Combo_Databit2.ListIndex = 1
    End If

    ret = -1
    For i = 0 To Combo_Startbit2.ListCount - 1
        If gSetup2.gStartBit = Trim(Combo_Startbit2.List(i)) Then
            Combo_Startbit2.ListIndex = i
            ret = 1
            Exit For
        End If
    Next i
    If ret = -1 Then
        Combo_Startbit2.ListIndex = 0
    End If
    
    ret = -1
    For i = 0 To Combo_Stopbit2.ListCount - 1
        If gSetup2.gStopBit = Trim(Combo_Stopbit2.List(i)) Then
            Combo_Stopbit2.ListIndex = i
            ret = 1
            Exit For
        End If
    Next i
    If ret = -1 Then
        Combo_Stopbit2.ListIndex = 0
    End If
    
    ret = -1
    For i = 0 To Combo_Parity2.ListCount - 1
        If gSetup2.gParity = Trim(Combo_Parity2.List(i)) Then
            Combo_Parity2.ListIndex = i
            ret = 1
            Exit For
        End If
    Next i
    If ret = -1 Then
        Combo_Parity2.ListIndex = 0
    End If
    
    If Trim(gSetup2.ACKUse) = "1" Then
        chkACK2.Value = 1
    Else
        chkACK2.Value = 0
    End If
    
    SSPanel1.Visible = True
    SSPanel3.Visible = False
    
    
End Sub

Private Sub spPort1_Click()
    SSPanel1.Visible = True
    SSPanel3.Visible = False
    
End Sub

Private Sub SSPanel2_Click()
    SSPanel1.Visible = False
    SSPanel3.Visible = True
    
End Sub
