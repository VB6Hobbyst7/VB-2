VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmConfig 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '단일 고정
   Caption         =   "포트설정"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
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
   ScaleHeight     =   2235
   ScaleWidth      =   3735
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox txtResultPort 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1680
      TabIndex        =   25
      Top             =   750
      Width           =   1665
   End
   Begin VB.TextBox txtOrderPort 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1680
      TabIndex        =   23
      Top             =   330
      Width           =   1665
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2040
      TabIndex        =   22
      Top             =   1470
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   660
      TabIndex        =   21
      Top             =   1470
      Width           =   1335
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   1845
      Left            =   6240
      TabIndex        =   16
      Top             =   1050
      Visible         =   0   'False
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   3254
      _Version        =   131072
      Begin VB.TextBox txtTestIdName 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1050
         TabIndex        =   20
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox txtTestWay 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1050
         TabIndex        =   19
         Top             =   390
         Width           =   4095
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "보고자"
         Height          =   195
         Index           =   7
         Left            =   -150
         TabIndex        =   18
         Top             =   930
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "검사방법"
         Height          =   195
         Index           =   6
         Left            =   -150
         TabIndex        =   17
         Top             =   480
         Width           =   1155
      End
   End
   Begin IF_Cobas4800국립암센터.MDButton cmdExit 
      Height          =   465
      Left            =   8610
      TabIndex        =   15
      Top             =   5190
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   820
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "종료"
   End
   Begin IF_Cobas4800국립암센터.MDButton cmdConfirm 
      Height          =   465
      Left            =   7170
      TabIndex        =   14
      Top             =   5190
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   820
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "확인"
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   735
      Left            =   7830
      TabIndex        =   0
      Top             =   3210
      Visible         =   0   'False
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   1296
      _Version        =   131072
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox Combo_Parity 
         Height          =   315
         ItemData        =   "frmConfig.frx":0000
         Left            =   1785
         List            =   "frmConfig.frx":0002
         TabIndex        =   7
         Top             =   2580
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox Combo_Stopbit 
         Height          =   315
         Left            =   1785
         TabIndex        =   6
         Top             =   2100
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox Combo_Startbit 
         Height          =   315
         Left            =   1785
         TabIndex        =   5
         Top             =   1650
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox Combo_Databit 
         Height          =   315
         ItemData        =   "frmConfig.frx":0004
         Left            =   1785
         List            =   "frmConfig.frx":0006
         TabIndex        =   4
         Top             =   1200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox Combo_BPS 
         Height          =   315
         ItemData        =   "frmConfig.frx":0008
         Left            =   1785
         List            =   "frmConfig.frx":000A
         TabIndex        =   3
         Top             =   750
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox Combo_Port 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frmConfig.frx":000C
         Left            =   1065
         List            =   "frmConfig.frx":000E
         TabIndex        =   2
         Top             =   210
         Width           =   1695
      End
      Begin VB.CheckBox chkACK 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00E0E0E0&
         Caption         =   "ACK 사용"
         Height          =   315
         Left            =   690
         TabIndex        =   1
         Top             =   3060
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "패리티"
         Height          =   195
         Index           =   5
         Left            =   405
         TabIndex        =   13
         Top             =   2640
         Visible         =   0   'False
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
         Top             =   2160
         Visible         =   0   'False
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
         Top             =   1710
         Visible         =   0   'False
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
         Top             =   1260
         Visible         =   0   'False
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
         Top             =   810
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "COM PORT"
         Height          =   195
         Index           =   0
         Left            =   -150
         TabIndex        =   8
         Top             =   270
         Width           =   1155
      End
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '투명
      Caption         =   "결과 Port :"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   390
      TabIndex        =   26
      Top             =   840
      Width           =   1305
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '투명
      Caption         =   "오더 Port :"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   390
      TabIndex        =   24
      Top             =   420
      Width           =   1125
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    
    Unload Me

End Sub

'Serial Setting

Private Sub cmdConfirm_Click()
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
        gSetup.gTestWay = txtTestWay.Text
        gSetup.gTestIdName = txtTestIdName.Text
        
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
        
        
        Call WritePrivateProfileString("config", "gPort", gSetup.gPort, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config", "gSpeed", gSetup.gSpeed, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config", "gParity", gSetup.gParity, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config", "gDataBit", gSetup.gDataBit, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config", "gStopBit", gSetup.gStopBit, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config", "gStartBit", gSetup.gStartBit, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config", "ACKUse", gSetup.ACKUse, App.Path & "\Interface.ini")
        
        Call WritePrivateProfileString("config", "gTestWay", gSetup.gTestWay, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config", "gTestIdName", gSetup.gTestIdName, App.Path & "\Interface.ini")
        
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

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    
    If MsgBox("설정을 저장하시겠습니까?", vbCritical + vbOKCancel + vbDefaultButton2, "확인!") = vbCancel Then
        Unload Me
        Exit Sub
    Else
        Call WritePrivateProfileString("config", "gPort", txtOrderPort.Text, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config", "gPort2", txtResultPort.Text, App.Path & "\Interface.ini")
        
    End If

End Sub

Private Sub Form_Load()
'    Dim i As Integer
'    Dim Ret As Integer
'
'    'SSPanel_machine.Caption = gEquip
'
''''    Combo_Port.AddItem ("1")
''''    Combo_Port.AddItem ("2")
''''    Combo_Port.AddItem ("3")
''''    Combo_Port.AddItem ("4")
''''    Combo_Port.AddItem ("5")
'
'    Combo_Port.AddItem ("5555")
'    Combo_Port.AddItem ("10000")
'
'    Combo_BPS.AddItem ("150")
'    Combo_BPS.AddItem ("300")
'    Combo_BPS.AddItem ("600")
'    Combo_BPS.AddItem ("1200")
'    Combo_BPS.AddItem ("2400")
'    Combo_BPS.AddItem ("4800")
'    Combo_BPS.AddItem ("9600")
'    Combo_BPS.AddItem ("14400")
'    Combo_BPS.AddItem ("19200")
'
'    Combo_Databit.AddItem ("7")
'    Combo_Databit.AddItem ("8")
'
'    Combo_Startbit.AddItem ("1")
'    Combo_Startbit.AddItem ("2")
'
'    Combo_Stopbit.AddItem ("1")
'    Combo_Stopbit.AddItem ("1.5")
'    Combo_Stopbit.AddItem ("2")
'
'    Combo_Parity.AddItem ("N")
'    Combo_Parity.AddItem ("E")
'    Combo_Parity.AddItem ("O")
'
'
''    Combo_Port.ListIndex = 0
'
'    Ret = -1
'    For i = 0 To Combo_Port.ListCount - 1
'        If gSetup.gPort = Trim(Combo_Port.List(i)) Then
'            Combo_Port.ListIndex = i
'            Ret = 1
'            Exit For
'        End If
'    Next i
'    If Ret = -1 Then
'        Combo_Port.ListIndex = 1
'    End If
'
'    Ret = -1
'    For i = 0 To Combo_BPS.ListCount - 1
'        If gSetup.gSpeed = Trim(Combo_BPS.List(i)) Then
'            Combo_BPS.ListIndex = i
'            Ret = 1
'            Exit For
'        End If
'    Next i
'    If Ret = -1 Then
'        Combo_BPS.ListIndex = 4
'    End If
'
'    Ret = -1
'    For i = 0 To Combo_Databit.ListCount - 1
'        If gSetup.gDataBit = Trim(Combo_Databit.List(i)) Then
'            Combo_Databit.ListIndex = i
'            Ret = 1
'            Exit For
'        End If
'    Next i
'    If Ret = -1 Then
'        Combo_Databit.ListIndex = 1
'    End If
'
'    Ret = -1
'    For i = 0 To Combo_Startbit.ListCount - 1
'        If gSetup.gStartBit = Trim(Combo_Startbit.List(i)) Then
'            Combo_Startbit.ListIndex = i
'            Ret = 1
'            Exit For
'        End If
'    Next i
'    If Ret = -1 Then
'        Combo_Startbit.ListIndex = 0
'    End If
'
'    Ret = -1
'    For i = 0 To Combo_Stopbit.ListCount - 1
'        If gSetup.gStopBit = Trim(Combo_Stopbit.List(i)) Then
'            Combo_Stopbit.ListIndex = i
'            Ret = 1
'            Exit For
'        End If
'    Next i
'    If Ret = -1 Then
'        Combo_Stopbit.ListIndex = 0
'    End If
'
'    Ret = -1
'    For i = 0 To Combo_Parity.ListCount - 1
'        If gSetup.gParity = Trim(Combo_Parity.List(i)) Then
'            Combo_Parity.ListIndex = i
'            Ret = 1
'            Exit For
'        End If
'    Next i
'    If Ret = -1 Then
'        Combo_Parity.ListIndex = 0
'    End If
'
'    If Trim(gSetup.ACKUse) = "1" Then
'        chkACK.Value = 1
'    Else
'        chkACK.Value = 0
'    End If
    
'    txtTestWay.Text = gSetup.gTestWay
'    txtTestIdName.Text = gSetup.gTestIdName
    
    txtOrderPort.Text = gSetup.gPort
    txtResultPort.Text = gSetup.gPort2
    
    
End Sub
