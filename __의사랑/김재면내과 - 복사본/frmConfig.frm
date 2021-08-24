VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmConfig 
   Caption         =   "통신설정"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3525
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4755
   ScaleWidth      =   3525
   StartUpPosition =   2  '화면 가운데
   Begin Threed.SSPanel SSPanel1 
      Height          =   3435
      Left            =   30
      TabIndex        =   3
      Top             =   570
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   6059
      _Version        =   131072
      BackColor       =   13160660
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
      Begin VB.CheckBox chkDTR 
         Height          =   225
         Left            =   1545
         TabIndex        =   17
         Top             =   2940
         Width           =   345
      End
      Begin VB.CheckBox chkRTS 
         Height          =   225
         Left            =   1545
         TabIndex        =   16
         Top             =   2595
         Width           =   345
      End
      Begin VB.ComboBox Combo_Port 
         Height          =   315
         ItemData        =   "frmConfig.frx":0442
         Left            =   1545
         List            =   "frmConfig.frx":0444
         TabIndex        =   14
         Top             =   150
         Width           =   1695
      End
      Begin VB.ComboBox Combo_BPS 
         Height          =   315
         ItemData        =   "frmConfig.frx":0446
         Left            =   1545
         List            =   "frmConfig.frx":0448
         TabIndex        =   8
         Top             =   540
         Width           =   1695
      End
      Begin VB.ComboBox Combo_Databit 
         Height          =   315
         ItemData        =   "frmConfig.frx":044A
         Left            =   1545
         List            =   "frmConfig.frx":044C
         TabIndex        =   7
         Top             =   930
         Width           =   1695
      End
      Begin VB.ComboBox Combo_Startbit 
         Height          =   315
         Left            =   1545
         TabIndex        =   6
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox Combo_Stopbit 
         Height          =   315
         Left            =   1545
         TabIndex        =   5
         Top             =   1710
         Width           =   1695
      End
      Begin VB.ComboBox Combo_Parity 
         Height          =   315
         ItemData        =   "frmConfig.frx":044E
         Left            =   1545
         List            =   "frmConfig.frx":0450
         TabIndex        =   4
         Top             =   2130
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "DTR Enabled"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   19
         Top             =   2955
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "RTS Enabled"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   18
         Top             =   2610
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "COM PORT"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   210
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "전송속도"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "데이터 비트"
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   12
         Top             =   990
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "시작 비트"
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   11
         Top             =   1380
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "정지 비트"
         Height          =   195
         Index           =   4
         Left            =   225
         TabIndex        =   10
         Top             =   1770
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "패리티"
         Height          =   195
         Index           =   5
         Left            =   225
         TabIndex        =   9
         Top             =   2190
         Width           =   1155
      End
   End
   Begin Threed.SSPanel SSPanel_machine 
      Height          =   525
      Left            =   30
      TabIndex        =   2
      Top             =   30
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   926
      _Version        =   131072
      ForeColor       =   12582912
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "종료"
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
      Left            =   1868
      TabIndex        =   1
      Top             =   4200
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "확인"
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
      Left            =   540
      TabIndex        =   0
      Top             =   4200
      Width           =   1155
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
    Dim sRTS As String
    Dim sDTR As String
    
    On Error GoTo ErrorHandler
    
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
    If chkRTS.Value = 1 Then
        gSetup.gRTSEnable = True
        sRTS = "T"
    Else
        gSetup.gRTSEnable = False
        sRTS = "F"
    End If
    
    If chkDTR.Value = 1 Then
        gSetup.gDTREnable = True
        sDTR = "T"
    Else
        gSetup.gDTREnable = False
        sDTR = "F"
    End If
    
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
        
        Call WritePrivateProfileString("config", "gPort", gSetup.gPort, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config", "gSpeed", gSetup.gSpeed, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config", "gParity", gSetup.gParity, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config", "gDataBit", gSetup.gDataBit, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config", "gStopBit", gSetup.gStopBit, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config", "gStartBit", gSetup.gStartBit, App.Path & "\Interface.ini")
        
        Call WritePrivateProfileString("config", "gRTSEnable", CStr(gSetup.gRTSEnable), App.Path & "\interface.ini")
        Call WritePrivateProfileString("config", "gDTREnable", CStr(gSetup.gDTREnable), App.Path & "\interface.ini")
        
        Unload Me
        
    End If
        
    Exit Sub
 
ErrorHandler:
    If MsgBox("통신설정이 맞지 않습니다", vbCritical + vbOKCancel + vbDefaultButton2, "종료버튼") = vbCancel Then
        Exit Sub
    End If
    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim Ret As Integer
    
    Combo_Port.AddItem ("1")
    Combo_Port.AddItem ("2")
    
    Combo_BPS.AddItem ("150")
    Combo_BPS.AddItem ("300")
    Combo_BPS.AddItem ("600")
    Combo_BPS.AddItem ("1200")
    Combo_BPS.AddItem ("2400")
    Combo_BPS.AddItem ("4800")
    Combo_BPS.AddItem ("9600")
    Combo_BPS.AddItem ("14400")
    Combo_BPS.AddItem ("19200")
    
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
   
   
'    Combo_Port.ListIndex = 0

    Ret = -1
    For i = 0 To Combo_Port.ListCount - 1
        If gSetup.gPort = Trim(Combo_Port.List(i)) Then
            Combo_Port.ListIndex = i
            Ret = 1
            Exit For
        End If
    Next i
    If Ret = -1 Then
        Combo_Port.ListIndex = 1
    End If
    
    Ret = -1
    For i = 0 To Combo_BPS.ListCount - 1
        If gSetup.gSpeed = Trim(Combo_BPS.List(i)) Then
            Combo_BPS.ListIndex = i
            Ret = 1
            Exit For
        End If
    Next i
    If Ret = -1 Then
        Combo_BPS.ListIndex = 4
    End If
    
    Ret = -1
    For i = 0 To Combo_Databit.ListCount - 1
        If gSetup.gDataBit = Trim(Combo_Databit.List(i)) Then
            Combo_Databit.ListIndex = i
            Ret = 1
            Exit For
        End If
    Next i
    If Ret = -1 Then
        Combo_Databit.ListIndex = 1
    End If

    Ret = -1
    For i = 0 To Combo_Startbit.ListCount - 1
        If gSetup.gStartBit = Trim(Combo_Startbit.List(i)) Then
            Combo_Startbit.ListIndex = i
            Ret = 1
            Exit For
        End If
    Next i
    If Ret = -1 Then
        Combo_Startbit.ListIndex = 0
    End If
    
    Ret = -1
    For i = 0 To Combo_Stopbit.ListCount - 1
        If gSetup.gStopBit = Trim(Combo_Stopbit.List(i)) Then
            Combo_Stopbit.ListIndex = i
            Ret = 1
            Exit For
        End If
    Next i
    If Ret = -1 Then
        Combo_Stopbit.ListIndex = 0
    End If
    
    Ret = -1
    For i = 0 To Combo_Parity.ListCount - 1
        If gSetup.gParity = Trim(Combo_Parity.List(i)) Then
            Combo_Parity.ListIndex = i
            Ret = 1
            Exit For
        End If
    Next i
    If Ret = -1 Then
        Combo_Parity.ListIndex = 0
    End If
    
    If gSetup.gRTSEnable = True Then
        chkRTS.Value = 1
    Else
        chkRTS.Value = 0
    End If
    
    If gSetup.gDTREnable = True Then
        chkDTR.Value = 1
    Else
        chkDTR.Value = 0
    End If
End Sub


