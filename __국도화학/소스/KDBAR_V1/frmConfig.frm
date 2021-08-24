VERSION 5.00
Begin VB.Form frmConfig 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '단일 고정
   Caption         =   "통신설정"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8595
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   8595
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton cmdExit 
      Caption         =   "취소"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7170
      TabIndex        =   33
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "저장"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6000
      TabIndex        =   32
      Top             =   5280
      Width           =   1095
   End
   Begin VB.OptionButton optUse 
      BackColor       =   &H00FFFFFF&
      Caption         =   "사용"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Index           =   1
      Left            =   7080
      TabIndex        =   31
      Top             =   1140
      Width           =   1125
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '위 맞춤
      BackColor       =   &H00808000&
      BorderStyle     =   0  '없음
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   8595
      TabIndex        =   29
      Top             =   0
      Width           =   8595
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "통신정보 설정"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   12
         Left            =   210
         TabIndex        =   30
         Top             =   180
         Width           =   2625
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   90
         Top             =   90
         Width           =   2865
      End
   End
   Begin VB.OptionButton optUse 
      BackColor       =   &H00FFFFFF&
      Caption         =   "사용"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Index           =   0
      Left            =   2910
      TabIndex        =   0
      Top             =   1170
      Value           =   -1  'True
      Width           =   1125
   End
   Begin VB.TextBox txtSaveDay 
      Alignment       =   2  '가운데 맞춤
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   6090
      TabIndex        =   26
      Top             =   4290
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Caption         =   " 소켓 방식 "
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   4530
      TabIndex        =   18
      Top             =   900
      Width           =   3825
      Begin VB.OptionButton optType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Client"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   22
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton optType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Server"
         Height          =   255
         Index           =   0
         Left            =   1380
         TabIndex        =   21
         Top             =   960
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox txtPort 
         Alignment       =   2  '가운데 맞춤
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   1260
         TabIndex        =   20
         Text            =   "5050"
         Top             =   1890
         Width           =   1815
      End
      Begin VB.TextBox txtIP 
         Alignment       =   2  '가운데 맞춤
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   1260
         TabIndex        =   19
         Text            =   "127.0.0.1"
         Top             =   1380
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "Port"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   -270
         TabIndex        =   25
         Top             =   1950
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "IP"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   -255
         TabIndex        =   24
         Top             =   1425
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   -255
         TabIndex        =   23
         Top             =   975
         Width           =   1305
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 시리얼 방식 "
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4875
      Left            =   180
      TabIndex        =   1
      Top             =   900
      Width           =   4095
      Begin VB.CheckBox chkRTS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "True"
         Height          =   315
         Left            =   1830
         TabIndex        =   9
         Top             =   3720
         Value           =   1  '확인
         Width           =   1785
      End
      Begin VB.CheckBox chkDTR 
         BackColor       =   &H00FFFFFF&
         Caption         =   "True"
         Height          =   315
         Left            =   1830
         TabIndex        =   8
         Top             =   4230
         Value           =   1  '확인
         Width           =   1785
      End
      Begin VB.ComboBox Combo_Parity 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   7
         Top             =   3240
         Width           =   1575
      End
      Begin VB.ComboBox Combo_Stopbit 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   6
         Top             =   2790
         Width           =   1605
      End
      Begin VB.ComboBox Combo_Startbit 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   5
         Top             =   2340
         Width           =   1605
      End
      Begin VB.ComboBox Combo_Databit 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   4
         Top             =   1890
         Width           =   1605
      End
      Begin VB.ComboBox Combo_BPS 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   3
         Top             =   1440
         Width           =   1605
      End
      Begin VB.ComboBox Combo_Port 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   2
         Top             =   990
         Width           =   1605
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "RTSEnable"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   330
         TabIndex        =   17
         Top             =   3780
         Width           =   1035
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "DTREnable"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   48
         Left            =   330
         TabIndex        =   16
         Top             =   4290
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "Parity"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   60
         TabIndex        =   15
         Top             =   3315
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "Stop Bit"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   60
         TabIndex        =   14
         Top             =   2865
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "Start Bit"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   60
         TabIndex        =   13
         Top             =   2415
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "Data Bit"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   60
         TabIndex        =   12
         Top             =   1965
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "Speed"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   75
         TabIndex        =   11
         Top             =   1515
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "Com Port"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   11
         Left            =   75
         TabIndex        =   10
         Top             =   1065
         Width           =   1305
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "검사결과          로컬저장기간"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   390
      Index           =   7
      Left            =   4650
      TabIndex        =   28
      Top             =   4230
      Width           =   1380
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "일 저장"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   8
      Left            =   7350
      TabIndex        =   27
      Top             =   4380
      Width           =   645
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdConfirm_Click()
    
    On Error GoTo ErrorHandler
    
    If MsgBox("설정을 저장하시겠습니까?", vbCritical + vbOKCancel + vbDefaultButton2, "확인!") = vbCancel Then
        Unload Me
        Exit Sub
    Else
        If optUse(0).Value = True Then
            Call WritePrivateProfileString("COMM", "COMTYPE", "1", App.PATH & "\KDBAR.ini")
        ElseIf optUse(1).Value = True Then
            Call WritePrivateProfileString("COMM", "COMTYPE", "2", App.PATH & "\KDBAR.ini")
        End If
        
        If optUse(0).Value = True Then
            gComm.COMPORT = Combo_Port.Text
            gComm.SPEED = Combo_BPS.Text
            gComm.DATABIT = Combo_Databit.Text
            gComm.STARTBIT = Combo_Startbit.Text
            gComm.STOPBIT = Combo_Stopbit.Text
            If Combo_Parity.ListIndex = 0 Then
               gComm.Parity = "N"
            End If
            If Combo_Parity.ListIndex = 1 Then
               gComm.Parity = "E"
            End If
            If Combo_Parity.ListIndex = 2 Then
               gComm.Parity = "O"
            End If
            
            Call WritePrivateProfileString("COMM", "COMPORT", gComm.COMPORT, App.PATH & "\KDBAR.ini")
            Call WritePrivateProfileString("COMM", "SPEED", gComm.SPEED, App.PATH & "\KDBAR.ini")
            Call WritePrivateProfileString("COMM", "PARITY", gComm.Parity, App.PATH & "\KDBAR.ini")
            Call WritePrivateProfileString("COMM", "DATABIT", gComm.DATABIT, App.PATH & "\KDBAR.ini")
            Call WritePrivateProfileString("COMM", "STOPBIT", gComm.STOPBIT, App.PATH & "\KDBAR.ini")
            Call WritePrivateProfileString("COMM", "STARTBIT", gComm.STARTBIT, App.PATH & "\KDBAR.ini")
            
            If chkRTS.Value = "1" Then
                Call WritePrivateProfileString("COMM", "RTSEnable", "True", App.PATH & "\KDBAR.ini")
            Else
                Call WritePrivateProfileString("COMM", "RTSEnable", "False", App.PATH & "\KDBAR.ini")
            End If
            If chkRTS.Value = "1" Then
                Call WritePrivateProfileString("COMM", "DTREnable", "True", App.PATH & "\KDBAR.ini")
            Else
                Call WritePrivateProfileString("COMM", "DTREnable", "False", App.PATH & "\KDBAR.ini")
            End If
            
            If frmMain.comEqp.PortOpen = True Then
                frmMain.comEqp.PortOpen = False
            End If
            
            frmMain.comEqp.CommPort = gComm.COMPORT
            frmMain.comEqp.Settings = gComm.SPEED & "," & gComm.Parity & "," & gComm.DATABIT & "," & gComm.STOPBIT
            
            frmMain.comEqp.PortOpen = True
        
        ElseIf optUse(1).Value = True Then
            gComm.TCPIP = txtIP.Text
            gComm.TCPPORT = txtPort.Text
            
            If optType(0).Value = True Then
                gComm.TCPTYPE = "SERVER"
            Else
                gComm.TCPTYPE = "CLIENT"
            End If
    
            Call WritePrivateProfileString("COMM", "TCPTYPE", gComm.TCPTYPE, App.PATH & "\KDBAR.ini")
            Call WritePrivateProfileString("COMM", "TCPIP", gComm.TCPIP, App.PATH & "\KDBAR.ini")
            Call WritePrivateProfileString("COMM", "TCPPORT", gComm.TCPPORT, App.PATH & "\KDBAR.ini")
        
        End If
        
        Call WritePrivateProfileString("HOSP", "SAVEDAY", txtSaveDay.Text, App.PATH & "\KDBAR.ini")
        
        Unload Me
    End If
        
    Exit Sub
 
ErrorHandler:
    Resume Next
    If MsgBox("통신설정이 맞지 않습니다", vbCritical + vbOKCancel + vbDefaultButton2, "종료버튼") = vbCancel Then
        Unload Me
    End If

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()
    Dim intComPortExist As Long
    Dim i As Integer
    Dim Ret As Integer
    
    If gComm.COMTYPE = "1" Then
        optUse(0).Value = True
        
        Combo_Port.Clear
        For i = 1 To 16
            intComPortExist = EnumSerPorts(i)
            If intComPortExist > 0 Then
                Combo_Port.AddItem Trim(Str(i))
            End If
        Next
        
        Combo_BPS.AddItem ("150")
        Combo_BPS.AddItem ("300")
        Combo_BPS.AddItem ("600")
        Combo_BPS.AddItem ("1200")
        Combo_BPS.AddItem ("2400")
        Combo_BPS.AddItem ("4800")
        Combo_BPS.AddItem ("9600")
        Combo_BPS.AddItem ("14400")
        Combo_BPS.AddItem ("19200")
        Combo_BPS.AddItem ("38400")
        Combo_BPS.AddItem ("115200")
        
        
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
       
        Ret = -1
        For i = 0 To Combo_Port.ListCount - 1
            If gComm.COMPORT = Trim(Combo_Port.List(i)) Then
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
            If gComm.SPEED = Trim(Combo_BPS.List(i)) Then
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
            If gComm.DATABIT = Trim(Combo_Databit.List(i)) Then
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
            If gComm.STARTBIT = Trim(Combo_Startbit.List(i)) Then
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
            If gComm.STOPBIT = Trim(Combo_Stopbit.List(i)) Then
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
            If gComm.Parity = Trim(Combo_Parity.List(i)) Then
                Combo_Parity.ListIndex = i
                Ret = 1
                Exit For
            End If
        Next i
        If Ret = -1 Then
            Combo_Parity.ListIndex = 0
        End If
        
        If gComm.RTSEnable = True Then
            chkRTS.Value = "1"
        Else
            chkRTS.Value = "0"
        End If
        
        If gComm.DTREnable = True Then
            chkDTR.Value = "1"
        Else
            chkDTR.Value = "0"
        End If
            
    ElseIf gComm.COMTYPE = "2" Then
        optUse(1).Value = True
        
        If gComm.TCPTYPE = "SERVER" Then
            optType(0).Value = True
        Else
            optType(1).Value = True
        End If
        
        txtIP.Text = gComm.TCPIP
        txtPort.Text = gComm.TCPPORT
    End If
    
    txtSaveDay.Text = gKUKDO.SAVEDAY
    
End Sub

