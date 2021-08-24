VERSION 5.00
Begin VB.Form frmConfig 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '단일 고정
   Caption         =   "통신설정"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   4650
   StartUpPosition =   1  '소유자 가운데
   Begin VB.TextBox txtSaveDay 
      Alignment       =   2  '가운데 맞춤
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   2010
      TabIndex        =   19
      Top             =   4260
      Width           =   975
   End
   Begin VB.CheckBox chkRTS 
      BackColor       =   &H00FFFFFF&
      Caption         =   "True"
      Height          =   315
      Left            =   2040
      TabIndex        =   15
      Top             =   3240
      Value           =   1  '확인
      Width           =   1785
   End
   Begin VB.CheckBox chkDTR 
      BackColor       =   &H00FFFFFF&
      Caption         =   "True"
      Height          =   315
      Left            =   2040
      TabIndex        =   14
      Top             =   3750
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
      Left            =   2010
      TabIndex        =   7
      Top             =   2760
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
      Left            =   2010
      TabIndex        =   6
      Top             =   2310
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
      Left            =   2010
      TabIndex        =   5
      Top             =   1860
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
      Left            =   2010
      TabIndex        =   4
      Top             =   1410
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
      Left            =   2010
      TabIndex        =   3
      Top             =   960
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
      Left            =   2010
      TabIndex        =   2
      Top             =   510
      Width           =   1605
   End
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
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   4890
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
      Height          =   375
      Left            =   1950
      TabIndex        =   0
      Top             =   4890
      Width           =   1095
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
      Left            =   3120
      TabIndex        =   20
      Top             =   4320
      Width           =   645
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "로컬저장기간"
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
      Index           =   7
      Left            =   540
      TabIndex        =   18
      Top             =   4320
      Width           =   1170
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
      Left            =   540
      TabIndex        =   17
      Top             =   3300
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
      Left            =   540
      TabIndex        =   16
      Top             =   3810
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
      Left            =   270
      TabIndex        =   13
      Top             =   2835
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
      Left            =   270
      TabIndex        =   12
      Top             =   2385
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
      Left            =   270
      TabIndex        =   11
      Top             =   1935
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
      Index           =   2
      Left            =   270
      TabIndex        =   10
      Top             =   1485
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
      Index           =   1
      Left            =   285
      TabIndex        =   9
      Top             =   1035
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
      Index           =   0
      Left            =   285
      TabIndex        =   8
      Top             =   585
      Width           =   1305
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
        
        Call WritePrivateProfileString("COMM", "COMPORT", gComm.COMPORT, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("COMM", "SPEED", gComm.SPEED, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("COMM", "PARITY", gComm.Parity, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("COMM", "DATABIT", gComm.DATABIT, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("COMM", "STOPBIT", gComm.STOPBIT, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("COMM", "STARTBIT", gComm.STARTBIT, App.PATH & "\INI\" & gMACH & ".ini")
        
        If chkRTS.Value = "1" Then
            Call WritePrivateProfileString("COMM", "RTSEnable", "True", App.PATH & "\INI\" & gMACH & ".ini")
        Else
            Call WritePrivateProfileString("COMM", "RTSEnable", "False", App.PATH & "\INI\" & gMACH & ".ini")
        End If
        If chkRTS.Value = "1" Then
            Call WritePrivateProfileString("COMM", "DTREnable", "True", App.PATH & "\INI\" & gMACH & ".ini")
        Else
            Call WritePrivateProfileString("COMM", "DTREnable", "False", App.PATH & "\INI\" & gMACH & ".ini")
        End If
        
        Call WritePrivateProfileString("HOSP", "SAVEDAY", txtSaveDay.Text, App.PATH & "\INI\" & gMACH & ".ini")
        
        If frmMain.comEqp.PortOpen = True Then
            frmMain.comEqp.PortOpen = False
        End If
        
        frmMain.comEqp.CommPort = gComm.COMPORT
        frmMain.comEqp.Settings = gComm.SPEED & "," & gComm.Parity & "," & gComm.DATABIT & "," & gComm.STOPBIT
        
        frmMain.comEqp.PortOpen = True
        
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

Private Sub Form_Load()
    Dim intComPortExist As Long
    Dim i As Integer
    Dim Ret As Integer
    
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
    
    txtSaveDay.Text = gHOSP.SAVEDAY
    
End Sub
