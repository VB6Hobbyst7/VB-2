VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Communication Setting"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3810
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   3810
   StartUpPosition =   1  '소유자 가운데
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
      ItemData        =   "frmConfig.frx":1272
      Left            =   1620
      List            =   "frmConfig.frx":1274
      TabIndex        =   8
      Top             =   2760
      Width           =   1515
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
      Left            =   1620
      TabIndex        =   7
      Top             =   2310
      Width           =   1515
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
      Left            =   1620
      TabIndex        =   6
      Top             =   1860
      Width           =   1515
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
      ItemData        =   "frmConfig.frx":1276
      Left            =   1620
      List            =   "frmConfig.frx":1278
      TabIndex        =   5
      Top             =   1410
      Width           =   1515
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
      ItemData        =   "frmConfig.frx":127A
      Left            =   1620
      List            =   "frmConfig.frx":127C
      TabIndex        =   4
      Top             =   960
      Width           =   1515
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
      ItemData        =   "frmConfig.frx":127E
      Left            =   1620
      List            =   "frmConfig.frx":1280
      TabIndex        =   3
      Top             =   510
      Width           =   1515
   End
   Begin VB.CheckBox chkACK 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00E0E0E0&
      Caption         =   "ACK 사용"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   4950
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2040
      TabIndex        =   1
      Top             =   3420
      Width           =   1095
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   870
      TabIndex        =   0
      Top             =   3420
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Parity"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   14
      Top             =   2835
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Stop Bit"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   2385
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Start Bit"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   1935
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Data Bit"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   1485
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Baud Rate"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   135
      TabIndex        =   10
      Top             =   1035
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Com Port"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   135
      TabIndex        =   9
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
        
        
        Call WritePrivateProfileString("config", "gPort", gSetup.gPort, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config", "gSpeed", gSetup.gSpeed, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config", "gParity", gSetup.gParity, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config", "gDataBit", gSetup.gDataBit, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config", "gStopBit", gSetup.gStopBit, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config", "gStartBit", gSetup.gStartBit, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config", "ACKUse", gSetup.ACKUse, App.Path & "\Interface.ini")
        
        If frmInterface.comEqp.PortOpen = True Then
            frmInterface.comEqp.PortOpen = False
        End If
        frmInterface.comEqp.CommPort = gSetup.gPort
        frmInterface.comEqp.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit
        
        frmInterface.comEqp.PortOpen = True
        
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

Private Sub Form_Load()
    Dim i As Integer
    Dim Ret As Integer
    
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
    
    If Trim(gSetup.ACKUse) = "1" Then
        chkACK.Value = 1
    Else
        chkACK.Value = 0
    End If
    
    Me.Width = 4035
    Me.Height = 5040
End Sub
