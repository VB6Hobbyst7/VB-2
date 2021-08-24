VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Communication Setting"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9495
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
   ScaleHeight     =   6405
   ScaleWidth      =   9495
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame Frame2 
      Caption         =   " TCP-IP "
      Height          =   4785
      Left            =   270
      TabIndex        =   16
      Top             =   330
      Width           =   3795
      Begin VB.TextBox Text2 
         Alignment       =   2  '가운데 맞춤
         Height          =   375
         Left            =   1650
         TabIndex        =   28
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Height          =   375
         Left            =   2490
         TabIndex        =   27
         Top             =   2550
         Width           =   825
      End
      Begin VB.CommandButton Command2 
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
         Left            =   1110
         TabIndex        =   26
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
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
         Left            =   2280
         TabIndex        =   25
         Top             =   3840
         Width           =   1095
      End
      Begin VB.TextBox txtQuanPos 
         Alignment       =   2  '가운데 맞춤
         Height          =   375
         Left            =   1650
         TabIndex        =   24
         Top             =   3030
         Width           =   1665
      End
      Begin VB.TextBox txtQualPos 
         Alignment       =   2  '가운데 맞춤
         Height          =   375
         Left            =   1650
         TabIndex        =   23
         Top             =   2550
         Width           =   825
      End
      Begin VB.OptionButton Option2 
         Caption         =   "정성 Pos"
         Height          =   225
         Left            =   300
         TabIndex        =   22
         Top             =   3090
         Width           =   1305
      End
      Begin VB.OptionButton Option1 
         Caption         =   "정량 Pos"
         Height          =   225
         Left            =   300
         TabIndex        =   21
         Top             =   2670
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.TextBox txtPort 
         Alignment       =   2  '가운데 맞춤
         Height          =   375
         Left            =   1050
         TabIndex        =   19
         Top             =   1260
         Width           =   2295
      End
      Begin VB.TextBox txtIIP 
         Alignment       =   2  '가운데 맞춤
         Height          =   375
         Left            =   1050
         TabIndex        =   17
         Top             =   780
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "Barcode Pos"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   29
         Top             =   2130
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "Port"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   20
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "IP"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   330
         TabIndex        =   18
         Top             =   900
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " RS-232C "
      Height          =   5505
      Left            =   4620
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   4245
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
         Left            =   1050
         TabIndex        =   9
         Top             =   3240
         Width           =   1095
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
         Left            =   2220
         TabIndex        =   8
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CheckBox chkACK 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00E0E0E0&
         Caption         =   "ACK 사용"
         Height          =   375
         Left            =   2220
         TabIndex        =   7
         Top             =   4770
         Visible         =   0   'False
         Width           =   915
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
         ItemData        =   "frmConfig.frx":1272
         Left            =   1800
         List            =   "frmConfig.frx":1274
         TabIndex        =   6
         Top             =   330
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
         ItemData        =   "frmConfig.frx":1276
         Left            =   1800
         List            =   "frmConfig.frx":1278
         TabIndex        =   5
         Top             =   780
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
         ItemData        =   "frmConfig.frx":127A
         Left            =   1800
         List            =   "frmConfig.frx":127C
         TabIndex        =   4
         Top             =   1230
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
         Left            =   1800
         TabIndex        =   3
         Top             =   1680
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
         Left            =   1800
         TabIndex        =   2
         Top             =   2130
         Width           =   1515
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
         ItemData        =   "frmConfig.frx":127E
         Left            =   1800
         List            =   "frmConfig.frx":1280
         TabIndex        =   1
         Top             =   2580
         Width           =   1515
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
         Left            =   315
         TabIndex        =   15
         Top             =   405
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
         Left            =   315
         TabIndex        =   14
         Top             =   855
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
         Left            =   300
         TabIndex        =   13
         Top             =   1305
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
         Left            =   300
         TabIndex        =   12
         Top             =   1755
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
         Left            =   300
         TabIndex        =   11
         Top             =   2205
         Width           =   1305
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
         Left            =   300
         TabIndex        =   10
         Top             =   2655
         Width           =   1305
      End
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
    
        gSetup.gPORT = Combo_Port.Text
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
        
        
        Call WritePrivateProfileString("config", "gPort", gSetup.gPORT, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config", "gSpeed", gSetup.gSpeed, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config", "gParity", gSetup.gParity, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config", "gDataBit", gSetup.gDataBit, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config", "gStopBit", gSetup.gStopBit, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config", "gStartBit", gSetup.gStartBit, App.Path & "\Interface.ini")
        Call WritePrivateProfileString("config", "ACKUse", gSetup.ACKUse, App.Path & "\Interface.ini")
        
        If frmInterface.comEqp.PortOpen = True Then
            frmInterface.comEqp.PortOpen = False
        End If
        frmInterface.comEqp.CommPort = gSetup.gPORT
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
    Combo_Port.AddItem ("10")
    Combo_Port.AddItem ("11")
    Combo_Port.AddItem ("12")
    Combo_Port.AddItem ("13")
    Combo_Port.AddItem ("14")
    Combo_Port.AddItem ("15")
    Combo_Port.AddItem ("16")
    Combo_Port.AddItem ("17")
    Combo_Port.AddItem ("18")
    Combo_Port.AddItem ("19")
    
    Combo_BPS.AddItem ("150")
    Combo_BPS.AddItem ("300")
    Combo_BPS.AddItem ("600")
    Combo_BPS.AddItem ("1200")
    Combo_BPS.AddItem ("2400")
    Combo_BPS.AddItem ("4800")
    Combo_BPS.AddItem ("9600")
    Combo_BPS.AddItem ("14400")
    Combo_BPS.AddItem ("19200")
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
   
   
'    Combo_Port.ListIndex = 0

    Ret = -1
    For i = 0 To Combo_Port.ListCount - 1
        If gSetup.gPORT = Trim(Combo_Port.List(i)) Then
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
