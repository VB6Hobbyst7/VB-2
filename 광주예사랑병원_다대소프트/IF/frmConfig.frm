VERSION 5.00
Begin VB.Form frmConfig 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '���� ����
   Caption         =   "��ż���"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985
   BeginProperty Font 
      Name            =   "����ü"
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
   ScaleHeight     =   5910
   ScaleWidth      =   8985
   StartUpPosition =   1  '������ ���
   Begin VB.PictureBox Picture1 
      Align           =   1  '�� ����
      BackColor       =   &H00808000&
      BorderStyle     =   0  '����
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   8985
      TabIndex        =   32
      Top             =   0
      Width           =   8985
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackStyle       =   0  '����
         Caption         =   "��� ����"
         BeginProperty Font 
            Name            =   "���� ���"
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
         TabIndex        =   33
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
      Caption         =   "���"
      Height          =   345
      Index           =   1
      Left            =   5520
      TabIndex        =   0
      Top             =   150
      Width           =   1125
   End
   Begin VB.OptionButton optUse 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���"
      Height          =   345
      Index           =   0
      Left            =   1140
      TabIndex        =   1
      Top             =   150
      Value           =   -1  'True
      Width           =   1125
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   29
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7650
      TabIndex        =   28
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox txtSaveDay 
      Alignment       =   2  '��� ����
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   6540
      TabIndex        =   27
      Top             =   4290
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   " TCP-IP "
      Height          =   2775
      Left            =   4530
      TabIndex        =   19
      Top             =   900
      Width           =   4275
      Begin VB.OptionButton optType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Client"
         Height          =   255
         Index           =   1
         Left            =   2820
         TabIndex        =   23
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Server"
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   22
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox txtPort 
         Alignment       =   2  '��� ����
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   1680
         TabIndex        =   21
         Text            =   "5050"
         Top             =   1410
         Width           =   1815
      End
      Begin VB.TextBox txtIP 
         Alignment       =   2  '��� ����
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   1680
         TabIndex        =   20
         Text            =   "127.0.0.1"
         Top             =   900
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "Port"
         BeginProperty Font 
            Name            =   "����"
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
         Left            =   150
         TabIndex        =   26
         Top             =   1470
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "IP"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   165
         TabIndex        =   25
         Top             =   945
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "����"
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
         Left            =   165
         TabIndex        =   24
         Top             =   495
         Width           =   1305
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Serial "
      Height          =   4395
      Left            =   180
      TabIndex        =   2
      Top             =   900
      Width           =   4095
      Begin VB.CheckBox chkRTS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "True"
         Height          =   315
         Left            =   1770
         TabIndex        =   10
         Top             =   3120
         Value           =   1  'Ȯ��
         Width           =   1785
      End
      Begin VB.CheckBox chkDTR 
         BackColor       =   &H00FFFFFF&
         Caption         =   "True"
         Height          =   315
         Left            =   1770
         TabIndex        =   9
         Top             =   3630
         Value           =   1  'Ȯ��
         Width           =   1785
      End
      Begin VB.ComboBox Combo_Parity 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1740
         TabIndex        =   8
         Top             =   2640
         Width           =   1575
      End
      Begin VB.ComboBox Combo_Stopbit 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1740
         TabIndex        =   7
         Top             =   2190
         Width           =   1605
      End
      Begin VB.ComboBox Combo_Startbit 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1740
         TabIndex        =   6
         Top             =   1740
         Width           =   1605
      End
      Begin VB.ComboBox Combo_Databit 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1740
         TabIndex        =   5
         Top             =   1290
         Width           =   1605
      End
      Begin VB.ComboBox Combo_BPS 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1740
         TabIndex        =   4
         Top             =   840
         Width           =   1605
      End
      Begin VB.ComboBox Combo_Port 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1740
         TabIndex        =   3
         Top             =   390
         Width           =   1605
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "RTSEnable"
         BeginProperty Font 
            Name            =   "����"
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
         Left            =   270
         TabIndex        =   18
         Top             =   3180
         Width           =   1035
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "DTREnable"
         BeginProperty Font 
            Name            =   "����"
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
         Left            =   270
         TabIndex        =   17
         Top             =   3690
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "Parity"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   0
         TabIndex        =   16
         Top             =   2715
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "Stop Bit"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   0
         TabIndex        =   15
         Top             =   2265
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "Start Bit"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   0
         TabIndex        =   14
         Top             =   1815
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "Data Bit"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   0
         TabIndex        =   13
         Top             =   1365
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "Speed"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   15
         TabIndex        =   12
         Top             =   915
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "Com Port"
         BeginProperty Font 
            Name            =   "����"
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
         Left            =   15
         TabIndex        =   11
         Top             =   465
         Width           =   1305
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "�˻���          ��������Ⱓ"
      BeginProperty Font 
         Name            =   "����"
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
      Left            =   5070
      TabIndex        =   31
      Top             =   4230
      Width           =   1380
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "�� ����"
      BeginProperty Font 
         Name            =   "����"
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
      Left            =   7800
      TabIndex        =   30
      Top             =   4350
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
    
    If MsgBox("������ �����Ͻðڽ��ϱ�?", vbCritical + vbOKCancel + vbDefaultButton2, "Ȯ��!") = vbCancel Then
        Unload Me
        Exit Sub
    Else
        If optUse(0).Value = True Then
            Call WritePrivateProfileString("COMM", "COMTYPE", "1", App.PATH & "\INI\" & gMACH & ".ini")
        ElseIf optUse(1).Value = True Then
            Call WritePrivateProfileString("COMM", "COMTYPE", "2", App.PATH & "\INI\" & gMACH & ".ini")
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
    
            Call WritePrivateProfileString("COMM", "TCPTYPE", gComm.TCPTYPE, App.PATH & "\INI\" & gMACH & ".ini")
            Call WritePrivateProfileString("COMM", "TCPIP", gComm.TCPIP, App.PATH & "\INI\" & gMACH & ".ini")
            Call WritePrivateProfileString("COMM", "TCPPORT", gComm.TCPPORT, App.PATH & "\INI\" & gMACH & ".ini")
        
        End If
        
        Call WritePrivateProfileString("HOSP", "SAVEDAY", txtSaveDay.Text, App.PATH & "\INI\" & gMACH & ".ini")
        
        Unload Me
    End If
        
    Exit Sub
 
ErrorHandler:
    Resume Next
    If MsgBox("��ż����� ���� �ʽ��ϴ�", vbCritical + vbOKCancel + vbDefaultButton2, "�����ư") = vbCancel Then
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
    
    txtSaveDay.Text = gHOSP.SAVEDAY
    
End Sub

