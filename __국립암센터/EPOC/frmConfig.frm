VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmConfig 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  '���� ����
   Caption         =   "��ż���"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "����ü"
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
   ScaleHeight     =   5310
   ScaleWidth      =   4335
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.PictureBox cmdExit 
      Height          =   465
      Left            =   2250
      ScaleHeight     =   405
      ScaleWidth      =   1215
      TabIndex        =   15
      Top             =   4650
      Width           =   1275
   End
   Begin VB.PictureBox cmdConfirm 
      Height          =   465
      Left            =   810
      ScaleHeight     =   405
      ScaleWidth      =   1215
      TabIndex        =   14
      Top             =   4650
      Width           =   1275
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   735
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   6588
      _Version        =   131072
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
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
         Width           =   1695
      End
      Begin VB.ComboBox Combo_Stopbit 
         Height          =   315
         Left            =   1785
         TabIndex        =   6
         Top             =   2100
         Width           =   1695
      End
      Begin VB.ComboBox Combo_Startbit 
         Height          =   315
         Left            =   1785
         TabIndex        =   5
         Top             =   1650
         Width           =   1695
      End
      Begin VB.ComboBox Combo_Databit 
         Height          =   315
         ItemData        =   "frmConfig.frx":0004
         Left            =   1785
         List            =   "frmConfig.frx":0006
         TabIndex        =   4
         Top             =   1200
         Width           =   1695
      End
      Begin VB.ComboBox Combo_BPS 
         Height          =   315
         ItemData        =   "frmConfig.frx":0008
         Left            =   1785
         List            =   "frmConfig.frx":000A
         TabIndex        =   3
         Top             =   750
         Width           =   1695
      End
      Begin VB.ComboBox Combo_Port 
         Height          =   315
         ItemData        =   "frmConfig.frx":000C
         Left            =   1785
         List            =   "frmConfig.frx":000E
         TabIndex        =   2
         Top             =   300
         Width           =   1695
      End
      Begin VB.CheckBox chkACK 
         Alignment       =   1  '������ ����
         BackColor       =   &H00FFFFFF&
         Caption         =   "ACK ���"
         Height          =   315
         Left            =   690
         TabIndex        =   1
         Top             =   3060
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "�и�Ƽ"
         Height          =   195
         Index           =   5
         Left            =   405
         TabIndex        =   13
         Top             =   2640
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "���� ��Ʈ"
         Height          =   195
         Index           =   4
         Left            =   405
         TabIndex        =   12
         Top             =   2160
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "���� ��Ʈ"
         Height          =   195
         Index           =   3
         Left            =   405
         TabIndex        =   11
         Top             =   1710
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "������ ��Ʈ"
         Height          =   195
         Index           =   2
         Left            =   405
         TabIndex        =   10
         Top             =   1260
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "���ۼӵ�"
         Height          =   195
         Index           =   1
         Left            =   420
         TabIndex        =   9
         Top             =   810
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "COM PORT"
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   8
         Top             =   360
         Width           =   1155
      End
   End
   Begin VB.Label Label2 
      Appearance      =   0  '���
      BackColor       =   &H00E0E0E0&
      Caption         =   "ABL80 ��ż���"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   18
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   240
      TabIndex        =   16
      Top             =   180
      Width           =   3765
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

    
End Sub

Private Sub Command2_Click()
    
End Sub

Private Sub cmdConfirm_Click()
        Dim Parity As String
    Dim sEquipNo As String
    
    On Error GoTo ErrorHandler
    
    If MsgBox("������ �����Ͻðڽ��ϱ�?", vbCritical + vbOKCancel + vbDefaultButton2, "Ȯ��!") = vbCancel Then
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
'    If MsgBox("��ż����� ���� �ʽ��ϴ�", vbCritical + vbOKCancel + vbDefaultButton2, "�����ư") = vbCancel Then
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
    
End Sub
