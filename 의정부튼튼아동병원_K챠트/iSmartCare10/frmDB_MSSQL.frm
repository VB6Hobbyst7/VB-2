VERSION 5.00
Begin VB.Form frmDB_MSSQL 
   BackColor       =   &H00BF8B59&
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "�����ͺ��̽� ����"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   Icon            =   "frmDB_MSSQL.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   Begin VB.CommandButton cmdChange 
      Caption         =   "���� ����"
      Height          =   300
      Left            =   2940
      TabIndex        =   13
      Top             =   3390
      Width           =   2055
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1800
      TabIndex        =   12
      Top             =   4170
      Width           =   1545
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3450
      TabIndex        =   11
      Top             =   4170
      Width           =   1545
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '�� ����
      BackColor       =   &H00AE8B59&
      BorderStyle     =   0  '����
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   5655
      TabIndex        =   9
      Top             =   0
      Width           =   5655
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   90
         Top             =   90
         Width           =   2865
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackStyle       =   0  '����
         Caption         =   "SQL �����ͺ��̽� ����"
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
         Index           =   2
         Left            =   210
         TabIndex        =   10
         Top             =   180
         Width           =   2625
      End
   End
   Begin VB.TextBox txtPWD 
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  '��� ����
      Left            =   2910
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2730
      Width           =   2115
   End
   Begin VB.TextBox txtIP 
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2910
      TabIndex        =   2
      Top             =   1380
      Width           =   2115
   End
   Begin VB.TextBox txtUID 
      Appearance      =   0  '���
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2910
      TabIndex        =   1
      Top             =   2295
      Width           =   2115
   End
   Begin VB.TextBox txtDB 
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2910
      TabIndex        =   0
      Top             =   1830
      Width           =   2115
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "�����ͺ��̽� ���� : "
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   1
      Left            =   945
      TabIndex        =   8
      Top             =   3450
      Width           =   1860
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "��ȣ : "
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   0
      Left            =   2190
      TabIndex        =   7
      Top             =   2820
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "���� : "
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   8
      Left            =   2190
      TabIndex        =   5
      Top             =   1470
      Width           =   615
   End
   Begin VB.Label ����ڸ� 
      Alignment       =   1  '������ ����
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "�����ͺ��̽��� : "
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   9
      Left            =   1215
      TabIndex        =   4
      Top             =   1920
      Width           =   1590
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "����ڸ� : "
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   10
      Left            =   1785
      TabIndex        =   3
      Top             =   2355
      Width           =   1005
   End
End
Attribute VB_Name = "frmDB_MSSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChange_Click()
    Unload Me
    frmEMRInfo.Show vbModal
End Sub

Private Sub cmdSave_Click()
    Dim strIP   As String
    Dim strDB   As String
    Dim strUID  As String
    Dim strPWD  As String
    
    If Trim(txtIP) = "" Then
        MsgBox " SID�� �Է� �ϼ���"
        Exit Sub
    ElseIf Trim(txtDB) = "" Then
        MsgBox " �����ͺ��̽����� �Է� �ϼ���"
        Exit Sub
    ElseIf Trim(txtUID) = "" Then
        MsgBox " ����ڸ��� �Է� �ϼ���"
        Exit Sub
    ElseIf Trim(txtPWD) = "" Then
        MsgBox " ��й�ȣ�� �Է� �ϼ���"
        Exit Sub
    Else
        strIP = txtIP.Text
        strDB = txtDB.Text
        strUID = txtUID.Text
        strPWD = txtPWD.Text
        
        Call WritePrivateProfileString("DATABASE", "MSSQLIP", txtIP.Text, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("DATABASE", "MSSQLDB", txtDB.Text, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("DATABASE", "MSSQLUID", txtUID.Text, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("DATABASE", "MSSQLPWD", txtPWD.Text, App.PATH & "\INI\" & gMACH & ".ini")
        
        'Call GetSetup
        '-- MSSQL DB SET
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("DATABASE", "MSSQLIP", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gSQLDB.IP = Trim(strSetUp1)
    
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("DATABASE", "MSSQLDB", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gSQLDB.DB = Trim(strSetUp1)
    
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("DATABASE", "MSSQLUID", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gSQLDB.UID = Trim(strSetUp1)
        
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("DATABASE", "MSSQLPWD", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gSQLDB.PWD = Trim(strSetUp1)

        If DbConnect_SQL Then
            'labMsg.Caption = "����Ÿ ���̽��� ã���ֽ��ϴ�."
            Unload Me
        Else
            MsgBox "  ������� �ʾҽ��ϴ�. �ٽ� �õ� �Ͻʽÿ�."
            txtIP.Enabled = True
            txtIP.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()

    txtIP.Text = gSQLDB.IP
    txtDB.Text = gSQLDB.DB
    txtUID.Text = gSQLDB.UID
    txtPWD.Text = gSQLDB.PWD
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        End
    End If
End Sub


