VERSION 5.00
Begin VB.Form frmDB_PGSQL 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   " Postgres SQL �����ͺ��̽� ����"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6090
   Icon            =   "frmDB_PGSQL.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   Begin VB.CommandButton cmdChange 
      Caption         =   "���� ����"
      Height          =   300
      Left            =   2940
      TabIndex        =   11
      Top             =   3690
      Width           =   2055
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
      TabIndex        =   9
      Top             =   3030
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
      TabIndex        =   5
      Top             =   1680
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
      TabIndex        =   4
      Top             =   2595
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
      TabIndex        =   3
      Top             =   2130
      Width           =   2115
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '�� ����
      Appearance      =   0  '���
      BackColor       =   &H80000005&
      BorderStyle     =   0  '����
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   6090
      TabIndex        =   1
      Top             =   0
      Width           =   6090
      Begin VB.Label Label2 
         BackStyle       =   0  '����
         Caption         =   "Postgres SQL DB ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Top             =   540
         Width           =   3135
      End
      Begin VB.Image Image3 
         Height          =   1065
         Left            =   0
         Picture         =   "frmDB_PGSQL.frx":000C
         Top             =   0
         Width           =   12900
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  '�Ʒ� ����
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '����
      Height          =   930
      Left            =   0
      ScaleHeight     =   930
      ScaleWidth      =   6090
      TabIndex        =   0
      Top             =   4035
      Width           =   6090
      Begin VB.Image imgMenuCancel 
         Height          =   375
         Left            =   3300
         Picture         =   "frmDB_PGSQL.frx":174F
         Top             =   180
         Width           =   1725
      End
      Begin VB.Image imgMenuInsert 
         Height          =   375
         Left            =   1470
         Picture         =   "frmDB_PGSQL.frx":24A7
         Top             =   180
         Width           =   1725
      End
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
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   945
      TabIndex        =   12
      Top             =   3750
      Width           =   1860
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   4
      Left            =   690
      Picture         =   "frmDB_PGSQL.frx":32A3
      Top             =   3720
      Width           =   150
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   3
      Left            =   690
      Picture         =   "frmDB_PGSQL.frx":368D
      Top             =   3090
      Width           =   150
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
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   2190
      TabIndex        =   10
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   0
      Left            =   690
      Picture         =   "frmDB_PGSQL.frx":3A77
      Top             =   1740
      Width           =   150
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
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   8
      Left            =   2190
      TabIndex        =   8
      Top             =   1770
      Width           =   615
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   1
      Left            =   690
      Picture         =   "frmDB_PGSQL.frx":3E61
      Top             =   2190
      Width           =   150
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
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   9
      Left            =   1215
      TabIndex        =   7
      Top             =   2220
      Width           =   1590
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   2
      Left            =   690
      Picture         =   "frmDB_PGSQL.frx":424B
      Top             =   2625
      Width           =   150
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
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   10
      Left            =   1785
      TabIndex        =   6
      Top             =   2655
      Width           =   1005
   End
End
Attribute VB_Name = "frmDB_PGSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdChange_Click()
    Unload Me
    frmEMRInfo.Show vbModal
End Sub

Private Sub Form_Load()

    txtIP.Text = gPGSQLDB.IP
    txtDB.Text = gPGSQLDB.DB
    txtUID.Text = gPGSQLDB.UID
    txtPWD.Text = gPGSQLDB.PWD
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        End
    End If
End Sub

Private Sub imgMenuCancel_Click()
    End
End Sub

Private Sub imgMenuInsert_Click()
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
        
        Call WritePrivateProfileString("DATABASE", "PGSQLIP", txtIP.Text, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("DATABASE", "PGSQLDB", txtDB.Text, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("DATABASE", "PGSQLUID", txtUID.Text, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("DATABASE", "PGSQLPWD", txtPWD.Text, App.PATH & "\INI\" & gMACH & ".ini")
        
        '-- Postgres SQL DB SET
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("DATABASE", "PGSQLIP", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gPGSQLDB.IP = Trim(strSetUp1)
    
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("DATABASE", "PGSQLDB", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gPGSQLDB.DB = Trim(strSetUp1)
    
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("DATABASE", "PGSQLUID", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gPGSQLDB.UID = Trim(strSetUp1)
        
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("DATABASE", "PGSQLPWD", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gPGSQLDB.PWD = Trim(strSetUp1)

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

