VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmDB_Local 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   " ���� �����ͺ��̽� ����"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8955
   Icon            =   "frmDB_Local.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   Begin VB.PictureBox Picture2 
      Align           =   2  '�Ʒ� ����
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '����
      Height          =   720
      Left            =   0
      ScaleHeight     =   720
      ScaleWidth      =   8955
      TabIndex        =   9
      Top             =   3015
      Width           =   8955
      Begin VB.Image imgMenuCancel 
         Height          =   375
         Left            =   6630
         Picture         =   "frmDB_Local.frx":000C
         Top             =   180
         Width           =   1725
      End
      Begin VB.Image imgMenuInsert 
         Height          =   375
         Left            =   4800
         Picture         =   "frmDB_Local.frx":0D64
         Top             =   180
         Width           =   1725
      End
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7980
      TabIndex        =   8
      Top             =   1380
      Width           =   375
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
      ScaleWidth      =   8955
      TabIndex        =   6
      Top             =   0
      Width           =   8955
      Begin VB.Label Label2 
         BackStyle       =   0  '����
         Caption         =   "���� DB ����"
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
         TabIndex        =   7
         Top             =   540
         Width           =   3135
      End
      Begin VB.Image Image3 
         Height          =   1065
         Left            =   0
         Picture         =   "frmDB_Local.frx":1B60
         Top             =   0
         Width           =   12900
      End
   End
   Begin VB.TextBox txtUser 
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
      TabIndex        =   1
      Top             =   1830
      Width           =   2115
   End
   Begin VB.TextBox txtPasswd 
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
      IMEMode         =   3  '��� ����
      Left            =   2910
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2295
      Width           =   2115
   End
   Begin VB.TextBox txtFilename 
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
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1380
      Width           =   5055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7890
      Top             =   1770
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Index           =   10
      Left            =   2175
      TabIndex        =   5
      Top             =   2355
      Width           =   615
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   2
      Left            =   690
      Picture         =   "frmDB_Local.frx":32A3
      Top             =   2325
      Width           =   150
   End
   Begin VB.Label ����ڸ� 
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
      Index           =   9
      Left            =   1800
      TabIndex        =   4
      Top             =   1920
      Width           =   1005
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   1
      Left            =   690
      Picture         =   "frmDB_Local.frx":368D
      Top             =   1890
      Width           =   150
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "�����ͺ��̽� ��� : "
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
      Left            =   945
      TabIndex        =   3
      Top             =   1470
      Width           =   1860
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   0
      Left            =   690
      Picture         =   "frmDB_Local.frx":3A77
      Top             =   1440
      Width           =   150
   End
End
Attribute VB_Name = "frmDB_Local"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFind_Click()
    
    With CommonDialog1
      .CancelError = True
      On Error GoTo ErrHandler
      .Flags = cdlOFNHideReadOnly
      .InitDir = App.PATH
      .Filter = "MS Access Files (*.MDB)|*.MDB|All Files (*.*)|*.*|"
      .FilterIndex = 1
      .FileName = "Interface.mdb"
      .ShowOpen
      txtFilename = .FileName
    End With

Exit Sub
  
ErrHandler:
  ' ����ڰ� [���] ���߸� �������ϴ�.
Exit Sub

End Sub

Private Sub Form_Load()

    txtFilename.Text = gLocalDB.PATH
    txtUser.Text = gLocalDB.UID
    txtPasswd.Text = gLocalDB.PWD
    
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
    Dim strPath As String
    Dim strUID  As String
    Dim strPWD  As String
    
    If Trim(txtFilename) = "" Then
        MsgBox " ����Ÿ ���̽��� ���� �ϼ���"
        Exit Sub
    ElseIf Trim(txtUser) = "" Then
        MsgBox " ����ڸ��� �Է� �ϼ���"
        Exit Sub
    ElseIf Trim(txtPasswd) = "" Then
        MsgBox " ��й�ȣ�� �Է� �ϼ���"
        Exit Sub
    Else
        strPath = txtFilename.Text
        strUID = txtUser.Text
        strPWD = txtPasswd.Text
        
        Call WritePrivateProfileString("DATABASE", "LOCALPATH", strPath, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
        Call WritePrivateProfileString("DATABASE", "LOCALUID", strUID, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
        Call WritePrivateProfileString("DATABASE", "LOCALPWD", strPWD, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
        
        'Call GetSetup
        '-- LOCAL DB GET
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("DATABASE", "LOCALPATH", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gLocalDB.PATH = Trim(strSetUp1)
    
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("DATABASE", "LOCALUID", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gLocalDB.UID = Trim(strSetUp1)
    
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("DATABASE", "LOCALPWD", "", strSetup, 100, App.PATH & "\INI\" & gHOSP.APPNM & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gLocalDB.PWD = Trim(strSetUp1)

        If DbConnect_Local Then
            'labMsg.Caption = "����Ÿ ���̽��� ã���ֽ��ϴ�."
            Unload Me
        Else
            MsgBox "  ������� �ʾҽ��ϴ�. �ٽ� �õ� �Ͻʽÿ�."
            txtFilename.Enabled = True
            txtFilename.SetFocus
        End If
    End If
    
End Sub
