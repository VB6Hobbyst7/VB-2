VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmDB_Local 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   " �� ���� ��"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8730
   Icon            =   "frmDB_Local.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
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
      Left            =   8010
      TabIndex        =   10
      Top             =   1410
      Width           =   375
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
      Left            =   5220
      TabIndex        =   9
      Top             =   3180
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
      Left            =   6870
      TabIndex        =   8
      Top             =   3180
      Width           =   1545
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '�� ����
      BackColor       =   &H00004080&
      BorderStyle     =   0  '����
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   8730
      TabIndex        =   6
      Top             =   0
      Width           =   8730
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackStyle       =   0  '����
         Caption         =   "���� �����ͺ��̽� ����"
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
         Index           =   0
         Left            =   210
         TabIndex        =   7
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
End
Attribute VB_Name = "frmDB_Local"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdFind_Click()
    
    With CommonDialog1
      .CancelError = True
      On Error GoTo ErrHandler
      .Flags = cdlOFNHideReadOnly
      .InitDir = App.PATH
      .Filter = "MS Access Files (*.MDB)|*.MDB|All Files (*.*)|*.*|"
      .FilterIndex = 1
      .Filename = "Interface.mdb"
      .ShowOpen
      txtFilename = .Filename
    End With

Exit Sub
  
ErrHandler:
  ' ����ڰ� [���] ���߸� �������ϴ�.
Exit Sub

End Sub

Private Sub cmdSave_Click()
    Dim strPath As String
    Dim strUID  As String
    Dim strPWD  As String
    Dim blnLUN As Boolean
    
    Dim intYear As Integer
    Dim intMon  As Integer
    Dim intDay  As Integer
    
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
        
        intYear = Year(Now)
        intMon = Month(Now)
        intDay = Day(Now)
        
        If Not GetSOL2LUN(intYear, intMon, intDay, strPWD) Then
            MsgBox "��й�ȣ�� Ʋ���ϴ�."
            Exit Sub
        End If
        
        'If strPWD <> Format(Now, "yyyymmdd") - 503 Then
        '    MsgBox "��й�ȣ�� Ʋ���ϴ�."
        '    Exit Sub
        'End If
        
        Call WritePrivateProfileString("DATABASE", "LOCALPATH", strPath, App.PATH & "\KDBAR.ini")
        Call WritePrivateProfileString("DATABASE", "LOCALUID", strUID, App.PATH & "\KDBAR.ini")
        Call WritePrivateProfileString("DATABASE", "LOCALPWD", strPWD, App.PATH & "\KDBAR.ini")
                
        'Call GetSetup
        '-- LOCAL DB GET
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("DATABASE", "LOCALPATH", "", strSetup, 100, App.PATH & "\KDBAR.ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gLocalDB.PATH = Trim(strSetUp1)
    
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("DATABASE", "LOCALUID", "", strSetup, 100, App.PATH & "\KDBAR.ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gLocalDB.UID = Trim(strSetUp1)
    
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("DATABASE", "LOCALPWD", "", strSetup, 100, App.PATH & "\KDBAR.ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gLocalDB.PWD = Trim(strSetUp1)

        If DbConnect_Local Then
            'labMsg.Caption = "����Ÿ ���̽��� ã���ֽ��ϴ�."
'            Call LetEqpMaster(gKUKDO.MACHCD)
            Unload Me
        Else
            MsgBox "  ������� �ʾҽ��ϴ�. �ٽ� �õ� �Ͻʽÿ�."
            txtFilename.Enabled = True
            txtFilename.SetFocus
        End If
    End If
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

