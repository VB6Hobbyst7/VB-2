VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmDB_Local 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   " �� ���� �����ͺ��̽� ���� ��"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8955
   Icon            =   "frmDB_Local.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
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
      Left            =   7860
      TabIndex        =   6
      Top             =   360
      Width           =   375
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
      Left            =   2790
      TabIndex        =   1
      Top             =   810
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
      Left            =   2790
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1275
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
      Left            =   2790
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   5055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7770
      Top             =   750
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgMenuInsert 
      Height          =   375
      Left            =   4230
      Picture         =   "frmDB_Local.frx":000C
      Top             =   1920
      Width           =   1725
   End
   Begin VB.Image imgMenuCancel 
      Height          =   375
      Left            =   6060
      Picture         =   "frmDB_Local.frx":0E08
      Top             =   1920
      Width           =   1725
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
      Left            =   2055
      TabIndex        =   5
      Top             =   1335
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
      Left            =   1680
      TabIndex        =   4
      Top             =   900
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
      Left            =   825
      TabIndex        =   3
      Top             =   450
      Width           =   1860
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
      .Filename = "Interface.mdb"
      .ShowOpen
      txtFilename = .Filename
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
        
        Call WritePrivateProfileString("DATABASE", "LOCALPATH", strPath, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("DATABASE", "LOCALUID", strUID, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("DATABASE", "LOCALPWD", strPWD, App.PATH & "\INI\" & gMACH & ".ini")
                
        'Call GetSetup
        '-- LOCAL DB GET
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("DATABASE", "LOCALPATH", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gLocalDB.PATH = Trim(strSetUp1)
    
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("DATABASE", "LOCALUID", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gLocalDB.UID = Trim(strSetUp1)
    
        strSetup = "":    strSetUp1 = ""
        Call GetPrivateProfileString("DATABASE", "LOCALPWD", "", strSetup, 100, App.PATH & "\INI\" & gMACH & ".ini")
        strSetUp1 = Trim(strSetup)
        strSetUp1 = Mid(strSetUp1, 1, Len(strSetUp1) - 1)
        gLocalDB.PWD = Trim(strSetUp1)

        If DbConnect_Local Then
            'labMsg.Caption = "����Ÿ ���̽��� ã���ֽ��ϴ�."
            Call LetEqpMaster(gHOSP.MACHCD)
            Unload Me
        Else
            MsgBox "  ������� �ʾҽ��ϴ�. �ٽ� �õ� �Ͻʽÿ�."
            txtFilename.Enabled = True
            txtFilename.SetFocus
        End If
    End If
    
End Sub
