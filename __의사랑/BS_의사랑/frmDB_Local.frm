VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Begin VB.Form frmDB_Local 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "�����ͺ��̽� ����"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8055
   Icon            =   "frmDB_Local.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   Begin VB.PictureBox Picture1 
      Align           =   1  '�� ����
      BackColor       =   &H00F8E4D8&
      BorderStyle     =   0  '����
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   8055
      TabIndex        =   6
      Top             =   0
      Width           =   8055
      Begin VB.Label Label1 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   0
         Left            =   210
         TabIndex        =   7
         Top             =   180
         Width           =   2625
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
      Left            =   2370
      TabIndex        =   1
      Top             =   1440
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
      Left            =   2370
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1905
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
      Left            =   2370
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   990
      Width           =   4605
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7350
      Top             =   1380
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin HSCotrol.CButton cmdSave 
      Height          =   495
      Left            =   4680
      TabIndex        =   8
      Top             =   2730
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   873
      BackColor       =   12632256
      Caption         =   " ��������"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmDB_Local.frx":0CB2
      MaskColor       =   0
      PicCapAlign     =   2
      BorderStyle     =   1
      BorderColor     =   16777215
      HoverColor      =   -2147483630
   End
   Begin HSCotrol.CButton cmdCancel 
      Height          =   495
      Left            =   6150
      TabIndex        =   9
      Top             =   2730
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   873
      BackColor       =   12632256
      Caption         =   " ��    ��"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmDB_Local.frx":0E0C
      MaskColor       =   0
      PicCapAlign     =   2
      BorderStyle     =   1
      BorderColor     =   16777215
      HoverColor      =   -2147483630
   End
   Begin HSCotrol.CButton cmdFind 
      Height          =   345
      Left            =   6990
      TabIndex        =   10
      Top             =   990
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   609
      BackColor       =   12632256
      Caption         =   "ã��"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      BorderStyle     =   1
      BorderColor     =   16777215
      HoverColor      =   -2147483630
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
      ForeColor       =   &H00404040&
      Height          =   180
      Index           =   10
      Left            =   1635
      TabIndex        =   5
      Top             =   1965
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
      ForeColor       =   &H00404040&
      Height          =   180
      Index           =   9
      Left            =   1260
      TabIndex        =   4
      Top             =   1530
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
      ForeColor       =   &H00404040&
      Height          =   180
      Index           =   8
      Left            =   405
      TabIndex        =   3
      Top             =   1080
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
    Unload Me
End Sub

Private Sub cmdFind_Click()
    
On Error GoTo ErrHandler
    
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
        
        'Call WritePrivateProfileString("DATABASE", "LOCALPATH", strPath, App.PATH & "\INI\" & gMACH & ".ini")
        'Call WritePrivateProfileString("DATABASE", "LOCALUID", strUID, App.PATH & "\INI\" & gMACH & ".ini")
        'Call WritePrivateProfileString("DATABASE", "LOCALPWD", strPWD, App.PATH & "\INI\" & gMACH & ".ini")
                
        Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "LOCALPATH", strPath)
        Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "LOCALUID", strUID)
        Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "LOCALPWD", strPWD)
                
        '-- LOCAL DB GET
        gLocalDB.PATH = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "MDBPATH")
        gLocalDB.UID = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "MDBUID")
        gLocalDB.PWD = GetString(HKEY_CURRENT_USER, REG_MACH & "\" & "DATABASE", "MDBPWD")
        
        If DbConnect_Local Then
            Call LetEqpMaster(gHOSP.MACHCD)
            Unload Me
        Else
            MsgBox "������� �ʾҽ��ϴ�. �ٽ� �õ� �Ͻʽÿ�."
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

