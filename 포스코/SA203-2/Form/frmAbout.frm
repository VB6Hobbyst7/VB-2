VERSION 5.00
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "MyApp ����"
   ClientHeight    =   3555
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5910
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2453.724
   ScaleMode       =   0  '�����
   ScaleWidth      =   5549.795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   Begin BHButton.BHImageButton cmdSysInfo 
      Height          =   420
      Left            =   4455
      TabIndex        =   0
      Top             =   2835
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   741
      Caption         =   "System Info"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton cmdOK 
      Height          =   420
      Left            =   4455
      TabIndex        =   1
      Top             =   2340
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   741
      Caption         =   "Ȯ��"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImgOutLineSize  =   3
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "���: ..."
      ForeColor       =   &H00000000&
      Height          =   1005
      Left            =   150
      TabIndex        =   5
      Top             =   2265
      Width           =   3705
   End
   Begin VB.Label lblVersion 
      Caption         =   "����"
      Height          =   225
      Left            =   1260
      TabIndex        =   4
      Top             =   600
      Width           =   4500
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1470.164
      Y2              =   1470.164
   End
   Begin VB.Label lblTitle 
      Caption         =   "���� ���α׷� ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1260
      TabIndex        =   3
      Top             =   60
      Width           =   4500
   End
   Begin VB.Label lblDescription 
      Caption         =   "���� ���α׷� ����"
      ForeColor       =   &H00000000&
      Height          =   1065
      Left            =   1260
      TabIndex        =   2
      Top             =   945
      Width           =   4500
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '���� �ܻ�
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1459.811
      Y2              =   1459.811
   End
   Begin VB.Image Image2 
      Height          =   2070
      Left            =   0
      Picture         =   "frmAbout.frx":27A2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1050
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Private Const gREGVALSYSINFOLOC = "MSINFO"
Private Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Private Const gREGVALSYSINFO = "PATH"

Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOk_Click()
  Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
        Case vbKeyReturn
            Unload Me
        Case Else
        
    End Select
    
End Sub

Private Sub Form_Load()
    Caption = App.Title & " ����"
    lblVersion.Caption = "���� " & App.major & "." & App.minor & "." & App.Revision
    lblTitle.Caption = App.Title
    lblDescription = App.Comments
    lblDisclaimer = "���:�� ��ǻ�� ���α׷��� ���۱� ��ȣ���� ���� ���࿡ ���� ��ȣ�˴ϴ�.�� ���α׷��� ���γ� �Ϻθ� �������� �����ϰų�, �����ϴ� ��쿡�� ���۱��� ħ�ذ��ǹǷ� �����Ͻñ� �ٶ��ϴ�."
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' �ý��� ���� ���α׷��� ��ο� �̸��� ������Ʈ������ ���� �ɴϴ�...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) <> "" Then
    '  �ý��� ���� ���α׷��� ��θ� ������Ʈ�������� ���� �ɴϴ�...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) <> "" Then
        ' �˷��� 32��Ʈ ���� ������ ���� ���θ� Ȯ���մϴ�.
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
        ' ���� - ������ ã�� �� �����ϴ�...
        Else
            GoTo SysInfoErr
        End If
    ' ���� - ������Ʈ�� �׸��� ã�� �� �����ϴ�...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "������ �ý��� ������ ����� �� �����ϴ�.", vbOKOnly
End Sub

