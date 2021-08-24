VERSION 5.00
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "MyApp ����"
   ClientHeight    =   3555
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2453.724
   ScaleMode       =   0  '�����
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   Begin HSCotrol.CButton cmdOK 
      Height          =   360
      Left            =   4260
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2580
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   635
      Caption         =   "OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      BorderStyle     =   1
      BorderColor     =   -2147483632
   End
   Begin HSCotrol.CButton cmdSysInfo 
      Height          =   360
      Left            =   4260
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3075
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   635
      Caption         =   "System Inf"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      BorderStyle     =   1
      BorderColor     =   -2147483632
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '���� �ܻ�
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1584.05
      Y2              =   1584.05
   End
   Begin VB.Label lblDescription 
      Caption         =   "���� ���α׷� ����"
      ForeColor       =   &H00000000&
      Height          =   1065
      Left            =   1050
      TabIndex        =   0
      Top             =   1125
      Width           =   4500
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
      Left            =   1050
      TabIndex        =   2
      Top             =   240
      Width           =   4500
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1594.403
      Y2              =   1594.403
   End
   Begin VB.Label lblVersion 
      Caption         =   "����"
      Height          =   225
      Left            =   1050
      TabIndex        =   3
      Top             =   780
      Width           =   4500
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "���: ..."
      ForeColor       =   &H00000000&
      Height          =   1005
      Left            =   150
      TabIndex        =   1
      Top             =   2445
      Width           =   3705
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
