VERSION 5.00
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmAbout 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "MyApp 정보"
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
   ScaleMode       =   0  '사용자
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin BHButton.BHImageButton cmdSysInfo 
      Height          =   420
      Left            =   4140
      TabIndex        =   4
      Top             =   3015
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   741
      Caption         =   "System Info"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
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
      Left            =   4140
      TabIndex        =   5
      Top             =   2520
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   741
      Caption         =   "확인"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImgOutLineSize  =   3
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   4500
      Picture         =   "frmAbout.frx":0000
      Top             =   180
      Width           =   1185
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '내부 단색
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1584.05
      Y2              =   1584.05
   End
   Begin VB.Label lblDescription 
      Caption         =   "응용 프로그램 설명"
      ForeColor       =   &H00000000&
      Height          =   1065
      Left            =   150
      TabIndex        =   0
      Top             =   1125
      Width           =   4500
   End
   Begin VB.Label lblTitle 
      Caption         =   "응용 프로그램 제목"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   150
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
      Caption         =   "버전"
      Height          =   225
      Left            =   150
      TabIndex        =   3
      Top             =   780
      Width           =   4500
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "경고: ..."
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
    Caption = App.Title & " 정보"
    lblVersion.Caption = "버전 " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    lblDescription = App.Comments
    lblDisclaimer = "경고:이 컴퓨터 프로그램은 저작권 보호법과 국제 협약에 의해 보호됩니다.이 프로그램의 전부나 일부를 무단으로 복제하거나, 배포하는 경우에는 저작권의 침해가되므로 주의하시기 바랍니다."
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' 시스템 정보 프로그램의 경로와 이름을 레지스트리에서 가져 옵니다...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) <> "" Then
    '  시스템 정보 프로그램의 경로를 레지스트리에서만 가져 옵니다...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) <> "" Then
        ' 알려진 32비트 파일 버전의 존재 여부를 확인합니다.
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
        ' 오류 - 파일을 찾을 수 없습니다...
        Else
            GoTo SysInfoErr
        End If
    ' 오류 - 레지스트리 항목을 찾을 수 없습니다...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "지금은 시스템 정보를 사용할 수 없습니다.", vbOKOnly
End Sub
