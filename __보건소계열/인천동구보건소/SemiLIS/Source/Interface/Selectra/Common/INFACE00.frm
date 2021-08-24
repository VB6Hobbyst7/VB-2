VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.MDIForm INTmain00 
   BackColor       =   &H8000000C&
   Caption         =   "########  인터페이스 초기 화면"
   ClientHeight    =   8535
   ClientLeft      =   2895
   ClientTop       =   1740
   ClientWidth     =   11400
   Icon            =   "INFACE00.frx":0000
   WindowState     =   2  '최대화
   Begin Threed.SSPanel pnlMain 
      Align           =   1  '위 맞춤
      Height          =   1005
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11400
      _Version        =   65536
      _ExtentX        =   20108
      _ExtentY        =   1773
      _StockProps     =   15
      ForeColor       =   -2147483630
      BackColor       =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BorderWidth     =   2
      BevelOuter      =   0
      Alignment       =   8
      Autosize        =   1
      Begin Threed.SSCommand SSCommand1 
         Height          =   375
         Left            =   9045
         TabIndex        =   8
         Top             =   225
         Visible         =   0   'False
         Width           =   195
         _Version        =   65536
         _ExtentX        =   344
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "SSCommand1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   45
         Left            =   0
         ScaleHeight     =   100
         ScaleMode       =   0  '사용자
         ScaleWidth      =   45
         TabIndex        =   7
         Top             =   0
         Width           =   45
      End
      Begin Threed.SSCommand cmdMenu 
         Height          =   870
         Index           =   0
         Left            =   240
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "장비로부터 검사결과를 전송받습니다."
         Top             =   50
         Width           =   1785
         _Version        =   65536
         _ExtentX        =   3149
         _ExtentY        =   1535
         _StockProps     =   78
         Caption         =   "검사 인터페이스"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   0   'False
         Picture         =   "INFACE00.frx":0442
      End
      Begin Threed.SSCommand cmdMenu 
         Height          =   870
         Index           =   1
         Left            =   2040
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "수신된 결과를 조회 및 SERVER에 등록합니다."
         Top             =   50
         Width           =   1785
         _Version        =   65536
         _ExtentX        =   3149
         _ExtentY        =   1535
         _StockProps     =   78
         Caption         =   "결과조회 및 수정"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   0   'False
         MouseIcon       =   "INFACE00.frx":0D1C
         Picture         =   "INFACE00.frx":0D38
      End
      Begin Threed.SSCommand cmdMenu 
         Height          =   870
         Index           =   2
         Left            =   3840
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "장비에서 검사할 검사항목을 입력합니다."
         Top             =   50
         Width           =   1785
         _Version        =   65536
         _ExtentX        =   3149
         _ExtentY        =   1535
         _StockProps     =   78
         Caption         =   "검사명 설정"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   0   'False
         Picture         =   "INFACE00.frx":1612
      End
      Begin Threed.SSCommand cmdMenu 
         Height          =   870
         Index           =   3
         Left            =   5640
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "장비와의 통신환경을 설정합니다."
         Top             =   50
         Width           =   1785
         _Version        =   65536
         _ExtentX        =   3149
         _ExtentY        =   1535
         _StockProps     =   78
         Caption         =   "통신환경설정"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   0   'False
         Picture         =   "INFACE00.frx":1EEC
      End
      Begin Threed.SSCommand cmdMenu 
         Height          =   870
         Index           =   4
         Left            =   7440
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "결과자료의 삭제기간을 설정합니다."
         Top             =   50
         Width           =   1785
         _Version        =   65536
         _ExtentX        =   3149
         _ExtentY        =   1535
         _StockProps     =   78
         Caption         =   "파일 정리"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   0   'False
         Picture         =   "INFACE00.frx":27C6
      End
      Begin Threed.SSCommand cmdMenu 
         Height          =   870
         Index           =   5
         Left            =   9240
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "인터페이스 프로그램을 종료합니다."
         Top             =   50
         Width           =   1785
         _Version        =   65536
         _ExtentX        =   3149
         _ExtentY        =   1535
         _StockProps     =   78
         Caption         =   "종   료"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   0   'False
         Picture         =   "INFACE00.frx":30A0
      End
   End
End
Attribute VB_Name = "INTmain00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub Unload_form()
    Select Case FrmFlag
        Case 10
            Unload INTcomm10
        Case 20
            Unload INTclear20
        Case 30
            Unload INTtname30
        Case 40
            Unload INTface40
        Case 41
            Unload INTface41
        Case 50
            Unload INTsearch50
    End Select
End Sub


Private Sub cmdMenu_Click(Index As Integer)
    
    On Error GoTo Load_Err
    
'    FrmTag = Index + 1
    Picture1.SetFocus
    ImgClickkey = True
    Screen.MousePointer = 11
    Select Case Index
        Case 0      '결과받기
'            If TypeOf Screen.ActiveForm Is MDIForm Then
            If FrmFlag = 0 Then
                Load INTface40
                INTface40.Show
            Else
                Unload_form
                'Unload Screen.ActiveForm
                Load INTface40
                INTface40.Show
            End If
        Case 1      '결과검색/등록
'            If TypeOf Screen.ActiveForm Is MDIForm Then
            If FrmFlag = 0 Then
                Load INTsearch50
                INTsearch50.Show
            Else
                Unload_form
'                Unload Screen.ActiveForm
                Load INTsearch50
                INTsearch50.Show
            End If
        Case 2      '검사명설정
'            If TypeOf Screen.ActiveForm Is MDIForm Then
            If FrmFlag = 0 Then
                Load INTtname30
                INTtname30.Show
            Else
                Unload_form
'                Unload Screen.ActiveForm
                Load INTtname30
                INTtname30.Show
            End If
        Case 3      '통신환경설정
'            If TypeOf Screen.ActiveForm Is MDIForm Then
            If FrmFlag = 0 Then
                Load INTcomm10
                INTcomm10.Show
            Else
                Unload_form
'                Unload Screen.ActiveForm
                Load INTcomm10
                INTcomm10.Show
            End If
        Case 4      '삭제일설정
'            If TypeOf Screen.ActiveForm Is MDIForm Then
            If FrmFlag = 0 Then
                Load INTclear20
                INTclear20.Show
            Else
                Unload_form
'                Unload Screen.ActiveForm
                Load INTclear20
                INTclear20.Show
            End If
        Case 5      '종  료
            End
    End Select
    Screen.MousePointer = 0
        
    Exit Sub
    
Load_Err:
End Sub

Private Sub cmdMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i   As Integer
    
    For i = 0 To 5
        With cmdMenu(i)
            If Index = i Then
                If .BevelWidth = 0 Then
                    .BevelWidth = 1
                    .Font3D = 2
                End If
            Else
                If .BevelWidth = 1 Then
                    .BevelWidth = 0
                    .Font3D = 0
                End If
            End If
        End With
    Next i
    
End Sub

Private Sub MDIForm_Activate()

    If App.PrevInstance = True Then
        MDIactivekey = False
        Unload Me
        AppActivate machstr
        SendKeys "%{ }"
    End If
    
End Sub

Private Sub MDIForm_Load()
    
    Call Create_Code_DB
    
    Call MachineConfig
'    Load frmLogOn
'    frmLogOn.Show 1
'    FrmFlag = 0
    'FrmTag = 0
    
'    If UCase(D0COM_USERID) = "SUPER" Then SSCommand1.Visible = True
    
    
End Sub


Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
'    Dim EXEStr      As String
'    Dim ret         As Variant
'
'    If CallLabKey = True Then
'        EXEStr = "C:\Lab\lab_pb.exe " & FileName & machstr & ".txt"
'        ret = Shell(EXEStr, 1)
'    Else
'        End
'    End If
    
End Sub

Private Sub pnlMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i    As Integer
    
    For i = 0 To 5
        With cmdMenu(i)
            If .BevelWidth = 1 Then
                .BevelWidth = 0
                .Font3D = 0
            End If
        End With
    Next i
            
End Sub

Private Sub SSCommand1_Click()

    frmDump.Show 1
    
End Sub


