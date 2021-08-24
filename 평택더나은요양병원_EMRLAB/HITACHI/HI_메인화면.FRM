VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.MDIForm Hi_Main 
   BackColor       =   &H8000000C&
   Caption         =   "HITACHI 7060  인터페이스 초기 화면"
   ClientHeight    =   8535
   ClientLeft      =   315
   ClientTop       =   345
   ClientWidth     =   11400
   Icon            =   "Hi_메인화면.frx":0000
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel pnlBack 
      Align           =   1  'Align Top
      Height          =   1005
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11400
      _Version        =   65536
      _ExtentX        =   20108
      _ExtentY        =   1773
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Left            =   11550
         ScaleHeight     =   45
         ScaleWidth      =   45
         TabIndex        =   0
         Top             =   90
         Width           =   45
      End
      Begin Threed.SSCommand cmdMenu 
         Height          =   870
         Index           =   0
         Left            =   360
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "장비로부터 검사결과를 전송받습니다."
         Top             =   70
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   0   'False
         Picture         =   "Hi_메인화면.frx":0442
      End
      Begin Threed.SSCommand cmdMenu 
         Height          =   870
         Index           =   1
         Left            =   2160
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "수신된 결과를 조회 및 SERVER에 등록합니다."
         Top             =   70
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   0   'False
         MouseIcon       =   "Hi_메인화면.frx":0D1C
         Picture         =   "Hi_메인화면.frx":0D38
      End
      Begin Threed.SSCommand cmdMenu 
         Height          =   870
         Index           =   2
         Left            =   3960
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "장비에서 검사할 검사항목을 입력합니다."
         Top             =   70
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   0   'False
         Picture         =   "Hi_메인화면.frx":1612
      End
      Begin Threed.SSCommand cmdMenu 
         Height          =   870
         Index           =   3
         Left            =   5760
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "장비와의 통신환경을 설정합니다."
         Top             =   70
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   0   'False
         Picture         =   "Hi_메인화면.frx":1EEC
      End
      Begin Threed.SSCommand cmdMenu 
         Height          =   870
         Index           =   4
         Left            =   7560
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "결과자료의 삭제기간을 설정합니다."
         Top             =   70
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   0   'False
         Picture         =   "Hi_메인화면.frx":27C6
      End
      Begin Threed.SSCommand cmdMenu 
         Height          =   870
         Index           =   5
         Left            =   9450
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "인터페이스 프로그램을 종료합니다."
         Top             =   70
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         RoundedCorners  =   0   'False
         Outline         =   0   'False
         Picture         =   "Hi_메인화면.frx":30A0
      End
   End
End
Attribute VB_Name = "Hi_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    
    
Private Sub cmdMenu_Click(Index As Integer)
    
    On Error GoTo Load_Err
    
    If FrmTag <> 0 Then
        Select Case FrmTag
            Case 1
                Unload frm화일명지정
            Case 2
                Unload frm결과등록
            Case 3
                Unload frm검사명설정
            Case 4
                Unload frm통신설정
            Case 5
                Unload frm삭제일설정
        End Select
    End If
    
    FrmTag = Index + 1
    Picture1.SetFocus
    
    Select Case Index
        Case 0      '결과받기
            Load frm화일명지정
            frm화일명지정.Show
        Case 1      '결과검색/등록
            Load frm결과등록
            frm결과등록.Show
        Case 2      '검사명설정
            Load frm검사명설정
            frm검사명설정.Show
        Case 3      '통신환경설정
            Load frm통신설정
            frm통신설정.Show
        Case 4      '삭제일설정
            Load frm삭제일설정
            frm삭제일설정.Show
        Case 5      '종  료
            Unload Me
    End Select
        
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
'        MDIactivekey = False
        Unload Me
'        AppActivate machstr
        SendKeys "%{ }"
    End If
    
End Sub

Sub Format_Menu()

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


Private Sub MDIForm_Load()
            
    strLAB_PATH = CurDir("C") & "\"
    ResDb_Path = strLAB_PATH & "Data\"
        
'    Set_EqName = "Hitachi"  '나중에 장비선택 버튼에서 넘어온 값 이용
    sEq_Name = "Hi"

'    '----- 사용자 ID 확인
    Load frmLogOn
    frmLogOn.Show 1
    If Chk_LogID <> True Then
        Unload Me
        Exit Sub
'    Else
'        BarMain.Panels(3) = UserID
'        BarMain.Panels(4) = UserNM
    End If
    
    '----- 통신환경과 검사항목 조회/저장
    If Set_Init() <> True Then
        Unload Me
        Exit Sub
    End If
    
    '----- 기간이 지난 MDB삭제
    Call DelCheck(ResDb_Path)
    
    FrmTag = 0
    
End Sub

'
'   Main Form이 Load시 통신환경과 검사항목을 조회하여 메모리에 저장한다.
'
Function Set_Init() As Integer

    Dim FileName    As String
    
    Set_Init = False
    On Error GoTo Error_Cfg
        
    FileName = strLAB_PATH & "System\Config.mdb"
       
    If iFileExists(FileName) = False Then
        MsgBox "기초자료가 입력되어 있지 않습니다." & Chr(13) & Chr(10) & _
                "확인하여 주십시오", vbExclamation
        Exit Function
    End If
    
    Set Db = OpenDatabase(FileName)
    
    '----- 검사항목
    Set TableCd = Db.OpenRecordset("Test_Item")
    
    TableCd.MoveFirst
    Do Until TableCd.EOF
'        If TableCd!SeqNo <> "" Then
'           TestNameTable(TableCd!SeqNo).SeqNo = TableCd!SeqNo
'        End If
        If TableCd!Eq_Code <> "" Then
           TestNameTable(TableCd!SeqNo).EqCd = TableCd!Eq_Code
        End If
        If TableCd!OrdCd <> "" Then
           TestNameTable(TableCd!SeqNo).Code = TableCd!OrdCd
        End If
        If TableCd!ordnm <> "" Then
           TestNameTable(TableCd!SeqNo).Name = TableCd!ordnm
        End If
                
        TableCd.MoveNext
    Loop
    TableCd.Close
    '==============
    
    '----- 통신환경
    Set TbComm = Db.OpenRecordset("Config")
    
    If TbComm.RecordCount = 0 Then
        CommCfg.ComPort = 1
        CommCfg.BaudRate = "9600"
        CommCfg.DataBit = "7"
        CommCfg.StopBit = "1"
        CommCfg.Parity = "N"
        
        dDate = "30"    '저장일자 Default Setting
        
        TbComm.Close
        Db.Close
        
        Set_Init = True
        Exit Function
    End If
        
    ' 포트
    If TbComm!ComPort = 1 Then
       CommCfg.ComPort = 1
    Else
       CommCfg.ComPort = 2
    End If
    ' 전송속도
    Select Case TbComm!BaudRate
        Case 1200
            CommCfg.BaudRate = "1200"
        Case 2400
            CommCfg.BaudRate = "2400"
        Case 4800
            CommCfg.BaudRate = "4800"
        Case 9600
            CommCfg.BaudRate = "9600"
    End Select
    ' 데이타 비트
    If TbComm!DataBit = 7 Then
        CommCfg.DataBit = "7"
    Else
        CommCfg.DataBit = "8"
    End If
    ' 스톱비트
    If TbComm!StopBit = 1 Then
        CommCfg.StopBit = "1"
    Else
        CommCfg.StopBit = "2"
    End If
    ' 패리티 비트
    Select Case TbComm!Parity
        Case "N"
            CommCfg.Parity = "N"
        Case "O"
            CommCfg.Parity = "O"
        Case "E"
            CommCfg.Parity = "E"
    End Select
    '----- Data 삭제일자 메모리에 저장
    dDate = TbComm!SaveDate
    
    TbComm.Close
    '=======================
    Db.Close
    Set_Init = True

    Exit Function
    
Error_Cfg:
    Resume Next
    Exit Function

End Function

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'    Dim EXEStr      As String
'    Dim ret         As Variant
'
'    If CallLabKey = True Then
'        EXEStr = "C:\Lab\lab_pb.exe " & Filename & machstr & ".txt"
'        ret = Shell(EXEStr, 1)
'    Else
'        End
'    End If
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    If Forms.count = 2 Then
        Unload frmHi_결과받기
    End If
    
End Sub

Private Sub pnlBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call Format_Menu

End Sub




