VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEQ공용_Login 
   BorderStyle     =   1  '단일 고정
   Caption         =   "로그인"
   ClientHeight    =   3735
   ClientLeft      =   6135
   ClientTop       =   2595
   ClientWidth     =   6675
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEQ공용_Login.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6675
   Begin VB.TextBox txtUserPW 
      Height          =   315
      IMEMode         =   3  '사용 못함
      Left            =   5160
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1980
      Width           =   1335
   End
   Begin VB.TextBox txtUserID 
      Height          =   315
      IMEMode         =   8  '영문
      Left            =   5160
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1620
      Width           =   1335
   End
   Begin VB.Timer tmr명암 
      Interval        =   2
      Left            =   4980
      Top             =   2820
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   435
      Left            =   5520
      TabIndex        =   5
      Top             =   2820
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   767
      _Version        =   393216
      AutoPlay        =   -1  'True
      BackStyle       =   1
      FullWidth       =   65
      FullHeight      =   29
   End
   Begin VB.Label lblUserPW 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "User PW"
      Height          =   180
      Left            =   4440
      TabIndex        =   13
      Top             =   2040
      Width           =   630
   End
   Begin VB.Label lblUserID 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "User ID"
      Height          =   180
      Left            =   4440
      TabIndex        =   12
      Top             =   1680
      Width           =   630
   End
   Begin VB.Label lbl설명 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Interface For Medical Machine"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   11
      Top             =   540
      Width           =   4065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "HIS DataBase Info"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   4
      Left            =   540
      TabIndex        =   10
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Shape shpNo2 
      BorderColor     =   &H00000000&
      BorderStyle     =   0  '투명
      FillColor       =   &H000000FF&
      FillStyle       =   0  '단색
      Height          =   255
      Left            =   180
      Shape           =   3  '원형
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Local DataBase Info"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   3
      Left            =   540
      TabIndex        =   9
      Top             =   1680
      Width           =   2565
   End
   Begin VB.Shape shpNo1 
      BorderColor     =   &H00000000&
      BorderStyle     =   0  '투명
      FillColor       =   &H000000FF&
      FillStyle       =   0  '단색
      Height          =   255
      Left            =   180
      Shape           =   3  '원형
      Top             =   1680
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderStyle     =   0  '투명
      FillColor       =   &H000000FF&
      FillStyle       =   0  '단색
      Height          =   255
      Left            =   4860
      Shape           =   3  '원형
      Top             =   360
      Width           =   255
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   1560
      X2              =   4920
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblState 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "DataBase 접속 중..."
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Top             =   900
      Width           =   2700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Interface EQ"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   21.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   435
      Index           =   6
      Left            =   1560
      TabIndex        =   7
      Top             =   0
      Width           =   2715
   End
   Begin VB.Label lbl회사이름 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "메디메이트(주)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   180
      TabIndex        =   6
      Top             =   3000
      Width           =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      Index           =   5
      X1              =   1500
      X2              =   6600
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   4
      X1              =   1680
      X2              =   6600
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   3
      X1              =   2400
      X2              =   6600
      Y1              =   1500
      Y2              =   1500
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   2
      X1              =   180
      X2              =   6480
      Y1              =   3300
      Y2              =   3300
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      Index           =   1
      X1              =   2160
      X2              =   6600
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   1920
      X2              =   6600
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Hi"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   72
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   1440
      Index           =   2
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   1410
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Version ?.?"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   5460
      TabIndex        =   3
      Top             =   3360
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Copyright ⓒ 2010 Medimate Co., Ltd."
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   2
      Top             =   3360
      Width           =   3165
   End
   Begin VB.Shape Shp반짝이 
      BackColor       =   &H00C000C0&
      BackStyle       =   1  '투명하지 않음
      Height          =   315
      Index           =   0
      Left            =   6300
      Top             =   60
      Width           =   315
   End
End
Attribute VB_Name = "frmEQ공용_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbl명암증가 As Double
Dim dbl명암속도 As Double

'/폼 투명 효과----------------------------------------------------------------------------------------------------------------------------------------------------------------/
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2
'/폼 투명 효과----------------------------------------------------------------------------------------------------------------------------------------------------------------/

'/폼 투명 효과
Private Function MakeLayeredWnd(hwnd As Long) As Long
     Dim WndStyle As Long

     WndStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
     WndStyle = WndStyle Or WS_EX_LAYERED
     MakeLayeredWnd = SetWindowLong(hwnd, GWL_EXSTYLE, WndStyle)
End Function

Public Sub SUB_MM_INITIAL()
    '/STEP1.Local DataBase 연결
    GoSub RTN_LOCALDB_CONNECT
    
    '/STEP2.HIS DataBase 연결
    GoSub RTN_HISDB_CONNECT
Exit Sub

'/----------------------------------------------------------------------------------------------------/

'/STEP1.Local DataBase 연결
'/- 1.INTERFACE.MDB 화일의 존재유무 확인한다.
'/    -> 화일 없으면 프로그램 종료
'/    -> 화일 있으면 다음 Step 진행
'/- 2.Local DataBase Connect 을 시행한다
'/    -> 연결실패 시 프로그램 종료
'/    -> 연결성공 시 기본정보 인식

RTN_LOCALDB_CONNECT:
    lblState = "Local DataBase 접속 중": DoEvents
    
    If Dir(App.Path & "\INTERFACE.MDB") = "" Then
        MsgBox "INTERFACE.MDB 화일이 실행화일과 같은 폴더에 존재하지 않습니다" & vbCrLf & _
               "전산실 혹은 공급업체에 연락주시기 바랍니다.", vbCritical, "프로그램 종료"
        
        End
    End If

    If ConnDB_LOC = False Then
        lblState = "Local DataBase 접속 실패!!!": DoEvents
            
        MsgBox "▶▶▶Local DataBase Info" & vbCrLf & vbCrLf & _
               "Local DataBase 를 접속할 수 없습니다." & vbCrLf & _
               "정상적인 프로그램 운용을 위해 전산실 혹은 공급업체에 연락주시기 바랍니다.", vbInformation, "확인"
    Else
        shpNo1.FillColor = RGB(0, 0, 255)
        
        gstrQuy = "SELECT * "
        gstrQuy = gstrQuy & vbCrLf & "  FROM CUS_MST "
        If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then End
        
        If Not ADR_LOC Is Nothing Then
            gtypHIS_CNN_INFO.ID = Trim(ADR_LOC!HISDB_ID & "")
            gtypHIS_CNN_INFO.PW = Trim(ADR_LOC!HISDB_PW & "")
            gtypHIS_CNN_INFO.SV = Trim(ADR_LOC!HISDB_SERVER & "")
            gtypHIS_CNN_INFO.DBNM = Trim(ADR_LOC!HISDB_DBNM & "") '/DBNM 명(SQL Server 일 경우)
            gtypHIS_CNN_INFO.TYPE = Trim(ADR_LOC!HISDB_TYPE & "") '/DB 종류
        
            ADR_LOC.Close: Set ADR_LOC = Nothing
        End If
        
        gstrQuy = "SELECT * "
        gstrQuy = gstrQuy & vbCrLf & "  FROM EQ_CONF "
        If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then End
        
        If Not ADR_LOC Is Nothing Then
            gtypEQ_INFO.SERIALPORT = Trim(ADR_LOC!SERIALPORT & "")
            gtypEQ_INFO.SERIALBAUD = Trim(ADR_LOC!SERIALBAUD & "")
            gtypEQ_INFO.SERIALDATABIT = Trim(ADR_LOC!SERIALDATABIT & "")
            gtypEQ_INFO.SERIALSTARTBIT = Trim(ADR_LOC!SERIALSTARTBIT & "")
            gtypEQ_INFO.SERIALSTOPBIT = Trim(ADR_LOC!SERIALSTOPBIT & "")
            gtypEQ_INFO.SERIALPARITY = Trim(ADR_LOC!SERIALPARITY & "")
            gtypEQ_INFO.SERIALRTS = Trim(ADR_LOC!SERIALRTS & "")
            gtypEQ_INFO.SERIALDTR = Trim(ADR_LOC!SERIALDTR & "")
            gtypEQ_INFO.WORKLISTGB = Trim(ADR_LOC!WORKLISTGB & "")
            gtypEQ_INFO.AUTOGB = Trim(ADR_LOC!AUTOGB & "")
        
            ADR_LOC.Close: Set ADR_LOC = Nothing
        End If
        
        Call CloseDB_LOC
    End If
Return

'/----------------------------------------------------------------------------------------------------/

'/STEP2.HIS DataBase 연결
'/- 1.Local DataBase Connect 을 시행한다
'/    -> 연결실패 시 프로그램 종료
'/    -> 연결성공 시 Login Process 진행

RTN_HISDB_CONNECT:
    lblState = "HIS Database 접속 중": DoEvents

    If ConnDB_HIS = True Then
        shpNo2.FillColor = RGB(0, 0, 255)
        
        Call CloseDB_HIS
    Else
        lblState = "HIS Database 실패!!!": DoEvents
        
        If MsgBox("▶▶▶HIS DataBase Info" & vbCrLf & vbCrLf & _
                  "HIS DB Connection Information이 올바르지 않습니다." & vbCrLf & _
                  "(재)설정하겠습니까?", vbQuestion + vbYesNo, "질의") = vbNo Then
            
            MsgBox "▶▶▶HIS DataBase Info" & vbCrLf & vbCrLf & _
                   "HIS DataBase 를 접속할 수 없습니다." & vbCrLf & _
                   "정상적인 프로그램 운용을 위해 전산실 혹은 공급업체에 연락주시기 바랍니다.", vbInformation, "확인"
                
            '/ID와 암호를 자사 지정 정보로 할 경우 End를 막고 아래 주석을 푼다.
            '''MsgBox "▶▶▶HIS DataBase Info" & vbCrLf & vbCrLf & _
                   "계속 진행할 경우 일부 기능이 제한됩니다." & vbCrLf & _
                   "정상적인 프로그램 운용을 위해 전산실 혹은 공급업체에 연락주시기 바랍니다.", vbInformation, "확인"
                
            End
            '''GoTo RTN_HISDB_CONNECT_SKIP'/ID와 암호를 자사 지정 정보로 할 경우 End를 막고 본 라인을 푼다.
        Else
            gstrArgTemp1 = "HIS": frmEQ공용_Set_DB.Show vbModal
        End If

        GoTo RTN_HISDB_CONNECT
        
RTN_HISDB_CONNECT_SKIP:
    
    End If
Return
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    '/중복실행 방지 루틴----------------------------------------------------------------------------------------------------/
'    If PrevInstance Then
'        MsgBox "프로그램이 이미 구동중입니다", vbExclamation, "이미 구동중"
'        End
'    End If
    '/중복실행 방지 루틴----------------------------------------------------------------------------------------------------/

    Me.Height = 4215
    Me.Width = 6795
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    txtUserID = ""
    txtUserPW = ""
    
    lblState = ""
    lbl설명 = "Interface For " & App.FileDescription
    
    lbl회사이름 = App.CompanyName
    lblVersion = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    shpNo1.FillColor = RGB(255, 0, 0)
    shpNo2.FillColor = RGB(255, 0, 0)
    
    
On Error Resume Next
    Animation1.Open App.Path & "\Login1.avi"
On Error GoTo 0

    DoEvents
    DoEvents
    DoEvents
    
    '/화면 명암관련 루틴----------------------------------------------------------------------------------------------------/
    Me.Visible = False
    tmr명암.Enabled = False
    
    MakeLayeredWnd Me.hwnd
    SetLayeredWindowAttributes Me.hwnd, 0, 255 * (0), LWA_ALPHA
    
    dbl명암속도 = 0.01
    
    tmr명암.Enabled = True
    tmr명암.Interval = 2
    
    Me.Visible = True
        
    DoEvents
    DoEvents
    DoEvents
    '/화면 명암관련 루틴----------------------------------------------------------------------------------------------------/
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CloseDB_LOC
    Call CloseDB_HIS
    Call CloseDB_ETC
    
    Set frmEQ공용_Login = Nothing
End Sub

Private Sub tmr명암_Timer()
    dbl명암증가 = dbl명암증가 + dbl명암속도 '0.03
    
    If dbl명암증가 > 1 Then
        dbl명암증가 = 1
    
        tmr명암.Enabled = False
        tmr명암.Interval = 0
    
        Call SUB_MM_INITIAL
        
        'txtUserID = "800042"
        'txtUserPW = "1"
        
        lblState = "ID 와 Password 를 입력하십시오!": DoEvents
        txtUserID.SetFocus
    Else
        MakeLayeredWnd Me.hwnd
        SetLayeredWindowAttributes Me.hwnd, 0, 255 * (dbl명암증가), LWA_ALPHA
    End If
End Sub

Private Sub txtUserID_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtUserID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtUserPW_GotFocus()
    
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtUserPW_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtUserID) = "" Then
            MsgBox "User ID를 (재)입력하십시오!", vbCritical, "로그인 실패"
            txtUserID.SetFocus
            Exit Sub
        End If
        If Trim(txtUserPW) = "" Then
            MsgBox "User PW를 (재)입력하십시오!", vbCritical, "로그인 실패"
            txtUserPW.SetFocus
            Exit Sub
        End If
        
        With gtypUSER
            .USERID = "" '/사용자ID
            .USERNM = "" '/사용자명
            .USERPW = "" '/사용자PW
        End With
        
'''        '/----------------------------------------------------------------------------------------------------/
'''        '/기본 Login Pass 부분
'''        '/----------------------------------------------------------------------------------------------------/
'''        If ConnDB_LOC = True Then
'''            gstrQuy = "SELECT USER_ID, USER_PW, USER_NM "
'''            gstrQuy = gstrQuy & vbCrLf & "  FROM USER_MST " '/HIS 사용자마스터 테이블(서북병원)
'''            gstrQuy = gstrQuy & vbCrLf & " WHERE USER_ID = '" & Trim(txtUserID) & "' "
'''            If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then Call CloseDB_LOC: End
'''
'''            If Not ADR_LOC Is Nothing Then
'''                If Trim(txtUserPW) = Trim(ADR_LOC!USER_PW & "") Then
'''                    gtypUSER.USERID = Trim(ADR_LOC!USER_ID & "")
'''                    gtypUSER.USERNM = Trim(ADR_LOC!USER_NM & "")
'''                    gtypUSER.USERPW = Trim(ADR_LOC!USER_PW & "")
'''
'''                    ADR_LOC.Close: Set ADR_LOC = Nothing
'''
'''                    Unload Me
'''
'''                    Call Main
'''                Else
'''                    ADR_LOC.Close: Set ADR_LOC = Nothing
'''
'''                    MsgBox "User PW가 맞지 않습니다!", vbCritical, "로그인 실패": Exit Sub
'''                End If
'''            Else
'''                MsgBox "등록되지 않은 ID 입니다!", vbCritical, "로그인 실패": Exit Sub
'''            End If
'''
'''            Call CloseDB_LOC
'''        End If
'''        '/----------------------------------------------------------------------------------------------------/
'''        '/기본 Login Pass 부분
'''        '/----------------------------------------------------------------------------------------------------/
        
        
        '/----------------------------------------------------------------------------------------------------/
        '/사용기관별 Login Pass 부분
        '/----------------------------------------------------------------------------------------------------/
        If ConnDB_HIS = True Then
            gstrQuy = "SELECT UID_1,USERENAME,UPASSWD "
            gstrQuy = gstrQuy & vbCrLf & "FROM USERMASTER"
            gstrQuy = gstrQuy & vbCrLf & "WHERE UID_1 = '" & Trim(txtUserID) & "' "
'            gstrQuy = "SELECT USER_ID, USER_NM, PWD "
'            gstrQuy = gstrQuy & vbCrLf & "  FROM TZUSERMSTN " '/HIS 사용자마스터 테이블(서북병원)
'            gstrQuy = gstrQuy & vbCrLf & " WHERE USER_ID = '" & Trim(txtUserID) & "' "
            If ReadSQL_HIS(gstrQuy, ADR_HIS) = False Then Call CloseDB_HIS: End

            If Not ADR_HIS Is Nothing Then
                If Trim(txtUserPW) = Trim(ADR_HIS!UPASSWD & "") Then
                    gtypUSER.USERID = Trim(ADR_HIS!UID_1 & "")
                    gtypUSER.USERNM = Trim(ADR_HIS!USERENAME & "")
                    gtypUSER.USERPW = Trim(ADR_HIS!UPASSWD & "")

                    ADR_HIS.Close: Set ADR_HIS = Nothing

                    Unload Me

                    Call Main
                Else
                    ADR_HIS.Close: Set ADR_HIS = Nothing

                    MsgBox "User PW가 맞지 않습니다!", vbCritical, "로그인 실패": Exit Sub
                End If
            Else
                MsgBox "등록되지 않은 ID 입니다!", vbCritical, "로그인 실패": Exit Sub
            End If

            Call CloseDB_HIS
        End If
        '/----------------------------------------------------------------------------------------------------/
        '/사용기관별 Login Pass 부분
        '/----------------------------------------------------------------------------------------------------/
    End If
End Sub
