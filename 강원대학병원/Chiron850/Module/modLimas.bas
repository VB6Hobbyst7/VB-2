Attribute VB_Name = "modLimas"
Option Explicit

Global Const REG_EQPCODE    As String = "INSCODE"
Global Const REG_EQPNAME    As String = "INSNAME"
Global Const REG_POSITION   As String = "Software\LIS_PAIK\" & REG_INSNAME

'공용테이블 장비 인덱스
Public Const IDX_STA        As String = "C202" 'LIMAS032 워크스테이션
Public Const IDX_SPC        As String = "C203" 'LIMAS032 검체
Public Const IDX_EQP        As String = "C209" 'LIMAS032 장비리스트
Public Const IDX_ROOM       As String = "C252" 'LIMAS032 검사실
Public Const IDX_SITE       As String = "C261" 'LIMAS032 사업장
Public Const IDX_TST        As String = "C604" 'LIMAS032 장비별 검사코드

'Visual Basic Color
Global Const vbLockColor = &HE0E0E0

'검사 타입
Public Const MSG_GEN As String = "G"        '일반
Public Const MSG_QCT As String = "Q"        'QC
Public Const MSG_ETC As String = "E"        '기타

'현재 사용자 정보
Public Const ELVELS_SUP  As String = "모든 권한"
Public Const ELVELS_RED  As String = "읽   기"
Public Const ELVELS_WRI  As String = "쓰   기"
Public Const ELVELS_RW   As String = "읽기,쓰기"
Public Const ELVELS_NOT  As String = "권한 없음"

Public Type UserInfo
    CuUserID    As String '사용자 ID
    CuUserNM    As String '사용자 이름
    CuUserPW    As String '사용자 비밀번호
    CuPower     As Authority  '사용자 권한
End Type

' 권 한
Public Enum Authority
    ELVEL_SUP = 1
    ELVEL_RED = 2
    ELVEL_WRI = 3
    ELVEL_RW = 4
    ELVEL_NOT = 5
End Enum

'현재 사용자 정보
Public CurrUser             As UserInfo
Public INS_CODE             As String       '장비코드
Public INS_NAME             As String       '장비명

Public DirPath              As String
Public MainForm             As MDIMain
Private TimerID             As Long

Sub Main()
    Dim strMsg      As String
    Dim rv          As Long
    Dim LocalPath   As String
    
    '두번 실행 하지 않음
    If App.PrevInstance Then
       MsgBox "     Now Excute twice!", vbExclamation
       End
    End If
    
    'Registree Scan
    If Len(GetString(HKEY_CURRENT_USER, REG_JETDB, REG_DATABASE)) = 0 Then
        frmDB_JET.Show vbModal
    End If
'
'    If Len(GetString(HKEY_CURRENT_USER, REG_MSSQLDB, REG_SERVER)) = 0 Then
'        frmDB_SQL.Show vbModal
'    End If
    
    If Not DbConnect_Jet Then
        strMsg = "Local Batabase Not found! Do you want database search it? "
        If vbYes = MsgBox(strMsg, vbCritical + vbYesNo) Then
            frmDB_JET.Show vbModal
        Else
            End
        End If
    End If
       
    rv = dce_setenv(App.Path & "\sl.env", "", "")
    If (rv = 0) Then
        MsgBox "DB연결 실패", vbOKOnly, MDIMain.Caption
        Exit Sub
    Else
         'rv = dce_error("msg")
         'MessageDlg('TCP Error: ' + msg, mtInformation,[mbOK],0);
    End If
    
        
    '실행 위치 저장
    DirPath = App.Path
    If Right(DirPath, 1) <> "\" Then DirPath = DirPath & "\"
    
    'Login Form 나타남
'    frmLogin.Show vbModal
    Set MainForm = New MDIMain
    MainForm.Show
    
End Sub

'Progressbar 설정
Public Sub SetProgress(ByVal lngMax As Long, ByVal CapStyle As CaptionStyles, ByVal strCaption As String, ByVal blnVisible As Boolean)
    'lngMax         : 최대값
    'CapStyle       : 켑션 스타일
    'strCaption     : 켑션
    'blnVisible     : 보임

    With MainForm.pgbMain
        .Max = lngMax
        .Visible = blnVisible
        .CaptionStyle = CapStyle
        .Caption = strCaption
        .Value = 0
    End With
End Sub

'Progressbar 값 설정
Public Sub ShowProgress(ByVal Values As Long, ByVal strCaption As String, ByVal blnVisible As Boolean)
    'Values         : 값
    'strCaption     : 켑션
    'blnVisible     : 나타남
    
    With MainForm.pgbMain
        .Visible = blnVisible
        .Caption = strCaption
        .Value = Values
    End With
End Sub

'상태 표시줄에 메시지 자동 지우기
Public Sub TimerProc(ByVal hwnd&, ByVal msg&, ByVal ID&, ByVal nTime&)
    Call KillTimer(MainForm.hwnd, TimerID)
    With MainForm.stbMain
        .Panels("Output").Text = ""
    End With
End Sub

'상태 표시줄에 메시지 나타내기
Public Sub ShowMessage(ByVal strMessage As String)
    'strMessage : 켑션
    
    Call KillTimer(MainForm.hwnd, TimerID)
    Call SetTimer(MainForm.hwnd, TimerID, 5000, AddressOf TimerProc)
    
    With MainForm
        With .pgbMain
            .Visible = False
        End With
        With .stbMain
            .Panels("Output").Text = strMessage
        End With
    End With
    
End Sub

