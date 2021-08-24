Attribute VB_Name = "modLimas"
Option Explicit

Global Const REG_EQPCODE    As String = "INSCODE"
Global Const REG_EQPNAME    As String = "INSNAME"
Global Const REG_POSITION   As String = "Software\KMI_INTERFACE\" & REG_INSNAME

'Visual Basic Color
Global Const vbLockColor = &HE0E0E0

'검사 타입
Public Const MSG_GEN As String = "G"        '일반
Public Const MSG_QCT As String = "Q"        'QC
Public Const MSG_ETC As String = "E"        '기타

Public INS_CODE             As String       '장비코드
Public INS_NAME             As String       '장비명
Public Const HOS_NAME       As String = ""      '병원명

Public DirPath              As String
Public MainForm             As MDIMain
Private TimerID             As Long

Sub Main()

    Dim strMsg As String
    Dim lngConnect  As Long
    
    '두번 실행 하지 않음
    If App.PrevInstance Then
       MsgBox "     Now Excute twice!", vbExclamation, INS_NAME
       End
    End If

    'Registree Scan
    If Len(GetString(HKEY_CURRENT_USER, REG_JETDB, REG_DATABASE)) = 0 Then
        frmDB_JET.Show vbModal
    End If
    
    If Len(GetString(HKEY_CURRENT_USER, REG_MSSQLDB, REG_SERVER)) = 0 Then
        frmDB_SQL.Show vbModal
    End If

    If Not DbConnect_Jet Then
        strMsg = "Local Batabase Not found! Do you want database search it? "
        If vbYes = MsgBox(strMsg, vbCritical + vbYesNo, INS_NAME) Then
            frmDB_JET.Show vbModal
        Else
            End
        End If
    End If
     
    If Not DbConnect_SQL Then
        strMsg = "SQL Batabase Not found! Do you want database search it?   "
        If vbYes = MsgBox(strMsg, vbCritical + vbYesNo, INS_NAME) Then
            frmDB_SQL.Show vbModal
        Else
            End
        End If
    End If
    
    '실행 위치 저장
    DirPath = App.Path
    If Right(DirPath, 1) <> "\" Then DirPath = DirPath & "\"
    
    UpdateODBCMDB DirPath & "Database\" & "Interface.mdb"
    
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


