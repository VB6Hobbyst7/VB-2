Attribute VB_Name = "modLimas"
Option Explicit

Global Const REG_EQPCODE    As String = "INSCODE"
Global Const REG_EQPNAME    As String = "INSNAME"
Global Const REG_POSITION   As String = "Software\DIDIM_Interface\�������պ���\" & REG_INSNAME

'�������̺� ��� �ε���
Public Const IDX_STA        As String = "C202" 'LIMAS032 ��ũ�����̼�
Public Const IDX_SPC        As String = "C203" 'LIMAS032 ��ü
Public Const IDX_EQP        As String = "C209" 'LIMAS032 ��񸮽�Ʈ
Public Const IDX_ROOM       As String = "C252" 'LIMAS032 �˻��
Public Const IDX_SITE       As String = "C261" 'LIMAS032 �����
Public Const IDX_TST        As String = "C604" 'LIMAS032 ��� �˻��ڵ�

'Visual Basic Color
Global Const vbLockColor = &HE0E0E0

'�˻� Ÿ��
Public Const MSG_GEN As String = "G"        '�Ϲ�
Public Const MSG_QCT As String = "Q"        'QC
Public Const MSG_ETC As String = "E"        '��Ÿ

'���� ����� ����
Public Const ELVELS_SUP  As String = "��� ����"
Public Const ELVELS_RED  As String = "��   ��"
Public Const ELVELS_WRI  As String = "��   ��"
Public Const ELVELS_RW   As String = "�б�,����"
Public Const ELVELS_NOT  As String = "���� ����"

Public Type UserInfo
    CuUserID    As String '����� ID
    CuUserNM    As String '����� �̸�
    CuUserPW    As String '����� ��й�ȣ
    CuPower     As Authority  '����� ����
End Type

' �� ��
Public Enum Authority
    ELVEL_SUP = 1
    ELVEL_RED = 2
    ELVEL_WRI = 3
    ELVEL_RW = 4
    ELVEL_NOT = 5
End Enum

'���� ����� ����
Public CurrUser             As UserInfo
Public INS_CODE             As String       '����ڵ�
Public INS_NAME             As String       '����

Public DirPath              As String
Public MainForm             As MDIMain
Private TimerID             As Long


'-- �ش� Ƽ�ƽ� ���� : 3.1.4
'-- Ƽ�ƽ� DLL  ���� : TMAX4GL.DLL
Public initcheck As Boolean

Public Function getSvrcInfo(ByVal strSrcNm As String, ByVal strArg1 As String, _
                            Optional ByVal strArg2 As String, Optional ByVal strArg3 As String, _
                            Optional ByVal strArg4 As String, Optional ByVal strArg5 As String, _
                            Optional ByVal strArg6 As String, Optional ByVal strArg7 As String, _
                            Optional ByVal strArg8 As String, Optional ByVal strArg9 As String) As String
    
    Dim iRet As Integer
    Dim strText1 As String, strText2 As String, strText3 As String, strText4 As String, strText5 As String
    Dim strText6 As String, strText7 As String, strText8 As String, strText9 As String, strText10 As String
    Dim ErrMsg As String
    Dim sendBuf As tuxbuf
    Dim rcvbuf As tuxbuf
    Dim rbuflen  As Long
    Dim strSvrNm As String
    Dim iRcvLen As Long
    Dim strGet(100) As String
    Dim intCnt As Integer
    Dim IntRow As Integer
    Dim intRecord As Integer
    
    Dim sendBuf1 As tuxbuf
    Dim rcvbuf1 As tuxbuf
    
    
    Dim Isendbuf As Long
    
    If initcheck = False Then
        MsgBox "TMAX������ ������ �����Դϴ�. ������ �� �����ϴ�"
        Exit Function
    End If
    
    sendBuf.bufptr = tpalloc("FIELD", "", ByVal 1024&)
    rcvbuf.bufptr = tpalloc("FIELD", "", ByVal 1024&)
        
    If sendBuf.bufptr = 0 Or rcvbuf.bufptr = 0 Then
        ErrMsg = tmaxerrdesc(gettperrno())
        Exit Function
    End If
    
    getSvrcInfo = ""
    strSvrNm = strSrcNm
    strText1 = strArg1
    strText2 = strArg2
    strText3 = strArg3
    strText4 = strArg4
    strText5 = strArg5
    strText6 = strArg6
    strText7 = strArg7
    strText8 = strArg8
    strText9 = strArg9
    
    Select Case strSvrNm
        '-- Dummy
        Case "CC_SYSDATE_S"
            IntRow = 1
            If fbput(ByVal sendBuf.bufptr, fbget_fldkey("GW_LOCATE"), ByVal "3", 0) = -1 Then
                ErrMsg = tmaxerrdesc(gettperrno())
                MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX ����"
                GoTo memoryfree
            End If
        
        '-- ����� ����
        Case "SL_USERM_L1"
            IntRow = 2
            If fbput(ByVal sendBuf.bufptr, fbget_fldkey("GW_LOCATE"), ByVal "3", 0) = -1 Then
                ErrMsg = tmaxerrdesc(gettperrno())
                MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX ����"
                GoTo memoryfree
            End If
            
            If fbput(ByVal sendBuf.bufptr, fbget_fldkey("S_IDNUM1"), ByVal strText1, 0) = -1 Then
                ErrMsg = tmaxerrdesc(gettperrno())
                MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX ����"
                GoTo memoryfree
            End If
        
        '-- �������� download (������)
        Case "SL_INTFC_L2"
            IntRow = 15
            If fbput(ByVal sendBuf.bufptr, fbget_fldkey("GW_LOCATE"), ByVal "3", 0) = -1 Then
                ErrMsg = tmaxerrdesc(gettperrno())
                MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX ����"
                GoTo memoryfree
            End If
            
            If fbput(ByVal sendBuf.bufptr, fbget_fldkey("S_IDNUM1"), ByVal strText1, 0) = -1 Then
                ErrMsg = tmaxerrdesc(gettperrno())
                MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX ����"
                GoTo memoryfree
            End If
            
        '-- ���� ����
        Case "SL_SPCMD_L12"
            IntRow = 2
            If fbput(ByVal sendBuf.bufptr, fbget_fldkey("GW_LOCATE"), ByVal "3", 0) = -1 Then
                ErrMsg = tmaxerrdesc(gettperrno())
                MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX ����"
                GoTo memoryfree
            End If
            
            
            If fbput(ByVal sendBuf.bufptr, fbget_fldkey("S_FLAG1"), ByVal strText1, 0) = -1 Then
                ErrMsg = tmaxerrdesc(gettperrno())
                MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX ����"
                GoTo memoryfree
            End If
            
            If strText2 <> "" Then
                If fbput(ByVal sendBuf.bufptr, fbget_fldkey("S_IDNUM1"), ByVal strText2, 0) = -1 Then
                    ErrMsg = tmaxerrdesc(gettperrno())
                    MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX ����"
                    GoTo memoryfree
                End If
            End If
        
        Case "SL_RSLTUP_M1"
            IntRow = 1
            If fbput(ByVal sendBuf.bufptr, fbget_fldkey("S_STRING1"), ByVal strText1, 0) = -1 Then      '-- ����ڵ�
                ErrMsg = tmaxerrdesc(gettperrno())
                MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX ����"
                GoTo memoryfree
            End If
            
            If strText2 <> "" Then
                If fbput(ByVal sendBuf.bufptr, fbget_fldkey("S_IDNUM1"), ByVal strText2, 0) = -1 Then   '-- ��ü��ȣ
                    ErrMsg = tmaxerrdesc(gettperrno())
                    MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX ����"
                    GoTo memoryfree
                End If
            End If
        
            If strText3 <> "" Then
                If fbput(ByVal sendBuf.bufptr, fbget_fldkey("S_CODE1"), ByVal strText3, 0) = -1 Then    '-- �˻��ڵ�
                    ErrMsg = tmaxerrdesc(gettperrno())
                    MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX ����"
                    GoTo memoryfree
                End If
            End If
        
            If strText4 <> "" Then
                If fbput(ByVal sendBuf.bufptr, fbget_fldkey("S_NO1"), ByVal strText4, 0) = -1 Then      '-- ó�����
                    ErrMsg = tmaxerrdesc(gettperrno())
                    MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX ����"
                    GoTo memoryfree
                End If
            End If
        
            If strText5 <> "" Then
                If fbput(ByVal sendBuf.bufptr, fbget_fldkey("S_TEXT1"), ByVal strText5, 0) = -1 Then    '-- �˻���
                    ErrMsg = tmaxerrdesc(gettperrno())
                    MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX ����"
                    GoTo memoryfree
                End If
            End If
        
            If strText6 <> "" Then
                If fbput(ByVal sendBuf.bufptr, fbget_fldkey("LOG_USERID"), ByVal strText6, 0) = -1 Then '-- ���
                    ErrMsg = tmaxerrdesc(gettperrno())
                    MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX ����"
                    GoTo memoryfree
                End If
            End If
        
            If strText7 <> "" Then
                If fbput(ByVal sendBuf.bufptr, fbget_fldkey("LOG_CLTNAME"), ByVal strText7, 0) = -1 Then '-- ���� IP
                    ErrMsg = tmaxerrdesc(gettperrno())
                    MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX ����"
                    GoTo memoryfree
                End If
            End If

    End Select
        
    iRet = tpcall(strSvrNm, ByVal sendBuf.bufptr, ByVal 0, rcvbuf.bufptr, iRcvLen, ByVal 0)
    
    If iRet = -1 Then
        ErrMsg = tmaxerrdesc(gettperrno())
        getSvrcInfo = ""
        GoTo memoryfree
    End If
    
    If strSvrNm = "CC_SYSDATE_S" Then
        iRet = fbkeyoccur(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_DATETIME1"))
        For intRecord = 1 To iRet
            For intCnt = 1 To IntRow
                strGet(intCnt) = String$(1024, Chr$(0))
                            
                iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_DATETIME1"), ByVal strGet(intCnt), 0)   '-- ����
                
                If iRet > 0 And Trim(strGet(intCnt)) <> "" Then
                    getSvrcInfo = getSvrcInfo & Mid(strGet(intCnt), 1, InStr(strGet(intCnt), Chr(0)) - 1)
                End If
            Next
        Next
    
    ElseIf strSvrNm = "SL_USERM_L1" Then
        iRet = fbkeyoccur(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_NAME1"))
    
    
        For intRecord = 1 To iRet
            For intCnt = 1 To IntRow
                strGet(intCnt) = String$(1024, Chr$(0))
                            
                If intCnt = 1 Then
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_NAME1"), ByVal strGet(intCnt), 0)   '-- ����
                ElseIf intCnt = 2 Then
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_TEXT1"), ByVal strGet(intCnt), 0)   '-- ��й�ȣ
                End If
                
                If iRet > 0 And Trim(strGet(intCnt)) <> "" Then
                    getSvrcInfo = getSvrcInfo & Mid(strGet(intCnt), 1, InStr(strGet(intCnt), Chr(0)) - 1) & "|"
                End If
            Next
        Next

    
    ElseIf strSvrNm = "SL_INTFC_L2" Then
        iRet = fbkeyoccur(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_DATE1"))
    
    
        For intRecord = 1 To iRet
            For intCnt = 1 To IntRow
                strGet(intCnt) = String$(1024, Chr$(0))
                                            
                Select Case intCnt
                Case 1
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_DATE1"), ByVal strGet(intCnt), 0)       '-- �˻�����[yyyymmdd]
                Case 2
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_NUMVAL2"), ByVal strGet(intCnt), 0)     '-- ���Ź�ȣ
                Case 3
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_NO1"), ByVal strGet(intCnt), 0)         '-- ó�����
                Case 4
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_CODE2"), ByVal strGet(intCnt), 0)       '-- �˻��ڵ�
                Case 5
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_NAME1"), ByVal strGet(intCnt), 0)       '-- �˻��
                Case 6
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_CODE3"), ByVal strGet(intCnt), 0)       '-- ��ü�ڵ�
                Case 7
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_NAME2"), ByVal strGet(intCnt), 0)       '-- ��ü��
                Case 8
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_PATNO"), ByVal strGet(intCnt), 0)       '-- ��Ϲ�ȣ
                Case 9
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_PATNAME"), ByVal strGet(intCnt), 0)     '-- ����
                Case 10
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_STRING1"), ByVal strGet(intCnt), 0)     '-- ����
                Case 11
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_STRING2"), ByVal strGet(intCnt), 0)     '-- ����
                Case 12
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_CODE4"), ByVal strGet(intCnt), 0)       '-- �����
                Case 13
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_TEXT1"), ByVal strGet(intCnt), 0)       '-- �������
                Case 14
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_TEXT2"), ByVal strGet(intCnt), 0)       '-- �˻���
                Case 15
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_DATE2"), ByVal strGet(intCnt), 0)       '-- ������������Ͻ�[yyyymmdd hh24:mi:ss]
                End Select
                
                If iRet > 0 And Trim(strGet(intCnt)) <> "" Then
                    getSvrcInfo = getSvrcInfo & Mid(strGet(intCnt), 1, InStr(strGet(intCnt), Chr(0)) - 1) & "|"
                End If
            Next
            getSvrcInfo = getSvrcInfo '& vbCrLf
        Next
    
    ElseIf strSvrNm = "SL_SPCMD_L12" Then
        iRet = fbkeyoccur(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_STAT1"))
    
        For intRecord = 1 To iRet
            For intCnt = 1 To IntRow
                strGet(intCnt) = String$(1024, Chr$(0))
                
                If intCnt = 1 Then
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_CODE6"), ByVal strGet(intCnt), 0)
                    
                    If iRet > 0 And Trim(strGet(intCnt)) <> "" Then
                        getSvrcInfo = getSvrcInfo & Mid(strGet(intCnt), 1, InStr(strGet(intCnt), Chr(0)) - 1) & "|"    '-- �˻��ڵ�
                    End If
                ElseIf intCnt = 2 Then
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_STAT1"), ByVal strGet(intCnt), 0)
                    
                    If iRet > 0 And Trim(strGet(intCnt)) <> "" Then
                        getSvrcInfo = getSvrcInfo & Mid(strGet(intCnt), 1, InStr(strGet(intCnt), Chr(0)) - 1) & "|"    '-- ���°� B:ä��/����,C:����,E:�˻�,N:����
                    End If
                End If
                
                
            Next
        Next

    ElseIf strSvrNm = "SL_RSLTUP_M1" Then
        iRet = fbkeyoccur(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_DATE1"))
    
    
        For intRecord = 1 To iRet
            For intCnt = 1 To IntRow
                strGet(intCnt) = String$(1024, Chr$(0))
                                            
                Select Case intCnt
                Case 1
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_DATE1"), ByVal strGet(intCnt), 0)       '--
                Case 2
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_NUMVAL2"), ByVal strGet(intCnt), 0)
                Case 3
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_NO1"), ByVal strGet(intCnt), 0)
                Case 4
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_CODE2"), ByVal strGet(intCnt), 0)
                Case 5
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_NAME1"), ByVal strGet(intCnt), 0)
                Case 6
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_CODE3"), ByVal strGet(intCnt), 0)
                Case 7
                    iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_NAME2"), ByVal strGet(intCnt), 0)
                End Select
                
                If iRet > 0 And Trim(strGet(intCnt)) <> "" Then
                    getSvrcInfo = getSvrcInfo & Mid(strGet(intCnt), 1, InStr(strGet(intCnt), Chr(0)) - 1) & "|"
                End If
            Next
            getSvrcInfo = getSvrcInfo & vbCrLf
        Next
    End If
    
'    Debug.Print getSvrcInfo

memoryfree:
    
    Call tpfree(ByVal sendBuf.bufptr)
    Call tpfree(ByVal rcvbuf.bufptr)
    
    Call tmaxexit

End Function



Sub Main()
    Dim ret     As Integer
    Dim lsndbuf As Long
    Dim tpinfo  As tpstart_t
    Dim strPtr  As Long
    Dim iRet    As Integer
    Dim lret    As Long
    Dim retVal  As String
    Dim strPath As String
    Dim strLbl  As String
    Dim ErrMsg  As String
    Dim strMsg As String
    Dim strSvrcCnt  As Variant
    Dim strSvrcData As Variant
    
    Dim lngConnect  As Long
    
    '�ι� ���� ���� ����
    If App.PrevInstance Then
       MsgBox "     Now Excute twice!", vbExclamation
       End
    End If
        
    'Registree Scan
    If Len(GetString(HKEY_CURRENT_USER, REG_JETDB, REG_DATABASE)) = 0 Then
        frmDB_JET.Show vbModal
    End If
    
    If Len(GetString(HKEY_CURRENT_USER, REG_ORACLEDB, REG_SERVER)) = 0 Then
        frmDB_ORACLE.Show vbModal
    End If
    
    If Not DbConnect_Jet Then
        strMsg = "Local Batabase Not found! Do you want database search it? "
        If vbYes = MsgBox(strMsg, vbCritical + vbYesNo) Then
            frmDB_JET.Show vbModal
        Else
            End
        End If
    End If
    
    If Not DbConnect_ORACLE Then
        strMsg = "Oracle Batabase Not found! Do you want database search it?   "
        If vbYes = MsgBox(strMsg, vbCritical + vbYesNo) Then
            frmDB_ORACLE.Show vbModal
        Else
            End
        End If
    End If
    
    '���� ��ġ ����
    DirPath = App.Path
    If Right(DirPath, 1) <> "\" Then DirPath = DirPath & "\"
        
    'Login Form ��Ÿ��
'    frmLogin.Show vbModal
    Set MainForm = New MDIMain
    MainForm.Show
    
End Sub

'Progressbar ����
Public Sub SetProgress(ByVal lngMax As Long, ByVal CapStyle As CaptionStyles, ByVal strCaption As String, ByVal blnVisible As Boolean)
    'lngMax         : �ִ밪
    'CapStyle       : �μ� ��Ÿ��
    'strCaption     : �μ�
    'blnVisible     : ����

    With MainForm.pgbMain
        .Max = lngMax
        .Visible = blnVisible
        .CaptionStyle = CapStyle
        .Caption = strCaption
        .Value = 0
    End With
End Sub

'Progressbar �� ����
Public Sub ShowProgress(ByVal Values As Long, ByVal strCaption As String, ByVal blnVisible As Boolean)
    'Values         : ��
    'strCaption     : �μ�
    'blnVisible     : ��Ÿ��
    
    With MainForm.pgbMain
        .Visible = blnVisible
        .Caption = strCaption
        .Value = Values
    End With
End Sub

'���� ǥ���ٿ� �޽��� �ڵ� �����
Public Sub TimerProc(ByVal hwnd&, ByVal MSG&, ByVal ID&, ByVal nTime&)
    Call KillTimer(MainForm.hwnd, TimerID)
    With MainForm.stbMain
        .Panels("Output").text = ""
    End With
End Sub

'���� ǥ���ٿ� �޽��� ��Ÿ����
Public Sub ShowMessage(ByVal strMessage As String)
    'strMessage : �μ�
    
    Call KillTimer(MainForm.hwnd, TimerID)
    Call SetTimer(MainForm.hwnd, TimerID, 5000, AddressOf TimerProc)
    
    With MainForm
        With .pgbMain
            .Visible = False
        End With
        With .stbMain
            .Panels("Output").text = strMessage
        End With
    End With
    
End Sub

