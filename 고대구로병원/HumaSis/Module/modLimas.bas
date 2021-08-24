Attribute VB_Name = "modLimas"
Option Explicit

Global Const REG_EQPCODE    As String = "INSCODE"
Global Const REG_EQPNAME    As String = "INSNAME"
Global Const REG_POSITION   As String = "Software\�޵����Ʈ\��뱸�κ���\" & REG_INSNAME

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

'***************************************************************
'   Function: Put The String Data Into FDL Buffer
'   Fdlptr  : FDL BUFFER POINTER
'   Field   : FIELD NAME
'   idx     : OCCURRENCE OF DATA
'   text    : STRING DATA
'***************************************************************
Function PUTVAR(ByVal Fdlptr&, Field As String, idx As Long, text As String) As Integer

    Dim ret As Long
    Dim lret As Long
    Dim tp_err_no As Integer
    Dim err_ret As Integer
    
    ret = fbchg_tu(ByVal Fdlptr&, ByVal fbget_fldkey(Field), idx, ByVal text$, 0)
   
    If ret = -1 Then
        err_ret = FdlErrorMsg("fbchg_tu")
    End If
    
End Function

Public Function getSvrcInfo(ByVal strSrcNm As String, ByVal strArg1 As String, ByVal strArg2 As String, ByVal strArg3 As String) As String
    
    Dim iRet As Integer
    Dim strText1, strText2, strText3
    
    
    Dim ErrMsg As String
    
    Dim sendBuf As tuxbuf
    Dim rcvbuf As tuxbuf
    Dim rbuflen  As Long
    Dim strSvrNm As String
    Dim iRcvLen As Long
    Dim strGet(100) As String
    Dim intCnt As Integer
    Dim intRow As Integer
    Dim intRecord As Integer
    
    Dim sendBuf1 As tuxbuf
    Dim rcvbuf1 As tuxbuf
    
    Dim varText1 As Variant, varText2 As Variant, varText3 As Variant
    
    
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


    varText1 = strArg1
    varText2 = strArg2
'    strText3 = "2"
    
    If fbput(ByVal sendBuf.bufptr, ByVal fbget_fldkey("S_TYPE1"), ByVal strArg1, 0) = -1 Then
        ErrMsg = tmaxerrdesc(gettperrno())
        MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX ����"
        GoTo memoryfree
    End If
    
    If fbput(ByVal sendBuf.bufptr, ByVal fbget_fldkey("S_TYPE2"), ByVal strArg2, 0) = -1 Then
        ErrMsg = tmaxerrdesc(gettperrno())
        MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX ����"
        GoTo memoryfree
    End If
    
    '-- 2012.08.22 ���Ⱦ����� ȫâ�� ��ȭ
    '-- S_TYPE3�� '2'�� �Ѱܾ� ��
    If fbput(ByVal sendBuf.bufptr, ByVal fbget_fldkey("S_TYPE3"), ByVal strArg3, 0) = -1 Then
        ErrMsg = tmaxerrdesc(gettperrno())
        MsgBox "Tmax Error Number : " & gettperrno() & vbNewLine & ErrMsg, vbOKOnly + vbCritical, "TMAX ����"
        GoTo memoryfree
    End If
    
    iRet = tx_begin
        
    iRet = tpcall(strSvrNm, ByVal sendBuf.bufptr, ByVal 0, rcvbuf.bufptr, iRcvLen, ByVal 0)
    
    
    If iRet = -1 Then
        ErrMsg = tmaxerrdesc(gettperrno())
        getSvrcInfo = ""
        getSvrcInfo = ErrMsg
        
        GoTo memoryfree
        iRet = tx_rollback
    Else
        iRet = tx_commit
    
    End If
    
    iRet = fbkeyoccur(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_STRING1"))
    For intRecord = 1 To iRet
        For intCnt = 1 To intRow
            strGet(intCnt) = String$(1024, Chr$(0))
                        
            iRet = fbget(ByVal rcvbuf.bufptr, ByVal fbget_fldkey("S_STRING1"), ByVal strGet(intCnt), 0)   '-- ����
            
            If iRet > 0 And Trim(strGet(intCnt)) <> "" Then
                getSvrcInfo = getSvrcInfo & Mid(strGet(intCnt), 1, InStr(strGet(intCnt), Chr(0)) - 1)
            End If
        Next
    Next


    
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
    
'    If Format(Now, "yyyymmdd") > "20120430" Then
'       MsgBox "     ���Ⱓ�� ����� �����Դϴ�.", vbExclamation
'       End
'    End If
    
    '�ι� ���� ���� ����
    If App.PrevInstance Then
       MsgBox "     Now Excute twice!", vbExclamation
       End
    End If

    'Registree Scan
    If Len(GetString(HKEY_CURRENT_USER, REG_JETDB, REG_DATABASE)) = 0 Then
        frmDB_JET.Show vbModal
    End If
    
    If Not DbConnect_Jet Then
        strMsg = "Local Batabase Not found! Do you want database search it? "
        If vbYes = MsgBox(strMsg, vbCritical + vbYesNo) Then
            frmDB_JET.Show vbModal
        Else
            End
        End If
    End If
    
    '���� ��ġ ����
    DirPath = App.Path
    If Right(DirPath, 1) <> "\" Then DirPath = DirPath & "\"
    
'    If TmaxConnect Then
'        initcheck = True
'    End If
    
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

