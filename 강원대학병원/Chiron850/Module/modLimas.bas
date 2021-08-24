Attribute VB_Name = "modLimas"
Option Explicit

Global Const REG_EQPCODE    As String = "INSCODE"
Global Const REG_EQPNAME    As String = "INSNAME"
Global Const REG_POSITION   As String = "Software\LIS_PAIK\" & REG_INSNAME

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

Sub Main()
    Dim strMsg      As String
    Dim rv          As Long
    Dim LocalPath   As String
    
    '�ι� ���� ���� ����
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
        MsgBox "DB���� ����", vbOKOnly, MDIMain.Caption
        Exit Sub
    Else
         'rv = dce_error("msg")
         'MessageDlg('TCP Error: ' + msg, mtInformation,[mbOK],0);
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
Public Sub TimerProc(ByVal hwnd&, ByVal msg&, ByVal ID&, ByVal nTime&)
    Call KillTimer(MainForm.hwnd, TimerID)
    With MainForm.stbMain
        .Panels("Output").Text = ""
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
            .Panels("Output").Text = strMessage
        End With
    End With
    
End Sub

