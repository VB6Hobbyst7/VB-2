Attribute VB_Name = "modLimas"
Option Explicit

Global Const REG_EQPCODE    As String = "INSCODE"
Global Const REG_EQPNAME    As String = "INSNAME"
Global Const REG_POSITION   As String = "Software\KMI_INTERFACE\" & REG_INSNAME

'Visual Basic Color
Global Const vbLockColor = &HE0E0E0

'�˻� Ÿ��
Public Const MSG_GEN As String = "G"        '�Ϲ�
Public Const MSG_QCT As String = "Q"        'QC
Public Const MSG_ETC As String = "E"        '��Ÿ

Public INS_CODE             As String       '����ڵ�
Public INS_NAME             As String       '����
Public Const HOS_NAME       As String = "�λ��ڸ𺴿�"      '������

Public DirPath              As String

'2011-03-21 ������ LIS�� �Ź��� LIS ���� �������̽� ó��
Public gLisVer              As String

Public MainForm             As MDIMain
Private TimerID             As Long

Private Sub VersionSelect()
Dim vntRs       As Variant
Dim objTestItem As clsCommon

    Set objTestItem = New clsCommon
       
    With objTestItem
        Call .SetAdoCn(AdoCn_SQL)
        vntRs = .Get_LisVer()
    End With
    gLisVer = vntRs(0, 0)
    
    Set objTestItem = Nothing
End Sub

Sub Main()

    Dim strMsg As String
    Dim lngConnect  As Long
    
    '�ι� ���� ���� ����
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
    
    'LIS ���� Ȯ��
    Call VersionSelect
    
    '���� ��ġ ����
    DirPath = App.Path
    If Right(DirPath, 1) <> "\" Then DirPath = DirPath & "\"
    
    UpdateODBCMDB DirPath & "Database\" & "Interface.mdb"
    
    Set MainForm = New MDIMain
    MainForm.Show
    
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


