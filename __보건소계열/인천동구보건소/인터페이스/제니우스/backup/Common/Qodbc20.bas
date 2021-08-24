Attribute VB_Name = "QODBC20"
'=========================================================================='
'             Copyright(c) 95, KICO(Korea Information Computing)           '
'--------------------------------------------------------------------------'
'                                                                          '
'                      Quick ODBC 2.0 for Visual Basic                     '
'                      -------------------------------                     '
'  �� DLL�� Visual Basic�� �̿��� ODBC�Լ����� ���� Accessó���� �����ϴ�  '
'  �Լ��Դϴ�.                                                             '
'=========================================================================='
Option Explicit

Public Index%

'*****************************************
'*** PROCESS DEFINE
'*****************************************
Public QsqlConn  As Long
Public QsqlCon1  As Long
Public QsqlCon2  As Long
Public QsqlCode  As Long
Public QsqlCod1  As Long
Public QsqlCod2  As Long

Public Image_Flg%
Public TEMPFILE As String * 50         ' Load�� Image File��
Public IMAGEFILE$                      ' ���õ� �̹��� ȭ�ϸ� (full path)
Public record$                         ' DB�κ��� ������ ���ڵ尡 ����� ����(QSqlSetRecord�� ȣ���Ͽ� Column�� ���ؾ�!)
Public col_cnt%                        ' DB�κ��� ������ ���ڵ��� Column ��
Public COL_LENGTH%                     ' QSqlSetRecord���� ����ϴ� ���ڵ� Column�� ����
Public QSqlData() As String            ' �� ���ڵ��� Cloumn������� ���� �迭(QSqlSetRecord���� ���)

Public FETCH_RECORD$                   ' DB�κ��� ������ ���ڵ尡 ����� ����(QSqlFetchRecord�� ȣ���Ͽ� Column�� ���ؾ�!)
Public FETCH_COL_CNT%                  ' DB�κ��� ������ ���ڵ��� Column ��
Public FETCH_COL_LENGTH%               ' QSqlFetchRecord���� ����ϴ� ���ڵ� Column�� ����
Public QSqlFetchData() As String       ' �� ���ڵ��� Cloumn������� ���� �迭(QSqlFetchRecord���� ���)

Public Const MAX_DB_USE = 10           ' �ִ� �̿� DB��
Public Const QSQL_NO_IMAGE = 99        ' �ش� �̹����� ���� ���
Public Const QSQL_NO_DATA_FOUND = 100  ' �ش� ����Ÿ�� ���� ���

Public Const QSQL_SUCCESS = 0          ' ���������� �������� ���
Public Const QSQL_NOUPDATE = 2         ' UPDATE ROW ZERO
Public Const QSQL_ERROR = 1            ' �������� ���
Public Const QSQL_ALLOC_ERROR = -2     ' �ش� ����Ÿ�� ���� ���
Public Const QSQL_TRANS_ERROR = -3     ' Ʈ�����ó���� ������ ���
Public Const QSQL_DUPINDEX = -4        ' DB Open�� index �ߺ�����
Public Const QSQL_NOTINDEX = -5        ' DB Close�� index ����

Public Const ONECLOSE = 0              ' DB 1���� Close
Public Const ALLCLOSE = 1              ' DB ��� Close

Public Const QSQL_FETCH_NEXT = 1       ' ���� ���ڵ带 ã�� ���
Public Const QSQL_FETCH_FIRST = 2      ' �� ó�� ���ڵ带 ã�� ���
Public Const QSQL_FETCH_LAST = 3       ' �� ������ ���ڵ带 ã�� ���
Public Const QSQL_FETCH_PREV = 4       ' ���� ���ڵ带 ã�� ���
Public Const QSQL_FETCH_ABSOLUTE = 5   ' irow �� ������ ���ڵ�� �̵��� ���
Public Const QSQL_FETCH_RELATIVE = 6   ' irow �� ������ ���� ���ڵ�� �̵��� ���

Public Const QSQL_MS_ACCESS = 0        ' Image Select�� Microsoft Access DB����
Public Const QSQL_SQL_SERVER = 1       ' Image Select�� SQL_SERVER DB����

'Declare Function QSqlOpen Lib "QODBC20.DLL" (ByVal Server$, ByVal HWND%, Index%) As Integer
'Declare Function Qsqlclose Lib "QODBC20.DLL" (ByVal Index%, ByVal finish%) As Integer

'Declare Function QSqlSelect Lib "QODBC20.DLL" (ByVal sStr$, RECORD$, COL_CNT%, COL_LENGTH%, ByVal maxrows%, ByVal Index%) As Integer
'Declare Function QSqlSelectFree Lib "QODBC20.DLL" (ByVal Index%) As Integer

'Declare Function QSqlFetchOpen Lib "QODBC20.DLL" (ByVal Server$, ByVal HWND%, ByVal sStr$, FETCH_COL_CNT%, FETCH_COL_LENGTH%, Index%) As Integer
'Declare Function QSqlRefresh Lib "QODBC20.DLL" (FETCH_RECORD$) As Integer
'Declare Function QSqlFetch Lib "QODBC20.DLL" (FETCH_RECORD$, ByVal fetch%, ByVal irow%, ByVal Index%) As Integer
'Declare Function QSqlFetchClose Lib "QODBC20.DLL" (ByVal Index%) As Integer

'Declare Function QSqlDBExec Lib "QODBC20.DLL" (ByVal sStr$, ByVal Index%) As Integer
'Declare Function QSqlGetRow Lib "QODBC20.DLL" (RECORD$, ByVal Index%) As Integer

'Declare Function QSqlDBExec Lib "QODBC20.DLL" (ByVal sStr$, ByVal Index%) As Integer
'Declare Function QSqlDelete Lib "QODBC20.DLL" (ByVal sStr$, ByVal Index%) As Integer
'Declare Function QSqlUpdate Lib "QODBC20.DLL" (ByVal sStr$, ByVal Index%) As Integer

'Declare Function QSqlImgSelect Lib "QODBC20.DLL" (ByVal sStr$, ByVal TEMPBMP$, ByVal Index%) As Integer
'Declare Function QSqlImgInsert Lib "QODBC20.DLL" (ByVal sStr$, ByVal TEMPBMP$, ByVal Index%) As Integer
'Declare Function QSqlImgUpdate Lib "QODBC20.DLL" (ByVal sStr$, ByVal TEMPBMP$, ByVal Index%) As Integer

'Declare Function QSqlBeginTrans Lib "QODBC20.DLL" () As Integer
'Declare Function QSqlRollBack Lib "QODBC20.DLL" () As Integer
'Declare Function QSqlCommitTrans Lib "QODBC20.DLL" () As Integer

Function QSqlError(iError As Integer) As String

    Select Case iError
        Case -5: QSqlError = " ��ȿ���� ���� �ε����Դϴ�."
        Case -4: QSqlError = " Open�� DB�ε����� �ƴմϴ�."
        Case -3: QSqlError = " �޸� ����"
        Case -2: QSqlError = " �޸� ����"
        Case -1: QSqlError = " DB Server ���ӿ���"
        Case 0: QSqlError = " �����Դϴ�.(Ret = 0, ����� ����)"
        Case 1: QSqlError = " �ش��ڷᰡ �̹� �����մϴ�."
        Case 2: QSqlError = " �ش��ڷᰡ �������� �ʽ��ϴ�. ��ȸ�� �����Ͻʽÿ�."
        Case 3: QSqlError = " �Է��ڷᰡ ��ȿ���� ����(Constraint Error) "
        Case 10: QSqlError = " �̹� Open�� DB�ε����Դϴ�."
        Case 11: QSqlError = " Syntex Error"
        Case 12: QSqlError = " Column ������ Value ������ ��ġ���� �ʽ��ϴ�."
        Case 13: QSqlError = " Table �Ǵ� View�� �������� �ʽ��ϴ�."
        Case 14: QSqlError = " �������� �ʴ� Column���Դϴ�."
        Case 15: QSqlError = " ������ ������ ������ϴ�."
        Case 16: QSqlError = " Table �Ǵ� View�� �̹� �����մϴ�."
        Case 17: QSqlError = " �ε����� �̹� �����մϴ�."
        Case 18: QSqlError = " �ε����� ã���� �����ϴ�."
        Case 19: QSqlError = " DateTime Field Overflow"
        Case 21: QSqlError = " ��ſ��ῡ �����߻�"
        Case 22: QSqlError = " �޸� �Ҵ� ����"
        Case 23: QSqlError = " Function Sequence Error"
        Case 24: QSqlError = " Time Out"
        Case 31: QSqlError = " ����� ����"
        Case 100: QSqlError = " �ش��ڷᰡ �������� �ʽ��ϴ�."
        Case Else: QSqlError = " �˼� ���� �����߻�"
    End Select
            
End Function

Sub QSqlFetchRecord()
    
    Dim i%, II%, K%
    ReDim QSqlFetchData(1 To FETCH_COL_CNT) As String

    II = 1: K = 1
    For i = 1 To FETCH_COL_CNT
        II = InStr(K, FETCH_RECORD, Chr(5), 1)
        QSqlFetchData(i) = Trim(Mid$(FETCH_RECORD, K, II - K))
        K = II + 1
    Next i
End Sub

Sub QSqlGetField(nCols As Integer, ByVal sStr As String, SData() As String)
    
    Dim i%, II%, K%
    Dim P%, c%
    ReDim SData(1 To nCols) As String
    
    II = 1: K = 1
    For i = 1 To nCols
        II = InStr(K, sStr, Chr(5), 1)
        SData(i) = Trim(Mid$(sStr, K, II - K))
        '-------------------
        '   Ascii Code �� 0�� Character �����
        '   97.05. KHJ
        '-------------------
        P = InStr(1, SData(i), Chr(0))
        If P <> 0 Then
            SData(i) = Left(SData(i), P - 1)
            For c = P To Len(SData(i))
                SData(i) = SData(i) & " "
            Next
        End If
        
        K = II + 1
    Next i
    
End Sub

Sub QSqlSetRecord()
    
    Dim i%, II%, K%
    ReDim QSqlData(1 To col_cnt) As String

    II = 1: K = 1
    For i = 1 To col_cnt
        II = InStr(K, record, Chr(5), 1)
        QSqlData(i) = Trim(Mid$(record, K, II - K))
        K = II + 1
    Next i
End Sub

