Attribute VB_Name = "QODBC20"
'=========================================================================='
'             Copyright(c) 95, KICO(Korea Information Computing)           '
'--------------------------------------------------------------------------'
'                                                                          '
'                      Quick ODBC 2.0 for Visual Basic                     '
'                      -------------------------------                     '
'  본 DLL은 Visual Basic을 이용한 ODBC함수에서 빠른 Access처리를 지원하는  '
'  함수입니다.                                                             '
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
Public TEMPFILE As String * 50         ' Load할 Image File명
Public IMAGEFILE$                      ' 선택된 이미지 화일명 (full path)
Public record$                         ' DB로부터 가져온 레코드가 저장된 변수(QSqlSetRecord를 호출하여 Column을 구해야!)
Public col_cnt%                        ' DB로부터 가져온 레코드의 Column 수
Public COL_LENGTH%                     ' QSqlSetRecord에서 사용하는 레코드 Column의 길이
Public QSqlData() As String            ' 한 레코드의 Cloumn내용들을 가진 배열(QSqlSetRecord에서 사용)

Public FETCH_RECORD$                   ' DB로부터 가져온 레코드가 저장된 변수(QSqlFetchRecord를 호출하여 Column을 구해야!)
Public FETCH_COL_CNT%                  ' DB로부터 가져온 레코드의 Column 수
Public FETCH_COL_LENGTH%               ' QSqlFetchRecord에서 사용하는 레코드 Column의 길이
Public QSqlFetchData() As String       ' 한 레코드의 Cloumn내용들을 가진 배열(QSqlFetchRecord에서 사용)

Public Const MAX_DB_USE = 10           ' 최대 이용 DB수
Public Const QSQL_NO_IMAGE = 99        ' 해당 이미지가 없을 경우
Public Const QSQL_NO_DATA_FOUND = 100  ' 해당 데이타가 없을 경우

Public Const QSQL_SUCCESS = 0          ' 성공적으로 수행했을 경우
Public Const QSQL_NOUPDATE = 2         ' UPDATE ROW ZERO
Public Const QSQL_ERROR = 1            ' 실패했을 경우
Public Const QSQL_ALLOC_ERROR = -2     ' 해당 데이타가 없을 경우
Public Const QSQL_TRANS_ERROR = -3     ' 트랜잭션처리에 실패한 경우
Public Const QSQL_DUPINDEX = -4        ' DB Open시 index 중복에러
Public Const QSQL_NOTINDEX = -5        ' DB Close시 index 에러

Public Const ONECLOSE = 0              ' DB 1개만 Close
Public Const ALLCLOSE = 1              ' DB 모두 Close

Public Const QSQL_FETCH_NEXT = 1       ' 다음 레코드를 찾을 경우
Public Const QSQL_FETCH_FIRST = 2      ' 맨 처음 레코드를 찾을 경우
Public Const QSQL_FETCH_LAST = 3       ' 맨 마지막 레코드를 찾을 경우
Public Const QSQL_FETCH_PREV = 4       ' 이전 레코드를 찾을 경우
Public Const QSQL_FETCH_ABSOLUTE = 5   ' irow 로 지정한 레코드로 이동할 경우
Public Const QSQL_FETCH_RELATIVE = 6   ' irow 로 지정한 다음 레코드로 이동할 경우

Public Const QSQL_MS_ACCESS = 0        ' Image Select시 Microsoft Access DB에서
Public Const QSQL_SQL_SERVER = 1       ' Image Select시 SQL_SERVER DB에서

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
        Case -5: QSqlError = " 유효하지 않은 인덱스입니다."
        Case -4: QSqlError = " Open된 DB인덱스가 아닙니다."
        Case -3: QSqlError = " 메모리 부족"
        Case -2: QSqlError = " 메모리 부족"
        Case -1: QSqlError = " DB Server 접속에러"
        Case 0: QSqlError = " 정상입니다.(Ret = 0, 사용자 에러)"
        Case 1: QSqlError = " 해당자료가 이미 존재합니다."
        Case 2: QSqlError = " 해당자료가 존재하지 않습니다. 조회를 실행하십시오."
        Case 3: QSqlError = " 입력자료가 유효하지 않음(Constraint Error) "
        Case 10: QSqlError = " 이미 Open된 DB인덱스입니다."
        Case 11: QSqlError = " Syntex Error"
        Case 12: QSqlError = " Column 갯수와 Value 갯수가 일치하지 않습니다."
        Case 13: QSqlError = " Table 또는 View가 존재하지 않습니다."
        Case 14: QSqlError = " 존재하지 않는 Column명입니다."
        Case 15: QSqlError = " 숫자의 범위를 벗어났습니다."
        Case 16: QSqlError = " Table 또는 View가 이미 존재합니다."
        Case 17: QSqlError = " 인덱스가 이미 존재합니다."
        Case 18: QSqlError = " 인덱스를 찾을수 없습니다."
        Case 19: QSqlError = " DateTime Field Overflow"
        Case 21: QSqlError = " 통신연결에 에러발생"
        Case 22: QSqlError = " 메모리 할당 에러"
        Case 23: QSqlError = " Function Sequence Error"
        Case 24: QSqlError = " Time Out"
        Case 31: QSqlError = " 사용자 에러"
        Case 100: QSqlError = " 해당자료가 존재하지 않습니다."
        Case Else: QSqlError = " 알수 없는 에러발생"
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
        '   Ascii Code 가 0인 Character 지우기
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

