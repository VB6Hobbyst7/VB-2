VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISSqlStatementsT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''%  45. 바코드라벨 재출력을 위해 해당 환자의 처방내역을 조회한다.
''%       - Calling FROM [frm157BarReprint] : 라벨재발행
''%       - 검사구분, WorkArea, 검체코드, 보관구분, 응급여부, 희망채혈일시,
''%         복수검체여부, 검체군, 처방일, 처방번호, 처방Seq, 처방코드
''%       - Lab201 : 채혈접수내역
''%       - Lab102 : 처방Body내역
''%       - Lab001 : 검사항목 내역
''%       - Lab032 : 검체명, 건물명
Public Function SqlBarReprint(ByVal intOption As Integer, ParamArray strKeys() As Variant) As String

    Dim tmpStr As String
    Dim tmpStr1 As String
    Dim tmpSql(3) As String
    Dim blnAllFg As Boolean

    tmpStr = ""
    blnAllFg = False
    If intOption = 1 Then  '환자ID-처방일 기준
        tmpStr = tmpStr & DBW("b.ptid = ", strKeys(0)) & " AND " & DBW("b.orddt =", strKeys(1)) & " AND " & DBW("b.donefg >", "0") & " AND " & DBW("b.dcfg =", "") & " "
        tmpStr1 = "AND e.ptid = b.ptid AND e.orddt = b.orddt " & _
                  "AND e.ordno = b.ordno AND e.ordseq = b.ordseq " & _
                  "AND a.workarea = e.workarea AND a.accdt = e.accdt " & _
                  "AND a.accseq = e.accseq "
        If UBound(strKeys) = 2 Then
            If strKeys(2) = "" Then blnAllFg = True
        End If
    Else        '접수번호 기준
        tmpStr = tmpStr & DBW("a.workarea = ", strKeys(0)) & " AND " & DBW("a.accdt = ", strKeys(1)) & " AND " & DBW("a.accseq = ", strKeys(2)) & " "
        tmpStr1 = "AND e.workarea = a.workarea AND e.accdt = a.accdt " & _
                  "AND e.accseq = a.accseq AND b.ptid = e.ptid " & _
                  "AND b.orddt = e.orddt AND b.ordno = e.ordno " & _
                  "AND b.ordseq = e.ordseq   AND    b.dcfg = '' "
    End If

    '혈액은행
    tmpSql(3) = "SELECT c.testnm, c.abbrnm5, a.ptid, 'XM' as workarea, '' as accdt, 0 as accseq, " & _
                "       a.spcyy, a.spcno, '' spccd, '' as storecd, '1' as stscd, b.statfg, a.buildcd, " & _
                "       i.deptcd, i.hosilid, '' as LabNo, '' as SpcNm, " & FUNC_SUBSTR & "(a.coldt,5,2)" & FUNC_CONCAT & "'/'" & FUNC_CONCAT & "" & _
                "       " & FUNC_SUBSTR & "(a.coldt,7,2) as ColDt, b.orddt, b.ordno, b.ordseq, b.ordcd, " & _
                "       e.field1 as BuildNm, f." & F_PTNM & " as PtNm, " & _
                "       i.reqdt, i.reqtm, i.orddiv, '' fzfg " & _
                "FROM " & T_BBS201 & " a, " & T_LAB102 & " b, " & T_BBS001 & " c, " & _
                          T_LAB032 & " e, " & T_HIS001 & " f, " & T_LAB101 & " i " & _
                "WHERE " & tmpStr & _
                "AND    i.ptid = b.ptid " & _
                "AND    i.orddt = b.orddt " & _
                "AND    i.ordno = b.ordno " & _
                "AND    " & DBW("i.orddiv = ", C_WORKAREA) & " " & _
                "AND    a.ptid = i.ptid " & _
                "AND    a.spcyy = (SELECT max(spcyy) FROM " & T_BBS201 & " WHERE ptid = i.ptid) " & _
                "AND    a.spcno = (SELECT max(spcno) FROM " & T_BBS201 & _
                                 " WHERE ptid = a.ptid AND spcyy = a.spcyy) " & _
                "AND    c.testcd = b.ordcd " & _
                "AND    " & DBW("e.cdindex = ", CD2_Buildings) & " " & _
                "AND    " & DBJ("e.cdval1 =* a.buildcd") & " " & _
                "AND    f." & F_PTID & " = a.ptid "

    If blnAllFg Then
        SqlBarReprint = tmpSql(1) & " UNION ALL " & tmpSql(2) & " UNION ALL " & tmpSql(3)
    Else
        Select Case strKeys(2)
        Case "A": SqlBarReprint = tmpSql(2)
        Case "B": SqlBarReprint = tmpSql(3)
        Case "L": SqlBarReprint = tmpSql(1)
        End Select
    End If
    SqlBarReprint = SqlBarReprint & " ORDER BY LabNo, b.ordno, b.ordcd "

End Function

