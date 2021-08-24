VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQueryABO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Public Function GetABOOrderList(ByVal ptid As String, _
'                                ByVal FrDt As String, ByVal ToDt As String, _
'                                ByVal dayfg As String) As Recordset
'
'    Dim strSQL As String
'
'    If ptid <> "" Then
'        strSQL = "SELECT a.orddoct,a.ordtm," & _
'                      "b.ptid,b.orddt,b.ordno,b.ordseq,b.ordcd,b.rcvdt,b.rcvtm,b.workarea,b.accdt,b.accseq,b.statfg,b.stscd," & _
'                      "c.testnm,d.field3 spcnm " & _
'               "FROM " & T_LAB032 & " d," & T_LAB102 & " b," & T_LAB001 & " c," & T_LAB101 & " a " & _
'               "WHERE " & DBW("a.orddiv=", "L") & " " & _
'               "AND " & DBW("a.ptid=", ptid) & " " & _
'               "AND a.ptid=b.ptid AND a.orddt=b.orddt AND a.ordno=b.ordno " & _
'               "AND b.ordcd=c.testcd AND " & DBW("c.testdiv=", "3") & " " & _
'               "AND b.spccd=d.cdval1 AND " & DBW("d.cdindex=", "C215") & " "
'    Else
'        strSQL = "SELECT a.orddoct,a.ordtm," & _
'                      "b.ptid,b.orddt,b.ordno,b.ordseq,b.ordcd,b.rcvdt,b.rcvtm,b.workarea,b.accdt,b.accseq,b.statfg,b.stscd," & _
'                      "c.testnm,d.field3 spcnm " & _
'               "FROM " & T_LAB032 & " d," & T_LAB102 & " b," & T_LAB001 & " c," & T_LAB101 & " a " & _
'               "WHERE " & DBW("a.orddiv=", "L") & " " & _
'               "AND a.ptid=b.ptid AND a.orddt=b.orddt AND a.ordno=b.ordno " & _
'               "AND b.ordcd=c.testcd AND " & DBW("c.testdiv=", "3") & " " & _
'               "AND b.spccd=d.cdval1 AND " & DBW("d.cdindex=", "C215") & " "
'    End If
'
'    If dayfg = 0 Then
'        '立荐老磊肺
'        strSQL = strSQL & " " & _
'               "AND " & DBW("b.rcvdt>=", FrDt) & " AND " & DBW("b.rcvdt<=", ToDt) & " "
'    Else
'        '贸规老磊肺
'        strSQL = strSQL & " " & _
'               "AND " & DBW("a.orddt>=", FrDt) & " AND " & DBW("a.orddt<=", ToDt) & " "
'    End If
'    strSQL = strSQL & " order by ptid"
'    Set GetABOOrderList = New Recordset
'    Call GetABOOrderList.Open(strSQL, DBConn)
'End Function
'
'Public Function GetABOResult(ByVal workarea As String, ByVal accdt As String, ByVal accseq As String) As Recordset
'    Dim strSQL As String
'
''    strSQL = "SELECT a.testcd,a.rstval,a.rstcd,a.rstunit," & _
''                  "b.testnm " & _
''           "FROM " & T_LAB302 & " a," & T_LAB001 & " b " & _
''           "WHERE " & DBW("a.workarea=", workarea) & " " & _
''           "AND " & DBW("a.accdt=", accdt) & " " & _
''           "AND " & DBW("a.accseq=", accseq) & " " & _
''           "AND a.testcd=b.testcd AND " & DBW("b.testdiv=", "3") & " "
'        strSQL = " SELECT a.testnm,b.workarea,b.accdt,b.accseq,b.testcd,b.rstcd,b.rstunit,c.field1 as RstCdNm " & _
'                " FROM " & T_LAB031 & " c," & T_LAB001 & " a," & T_LAB302 & " b" & _
'                " WHERE " & DBW("b.workarea=", workarea) & " " & _
'                " AND " & DBW("b.accdt=", accdt) & " " & _
'                " AND " & DBW("b.accseq=", accseq) & " " & _
'                " AND a.testcd=b.testcd " & _
'                " and a.testdiv='3' " & _
'                " AND     " & DBJ(DBW("c.cdindex", LC2_ItemResult, 2)) & _
'                " AND     " & DBJ("c.cdval1   =* b.testcd") & _
'                " AND     " & DBJ("c.cdval2   =* b.rstcd")
'
'    Set GetABOResult = New Recordset
'    Call GetABOResult.Open(strSQL, DBConn)
'End Function
'
'Public Function GetAccNoList(ByVal accdt As String) As String
'    Dim strSQL As String
'    Dim Rs   As Recordset
'
'    strSQL = " SELECT max(accseq) as maxseq,min(accseq) as minseq FROM " & T_BBS303 & _
'                   " WHERE " & DBW("accdt", accdt, 2)
'
'    Set Rs = New Recordset
'    Rs.Open strSQL, DBConn
'    If Not Rs.EOF Then
'        GetAccNoList = Rs.Fields("maxseq").Value & "" & COL_DIV & Rs.Fields("minseq").Value & ""
'    End If
'
'    Set Rs = Nothing
'End Function


