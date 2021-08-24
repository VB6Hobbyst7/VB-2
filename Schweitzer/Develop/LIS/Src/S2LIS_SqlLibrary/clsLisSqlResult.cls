VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisSqlResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"clsEquipMaster"
Option Explicit

'
'Dim OraSession  As Object

Dim m_RecordSet As Recordset
Dim m_Query     As String
Dim m_Fields    As Object
Dim m_SaveOK    As Boolean


'OO4O Error Message
Dim m_LastServerErr     As Long
Dim m_LastServerErrText As String


'/* ---------------------------------------------------------*/
'/*                         Common SQL Method                               */
'/* ---------------------------------------------------------*/
Public Property Get LastServerErr() As Long
'Last OO4O Error: Read Only Property
   '
    LastServerErr = m_LastServerErr
   '
End Property

Public Property Get LastServerErrText() As String
'Last OO4O Error Text: Read Only Property
   '
    LastServerErrText = m_LastServerErrText
   '
End Property

Public Property Get SaveOk() As Boolean
    SaveOk = m_SaveOK
End Property

Private Sub DBExec(ByVal SqlStmt As String)
'Execute the SQL Command
On Error GoTo DBExecError
   'LOCK TABLE table IN ROW SHARE MODE
    DBConn.BeginTrans
    DBConn.Execute SqlStmt
    DBConn.CommitTrans
   '
    Exit Sub
DBExecError:
'    Call Error_Routine
    MsgBox "처리 도중 오류가 발생하였습니다.", vbExclamation
End Sub

Private Sub DBMultiExec(ByRef arySql As Variant)
'Execute the Multi SQL Command
    Dim intSqlCnt As Integer
On Error GoTo DBExecError
   'LOCK TABLE table IN ROW SHARE MODE
    DBConn.BeginTrans
    For intSqlCnt = LBound(arySql) To UBound(arySql)
        DBConn.Execute (arySql(intSqlCnt))
    Next intSqlCnt
    DBConn.CommitTrans
   '
    Exit Sub
DBExecError:
'    Call Error_Routine
    MsgBox "처리 도중 오류가 발생하였습니다.", vbExclamation
End Sub

Private Sub DBOpenDyna(ByVal SqlStmt As String)
'Open OO4O-Dynaset based on the SQL Query
On Error GoTo DBOpenDynaError
   '
    Set m_RecordSet = Nothing
    Set m_RecordSet = New Recordset ' OpenRecordSet(SqlStmt)
    
    m_RecordSet.Open SqlStmt, DBConn
    
    Exit Sub
   
DBOpenDynaError:
'    Call Error_Routine
    MsgBox "처리 도중 오류가 발생하였습니다.", vbExclamation
End Sub

Private Function DBExists(ByVal SqlStmt As String) As Boolean
'Check Database with OO4O-Dynaset
    Dim OraDynaset As Recordset
On Error GoTo DBOpenDynaError
   '
    Set OraDynaset = New Recordset
    OraDynaset.Open SqlStmt, DBConn
    
    If OraDynaset.RecordCount > 0 Then
       DBExists = True
    Else
       DBExists = False
    End If
    Set OraDynaset = Nothing
    '
    Exit Function
DBOpenDynaError:
'    Call Error_Routine
    Set OraDynaset = Nothing
End Function

Private Sub DBErrorSet()
'OO4O Error Setting
    With DBConn
        m_LastServerErr = .Errors.Item(0).Number
        If m_LastServerErr > 0 Then
            m_LastServerErr = m_LastServerErr + 90000
            m_LastServerErrText = .Errors.Item(0).Description
        Else
            .Errors.Clear
        End If
    End With
   '
End Sub


Private Function DBStrConv(ByVal strValue As String, _
                           Optional ByVal optNum As Long) As String
'String Conversion For Database INSERT,UPDATE
    DBStrConv = "'" & CStr(strValue) & "'"
    If IsMissing(optNum) = False Then
        Select Case optNum
            Case 1
                DBStrConv = DBStrConv & ","
            Case 2
                DBStrConv = "=" & DBStrConv
            Case 3
                DBStrConv = "=" & DBStrConv & ","
            Case Else
        End Select
    End If
   '
End Function

Private Function DBDateConv(ByVal DateValue As String, _
                            Optional ByVal optNum As Long) As String
'Date Conversion For Database INSERT,UPDATE
    DBDateConv = "TO_DATE('" & DateValue & "','yy-mm-dd')"
    If IsMissing(optNum) = False Then
        Select Case optNum
            Case 1
                DBDateConv = DBDateConv & ","
            Case 2
                DBDateConv = "=" & DBDateConv
            Case 3
                DBDateConv = "=" & DBDateConv & ","
            Case Else
        End Select
    End If
End Function

Private Function DBNumConv(ByVal NumValue As String, _
                            Optional ByVal optNum As Long) As String
'Number Conversion For Database INSERT,UPDATE
    DBNumConv = CStr(NumValue)
    If DBNumConv = "" Then DBNumConv = "null"
    If IsMissing(optNum) = False Then
        Select Case optNum
            Case 1
                DBNumConv = DBNumConv & ","
            Case 2
                DBNumConv = " = " & DBNumConv
            Case 3
                DBNumConv = " = " & DBNumConv & ","
            Case Else
        End Select
    End If
   '
End Function

Private Sub DBErrClear()
'DB Error Clear
    m_LastServerErr = 0
    m_LastServerErrText = ""
   '
End Sub

Private Sub Class_Initialize()

    DBErrClear
   '
End Sub

Public Function GetLAB001(Optional pTestCd As String = "Default") As Recordset
    
    Dim strSQL As String
    strSQL = " SELECT testcd, testnm, applydt FROM " & T_LAB001 & " a " & _
             " WHERE  applydt IN(SELECT MAX(applydt) FROM " & T_LAB001 & " b " & _
             " WHERE  b.testcd = a.testcd )"
    If pTestCd <> "Default" Then
       strSQL = strSQL & " AND   " & DBW("testcd", pTestCd, 2) & _
                         " ORDER BY testcd ASC"
    Else
       strSQL = strSQL & " ORDER BY restcd ASC"
    End If
    '
    DBOpenDyna strSQL
    Set GetLAB001 = m_RecordSet
   '
End Function

Public Function GetLAB006(Optional ByVal EqpDiv As String = "Default", _
                          Optional ByVal EqpCd As String) As Recordset
    Dim strSQL As String
'Get the list of Equip Master
   '
    If EqpDiv = "Default" Then
        strSQL = "SELECT * FROM " & T_LAB006 & " ORDER BY eqpcd ASC"
    Else
        If EqpCd = "Default" Then
            strSQL = " SELECT * FROM " & T_LAB006 & " WHERE " & DBW("eqpdiv", EqpDiv, 2) & _
                     " ORDER BY eqpcd ASC"
        Else
            strSQL = " SELECT * FROM " & T_LAB006 & _
                     " WHERE  " & DBW("eqpdiv", EqpDiv, 2) & _
                     " AND    " & DBW("eqpcd", EqpCd, 2)
        End If
    End If
    DBOpenDyna strSQL
    Set GetLAB006 = m_RecordSet
   '
End Function

Public Function InsertLAB006(ByVal EqpCd As String, ByVal EqpNm As String, _
                            ByVal EqpDiv As String, ByVal ModelNm As String, ByVal PurchDt As String, _
                            ByVal VandCd As String, ByVal TempLow As Single, ByVal TempHigh As Single, _
                            Remark As String)
    Dim SqlStmt As String
'INSERT INTO LAB006
   '
    SqlStmt = DBV("eqpcd", EqpCd, 1) & _
              DBV("eqpnm", EqpNm, 1) & _
              DBV("eqpdiv", EqpDiv, 1) & _
              DBV("modelnm", ModelNm, 1) & _
              DBV("purchdt", PurchDt, 1) & _
              DBV("vandcd", VandCd, 1) & _
              DBV("templow", TempLow, 1) & _
              DBV("temphigh", TempHigh, 1) & _
              DBV("remark", Remark)
    DBErrClear
    DBExec "INSERT INTO LAB006(eqpcd,eqpnm,eqpdiv,modelnm,purchdt,vandcd,templow," & _
                                "temphigh,remark) VALUES(" & SqlStmt & ")"
   '
End Function

Public Function UpdateLAB006(ByVal EqpCd As String, ByVal EqpNm As String, _
                            ByVal EqpDiv As String, ByVal ModelNm As String, ByVal PurchDt As String, _
                            ByVal VandCd As String, ByVal TempLow As Single, ByVal TempHigh As Single, _
                            Remark As String)
    Dim SqlStmt As String
'UPDATE LAB006 :update lis06 SET modelnm = 'RWC' WHERE eqpcd = 'H020'
   '
   '
    SqlStmt = DBW("eqpnm", EqpNm, 3) & _
              DBW("eqpdiv", EqpDiv, 3) & _
              DBW("modelnm", ModelNm, 3) & _
              DBW("purchdt", PurchDt, 3) & _
              DBW("vandcd", VandCd, 3) & _
              DBW("templow", TempLow, 3) & _
              DBW("temphigh", TempHigh, 3) & _
              DBW("remark", Remark, 2)
    DBErrClear
    DBExec "UPDATE " & T_LAB006 & " SET " & SqlStmt & " WHERE " & DBW("eqpcd", EqpCd, 2)
   '
End Function

Public Function UpdatableLAB006(ByVal EqpCd As String) As Boolean
   '
    UpdatableLAB006 = DBExists("SELECT * FROM " & T_LAB006 & " WHERE " & DBW("eqpcd", EqpCd, 2) & "")
   '
End Function

Public Function GetLAB011(ByVal CtlCd As String, ByVal LotNo As String, _
                          ByVal CtlLevel As String) As Recordset
   '
    DBOpenDyna " SELECT a.ctlcd,a.lotno,a.ctllevel,a.startdt,a.expdt" & _
               "        ,a.eqpcd,a.vandcd,a.qcmtnm,a.remark,b.eqpnm,c.text1" & _
               " FROM   " & T_LAB031 & " c," & T_LAB006 & " b," & T_LAB011 & " a" & _
               " WHERE  " & DBW("a.ctlcd", CtlCd, 2) & _
               " AND    " & DBW("a.lotno", LotNo, 2) & _
               " AND    " & DBW("a.ctllevel", CtlLevel, 2) & _
               " AND    a.eqpcd = b.eqpcd" & _
               " AND    " & DBW("c.cdindex", "50", 2) & _
               " AND    " & DBW("c.cdval1", "E", 2) & _
               " AND    a.vandcd = c.cdval2"

    '
    Set GetLAB011 = m_RecordSet
   '
End Function

Public Function GetStartDtLAB011(ByVal CtlCd As String, ByVal LotNo As String, _
                                 ByVal CtlLevel As String, ByVal StartDt As String) As String
   '
    DBOpenDyna " SELECT MAX(StartDt) LASTDT FROM " & T_LAB011 & _
               " WHERE  " & DBW("ctlcd", CtlCd, 2) & _
               " AND    " & DBW("ctllevel", CtlLevel, 2)
    '
    If m_RecordSet.RecordCount < 1 Then
        GetStartDtLAB011 = ""
    Else
        GetStartDtLAB011 = "" & m_RecordSet.Fields("LASTDT")
    End If
   '
End Function

Public Function InsertLAB011(ByVal CtlCd As String, ByVal LotNo As String, _
                            ByVal CtlLevel As String, ByVal StartDt As String, ByVal ExpDt As String, _
                            EqpCd As String, VandCd As String, ByVal QcMtNm As String, _
                            ByVal Remark As String, ByVal LAB012String As String)
    Dim SqlStmt() As String
    Dim arySqlBuf() As String
    Dim arySql() As String
    Dim ii As Long
   '
    DBErrClear
    ReDim SqlStmt(0)
    SqlStmt(0) = DBV("ctlcd", CtlCd, 1) & _
                 DBV("lotno", LotNo, 1) & _
                 DBV("ctllevel", CtlLevel, 1) & _
                 DBV("startdt", StartDt, 1) & _
                 DBV("expdt", ExpDt, 1) & _
                 DBV("eqpcd", EqpCd, 1) & _
                 DBV("vandcd", VandCd, 1) & _
                 DBV("qcmtnm", QcMtNm, 1) & _
                 DBV("remark", Remark)
    SqlStmt(0) = "INSERT INTO LAB011(ctlcd,lotno,ctllevel,startdt,expdt,eqpcd" & _
       ",vandcd,qcmtnm,remark) VALUES(" & SqlStmt(0) & ")"
    '
    arySqlBuf = Split(LAB012String, Chr(13))
    ReDim Preserve SqlStmt(UBound(arySqlBuf))
   '
    For ii = 1 To (UBound(arySqlBuf))
        arySql = Split(arySqlBuf(ii - 1), Chr(9))
        'arySql(12):testcd, arySql(1):MeanVal,arySql(2):sdval,arySql(3):AvalVal
        'arySql(4):fromval,arySql(5):toval,arySql(6):txtval,arySql(7):RstUnit
        'arySql(8):cvval,arySql(9):minval,arySql(10):MaxVal
        SqlStmt(ii) = DBV("ctlcd", CtlCd, 1) & _
                      DBV("lotno", LotNo, 1) & _
                      DBV("ctllevel", CtlLevel, 1) & _
                      DBV("testcd", arySql(12), 1) & _
                      DBV("meanval", arySql(1), 1) & _
                      DBV("sdval", arySql(2), 1) & _
                      DBV("avalval", arySql(3), 1) & _
                      DBV("fromval", arySql(4), 1) & _
                      DBV("toval", arySql(5), 1) & _
                      DBV("txtval", arySql(6), 1) & _
                      DBV("rstunit", arySql(7), 1) & _
                      DBV("cvval", arySql(8), 1) & _
                      DBV("minval", arySql(9), 1) & _
                      DBV("maxval", arySql(10))
        SqlStmt(ii) = "INSERT INTO LAB012(ctlcd,lotno,ctllevel,testcd, meanval," & _
                    "sdval,avalval,fromval,toval,txtval,rstunit,cvval,minval,maxval" & _
                    ") VALUES(" & SqlStmt(ii) & ")"
    Next
    '
    DBMultiExec SqlStmt
   '
End Function

Public Function UpdateLAB011(ByVal CtlCd As String, ByVal LotNo As String, _
                            ByVal CtlLevel As String, ByVal StartDt As String, ByVal ExpDt As String, _
                            EqpCd As String, VandCd As String, ByVal QcMtNm As String, _
                            ByVal Remark As String, ByVal LAB012String As String)
    Dim SqlStmt() As String
    Dim arySqlBuf() As String
    Dim arySql() As String
    Dim ii As Integer
   '
    DBErrClear
    ReDim SqlStmt(0)
    SqlStmt(0) = DBW("StartDt", StartDt, 3) & _
                 DBW("ExpDt", ExpDt, 3) & _
                 DBW("EqpCd", EqpCd, 3) & _
                 DBW("VandCd", VandCd, 3) & _
                 DBW("QcMtNm", QcMtNm, 3) & _
                 DBW("Remark", Remark, 2)
    '
    SqlStmt(0) = " UPDATE LAB011 SET " & SqlStmt(0) & _
                 " WHERE  " & DBW("ctlcd", CtlCd, 2) & _
                 " AND    " & DBW("lotno", LotNo, 2) & _
                 " AND    " & DBW("ctllevel", CtlLevel, 2)
   '
    arySqlBuf = Split(LAB012String, Chr(13))
    ReDim Preserve SqlStmt(UBound(arySqlBuf))
    '
    For ii = 1 To (UBound(arySqlBuf))
        arySql = Split(arySqlBuf(ii - 1), Chr(9))
        'arySql(12):testcd, arySql(1):MeanVal,arySql(2):sdval,arySql(3):AvalVal
        'arySql(4):fromval,arySql(5):toval,arySql(6):txtval,arySql(7):RstUnit
        'arySql(8):cvval,arySql(9):minval,arySql(10):MaxVal
        SqlStmt(ii) = DBW("meanval", arySql(1), 3) & _
                      DBW("sdval", arySql(2), 3) & _
                      DBW("avalval", arySql(3), 3) & _
                      DBW("fromval", arySql(4), 3) & _
                      DBW("toval", arySql(5), 3) & _
                      DBW("txtval", arySql(6), 3) & _
                      DBW("rstunit", arySql(7), 3) & _
                      DBW("cvval", arySql(8), 3) & _
                      DBW("minval", arySql(9), 3) & _
                      DBW("maxval", arySql(10), 2)
        
        SqlStmt(ii) = " UPDATE LAB012 SET " & SqlStmt(ii) & _
                      " WHERE  " & DBW("ctlcd", CtlCd, 2) & _
                      " AND    " & DBW("lotno", LotNo, 2) & _
                      " AND    " & DBW("ctllevel", CtlLevel, 2) & _
                      " AND    " & DBW("testcd", arySql(12), 2)
    Next
   '
    DBMultiExec SqlStmt
   '
End Function

Public Function UpdatableLAB011(ByVal CtlCd As String, _
                                ByVal LotNo As String, ByVal CtlLevel As String) As Boolean
   '
    UpdatableLAB011 = DBExists(" SELECT * FROM " & T_LAB011 & _
                               " WHERE  " & DBW("ctlcd", CtlCd, 2) & _
                               " AND    " & DBW("lotno", LotNo, 2) & _
                               " AND    " & DBW("ctllevel", CtlLevel, 2) & "")
   '
End Function

Public Function GetLAB012QCItem(ByVal CtlCd As String, ByVal LotNo As String, _
                                ByVal CtlLevel As String) As Recordset
   '
    DBOpenDyna "SELECT c.testcd, c.meanval,c.sdval,c.avalval,c.fromval," & _
               "       c.toval,c.txtval,c.rstunit,c.cvval,c.minval,c.maxval," & _
               "       b.field1,a.testnm" & _
               " FROM  " & T_LAB031 & " b," & T_LAB001 & " a," & T_LAB012 & " c" & _
               " WHERE " & DBW("c.ctlcd", CtlCd, 2) & _
               " AND   " & DBW("c.lotno", LotNo, 2) & _
               " AND   " & DBW("c.ctllevel", CtlLevel, 2) & _
               " AND   c.testcd =  a.testcd" & _
               " AND   b.cdindex = " & LC2_QcControl & _
               " AND   b.cdval1 =  c.ctlcd" & _
               " AND   b.cdval2 =  c.testcd"
    '
    Set GetLAB012QCItem = m_RecordSet
   '
End Function

Public Function GetLAB031QCItem(ByVal CtlCd As String, ByVal LotNo As String, _
                                ByVal CtlLevel As String) As Recordset
   '
    DBOpenDyna " SELECT b.cdval2,b.field1,a.testnm" & _
               " FROM   " & T_LAB001 & " a," & T_LAB031 & " b" & _
               " WHERE  " & DBW("b.cdindex", LC2_QcControl, 2) & _
               " AND    " & DBW("b.cdval1", CtlCd, 2) & _
               " AND    b.cdval2 = a.testcd"
    '
    Set GetLAB031QCItem = m_RecordSet
   '
End Function

Public Function GetLAB014(Optional pCtlCd As String = "Default") As Recordset
    Dim strSQL As String
    If pCtlCd = "Default" Then
       strSQL = "SELECT * FROM " & T_LAB014 & " ORDER BY ctlcd ASC"
    Else
       strSQL = "SELECT * FROM " & T_LAB014 & " WHERE " & DBW("ctlcd", pCtlCd, 2)
    End If
    '
    DBOpenDyna strSQL
    Set GetLAB014 = m_RecordSet
    '
End Function

Public Function GetLAB014QcTest(ByVal pCtlCd As String) As Recordset
    
    Dim strSQL As String
    strSQL = " SELECT b.cdval2,b.field1,a.testnm" & _
             " FROM   " & T_LAB001 & " a," & T_LAB031 & " b  " & _
             " WHERE  " & DBW("b.cdindex", "10", 2) & _
             " AND    " & DBW("b.cdval1", pCtlCd, 2) & _
             " AND    b.cdval2 = a.testcd" & _
             " AND    a.applydt IN( SELECT MAX(applydt)" & _
                                  " FROM  " & T_LAB001 & " c" & _
                                  " WHERE a.testcd = c.testcd" & " )"
    '
    DBOpenDyna strSQL
    Set GetLAB014QcTest = m_RecordSet
   '
End Function






'Public Function InsertLAB014(ByVal CtlCd As String, ByVal CtlNm As String, _
'                            ByVal AutoFg As String, ByVal Duration As String, ByVal AutoDay As String, _
'                            ByVal AutoTime As String, AutoText As String, QcItems As String)
'    Dim SqlStmt() As String
'    Dim aryTmp() As String
'    Dim arySql() As String
'    Dim strTmp As String
'    Dim ii As Long
'    Dim lngCnt As Long
'    Dim lngSqlCnt As Long
'    Dim strQcItems As String
'   '
'    ReDim SqlStmt(0)
'    lngSqlCnt = 0
'    SqlStmt(0) = "INSERT INTO LAB014(ctlcd,ctlnm,duration,autofg,autoday,autotime," & _
'                 "autotext) VALUES(" & _
'                 DBV("ctlcd", CtlCd, 1) & _
'                 DBV("ctlnm", CtlNm, 1) & _
'                 DBV("duration", AutoFg, 1) & _
'                 DBV("autofg", Duration, 1) & _
'                 DBV("autoday", AutoDay, 1) & _
'                 DBV("autotime", AutoTime, 1) & _
'                 DBV("autotext", AutoText) & ")"
'    lngCnt = -1
'    strQcItems = GetQcItems(CtlCd)
'    aryTmp = Split(strQcItems, vbNewLine)
'    For ii = LBound(aryTmp) To UBound(aryTmp) - 1
'        If (QcItems Like "*" & medGetP(aryTmp(ii), 1, vbTab) & "*") = False Then
'            lngCnt = lngCnt + 1
'            ReDim Preserve arySql(lngCnt)
'            arySql(lngCnt) = medGetP(aryTmp(ii), 1, vbTab) & vbTab & vbTab & "D"
'        End If
'    Next ii
'    aryTmp = Split(QcItems, vbNewLine)
'    For ii = LBound(aryTmp) To UBound(aryTmp) - 1
'        lngCnt = lngCnt + 1
'        ReDim Preserve arySql(lngCnt)
'        arySql(lngCnt) = medGetP(aryTmp(ii), 1, vbTab) & vbTab & _
'                        medGetP(aryTmp(ii), 3, vbTab) & vbTab & _
'                        medGetP(aryTmp(ii), 4, vbTab)
'    Next
'    If lngCnt > -1 Then
'        For ii = LBound(arySql) To UBound(arySql)
'            lngSqlCnt = lngSqlCnt + 1
'            SELECT Case medGetP(arySql(ii), 3, vbTab)
'                Case "D"
'                    strTmp = " DELETE " & T_LAB031 & _
'                             " WHERE  " & DBW("cdindex", LC2_QcControl, 2) & _
'                             " AND    " & DBW("cdval1", CtlCd, 2) & _
'                             " AND    " & DBW("cdval2", medGetP(arySql(ii), 1, vbTab), 2)
'                Case "U"
'                    strTmp = " UPDATE " & T_LAB031 & _
'                             " SET    " & _
'                                          DBW("Field1", medGetP(arySql(ii), 2, vbTab), 2) & _
'                             " WHERE  " & DBW("cdindex", LC2_QcControl, 2) & _
'                             " AND    " & DBW("cdval1", CtlCd, 2) & _
'                             " AND    " & DBW("cdval2", medGetP(arySql(ii), 1, vbTab), 2)
'                Case "I"
'                    strTmp = " INSERT INTO " & T_LAB031 & " (cdindex,cdval1,cdval2,field1)" & _
'                             " VALUES( " & _
'                             DBV("cdindex", LC2_QcControl, 1) & _
'                             DBV("cdval1", CtlCd, 1) & _
'                             DBV("cdval2", medGetP(arySql(ii), 1, vbTab), 1) & _
'                             DBV("field1", medGetP(arySql(ii), 2, vbTab)) & ")"
'            End SELECT
'            ReDim Preserve SqlStmt(lngSqlCnt)
'            SqlStmt(lngSqlCnt) = strTmp
'        Next ii
'    End If
'    DBErrClear
'   '
'    DBMultiExec SqlStmt
'   '
'End Function

'Public Function UpdateLAB014(ByVal CtlCd As String, ByVal CtlNm As String, _
'                            ByVal AutoFg As String, ByVal Duration As String, ByVal AutoDay As String, _
'                            ByVal AutoTime As String, ByVal AutoText As String, ByVal QcItems As String)
'    Dim SqlStmt() As String
'    Dim aryTmp() As String
'    Dim arySql() As String
'    Dim strTmp As String
'    Dim ii As Long
'    Dim lngCnt As Long
'    Dim lngSqlCnt As Long
'    Dim strQcItems As String
'   '
'    ReDim SqlStmt(0)
'    lngSqlCnt = 0
'    SqlStmt(0) = " UPDATE " & T_LAB014 & _
'                 " SET " & _
'                              DBW("ctlnm", CtlNm, 3) & _
'                              DBW("autofg", AutoFg, 3) & _
'                              DBW("duration", Duration, 3) & _
'                              DBW("autoday", AutoDay, 3) & _
'                              DBW("autotime", AutoTime, 3) & _
'                              DBW("autotext", AutoText, 2) & _
'                 " WHERE  " & DBW("ctlcd", CtlCd, 2)
'   '
'    lngCnt = -1
'    strQcItems = GetQcItems(CtlCd)
'    aryTmp = Split(strQcItems, vbNewLine)
'    For ii = LBound(aryTmp) To UBound(aryTmp) - 1
'        If (QcItems Like "*" & medGetP(aryTmp(ii), 1, vbTab) & "*") = False Then
'            lngCnt = lngCnt + 1
'            ReDim Preserve arySql(lngCnt)
'            arySql(lngCnt) = medGetP(aryTmp(ii), 1, vbTab) & vbTab & vbTab & "D"
'        End If
'    Next ii
'    aryTmp = Split(QcItems, vbNewLine)
'    For ii = LBound(aryTmp) To UBound(aryTmp) - 1
'        lngCnt = lngCnt + 1
'        ReDim Preserve arySql(lngCnt)
'        arySql(lngCnt) = medGetP(aryTmp(ii), 1, vbTab) & vbTab & _
'                        medGetP(aryTmp(ii), 3, vbTab) & vbTab & _
'                        medGetP(aryTmp(ii), 4, vbTab)
'    Next
'    If lngCnt > -1 Then
'        For ii = LBound(arySql) To UBound(arySql)
'            lngSqlCnt = lngSqlCnt + 1
'            SELECT Case medGetP(arySql(ii), 3, vbTab)
'                Case "D"
'                    strTmp = " DELETE " & T_LAB031 & _
'                             " WHERE  " & DBW("cdindex", "10", 2) & _
'                             " AND    " & DBW("cdval1", CtlCd, 2) & _
'                             " AND    " & DBW("cdval2", medGetP(arySql(ii), 1, vbTab), 2)
'                Case "U"
'                    strTmp = " UPDATE " & T_LAB031 & _
'                             " SET " & _
'                                      " Field1" & DBStrConv(medGetP(arySql(ii), 2, vbTab), 2) & _
'                             " WHERE  " & DBW("cdindex", LC2_QcControl, 2) & _
'                             " AND    " & DBW("cdval1", CtlCd, 2) & _
'                             " AND    " & DBW("cdval2", medGetP(arySql(ii), 1, vbTab), 2)
'                Case "I"
'                    strTmp = " INSERT INTO " & T_LAB031 & " (cdindex,cdval1,cdval2,field1)" & _
'                             " VALUES(" & _
'                             DBV("cdindex", LC2_QcControl, 1) & _
'                             DBV("cdval1", CtlCd, 1) & _
'                             DBV("cdval2", medGetP(arySql(ii), 1, vbTab), 1) & _
'                             DBV("field1", medGetP(arySql(ii), 2, vbTab)) & ")"
'            End SELECT
'            ReDim Preserve SqlStmt(lngSqlCnt)
'            SqlStmt(lngSqlCnt) = strTmp
'        Next ii
'    End If
'    DBErrClear
'   '
'    DBMultiExec SqlStmt
'   '
'End Function

'Public Function GetQcItems(ByVal CtlCd As String) As String
'    Dim objLab014 As New clsQCControls
'    With objLab014
'        GetQcItems = .GetQcItems(OraSession, dbconn, CtlCd)        'QC컨트롤 ITEM
'    End With
'End Function

'Public Function UpdatableLAB014(ByVal CtlCd As String) As Boolean
'   '
'    UpdatableLAB014 = DBExists("SELECT * FROM " & T_LAB014 & " WHERE " & DBW("ctlcd", CtlCd, 2) & "")
'   '
'End Function

Public Sub GetLAB031(ByRef pOraDynaset As Recordset, _
                    ByVal cdindex As String, ByVal CdVal1 As String, _
                    ByVal CdVal2 As String)
    Dim strSQL As String
   '
    strSQL = "SELECT * FROM " & T_LAB031 & " WHERE " & DBW("cdindex", cdindex, 2)
    If CdVal1 <> "" Then
        strSQL = strSQL & " AND " & DBW("cdval1", CdVal1, 2)
        If CdVal2 <> "" Then
            strSQL = strSQL & " AND " & DBW("cdval2", CdVal2, 2)
        End If
    End If
    DBOpenDyna strSQL
    Set pOraDynaset = m_RecordSet
   '
End Sub
 
Public Function BuildWorkSheet(ByVal WorkDt As String, _
                                ByVal WorkCd As String, ByVal WorkTm As String, _
                                ByVal StartWorkSeq As Long, ByVal WorkId As String, _
                                ByVal BuildList As String, _
                                ByVal pLastDt As String, ByVal pLastTm As String, _
                                Optional ByRef pPrgBar As Object)

    Dim SqlStmt()       As String
    Dim aryTmp()        As String
    Dim aryTmp1()       As String
    Dim aryOrd()        As String
    Dim strSQL          As String
    Dim tmpWorkSeq      As Long
    Dim strLastWorkSeq  As String
    Dim intSqlCnt       As Integer
    Dim ii As Long
    Dim jj As Long
    
    
   '
    m_SaveOK = False
    'DBOpenDyna strSql
    aryTmp = Split(BuildList, vbTab)
    ReDim SqlStmt(UBound(aryTmp))
    '/*워크쉬트내역 생성
    On Error GoTo DBExecError
   '
    strLastWorkSeq = CStr(GetWSLastSeq(WorkCd, WorkDt))
    If strLastWorkSeq = "1" Then

        strSQL = " INSERT INTO " & T_COM099 & "(noindex,divcd1,divcd2,divcd3,seq,remark) " & _
                 " VALUES( " & _
                 DBV("noindex", COM99_LIS_WorkNo, 1) & _
                 DBV("divcd1", "0", 1) & _
                 DBV("divcd2", WorkCd, 1) & _
                 DBV("divcd3", WorkDt, 1) & _
                 DBV("seq", "0", 1) & _
                 DBV("remark", "") & ")"
        DBExec strSQL
    End If
    StartWorkSeq = Val(strLastWorkSeq)
   
    pPrgBar.Max = (UBound(aryTmp) + 1) * 3
    pPrgBar.Value = 0
    DoEvents
   
    For ii = 0 To UBound(aryTmp)
        
        aryTmp1 = Split(aryTmp(ii), COL_DIV)
        
        tmpWorkSeq = Val(aryTmp1(1))
        If tmpWorkSeq = 0 Then
            tmpWorkSeq = StartWorkSeq
            StartWorkSeq = StartWorkSeq + 1
        End If
            
        SqlStmt(ii) = " INSERT INTO " & T_LAB301 & " (workdt,workcd,workseq,worktm,workid," & _
                      " workarea,accdt,accseq) VALUES(" & _
                      DBV("workdt", WorkDt, 1) & _
                      DBV("workcd", WorkCd, 1) & _
                      DBV("workseq", tmpWorkSeq, 1) & _
                      DBV("worktm", WorkTm, 1) & _
                      DBV("workid", WorkId, 1) & _
                      DBV("workarea", medGetP(aryTmp1(0), 1, "-"), 1) & _
                      DBV("accdt", medGetP(aryTmp1(0), 2, "-"), 1) & _
                      DBV("accseq", medGetP(aryTmp1(0), 3, "-")) & ")"
    Next ii
    
    '/* 번호부여정보 UPDATE
    ReDim Preserve SqlStmt(UBound(SqlStmt) + 1)
    
    SqlStmt(UBound(SqlStmt)) = " UPDATE " & T_COM099 & _
                               " SET    " & _
                                            DBW("seq", (StartWorkSeq - 1), 2) & _
                               " WHERE  " & DBW("noindex", COM99_LIS_WorkNo, 2) & _
                               " AND    " & DBW("divcd1", "0", 2) & _
                               " AND    " & DBW("divcd2", WorkCd, 2) & _
                                " AND    " & DBW("divcd3", WorkDt, 2)

    
    '/* Worksheet Name 내역에 마감일시 UPDATE
    ReDim Preserve SqlStmt(UBound(SqlStmt) + 1)
    SqlStmt(UBound(SqlStmt)) = " UPDATE " & T_LAB032 & _
                               " SET    " & _
                                            DBW("field3", pLastDt, 3) & _
                                            DBW("field4", pLastTm, 2) & _
                               " WHERE  " & DBW("cdindex", LC3_WorkSheetName, 2) & _
                               " AND    " & DBW("cdval1", WorkCd, 2)
    
    For ii = 0 To UBound(aryTmp)
       
        If pPrgBar.Value < pPrgBar.Max Then pPrgBar.Value = pPrgBar.Value + 1
        DoEvents
        
        aryTmp1 = Split(aryTmp(ii), COL_DIV)
        
        '/* 접수내역 Update
        ReDim Preserve SqlStmt(UBound(SqlStmt) + 1)
        strSQL = "UPDATE " & T_LAB201 & " SET " & _
                                   DBW("stscd", enStsCd.StsCd_LIS_InProcess, 2) & _
                 " WHERE " & _
                                   DBW("workarea = ", medGetP(aryTmp1(0), 1, "-")) & _
                         " AND " & DBW("accdt    = ", medGetP(aryTmp1(0), 2, "-")) & _
                         " AND " & DBW("accseq   = ", medGetP(aryTmp1(0), 3, "-")) & _
                         " AND " & DBW("stscd    = ", enStsCd.StsCd_LIS_Accession)
        SqlStmt(UBound(SqlStmt)) = strSQL
        
        
        '/* 결과내역 Update - 웤슅코드

        ReDim Preserve SqlStmt(UBound(SqlStmt) + 1)
        strSQL = "UPDATE " & T_LAB302 & " SET " & _
                                   DBW("workcd", WorkCd, 2) & _
                 " WHERE " & _
                                   DBW("workarea = ", medGetP(aryTmp1(0), 1, "-")) & _
                         " AND " & DBW("accdt    = ", medGetP(aryTmp1(0), 2, "-")) & _
                         " AND " & DBW("accseq   = ", medGetP(aryTmp1(0), 3, "-")) & _
                         " AND EXISTS (SELECT * FROM " & T_LAB008 & " a " & _
                                     " WHERE " & DBW("a.workcd = ", WorkCd) & _
                                     " AND   a.testcd = " & T_LAB302 & ".testcd " & _
                                     " AND   a.spccd  = " & T_LAB302 & ".spccd) "
        If aryTmp1(2) = "" Then '추가검체가 아니면,,,
            strSQL = strSQL & " AND (workcd = ''  or  workcd is null) "
        End If
                     
        SqlStmt(UBound(SqlStmt)) = strSQL
        
    Next ii
    '/*처방 Status Update
    For ii = 0 To UBound(aryTmp)
       
        If pPrgBar.Value < pPrgBar.Max Then pPrgBar.Value = pPrgBar.Value + 1
        DoEvents
        
        aryTmp1 = Split(aryTmp(ii), COL_DIV)
        aryOrd = Split(GetOrderForWSBuild(aryTmp1(0), WorkCd), vbTab)
        For jj = 0 To UBound(aryOrd)
            ReDim Preserve SqlStmt(UBound(SqlStmt) + 1)
            SqlStmt(UBound(SqlStmt)) = aryOrd(jj)
        Next
    Next ii
    '
    DBErrClear
    '
    'DBMultiExec SqlStmt
    DBConn.BeginTrans
    For intSqlCnt = LBound(SqlStmt) To UBound(SqlStmt)
        If pPrgBar.Value < pPrgBar.Max Then pPrgBar.Value = pPrgBar.Value + 1
        DoEvents
'        Debug.Print SqlStmt(intSqlCnt)
        DBConn.Execute (SqlStmt(intSqlCnt))
    Next intSqlCnt
    DBConn.CommitTrans
    '
    m_SaveOK = True
    Exit Function
DBExecError:
'    Call Error_Routine
    MsgBox "처리 도중 오류가 발생하였습니다.", vbExclamation
    
End Function
Public Function SetWorkSheetDelete(ByVal sWorkCd As String, ByVal sWorkDt As String, ByVal sWorkSeq As String) As Boolean
    Dim sSQL        As String
    Dim aryTmp()    As String
    Dim ii          As Integer
    
    aryTmp = Split(sWorkSeq, COL_DIV)
    
    On Error GoTo SAVE_ERROR
    DBConn.BeginTrans
    For ii = 0 To UBound(aryTmp())
        sSQL = " delete " & T_LAB301 & _
               " WHERE " & DBW("workdt=", sWorkDt) & _
               " AND " & DBW("workcd=", sWorkCd) & _
               " AND " & DBW("workseq=", aryTmp(ii))
        
        DBConn.Execute sSQL
    Next
    DBConn.CommitTrans
    SetWorkSheetDelete = True
    Exit Function
SAVE_ERROR:
    DBConn.RollbackTrans
End Function

Public Function GetOrderForWSBuild(ByVal AccNo As String, _
                                    WorkCd As String) As String
    Dim strSQL As String
    Dim strWorkArea As String
    Dim strAccDt As String
    Dim strAccSeq As String
    Dim strWrkDiv As String
    Dim arySql() As String
    Dim ii As Integer
   '
    strWorkArea = medGetP(AccNo, 1, "-")
    strAccDt = medGetP(AccNo, 2, "-")
    strAccSeq = medGetP(AccNo, 3, "-")
    GetOrderForWSBuild = ""
   '
    strSQL = " SELECT DISTINCT a.ptid,a.orddt,a.ordno,a.ordseq, c.wrkdiv" & _
             " FROM " & T_LAB102 & " c," & T_LAB201 & " b," & T_LAB302 & " a" & _
             " WHERE " & _
                     DBW("a.workarea = ", strWorkArea) & _
             " AND " & DBW("a.accdt = ", strAccDt) & _
             " AND " & DBW("a.accseq = ", strAccSeq) & _
             " AND (a.vfydt = ''  or  a.vfydt is null) " & _
             " AND c.ptid = a.ptid" & _
             " AND c.orddt = a.orddt" & _
             " AND c.ordno = a.ordno" & _
             " AND c.ordseq = a.ordseq" & _
             " AND " & DBW("c.stscd = ", enStsCd.StsCd_LIS_Accession) & _
             " AND b.workarea =  a.workarea" & _
             " AND b.accdt =  a.accdt" & _
             " AND b.accseq = a.accseq" & _
             " AND EXISTS(SELECT *" & _
                        " FROM " & T_LAB008 & " d WHERE " & DBW("d.workcd = ", WorkCd) & _
                        " AND d.testcd = a.testcd" & _
                        " AND d.spccd = b.spccd" & _
                        " AND d.workarea = b.workarea)"
   '
    DBOpenDyna strSQL
    With m_RecordSet
        If .RecordCount > 0 Then
            ReDim arySql(0): ii = 0
            .MoveFirst
            Do Until .EOF
                 If ii <> 0 Then
                    ReDim Preserve arySql(ii)
                 End If
                 
                 '** 원본 ==========================================================================
'                 arySql(ii) = "UPDATE " & T_LAB102 & " SET " & DBW("stscd =", enStsCd.StsCd_LIS_InProcess) & _
'                              " WHERE " & _
'                                        DBW("ptid  = ", "" & m_RecordSet.Fields("PTID").Value) & _
'                              " AND " & DBW("orddt = ", "" & m_RecordSet.Fields("ORDDT").Value) & _
'                              " AND " & DBW("ordno = ", "" & m_RecordSet.Fields("ORDNO").Value) & _
'                              " AND " & DBW("ordseq= ", "" & m_RecordSet.Fields("ORDSEQ").Value)
                '===================================================================================
                
                '** 전주 예수 병원 변경 루틴========================================================
                strWrkDiv = m_RecordSet.Fields("wrkdiv").Value & ""
                
                If strWrkDiv = "3" Then         '종건
                    arySql(ii) = "UPDATE su2examt SET " & DBW("stscd =", enStsCd.StsCd_LIS_InProcess) & _
                                 " WHERE patno = " & DBS(m_RecordSet.Fields("PTID").Value) & _
                                 " AND orddate = TO_DATE(" & DBS(m_RecordSet.Fields("ORDDT").Value) & ", 'yyyymmdd')" & _
                                 " AND ordseqno = " & DBN(m_RecordSet.Fields("ORDNO").Value)
                ElseIf strWrkDiv = "4" Then     '일건
                        arySql(ii) = "UPDATE sg2examt SET " & DBW("stscd =", enStsCd.StsCd_LIS_InProcess) & _
                                     " WHERE patno = " & DBS(m_RecordSet.Fields("PTID").Value) & _
                                     " AND orddate = TO_DATE(" & DBS(m_RecordSet.Fields("ORDDT").Value) & ", 'yyyymmdd')" & _
                                     " AND ordseqno = " & DBN(m_RecordSet.Fields("ORDNO").Value)
                Else
                    arySql(ii) = "UPDATE mdexmort SET " & DBW("stscd =", enStsCd.StsCd_LIS_InProcess) & _
                                 " WHERE patno = " & DBS(m_RecordSet.Fields("PTID").Value) & _
                                 " AND orddate = TO_DATE(" & DBS(m_RecordSet.Fields("ORDDT").Value) & ", 'yyyymmdd')" & _
                                 " AND ordseqno = " & DBN(m_RecordSet.Fields("ORDNO").Value)
                End If
                '===================================================================================
                
                 .MoveNext
                ii = ii + 1
            Loop
            GetOrderForWSBuild = Join(arySql, vbTab)
        End If
    End With
   '
End Function
  
Public Function GetRstCd(ByVal cdindex As String, ByVal CdVal1 As String, _
                         ByVal CdVal2 As String) As Recordset
    Dim SqlStmt As String
   '
    SqlStmt = " SELECT * FROM " & T_LAB031 & _
              " WHERE " & DBW("cdindex = ", cdindex) & _
              " AND   " & DBW("cdval1  = ", CdVal1) & _
              " AND   " & DBW("cdval2  = ", CdVal2)
    Set GetRstCd = New Recordset
    GetRstCd.Open SqlStmt, DBConn
End Function
     
Public Function UpdatableLAB031(ByVal cdindex As String, ByVal CdVal1 As String, _
                                ByVal CdVal2 As String) As Boolean
    Dim SqlStmt As String
   '
    SqlStmt = " SELECT * FROM " & T_LAB031 & _
              " WHERE " & DBW("cdindex = ", cdindex) & _
              " AND   " & DBW("cdval1  = ", CdVal1) & _
              " AND   " & DBW("cdval2  = ", CdVal2)
              
    UpdatableLAB031 = DBExists(SqlStmt)
   '
End Function
     
Public Sub GetLAB032(ByRef pOraDynaset As Recordset, _
                    ByVal CdIndex_A As String, ByVal CdIndex_B As String, _
                    ByVal CdVal1 As String)
    Dim strSQL As String
   '
    strSQL = "SELECT * FROM " & T_LAB032 & " WHERE"
    If CdIndex_B <> "" Then
        strSQL = strSQL & " cdindex in(" & DBV("cdindex", CdIndex_A) & "," & DBV("cdindex", CdIndex_B) & ")"
    Else
        strSQL = strSQL & " " & DBW("cdindex = ", CdIndex_A)
    End If
    If CdVal1 <> "" Then
        strSQL = strSQL & " AND " & DBW("cdval1 = ", CdVal1)
    End If
    DBOpenDyna strSQL
    Set pOraDynaset = m_RecordSet
   '
End Sub
 
Public Function UpdatableLAB032(ByVal cdindex As String, _
                                ByVal CdVal1 As String) As Boolean
    Dim SqlStmt As String
   '
    SqlStmt = " SELECT * FROM " & T_LAB032 & _
              " WHERE " & DBW("cdindex =", cdindex) & _
              " AND   " & DBW("cdval1 =", CdVal1)
              
    UpdatableLAB032 = DBExists(SqlStmt)
   '
End Function
     
Public Sub GetLAB034(ByRef pOraDynaset As Recordset, _
                    ByVal CdIndex_A As String, ByVal CdIndex_B As String, _
                    ByVal CdVal1 As String)
    Dim strSQL As String
   '
    strSQL = "SELECT * FROM " & T_LAB034 & " WHERE "
    If CdIndex_B <> "" Then
        strSQL = strSQL & " cdindex in(" & DBV("cdindex", CdIndex_A) & "," & DBV("cdindex", CdIndex_B) & ")"
    Else
        strSQL = strSQL & " " & DBW("cdindex = ", CdIndex_A)
    End If
    If CdVal1 <> "" Then
        strSQL = strSQL & " AND " & DBW("cdval1 = ", CdVal1)
    End If
    strSQL = strSQL & " order by cdval1 "
     
    DBOpenDyna strSQL
    Set pOraDynaset = m_RecordSet
   '
End Sub
     
Public Sub GetLAB301(ByRef pOraDynaset As Recordset, _
                     ByVal WorkDt As String, ByVal WorkCd As String, _
                     Optional FrWorkSeq As String = "Default", _
                     Optional ToWorkSeq As String = "Default", _
                     Optional pStatFg As String = "0")
    Dim strSQL As String

    '--2000.12.4 김미경 수정
    '  QC Data를 가져오기 위해 h7lab026을 UNION 걸어준다.
    '
    strSQL = " SELECT  a.workdt, a.workcd, a.workseq, a.worktm, a.workid, " & _
             "         a.workarea, a.accdt, a.accseq, b.statfg, b.rcvdt, b.rcvtm " & _
             " FROM  " & T_LAB201 & " b," & T_LAB301 & " a " & _
             " WHERE " & DBW("a.workdt = ", WorkDt) & _
             " AND   " & DBW("a.workcd = ", WorkCd)
   
    If FrWorkSeq <> "Default" Then
        strSQL = strSQL & " AND a.workseq BETWEEN " & DBV("workseq", FrWorkSeq) & " AND " & DBV("workseq", ToWorkSeq)
    End If
   
    strSQL = strSQL & " AND   b.workarea = a.workarea" & _
                      " AND   b.accdt    = a.accdt" & _
                      " AND   b.accseq   = a.accseq" & _
                      " AND   b.stscd IN (" & DBV("stscd", enStsCd.StsCd_LIS_InProcess) & "," & _
                                              DBV("stscd", enStsCd.StsCd_LIS_MidRst) & ")" & _
                      " AND   EXISTS(SELECT c.testcd FROM " & T_LAB302 & " c, " & T_LAB008 & " d " & _
                      " WHERE c.workarea = a.workarea" & _
                      " AND   c.accdt    = a.accdt" & _
                      " AND   c.accseq   = a.accseq" & _
                      " AND  (c.vfydt = '' or c.vfydt is null) " & _
                      " AND " & DBW("d.workcd = ", WorkCd) & _
                      " AND   d.testcd = c.testcd "
    strSQL = strSQL & " UNION " & _
                        " SELECT c.testcd " & _
                        "   FROM   " & T_LAB008 & " d, " & T_LAB351 & " c " & _
                        " WHERE  c.workarea = a.workarea" & _
                        " AND    c.accdt    = a.accdt" & _
                        " AND    c.accseq   = a.accseq" & _
                        " AND    " & DBW("d.workcd", WorkCd, 2) & _
                        " AND    c.testcd   = d.testcd  " & _
                        " AND    " & DBW("c.stscd", enStsCd.StsCd_LIS_Accession, 2) & " )"
   
    strSQL = strSQL & " UNION SELECT a.workdt, a.workcd, a.workseq, a.worktm, a.workid, " & _
                      "              a.workarea, a.accdt, a.accseq, b.statfg, b.rcvdt, b.rcvtm " & _
                      " FROM  " & T_LAB201 & " b," & T_LAB301 & " a " & _
                      " WHERE " & DBW("a.workdt = ", WorkDt) & _
                      " AND   " & DBW("a.workcd = ", WorkCd)
   
    If FrWorkSeq <> "Default" Then
        strSQL = strSQL & " AND a.workseq BETWEEN " & DBV("workseq", FrWorkSeq) & " AND " & DBV("workseq", ToWorkSeq)
    End If
   
    strSQL = strSQL & " AND   b.workarea = a.workarea" & _
                      " AND   b.accdt    = a.accdt" & _
                      " AND   b.accseq   = a.accseq" & _
                      " AND   b.stscd IN (" & DBV("stscd", enStsCd.StsCd_LIS_InProcess) & "," & _
                                              DBV("stscd", enStsCd.StsCd_LIS_MidRst) & ")" & _
                      " AND   EXISTS(SELECT c.testcd FROM " & T_LAB026 & " c, " & T_LAB008 & " d " & _
                      " WHERE c.workarea = a.workarea" & _
                      " AND   c.accdt    = a.accdt" & _
                      " AND   c.accseq   = a.accseq" & _
                      " AND  (c.vfydt = '' or c.vfydt is null) " & _
                      " AND " & DBW("d.workcd =", WorkCd) & " AND d.testcd = c.testcd ) "
   
    strSQL = strSQL & " order by workseq "
    DBOpenDyna strSQL
    Set pOraDynaset = m_RecordSet
   '
End Sub
  
Public Function UpdatableLAB301(ByVal WorkDt As String, _
                                ByVal WorkCd As String, ByVal WorkSeq As String) As Boolean
    Dim SqlStmt As String
   '
    SqlStmt = "SELECT * FROM " & T_LAB301 & _
              " WHERE " & DBW("workdt  = ", WorkDt) & _
              " AND   " & DBW("workcd  = ", WorkCd) & _
              " AND   " & DBW("workseq = ", WorkSeq)
    UpdatableLAB301 = DBExists(SqlStmt)
   '
End Function

Public Function GetIsEqpCd(ByVal EqpCd As String) As Boolean
    
    Dim strSQL As String
   '
    GetIsEqpCd = False
    strSQL = " SELECT COUNT(eqpcd) AS CNT" & _
             " FROM  " & T_LAB006 & _
             " WHERE " & DBW("eqpcd = ", EqpCd)
    '
    DBOpenDyna strSQL
    With m_RecordSet
        If Val("" & .Fields("CNT").Value) > 0 Then
            GetIsEqpCd = True
        End If
    End With
   '
End Function

Public Function GetEquipTransDtTm(ByRef pOraDynaset As Recordset, _
                                  ByVal EqpCd As String, ByVal pSpcNo As String) As String
    Dim strSQL As String
    Dim strTransDt As String
    Dim strTransTm As String
   '
    GetEquipTransDtTm = ""
    If pSpcNo = "" Then Exit Function
    strSQL = " SELECT * FROM " & T_LAB306 & " a " & _
             " WHERE " & DBW("a.eqpcd = ", EqpCd) & _
             " AND   " & DBW("a.spcyy = ", medGetP(pSpcNo, 1, "-")) & _
             " AND   " & DBW("a.spcno = ", Format(medGetP(pSpcNo, 2, "-"), "########"))
             
    DBOpenDyna strSQL
    Set pOraDynaset = m_RecordSet
    If m_RecordSet.RecordCount > 0 Then
        GetEquipTransDtTm = "" & m_RecordSet.Fields("TRANSDT").Value & vbTab & _
                            "" & m_RecordSet.Fields("TRANSTM").Value
    End If
    Debug.Print strSQL
End Function

Public Sub GetLAB306(ByRef pOraDynaset As Recordset, _
                     ByVal EqpCd As String, ByVal pTransDt, ByVal pTransTm As String, _
                     Optional ByVal pStatFg As String = "")
    Dim strSQL As String
   '
    strSQL = "SELECT distinct a.eqpcd, a.transdt, a.transtm, a.transno, a.spcyy, a.spcno, b.statfg, b.workarea, b.accdt, b.accseq " & _
             " FROM  " & T_LAB306 & " a, " & T_LAB201 & " b " & _
             " WHERE " & DBW("a.eqpcd    = ", EqpCd) & _
             " AND   " & DBW("a.transdt >= ", pTransDt) & _
             " AND   a.transdt " & FUNC_CONCAT & " a.transtm >=" & DBS(pTransDt & pTransTm) & _
             " AND   b.spcyy = a.spcyy " & _
             " AND   b.spcno = a.spcno " & _
             " AND   EXISTS(SELECT c.accseq FROM " & T_LAB302 & " c WHERE c.workarea = b.workarea" & _
             " AND   c.accdt = b.accdt  AND  c.accseq = b.accseq  AND  c.eqpcd = a.eqpcd " & _
             " AND  (c.vfydt = ''  or  c.vfydt is null))"
'--2000.12.6 김미경 추가 : QC 결과내역도 가져오기 위해 UNION 걸어줌
'    strSql = strSql & " UNION ALL SELECT a.eqpcd, a.transdt, a.transtm, a.transno, a.spcyy, a.spcno, b.statfg, b.workarea, b.accdt, b.accseq " & _
'             " FROM  " & T_LAB306 & " a, " & T_LAB201 & " b " & _
'             " WHERE " & DBW("a.eqpcd    = ", EqpCd) & _
'             " AND   " & DBW("a.transdt >= ", pTransDt) & _
'             " AND   a.transdt " & FUNC_CONCAT & " a.transtm >=" & DBS(pTransDt & pTransTm) & _
'             " AND   b.spcyy = a.spcyy " & _
'             " AND   b.spcno = a.spcno " & _
'             " AND   EXISTS(SELECT c.accseq FROM " & T_LAB026 & " c WHERE c.workarea = b.workarea" & _
'             " AND   c.accdt = b.accdt  AND  c.accseq = b.accseq  AND  c.eqpcd = a.eqpcd" & _
'             " AND  (c.vfydt = ''  or  c.vfydt is null))"
    If pStatFg = "1" Then
        strSQL = strSQL & " ORDER BY statfg desc, transdt, transtm"
    Else
        strSQL = strSQL & " ORDER BY transdt, transtm"
    End If
    DBOpenDyna strSQL
    Set pOraDynaset = m_RecordSet
    Debug.Print strSQL
End Sub
     
Public Function GetWardLocation(ByVal PtId As String, ByVal Bedindt As String) As String
    
    Dim strSQL As String
    '
    strSQL = " SELECT " & F_PTWARDID & " wardid, " & F_PTROOMID & " roomid, " & F_PTBEDID & " bedid " & _
             " FROM " & T_HIS002 & _
             " WHERE " & DBW(F_PTID & " = ", PtId) & _
             " AND   in_date = to_date(" & Bedindt & ",'yyyymmdd') "
       
    '
    DBOpenDyna strSQL
    With m_RecordSet
        If .RecordCount > 0 Then
            GetWardLocation = "" & m_RecordSet.Fields("WARDID").Value & "-" & _
                              "" & m_RecordSet.Fields("ROOMID").Value & "-" & _
                              "" & m_RecordSet.Fields("BEDID").Value
        Else
            GetWardLocation = ""
        End If
    End With
    '
End Function

Public Sub GetPtInfo(ByRef pOraDynaset As Recordset, _
                     ByVal WorkArea As String, ByVal AccDt As String, ByVal accseq As String)
    Dim strSQL As String
    '
    strSQL = "SELECT a.ptid,b.ptnm,b.ssn,b.dob,b.sex" & _
             " FROM  " & T_HIS001 & " b," & T_LAB201 & " a" & _
             " WHERE " & DBW("a.workarea = ", WorkArea) & _
             " AND   " & DBW("a.accdt    = ", AccDt) & _
             " AND   " & DBW("a.accseq   = ", accseq) & _
             " AND   a.ptid = b." & F_PTID
    '
    DBOpenDyna strSQL
    Set pOraDynaset = m_RecordSet
    '
End Sub
    
Public Function GetOcsPtInfo(ByVal PtId As String, ByVal DeptCd As String, _
                             ByVal OrdDoct As String) As String
    
    Dim strSQL As String
    Dim aryTmp() As String
'    Dim objData As clsBasisData
   
    '/* 병원쪽 OCS에서 환자의 기본정보를 가져올 경우 이용한다. */
    '/* 단, 파라메터가 널이면 정보를 구하지 않고 NULL을 RETURN한다. */
    ReDim aryTmp(4)
    If PtId <> "" Then
        strSQL = "SELECT " & F_PTNM & " as PtNm, " & F_DOB2 & " as dob, " & F_SEX2 & " as sex" & _
                 " FROM  " & T_HIS001 & _
                 " WHERE " & DBW(F_PTID & " = ", PtId)
          
        DBOpenDyna strSQL
        If m_RecordSet.RecordCount > 0 Then
            aryTmp(0) = "" & m_RecordSet.Fields("PTNM").Value
            aryTmp(1) = "" & m_RecordSet.Fields("DOB").Value
            
            aryTmp(2) = Trim(m_RecordSet.Fields("SEX").Value & "")
            If IsNumeric(aryTmp(2)) Then
                aryTmp(2) = Choose((Val(aryTmp(2)) Mod 2) + 1, "F", "M")
            End If
            


        End If
    End If
    
'    Set objData = New clsBasisData
    
    If DeptCd <> "" Then
        aryTmp(3) = GetDeptNm(DeptCd) 'GetDeptName(DeptCd)
    End If
    '
    aryTmp(4) = GetDoctNm(OrdDoct) 'GetDoctName(OrdDoct)
    
    GetOcsPtInfo = Join(aryTmp, vbTab)
    
'    Set objData = Nothing
End Function
    
Public Sub GetResultPtAccNo(ByVal QueryType As Integer, ByRef pOraDynaset As Recordset, _
                            ByVal WorkArea As String, ByVal AccDt As String, ByVal accseq As String)
    Dim strSQL As String
   '
    '** 검체명도 같이 가지고 온다.. 2000/2/21 by 김미경
    strSQL = " SELECT a.ptid,a.bedindt,a.deptcd,a.orddoct,a.majdoct,a.stscd," & _
             "        a.workarea,a.accdt,a.accseq,a.spccd,a.coldt,a.orgaccno," & _
             "        a.ageday,a.reqtotcnt,a.reqinputcnt,a.rmkcd," & _
             "        a.footnotefg,a.wardid,a.roomid,a.bedid, a.qcfg, b.field3 as SpcNm, a.rcvdt, a.rcvtm " & _
             " FROM   " & T_LAB201 & " a, " & T_LAB032 & " b " & _
             " WHERE  " & DBW("a.workarea", WorkArea, 2) & _
             " AND    " & DBW("a.accdt", AccDt, 2) & _
             " AND    " & DBW("a.accseq", accseq, 2) & _
             " AND    " & DBW("b.cdindex", LC3_Specimen, 2) & _
             " AND    b.cdval1 = a.spccd "
   '
    If QueryType <> RESULT_BY_MODIFY Then
       strSQL = strSQL & " AND a.stscd IN (" & DBV("stscd", enStsCd.StsCd_LIS_Accession) & "," & _
                                               DBV("stscd", enStsCd.StsCd_LIS_InProcess) & "," & _
                                               DBV("stscd", enStsCd.StsCd_LIS_MidRst) & ")"
    Else
       strSQL = strSQL & " AND a.stscd IN (" & DBV("stscd", enStsCd.StsCd_LIS_Accession) & "," & _
                                               DBV("stscd", enStsCd.StsCd_LIS_InProcess) & "," & _
                                               DBV("stscd", enStsCd.StsCd_LIS_MidRst) & "," & _
                                               DBV("stscd", enStsCd.StsCd_LIS_FinRst) & ")"
    End If
    
    DBOpenDyna strSQL
    Set pOraDynaset = m_RecordSet
   '
End Sub
    
Public Function GetEqpTrans(ByVal EqpCd As String, ByVal SpcYy As String, _
                            ByVal SpcNo As String, ByVal AutoFg As String) As String
    Dim strSQL As String
    Dim strEqpNm As String
    Dim strTransDt As String
    Dim strTransTm As String
   '
    If AutoFg = "1" Then
       strSQL = " SELECT a.transdt,a.transtm,b.eqpnm" & _
                " FROM   " & T_LAB006 & " b," & T_LAB306 & " a" & _
                " WHERE  " & DBW("a.eqpcd", EqpCd, 2) & _
                " AND    " & DBW("a.spcyy", SpcYy, 2) & _
                " AND    " & DBW("a.spcno", SpcNo, 2) & _
                " AND    b.eqpcd = a.eqpcd"
    Else
       strSQL = " SELECT eqpnm FROM " & T_LAB006 & _
                " WHERE  " & DBW("eqpcd", EqpCd, 2)
    End If
   '
    DBOpenDyna strSQL
    strEqpNm = "": strTransDt = "": strTransTm = ""
    If m_RecordSet.RecordCount > 0 Then
       strEqpNm = "" & m_RecordSet.Fields("EQPNM").Value
       If AutoFg = "1" Then
          strTransDt = "" & m_RecordSet.Fields("TRANSDT").Value
          strTransTm = "" & m_RecordSet.Fields("TRANSTM").Value
       End If
    End If
    GetEqpTrans = strEqpNm & vbTab & strTransDt & vbTab & strTransTm
   '
End Function
    
Public Sub GetWSNo(ByRef pOraDynaset As Recordset, _
                    ByVal PtType As Integer, ByVal WorkDt As String, _
                    ByVal WorkCd As String, ByVal strFrWSeq As String, _
                    ByVal strToWSeq As String)
    Dim strSQL As String
   '
    strSQL = " SELECT a.* FROM " & T_LAB301 & " a" & _
             " WHERE  " & DBW("a.workdt", WorkDt, 2) & _
             " AND    " & DBW("a.workcd", WorkCd, 2) & _
             " AND    a.workseq BETWEEN " & DBV("workseq", strFrWSeq) & " AND  " & DBV("workseq", strToWSeq)
    DBOpenDyna strSQL
    Set pOraDynaset = m_RecordSet
   '
End Sub


Public Sub GetWSBuild(ByRef pOraDynaset As Recordset, _
                      ByVal WorkCd As String, ByVal RcvDt As String, _
                      ByVal RcvTm As String, ByVal BuildCd As String, _
                      ByVal strLastDtTm As String, Optional ByVal WorkArea As String, _
                      Optional ByVal blnAllData As Boolean, _
                      Optional ByVal strStatFg As String)
    Dim strSQL As String
    Dim strRcvDtTm As String
    Dim strLastDt As String
    Dim strStatus As String

    '--2000.12.4 김미경 수정
    '  1. h7lab201_index2 를 타게 하기위해 조건 추가 (rcvdt >= , rcvdt <=)
    '  2. QC내역도 Worksheet에 추가하기 위해 h7lab026을 UNION으로 걸어줌
    
    '--2001.6.26 김미경
    '  3. lab302에 웤슅코드(workcd)를 관리해서 다른 웤슅에 한번 출력된 검사는
    '     다른 웤슅에 출력안되게 한다.
    '  4. 응급,비응급 구분을 두어 응급검체와 비응급검체를 분리하여 웤슅을 출력할수
    '     있되, 마감시간은 정확히 관리가 안되므로 사용자가 시간을 조절하여 작성한다.
    '     (전체일때는 무관함)
          
    strRcvDtTm = RcvDt & RcvTm
    strLastDt = Mid(strLastDtTm, 1, 8)
    
    If strStatFg <> "2" Then
        strStatus = strStatus & " AND  " & DBW("c.statfg = ", strStatFg)
    End If
    
    '
    '2002/08/21
    '기존은 10시간 이전 접수항목에 대해서만 worksheet를 작성하였지만,
    '앞으로는 접수가 되지 않은항목에 대해서 모두 워크쉬트를 작성한다.
    '접수일자는 선택된 날짜에서 2개월 전까지만 접수일자를 조회한다.
    '" AND    a.rcvdt " & FUNC_CONCAT & " a.rcvtm > " & DBS(strLastDtTm) & _ 삭제
    
    strSQL = " SELECT a.workarea,a.accdt,a.accseq,a.rcvdt,a.rcvtm " & _
             " FROM   " & T_LAB201 & " a " & _
             " WHERE  " & DBW("a.workarea", WorkArea, 2) & _
             " AND    " & DBW("a.rcvdt >= ", strLastDt) & _
             " AND    " & DBW("a.rcvdt <= ", RcvDt) & _
             " AND    a.rcvdt " & FUNC_CONCAT & " a.rcvtm <= " & DBS(strRcvDtTm) & _
             " AND    a.stscd in (" & DBV("stscd", enStsCd.StsCd_LIS_Accession) & ", " & _
                                      DBV("stscd", enStsCd.StsCd_LIS_InProcess) & ") " & _
             " AND    " & DBW("a.buildcd", BuildCd, 2) & _
             " AND    EXISTS(SELECT *" & _
                            " FROM   " & T_LAB008 & " b, " & T_LAB302 & " c " & _
                            " WHERE  c.workarea = a.workarea" & _
                            " AND    c.accdt    = a.accdt" & _
                            " AND    c.accseq   = a.accseq" & _
                            " AND    " & DBW("b.workcd", WorkCd, 2) & strStatus & _
                            " AND    c.testcd   = b.testcd AND c.spccd = b.spccd " & _
                            " AND   (c.workcd =''  or  c.workcd is null) " & _
                            " AND   (c.vfydt = ''  or  c.vfydt  is null) )"
    
    strSQL = strSQL & " UNION " & _
             " SELECT a.workarea,a.accdt,a.accseq,a.rcvdt,a.rcvtm " & _
             " FROM   " & T_LAB201 & " a " & _
             " WHERE  " & DBW("a.workarea", WorkArea, 2) & _
             " AND    " & DBW("a.rcvdt >= ", strLastDt) & _
             " AND    " & DBW("a.rcvdt <= ", RcvDt) & _
             " AND    a.rcvdt " & FUNC_CONCAT & " a.rcvtm <= " & DBS(strRcvDtTm) & _
             " AND    a.stscd in (" & DBV("stscd", enStsCd.StsCd_LIS_Accession) & ", " & _
                                      DBV("stscd", enStsCd.StsCd_LIS_InProcess) & ") " & _
             " AND    " & DBW("a.buildcd", BuildCd, 2) & _
             " AND    EXISTS(SELECT *" & _
                            " FROM   " & T_LAB008 & " b, " & T_LAB351 & " c " & _
                            " WHERE  c.workarea = a.workarea" & _
                            " AND    c.accdt    = a.accdt" & _
                            " AND    c.accseq   = a.accseq" & _
                            " AND    " & DBW("b.workcd", WorkCd, 2) & _
                            " AND    c.testcd   = b.testcd  " & _
                            " AND    " & DBW("c.stscd", enStsCd.StsCd_LIS_Accession, 2) & " )"
    '--> 나중에 QC 적용시 반영
'    strSql = strSql & " UNION " & _
'             " SELECT a.workarea,a.accdt,a.accseq,a.rcvdt,a.rcvtm " & _
'             " FROM   " & T_LAB201 & " a," & T_LAB008 & " b" & _
'             " WHERE  " & DBW("a.workarea", WorkArea, 2) & _
'             " AND    " & DBW("a.rcvdt >= ", strLastDt) & _
'             " AND    " & DBW("a.rcvdt <= ", RcvDt) & _
'             " AND    a.rcvdt " & FUNC_CONCAT & " a.rcvtm > " & DBS(strLastDtTm) & _
'             " AND    a.rcvdt " & FUNC_CONCAT & " a.rcvtm <= " & DBS(strRcvDtTm) & _
'             " AND    a.stscd in (" & strStatus & ") " & _
'             " AND    " & DBW("a.buildcd", BuildCd, 2) & _
'             " AND    " & DBW("a.qcfg", "1", 2) & _
'             " AND    EXISTS(SELECT *" & _
'                            " FROM   " & T_LAB008 & " b, " & T_LAB026 & " c " & _
'                            " WHERE  c.workarea = a.workarea" & _
'                            " AND    c.accdt = a.accdt" & _
'                            " AND    c.accseq = a.accseq" & _
'                            " AND    " & DBW("b.workcd", WorkCd, 2) & _
'                            " AND    c.testcd = b.testcd  " & _
'                            " AND   (c.vfydt = ''  or  c.vfydt is null) )"
    strSQL = strSQL & " order by rcvdt, rcvtm"               '시간순으로.. 2000.2.11
    '--------------------------------------------------------------------
    '   한 Lab No.가 다수의 워크쉬트코드에 해당될수 있으므로 status가
    '   접수/In-Process 모두 검색한다.
    '   대신, 마감일자 관리를 함으로써 워크쉬트의 중복생성을 방지한다.
    '   마감일자 : LAB032 - 'C228' (field2, field3)
    '   " AND a.stscd = '" & enStsCd.StsCd_LIS_Accession & "' " & _
    '                                                                       1999.09.08  by KMK
    '   모든 검색기준에서 TestDiv를 제외시킨다. 일반검사와 기타검사가
    '   한 LabNo로 묶일수 있기 때문이다.. (미생물까지도...)
    '   " AND a.testdiv = '" & TST_RouTest & "'" & _
    '                                                                       1999.10.08  by KMK
    '--------------------------------------------------------------------
    DBOpenDyna strSQL
    Set pOraDynaset = m_RecordSet
    '
End Sub


Public Sub GetWSBuild_NEW(ByRef pOraDynaset As Recordset, _
                          ByVal WorkCd As String, ByVal RcvDt As String, _
                          ByVal RcvTm As String, ByVal BuildCd As String, _
                          ByVal strLastDtTm As String, Optional ByVal WorkArea As String, _
                          Optional ByVal blnAllData As Boolean, _
                          Optional ByVal strStatFg As String, _
                          Optional ByVal chkFg As String = "1")
    Dim strSQL As String
    Dim strRcvDtTm As String
    Dim strLastDt As String
    Dim strStatus As String
    Dim strRangeDt As String
    

    '--2000.12.4 김미경 수정
    '  1. h7lab201_index2 를 타게 하기위해 조건 추가 (rcvdt >= , rcvdt <=)
    '  2. QC내역도 Worksheet에 추가하기 위해 h7lab026을 UNION으로 걸어줌
    
    '--2001.6.26 김미경
    '  3. lab302에 웤슅코드(workcd)를 관리해서 다른 웤슅에 한번 출력된 검사는
    '     다른 웤슅에 출력안되게 한다.
    '  4. 응급,비응급 구분을 두어 응급검체와 비응급검체를 분리하여 웤슅을 출력할수
    '     있되, 마감시간은 정확히 관리가 안되므로 사용자가 시간을 조절하여 작성한다.
    '     (전체일때는 무관함)
          
    strRcvDtTm = RcvDt & RcvTm
    
    '성바오로병원의 경우 2개월 전에 접수된 접수번호에 대해서도 웍쉬트를 만든다
    strLastDt = Replace(DateAdd("d", -60, Format(Mid(strLastDtTm, 1, 8), "####-##-##")), "-", "")
    

    
    If strStatFg <> "2" Then
        strStatus = strStatus & " AND  " & DBW("c.statfg = ", strStatFg)
    End If
    
    '-------------------
    '성바오로 요청 sql문 수정.
    '-----------------
    
'    If dbconn.Whatsthis = dbconn.ThisIsOracle Then
    If ObjSysInfo.dbtype = 0 Then 'oracle
        '## 5.0.26: 이상대(2005-06-01)
        '   - 조회속도가 느려 Rule Base로 쿼리수정
        strSQL = " SELECT /*+ RULE S2LAB201(S2LAB201_IDX2) */ a.workarea,a.accdt,a.accseq,a.rcvdt,a.rcvtm " & _
                         " FROM   " & T_LAB201 & " a " & _
                         " WHERE  " & DBW("a.rcvdt >= ", strLastDt) & _
                         " AND    " & DBW("a.rcvdt <= ", RcvDt) & _
                         " AND    " & DBW("a.workarea>=", WorkArea) & _
                         " AND    " & DBW("a.workarea<=", WorkArea) & _
                         " AND    a.rcvdt " & FUNC_CONCAT & " a.rcvtm <= " & DBS(strRcvDtTm) & _
                         " AND    a.stscd in (" & DBV("stscd", enStsCd.StsCd_LIS_Accession) & ", " & _
                                                  DBV("stscd", enStsCd.StsCd_LIS_InProcess) & ") " & _
                         " AND    " & DBW("a.buildcd", BuildCd, 2) & _
                         " AND    EXISTS(SELECT 'x' " & _
                                        " FROM   " & T_LAB008 & " b, " & T_LAB302 & " c " & _
                                        " WHERE  c.workarea = a.workarea" & _
                                        " AND    c.accdt    = a.accdt" & _
                                        " AND    c.accseq   = a.accseq" & _
                                        " AND    " & DBW("b.workcd", WorkCd, 2) & strStatus & _
                                        " AND    c.testcd   = b.testcd AND c.spccd = b.spccd " & _
                                        " AND   (c.workcd =''  or  c.workcd is null) " & _
                                        " AND   (c.vfydt = ''  or  c.vfydt  is null) "
                                        
                         strSQL = strSQL & " UNION " & _
                                           " SELECT 'x' " & _
                                           "   FROM   " & T_LAB008 & " b, " & T_LAB351 & " c " & _
                                           " WHERE  c.workarea = a.workarea" & _
                                           " AND    c.accdt    = a.accdt" & _
                                           " AND    c.accseq   = a.accseq" & _
                                           " AND    " & DBW("b.workcd", WorkCd, 2) & _
                                           " AND    c.testcd   = b.testcd  " & _
                                           " AND    " & DBW("c.stscd", enStsCd.StsCd_LIS_Accession, 2) & " )"
    
    Else
        
        'workarea 를 빼버렸음...index 안타서
        '" & DBW("a.workarea", WorkArea, 2) & _
                 " AND
            ' 원래 소스
            strSQL = " SELECT a.workarea,a.accdt,a.accseq,a.rcvdt,a.rcvtm " & _
                     " FROM   " & T_LAB201 & " a " & _
                     " WHERE  " & DBW("a.rcvdt >= ", strLastDt) & _
                     " AND    " & DBW("a.rcvdt <= ", RcvDt) & _
                     " AND    " & DBW("a.workarea>=", WorkArea) & _
                     " AND    " & DBW("a.workarea<=", WorkArea) & _
                     " AND    a.rcvdt " & FUNC_CONCAT & " a.rcvtm <= " & DBS(strRcvDtTm) & _
                     " AND    a.stscd in (" & DBV("stscd", enStsCd.StsCd_LIS_Accession) & ", " & _
                                              DBV("stscd", enStsCd.StsCd_LIS_InProcess) & ") " & _
                     " AND    " & DBW("a.buildcd", BuildCd, 2) & _
                     " AND    EXISTS(SELECT *" & _
                                    " FROM   " & T_LAB008 & " b, " & T_LAB302 & " c " & _
                                    " WHERE  c.workarea = a.workarea" & _
                                    " AND    c.accdt    = a.accdt" & _
                                    " AND    c.accseq   = a.accseq" & _
                                    " AND    " & DBW("b.workcd", WorkCd, 2) & strStatus & _
                                    " AND    c.testcd   = b.testcd AND c.spccd = b.spccd " & _
                                    " AND   (c.workcd =''  or  c.workcd is null) " & _
                                    " AND   (c.vfydt = ''  or  c.vfydt  is null) )"
        '원래 소스
            strSQL = strSQL & " UNION " & _
                     " SELECT a.workarea,a.accdt,a.accseq,a.rcvdt,a.rcvtm " & _
                     " FROM   " & T_LAB201 & " a " & _
                     " WHERE  " & DBW("a.rcvdt >= ", strLastDt) & _
                     " AND    " & DBW("a.rcvdt <= ", RcvDt) & _
                     " AND    " & DBW("a.workarea>=", WorkArea) & _
                     " AND    " & DBW("a.workarea<=", WorkArea) & _
                     " AND    a.rcvdt " & FUNC_CONCAT & " a.rcvtm <= " & DBS(strRcvDtTm) & _
                     " AND    a.stscd in (" & DBV("stscd", enStsCd.StsCd_LIS_Accession) & ", " & _
                                              DBV("stscd", enStsCd.StsCd_LIS_InProcess) & ") " & _
                     " AND    " & DBW("a.buildcd", BuildCd, 2) & _
                     " AND    EXISTS(SELECT *" & _
                                    " FROM   " & T_LAB008 & " b, " & T_LAB351 & " c " & _
                                    " WHERE  c.workarea = a.workarea" & _
                                    " AND    c.accdt    = a.accdt" & _
                                    " AND    c.accseq   = a.accseq" & _
                                    " AND    " & DBW("b.workcd", WorkCd, 2) & _
                                    " AND    c.testcd   = b.testcd  " & _
                                    " AND    " & DBW("c.stscd", enStsCd.StsCd_LIS_Accession, 2) & " )"
    End If
    
    
    '2002/08/21
    '기존은 10시간 이전 접수항목에 대해서만 worksheet를 작성하였지만,
    '앞으로는 접수가 되지 않은항목에 대해서 모두 워크쉬트를 작성한다.
    '접수일자는 선택된 날짜에서 2개월 전까지만 접수일자를 조회한다.
    '" AND    a.rcvdt " & FUNC_CONCAT & " a.rcvtm > " & DBS(strLastDtTm) & _ 삭제
    

    
    

    If chkFg = "1" Then
        strSQL = strSQL & " order by rcvdt, rcvtm"               '시간순으로.. 2000.2.11
    Else
        strSQL = strSQL & " order by workarea,accdt,accseq"
    End If
    '--------------------------------------------------------------------
    '   한 Lab No.가 다수의 워크쉬트코드에 해당될수 있으므로 status가
    '   접수/In-Process 모두 검색한다.
    '   대신, 마감일자 관리를 함으로써 워크쉬트의 중복생성을 방지한다.
    '   마감일자 : LAB032 - 'C228' (field2, field3)
    '   " AND a.stscd = '" & enStsCd.StsCd_LIS_Accession & "' " & _
    '                                                                       1999.09.08  by KMK
    '   모든 검색기준에서 TestDiv를 제외시킨다. 일반검사와 기타검사가
    '   한 LabNo로 묶일수 있기 때문이다.. (미생물까지도...)
    '   " AND a.testdiv = '" & TST_RouTest & "'" & _
    '                                                                       1999.10.08  by KMK
    '--------------------------------------------------------------------
    
    Debug.Print strSQL
    DBOpenDyna strSQL
    Set pOraDynaset = m_RecordSet
    '
End Sub

Public Sub GetWSBuild_NEW_2014(ByRef pOraDynaset As Recordset, _
                          ByVal Poct As String, ByVal WorkCd As String, ByVal RcvDt As String, _
                          ByVal RcvTm As String, ByVal BuildCd As String, _
                          ByVal strLastDtTm As String, Optional ByVal WorkArea As String, _
                          Optional ByVal blnAllData As Boolean, _
                          Optional ByVal strStatFg As String, _
                          Optional ByVal chkFg As String = "1")
    Dim strSQL As String
    Dim strRcvDtTm As String
    Dim strLastDt As String
    Dim strStatus As String
    Dim strRangeDt As String
    

    '--2000.12.4 김미경 수정
    '  1. h7lab201_index2 를 타게 하기위해 조건 추가 (rcvdt >= , rcvdt <=)
    '  2. QC내역도 Worksheet에 추가하기 위해 h7lab026을 UNION으로 걸어줌
    
    '--2001.6.26 김미경
    '  3. lab302에 웤슅코드(workcd)를 관리해서 다른 웤슅에 한번 출력된 검사는
    '     다른 웤슅에 출력안되게 한다.
    '  4. 응급,비응급 구분을 두어 응급검체와 비응급검체를 분리하여 웤슅을 출력할수
    '     있되, 마감시간은 정확히 관리가 안되므로 사용자가 시간을 조절하여 작성한다.
    '     (전체일때는 무관함)
          
    strRcvDtTm = RcvDt & RcvTm
    
    '성바오로병원의 경우 2개월 전에 접수된 접수번호에 대해서도 웍쉬트를 만든다
    strLastDt = Replace(DateAdd("d", -60, Format(Mid(strLastDtTm, 1, 8), "####-##-##")), "-", "")
    

    
    If strStatFg <> "2" Then
        strStatus = strStatus & " AND  " & DBW("c.statfg = ", strStatFg)
    End If
    
    '-------------------
    '성바오로 요청 sql문 수정.
    '-----------------
    
'    If dbconn.Whatsthis = dbconn.ThisIsOracle Then
    If ObjSysInfo.dbtype = 0 Then 'oracle
        '## 5.0.26: 이상대(2005-06-01)
        '   - 조회속도가 느려 Rule Base로 쿼리수정
        
        '2014-01-27 Poct = 0.전체 1.제활센타(94,95,96병동) 2.본관(94,95,96병동제외)
        Select Case Poct
         Case "1"   '제활센타(94,95,96 병동)
                strSQL = " SELECT /*+ RULE S2LAB201(S2LAB201_IDX2) */ a.workarea,a.accdt,a.accseq,a.rcvdt,a.rcvtm " & _
                         " FROM   " & T_LAB201 & " a " & _
                         " WHERE  " & DBW("a.rcvdt >= ", strLastDt) & _
                         " AND    " & DBW("a.rcvdt <= ", RcvDt) & _
                         " AND    " & DBW("a.workarea>=", WorkArea) & _
                         " AND    " & DBW("a.workarea<=", WorkArea) & _
                         " AND    a.rcvdt " & FUNC_CONCAT & " a.rcvtm <= " & DBS(strRcvDtTm) & _
                         " AND    a.stscd in (" & DBV("stscd", enStsCd.StsCd_LIS_Accession) & ", " & _
                                                  DBV("stscd", enStsCd.StsCd_LIS_InProcess) & ") " & _
                         " AND    " & DBW("a.buildcd", BuildCd, 2) & _
                         " AND    a.wardid   IN ('94','95','96') " & _
                         " AND    EXISTS(SELECT 'x' " & _
                                        " FROM   " & T_LAB008 & " b, " & T_LAB302 & " c " & _
                                        " WHERE  c.workarea = a.workarea" & _
                                        " AND    c.accdt    = a.accdt" & _
                                        " AND    c.accseq   = a.accseq" & _
                                        " AND    " & DBW("b.workcd", WorkCd, 2) & strStatus & _
                                        " AND    c.testcd   = b.testcd AND c.spccd = b.spccd " & _
                                        " AND   (c.workcd =''  or  c.workcd is null) " & _
                                        " AND   (c.vfydt = ''  or  c.vfydt  is null) "
                                        
                         strSQL = strSQL & " UNION " & _
                                           " SELECT 'x' " & _
                                           "   FROM   " & T_LAB008 & " b, " & T_LAB351 & " c " & _
                                           " WHERE  c.workarea = a.workarea" & _
                                           " AND    c.accdt    = a.accdt" & _
                                           " AND    c.accseq   = a.accseq" & _
                                           " AND    " & DBW("b.workcd", WorkCd, 2) & _
                                           " AND    c.testcd   = b.testcd  " & _
                                           " AND    " & DBW("c.stscd", enStsCd.StsCd_LIS_Accession, 2) & " )"
         Case "2"   '본관(94,95,96 병동제외)
                strSQL = " SELECT /*+ RULE S2LAB201(S2LAB201_IDX2) */ a.workarea,a.accdt,a.accseq,a.rcvdt,a.rcvtm " & _
                         " FROM   " & T_LAB201 & " a " & _
                         " WHERE  " & DBW("a.rcvdt >= ", strLastDt) & _
                         " AND    " & DBW("a.rcvdt <= ", RcvDt) & _
                         " AND    " & DBW("a.workarea>=", WorkArea) & _
                         " AND    " & DBW("a.workarea<=", WorkArea) & _
                         " AND    a.rcvdt " & FUNC_CONCAT & " a.rcvtm <= " & DBS(strRcvDtTm) & _
                         " AND    a.stscd in (" & DBV("stscd", enStsCd.StsCd_LIS_Accession) & ", " & _
                                                  DBV("stscd", enStsCd.StsCd_LIS_InProcess) & ") " & _
                         " AND    " & DBW("a.buildcd", BuildCd, 2) & _
                         " AND    a.wardid   NOT IN ('94','95','96') " & _
                         " AND    EXISTS(SELECT 'x' " & _
                                        " FROM   " & T_LAB008 & " b, " & T_LAB302 & " c " & _
                                        " WHERE  c.workarea = a.workarea" & _
                                        " AND    c.accdt    = a.accdt" & _
                                        " AND    c.accseq   = a.accseq" & _
                                        " AND    " & DBW("b.workcd", WorkCd, 2) & strStatus & _
                                        " AND    c.testcd   = b.testcd AND c.spccd = b.spccd " & _
                                        " AND   (c.workcd =''  or  c.workcd is null) " & _
                                        " AND   (c.vfydt = ''  or  c.vfydt  is null) "
                                        
                         strSQL = strSQL & " UNION " & _
                                           " SELECT 'x' " & _
                                           "   FROM   " & T_LAB008 & " b, " & T_LAB351 & " c " & _
                                           " WHERE  c.workarea = a.workarea" & _
                                           " AND    c.accdt    = a.accdt" & _
                                           " AND    c.accseq   = a.accseq" & _
                                           " AND    " & DBW("b.workcd", WorkCd, 2) & _
                                           " AND    c.testcd   = b.testcd  " & _
                                           " AND    " & DBW("c.stscd", enStsCd.StsCd_LIS_Accession, 2) & " )"
         Case "0"   '전체
            strSQL = " SELECT /*+ RULE S2LAB201(S2LAB201_IDX2) */ a.workarea,a.accdt,a.accseq,a.rcvdt,a.rcvtm " & _
                             " FROM   " & T_LAB201 & " a " & _
                             " WHERE  " & DBW("a.rcvdt >= ", strLastDt) & _
                             " AND    " & DBW("a.rcvdt <= ", RcvDt) & _
                             " AND    " & DBW("a.workarea>=", WorkArea) & _
                             " AND    " & DBW("a.workarea<=", WorkArea) & _
                             " AND    a.rcvdt " & FUNC_CONCAT & " a.rcvtm <= " & DBS(strRcvDtTm) & _
                             " AND    a.stscd in (" & DBV("stscd", enStsCd.StsCd_LIS_Accession) & ", " & _
                                                      DBV("stscd", enStsCd.StsCd_LIS_InProcess) & ") " & _
                             " AND    " & DBW("a.buildcd", BuildCd, 2) & _
                             " AND    EXISTS(SELECT 'x' " & _
                                            " FROM   " & T_LAB008 & " b, " & T_LAB302 & " c " & _
                                            " WHERE  c.workarea = a.workarea" & _
                                            " AND    c.accdt    = a.accdt" & _
                                            " AND    c.accseq   = a.accseq" & _
                                            " AND    " & DBW("b.workcd", WorkCd, 2) & strStatus & _
                                            " AND    c.testcd   = b.testcd AND c.spccd = b.spccd " & _
                                            " AND   (c.workcd =''  or  c.workcd is null) " & _
                                            " AND   (c.vfydt = ''  or  c.vfydt  is null) "
                                            
                             strSQL = strSQL & " UNION " & _
                                               " SELECT 'x' " & _
                                               "   FROM   " & T_LAB008 & " b, " & T_LAB351 & " c " & _
                                               " WHERE  c.workarea = a.workarea" & _
                                               " AND    c.accdt    = a.accdt" & _
                                               " AND    c.accseq   = a.accseq" & _
                                               " AND    " & DBW("b.workcd", WorkCd, 2) & _
                                               " AND    c.testcd   = b.testcd  " & _
                                               " AND    " & DBW("c.stscd", enStsCd.StsCd_LIS_Accession, 2) & " )"
        End Select
    Else
        
        'workarea 를 빼버렸음...index 안타서
        '" & DBW("a.workarea", WorkArea, 2) & _
                 " AND
            ' 원래 소스
            strSQL = " SELECT a.workarea,a.accdt,a.accseq,a.rcvdt,a.rcvtm " & _
                     " FROM   " & T_LAB201 & " a " & _
                     " WHERE  " & DBW("a.rcvdt >= ", strLastDt) & _
                     " AND    " & DBW("a.rcvdt <= ", RcvDt) & _
                     " AND    " & DBW("a.workarea>=", WorkArea) & _
                     " AND    " & DBW("a.workarea<=", WorkArea) & _
                     " AND    a.rcvdt " & FUNC_CONCAT & " a.rcvtm <= " & DBS(strRcvDtTm) & _
                     " AND    a.stscd in (" & DBV("stscd", enStsCd.StsCd_LIS_Accession) & ", " & _
                                              DBV("stscd", enStsCd.StsCd_LIS_InProcess) & ") " & _
                     " AND    " & DBW("a.buildcd", BuildCd, 2) & _
                     " AND    EXISTS(SELECT *" & _
                                    " FROM   " & T_LAB008 & " b, " & T_LAB302 & " c " & _
                                    " WHERE  c.workarea = a.workarea" & _
                                    " AND    c.accdt    = a.accdt" & _
                                    " AND    c.accseq   = a.accseq" & _
                                    " AND    " & DBW("b.workcd", WorkCd, 2) & strStatus & _
                                    " AND    c.testcd   = b.testcd AND c.spccd = b.spccd " & _
                                    " AND   (c.workcd =''  or  c.workcd is null) " & _
                                    " AND   (c.vfydt = ''  or  c.vfydt  is null) )"
        '원래 소스
            strSQL = strSQL & " UNION " & _
                     " SELECT a.workarea,a.accdt,a.accseq,a.rcvdt,a.rcvtm " & _
                     " FROM   " & T_LAB201 & " a " & _
                     " WHERE  " & DBW("a.rcvdt >= ", strLastDt) & _
                     " AND    " & DBW("a.rcvdt <= ", RcvDt) & _
                     " AND    " & DBW("a.workarea>=", WorkArea) & _
                     " AND    " & DBW("a.workarea<=", WorkArea) & _
                     " AND    a.rcvdt " & FUNC_CONCAT & " a.rcvtm <= " & DBS(strRcvDtTm) & _
                     " AND    a.stscd in (" & DBV("stscd", enStsCd.StsCd_LIS_Accession) & ", " & _
                                              DBV("stscd", enStsCd.StsCd_LIS_InProcess) & ") " & _
                     " AND    " & DBW("a.buildcd", BuildCd, 2) & _
                     " AND    EXISTS(SELECT *" & _
                                    " FROM   " & T_LAB008 & " b, " & T_LAB351 & " c " & _
                                    " WHERE  c.workarea = a.workarea" & _
                                    " AND    c.accdt    = a.accdt" & _
                                    " AND    c.accseq   = a.accseq" & _
                                    " AND    " & DBW("b.workcd", WorkCd, 2) & _
                                    " AND    c.testcd   = b.testcd  " & _
                                    " AND    " & DBW("c.stscd", enStsCd.StsCd_LIS_Accession, 2) & " )"
    End If
    
    
    '2002/08/21
    '기존은 10시간 이전 접수항목에 대해서만 worksheet를 작성하였지만,
    '앞으로는 접수가 되지 않은항목에 대해서 모두 워크쉬트를 작성한다.
    '접수일자는 선택된 날짜에서 2개월 전까지만 접수일자를 조회한다.
    '" AND    a.rcvdt " & FUNC_CONCAT & " a.rcvtm > " & DBS(strLastDtTm) & _ 삭제
    

    
    

    If chkFg = "1" Then
        strSQL = strSQL & " order by rcvdt, rcvtm"               '시간순으로.. 2000.2.11
    Else
        strSQL = strSQL & " order by workarea,accdt,accseq"
    End If
    '--------------------------------------------------------------------
    '   한 Lab No.가 다수의 워크쉬트코드에 해당될수 있으므로 status가
    '   접수/In-Process 모두 검색한다.
    '   대신, 마감일자 관리를 함으로써 워크쉬트의 중복생성을 방지한다.
    '   마감일자 : LAB032 - 'C228' (field2, field3)
    '   " AND a.stscd = '" & enStsCd.StsCd_LIS_Accession & "' " & _
    '                                                                       1999.09.08  by KMK
    '   모든 검색기준에서 TestDiv를 제외시킨다. 일반검사와 기타검사가
    '   한 LabNo로 묶일수 있기 때문이다.. (미생물까지도...)
    '   " AND a.testdiv = '" & TST_RouTest & "'" & _
    '                                                                       1999.10.08  by KMK
    '--------------------------------------------------------------------
    
    Debug.Print strSQL
    DBOpenDyna strSQL
    Set pOraDynaset = m_RecordSet
    '
End Sub
     
Public Function GetWSLastSeq(ByVal DivCd2 As String, _
                             ByVal DivCd3 As String) As Long
     Dim strSQL As String
     strSQL = " SELECT seq" & _
            " FROM   " & T_COM099 & _
            " WHERE  " & DBW("noindex", COM99_LIS_WorkNo, 2) & _
            " AND    " & DBW("divcd1", "0", 2) & _
            " AND    " & DBW("divcd2", DivCd2, 2) & _
            " AND    " & DBW("divcd3", DivCd3, 2)
     
     DBOpenDyna strSQL
     With m_RecordSet
        If .RecordCount = 0 Then
            GetWSLastSeq = 1
        Else
            GetWSLastSeq = Val("" & .Fields("SEQ").Value) + 1
        End If
     End With
    '
End Function
   
Public Function GetIsWorkCd(ByVal WorkCd As String) As Boolean
    
    Dim strSQL As String
    '
    GetIsWorkCd = False
    strSQL = " SELECT COUNT(workcd) AS CNT" & _
             " FROM  " & T_LAB008 & " " & _
             " WHERE " & DBW("workcd", WorkCd, 2)
    '
    DBOpenDyna strSQL
    With m_RecordSet
        If Val("" & .Fields("CNT").Value) > 0 Then
            GetIsWorkCd = True
        End If
    End With
    '
End Function

Public Function GetAddInfo(ByVal AccNo As String, _
                           ByVal WorkCd As String, ByVal RcvDt As String, _
                           ByVal RcvTm As String) As String
    Dim strSQL      As String
    Dim strWorkArea As String
    Dim strAccDt    As String
    Dim strAccSeq   As String
    Dim strWorkInfo As String
    
    Dim blnFirst    As Boolean
   
    strWorkArea = medGetP(AccNo, 1, "-")
    strAccDt = medGetP(AccNo, 2, "-")
    strAccSeq = medGetP(AccNo, 3, "-")
    '
    GetAddInfo = ""
    strWorkInfo = ""
    '
'    strSQL = " SELECT b.testcd, a.spccd,a.coldt" & _
'             " FROM   " & T_LAB302 & " b," & T_LAB201 & " a" & _
'             " WHERE  " & DBW("a.workarea", strWorkArea, 2) & _
'             " AND    " & DBW("a.accdt", strAccDt, 2) & _
'             " AND    " & DBW("a.accseq", strAccSeq, 2) & _
'             " AND    a.stscd IN(" & DBV("stscd", enStsCd.StsCd_LIS_Accession) & "," & _
'                                     DBV("stscd", enStsCd.StsCd_LIS_InProcess) & ")" & _
'             " AND    b.workarea = a.workarea" & _
'             " AND    b.accdt = a.accdt" & _
'             " AND    b.accseq = a.accseq" & _
'             " AND   (b.vfydt = ''  or  b.vfydt is null) " & _
'             " AND    EXISTS(SELECT c.testcd FROM " & T_LAB008 & " c" & _
'                            " WHERE  " & DBW("c.workcd", WorkCd, 2) & _
'                            " AND    c.testcd = b.testcd" & _
'                            " AND    c.spccd = b.spccd" & _
'                            " AND    c.workarea = b.workarea)"
       '" AND a.testdiv = '" & TST_RouTest & "'" & _

    strSQL = " SELECT b.testcd, a.spccd,a.coldt,c.workinfo" & _
             " FROM   " & T_LAB008 & " c," & T_LAB302 & " b," & T_LAB201 & " a" & _
             " WHERE  " & DBW("a.workarea", strWorkArea, 2) & _
             " AND    " & DBW("a.accdt", strAccDt, 2) & _
             " AND    " & DBW("a.accseq", strAccSeq, 2) & _
             " AND    a.stscd IN(" & DBV("stscd", enStsCd.StsCd_LIS_Accession) & "," & _
                                     DBV("stscd", enStsCd.StsCd_LIS_InProcess) & ")" & _
             " AND    b.workarea = a.workarea" & _
             " AND    b.accdt = a.accdt" & _
             " AND    b.accseq = a.accseq" & _
             " AND   (b.vfydt = ''  or  b.vfydt is null) " & _
             " AND  " & DBW("c.workcd", WorkCd, 2) & _
             " AND    c.testcd = b.testcd" & _
             " AND    c.spccd = b.spccd" & _
             " AND    c.workarea = b.workarea" & _
             " ORDER BY c.workinfo"
             
                            
             
             
             
             
             '& _
             " AND    EXISTS(SELECT c.testcd FROM " & T_LAB008 & " c" & _
                            " WHERE  " & DBW("c.workcd", WorkCd, 2) & _
                            " AND    c.testcd = b.testcd" & _
                            " AND    c.spccd = b.spccd" & _
                            " AND    c.workarea = b.workarea)"
    
    DBOpenDyna strSQL
    With m_RecordSet
        If .RecordCount > 0 Then
            .MoveFirst
            GetAddInfo = "" & m_RecordSet.Fields("SPCCD").Value & "," & _
                         "" & m_RecordSet.Fields("COLDT").Value & "$"
            blnFirst = True
            Do Until .EOF
                If blnFirst = True Then
                    GetAddInfo = GetAddInfo & "" & m_RecordSet.Fields("TESTCD").Value
                    blnFirst = False
                Else
                    GetAddInfo = GetAddInfo & vbTab & _
                                "" & m_RecordSet.Fields("TESTCD").Value
                End If
                strWorkInfo = strWorkInfo & Format(m_RecordSet.Fields("workinfo").Value & "", "00")
                .MoveNext
            Loop
            GetAddInfo = GetAddInfo & "$" & strWorkInfo
            
        End If
    End With
    '
End Function

Public Function GetTestNm(ByVal pTestCd As String, ByVal ApplyDt As String, _
                          Optional pTxtType As String = "Default") As String
    Dim strSQL As String
    strSQL = " SELECT a.testcd, a.testnm, a.txttype, a.applydt" & _
             " FROM   " & T_LAB001 & " a" & _
             " WHERE  " & DBW("a.testcd", pTestCd, 2) & _
             " AND    a.applydt =(SELECT MAX(b.applydt) FROM " & T_LAB001 & " b" & _
             " WHERE  b.testcd = a.testcd" & _
             " AND    " & DBW("b.applydt <= ", ApplyDt) & _
             " )"
    '
    DBOpenDyna strSQL
    If m_RecordSet.RecordCount > 0 Then
        GetTestNm = "" & m_RecordSet.Fields("TESTNM").Value
        If pTxtType <> "Default" Then
            GetTestNm = GetTestNm & vbTab & "" & m_RecordSet.Fields("TXTTYPE").Value
        End If
    Else
       GetTestNm = ""
    End If
    '
End Function

Public Function GetTestItemDataEntry(ByVal pTestCd As String, ByVal ApplyDt As String) As String
    
    Dim strSQL As String
    strSQL = " SELECT a.testcd, a.testnm" & _
             " FROM   " & T_LAB001 & " a" & _
             " WHERE  " & DBW("a.testcd", pTestCd, 2) & _
             " AND    a.applydt =(SELECT MAX(b.applydt) FROM " & T_LAB001 & " b" & _
                                " WHERE  b.testcd = a.testcd" & _
                                " AND    " & DBW("b.applydt <= ", ApplyDt) & _
                                " AND   (b.panelfg  = ''  or  b.panelfg is null) " & _
                                " AND   (b.detailfg = ''  or  b.detailfg is null) " & _
                                " )"
    '
    DBOpenDyna strSQL
    If m_RecordSet.RecordCount > 0 Then
        GetTestItemDataEntry = "" & m_RecordSet.Fields("TESTNM").Value
    Else
        GetTestItemDataEntry = ""
    End If
    '
End Function

Public Sub GetPanicDelta(ByRef pOraDynaset As Recordset, _
                         ByVal TestCd As String, ByVal SpcCd As String, ByVal ApplyDt As String)
    Dim strSQL As String
    strSQL = " SELECT a.avalval,a.panicfg,a.panicfrval,a.panictoval,a.deltafg,a.deltaval,a.deltaval2" & _
             " FROM   " & T_LAB004 & " a" & _
             " WHERE  " & DBW("a.testcd", TestCd, 2) & _
             " AND    " & DBW("a.spccd", SpcCd, 2) & _
             " AND    a.applydt IN(SELECT MAX(b.applydt) FROM " & T_LAB004 & " b" & _
             " WHERE  b.testcd = a.testcd" & _
             " AND    b.spccd  = a.spccd" & _
             " AND    " & DBW("b.applydt <= ", ApplyDt) & _
             " )"
    '
    DBOpenDyna strSQL
    Set pOraDynaset = m_RecordSet
    '
End Sub

Public Sub GetRefVal(ByRef pOraDynaset As Recordset, _
                     ByVal TestCd As String, ByVal SpcCd As String, ByVal ApplyDt As String, _
                     ByVal ApplySex As String, ByVal AgeDay As String)
    Dim strSQL As String
    
    strSQL = " SELECT a.applysex,a.applydt,a.refvalfrom,a.refcd,a.panicfrval,a.panictoval,a.arletfrval,a.arlettoval,a.arefvalfrom,a.refvalto,a.AMRFRVAL,a.AMRTOVAL  " & _
             " FROM   " & T_LAB005 & " a" & _
             " WHERE  " & DBW("a.testcd", TestCd, 2) & _
             " AND    " & DBW("a.spccd", SpcCd, 2) & _
             " AND    a.applydt = (SELECT MAX(b.applydt) FROM " & T_LAB005 & " b" & _
                                " WHERE b.testcd = a.testcd" & _
                                " AND   b.spccd = a.spccd AND b.applysex = a.applysex " & _
                                " AND   b.agefrom = a.agefrom AND b.ageto = a.ageto  " & _
                                " AND  (b.expdt is null or b.expdt = '' ))" & _
             " AND    " & DBW("a.applysex", ApplySex, 2) & _
             " AND    " & DBW("a.agefrom <=", AgeDay) & _
             " AND    " & DBW("a.ageto   >=", AgeDay) & _
             " AND    (a.expdt is null or a.expdt = '') " & _
             " ORDER BY a.applydt desc "
       '
    '-> 성은 기존의 성('M','F','-','A') 중 가져오되 없으면 다시 GetRefVal을 Call하여(ApplySex = 'B')
    '    'B'를 차선으로 택해서 가져온다.
    DBOpenDyna strSQL
    Set pOraDynaset = m_RecordSet
   '
End Sub

Public Function GetTempleteText1(ByVal cdindex As String, ByVal CdVal1 As String) As String
    
    Dim strSQL As String
    strSQL = " SELECT text1 FROM " & T_LAB034 & _
             " WHERE  " & DBW("cdindex", cdindex, 2) & _
             " AND    " & DBW("cdval1", CdVal1, 2)
    '
    DBOpenDyna strSQL
    If m_RecordSet.RecordCount > 0 Then
        GetTempleteText1 = "" & m_RecordSet.Fields("TEXT1").Value
    Else
        GetTempleteText1 = ""
    End If
    '
End Function

Public Function GetRstCdNm(ByVal TestCd As String, ByVal RstCd As String) As String
    
    Dim strSQL As String
    strSQL = " SELECT a.cdval2,a.field1" & _
             " FROM   " & T_LAB031 & " a " & _
             " WHERE  " & DBW("a.cdindex", LC2_ItemResult, 2) & _
             " AND    " & DBW("a.cdval1", TestCd, 2) & _
             " AND    " & DBW("a.cdval2", RstCd, 2)
    
    DBOpenDyna strSQL
    If m_RecordSet.RecordCount > 0 Then
        GetRstCdNm = "" & m_RecordSet.Fields("FIELD1").Value
    Else
        GetRstCdNm = ""
    End If
    '
End Function

Public Function GetEqpNm(ByVal pEqpCd As String) As String
    
    Dim strSQL As String
    
    strSQL = " SELECT eqpnm FROM " & T_LAB006 & _
             " WHERE  " & DBW("eqpcd", pEqpCd, 2)
    
    DBOpenDyna strSQL
    If Not m_RecordSet.EOF Then
        GetEqpNm = "" & m_RecordSet.Fields("EQPNM").Value
    Else
        GetEqpNm = ""
    End If
    '
End Function


Public Function IsWorkSeqExists(ByVal pWorkCd As String, ByVal pWorkDt As String, _
                                ByVal pWorkSeq As String) As Boolean
    
    Dim strSQL As String
    
    strSQL = " SELECT * FROM " & T_LAB301 & _
             " WHERE  " & DBW("workcd", pWorkCd, 2) & _
             " AND    " & DBW("workdt", pWorkDt, 2) & _
             " AND    " & DBW("workseq", pWorkSeq, 2)
    
    DBOpenDyna strSQL
    If Not m_RecordSet.EOF Then
        IsWorkSeqExists = True
    Else
        IsWorkSeqExists = False
    End If
    '
End Function

Public Function CheckWSData(ByRef pOraDynaset As Recordset, _
                            ByVal WorkArea As String, ByVal AccDt As String, _
                            ByVal accseq As String, Optional SetCd As String = "Default")

    Dim strSQL As String
    strSQL = " SELECT  a.testcd FROM " & T_LAB302 & " a, " & T_LAB008 & " b " & _
             " WHERE   a.workarea = '" & WorkArea & "' " & _
             " AND     a.accdt = '" & AccDt & "' " & _
             " AND     a.accseq = " & accseq & " " & _
             " AND     b.workcd = '" & SetCd & "' " & _
             " AND     b.testcd = a.testcd " & _
             " AND    (a.vfydt = '' or a.vfydt is null) "
    DBOpenDyna strSQL
    Set pOraDynaset = m_RecordSet

End Function

Public Function GetAccSeq(ByRef pOraDynaset As Recordset, _
                          ByVal WorkArea As String, ByVal AccDt As String, _
                          ByVal accseq As String, ByVal pOption As Integer)

    Dim strSQL As String
    If pOption = 1 Then
        strSQL = " SELECT  max(a.accseq) as NextAccSeq FROM " & T_LAB201 & " a " & _
                 " WHERE   " & DBW("a.workarea", WorkArea, 2) & _
                 " AND     " & DBW("a.accdt", AccDt, 2) & _
                 " AND     " & DBW("a.accseq < ", accseq) & _
                 " AND     " & DBW("a.stscd < ", enStsCd.StsCd_LIS_FinRst) & _
                 " AND     " & DBW("a.stscd > ", enStsCd.StsCd_LIS_Collection)
    Else
        strSQL = " SELECT  min(a.accseq) as NextAccSeq FROM " & T_LAB201 & " a " & _
                 " WHERE   " & DBW("a.workarea", WorkArea, 2) & _
                 " AND     " & DBW("a.accdt", AccDt, 2) & _
                 " AND     " & DBW("a.accseq > ", accseq) & _
                 " AND     " & DBW("a.stscd < ", enStsCd.StsCd_LIS_FinRst) & _
                 " AND     " & DBW("a.stscd > ", enStsCd.StsCd_LIS_Collection)
    End If
    DBOpenDyna strSQL
    Set pOraDynaset = m_RecordSet

End Function

Public Function GetCalType(ByRef pOraDynaset As Recordset, _
                          ByVal sTestCd As String, ByVal sSpcCd As String)

    Dim strSQL As String
    strSQL = " SELECT  field1 as CalType FROM " & T_LAB031 & " a " & _
             " WHERE   " & DBW("a.cdindex", LC2_Calculation, 2) & _
             " AND     " & DBW("a.cdval1", sTestCd, 2) & _
             " AND     " & DBW("a.cdval2", sSpcCd, 2)
    DBOpenDyna strSQL
    Set pOraDynaset = m_RecordSet

End Function

Public Sub GetResult(ByVal QueryType As Integer, _
                     ByRef pOraDynaset As Recordset, _
                     ByVal WorkArea As String, ByVal AccDt As String, _
                     ByVal accseq As String, Optional SetCd As String = "Default")
    
    Dim strSQL As String
    Dim strWSql As String
    Dim strWSql1 As String
    Dim strWSql2 As String
    
    '## 5.0.13: 이상대(2004-12-30)
    '   - 검사항목별 결과코드에 등록된 결과명도 조회
    strSQL = "SELECT a.workarea,a.accdt,a.accseq,a.testcd, a.rstval,a.rstcd," & _
             "       a.rstunit,a.hldiv,a.dpdiv,a.spccd,a.statfg,a.lastrst,b.field1 as lastrstval,a.lastvfydt,a.lastvfytm," & _
             "       a.lastvfyid,a.vfydt,a.vfytm,a.vfyid,a.mfyfg,a.grpfg,a.txtfg,a.rsttype,a.rstdiv," & _
             "       a.ptid,a.orddt,a.ordno,a.ordseq,a.detailfg,a.autofg,a.eqpcd,g.coldt,g.spcyy,g.spcno," & _
             "       h.testnm,h.txttype,h.rptseq," & _
             "       x.avalval,x.panicfg,x.panicfrval,x.panictoval,x.deltafg,x.deltaval,x.deltaval2," & _
             "       y.eqpnm " & _
             " FROM  " & T_LAB031 & " b, " & T_LAB001 & " h, " & T_LAB004 & " x, " & T_LAB006 & " y, " & _
                         T_LAB201 & " g, " & T_LAB302 & " a " & _
             " WHERE " & DBW("a.workarea", WorkArea, 2) & _
             " AND   " & DBW("a.accdt", AccDt, 2) & _
             " AND   " & DBW("a.accseq", accseq, 2) & _
             " AND   a.workarea = g.workarea" & _
             " AND   a.accdt    = g.accdt" & _
             " AND   a.accseq   = g.accseq" & _
             " AND   " & DBJ("y.eqpcd =* a.eqpcd")

    '*/ 결과등록 종류별 SQL조건문 기술
    Select Case QueryType
        Case RESULT_BY_ACCESSION:
            strWSql = " AND  (a.vfydt = '' or a.vfydt is null)"
            
        Case RESULT_BY_DIFFCOUNT:
            strWSql = " AND  (a.vfydt = '' or a.vfydt is null)" & _
                      " AND   EXISTS(SELECT f.field1 FROM " & T_LAB032 & " f " & _
                                    " WHERE " & DBW("f.cdindex", LC3_WBCDiffCode, 2) & _
                                    " AND   f.cdval1 = a.testcd )"
                                    
        Case RESULT_BY_WORKSHEET:
            strWSql1 = " AND  a.detailfg <> ' '  AND  a.detailfg is not null" & _
                       " AND  (a.vfydt = '' or a.vfydt is null)" & _
                       " AND   EXISTS(SELECT * FROM " & T_LAB302 & " k " & _
                                    " WHERE k.workarea = a.workarea " & _
                                    " AND   k.accdt    = a.accdt " & _
                                    " AND   k.accseq   = a.accseq " & _
                                    " AND   k.detailfg   = a.detailfg " & _
                                    " AND   " & DBW("k.workcd", SetCd, 2) & ") "
            strWSql = " AND  (a.vfydt = '' or a.vfydt is null)" & _
                      " AND  (a.detailfg = '' or a.detailfg is null)" & _
                      " AND   EXISTS(SELECT f.testcd FROM " & T_LAB008 & " f " & _
                                    " WHERE " & DBW("f.workcd", SetCd, 2) & _
                                    " AND   f.testcd = a.testcd " & _
                                    " AND   f.spccd = a.spccd " & _
                                    " AND   f.workarea = a.workarea )"
                                    
        Case RESULT_BY_EQUIPMENT:
            strWSql1 = " AND  a.detailfg <> ' '  AND  a.detailfg is not null" & _
                       " AND  (a.vfydt = '' or a.vfydt is null)" & _
                       " AND   EXISTS(SELECT j.eqpcd FROM " & T_LAB306 & " j " & _
                                    " WHERE " & DBW("j.eqpcd", SetCd, 2) & _
                                    " AND   j.spcyy = g.spcyy " & _
                                    " AND   j.spcno = g.spcno ) " & _
                       " AND   EXISTS(SELECT * FROM " & T_LAB302 & " k " & _
                                    " WHERE k.workarea = a.workarea " & _
                                    " AND   k.accdt    = a.accdt " & _
                                    " AND   k.accseq   = a.accseq " & _
                                    " AND   k.detailfg   = a.detailfg " & _
                                    " AND   " & DBW("k.eqpcd", SetCd, 2) & ") "
            strWSql = " AND   " & DBW("a.eqpcd = ", SetCd) & _
                      " AND  (a.detailfg = '' or a.detailfg is null)" & _
                      " AND  (a.vfydt = '' or a.vfydt is null)" & _
                      " AND   EXISTS(SELECT j.eqpcd FROM " & T_LAB306 & " j " & _
                                    " WHERE " & DBW("j.eqpcd", SetCd, 2) & _
                                    " AND   j.spcyy = g.spcyy " & _
                                    " AND   j.spcno = g.spcno ) "
        Case RESULT_BY_MODIFY:
            strWSql = " AND   a.vfydt <> ' ' "
             
    End Select
    '
    '*/ 화면 SORT
    If QueryType <> RESULT_BY_EQUIPMENT And _
       QueryType <> RESULT_BY_WORKSHEET Then
        strSQL = strSQL & strWSql & _
                 " AND   h.testcd = a.testcd" & _
                 " AND   h.applydt = (SELECT MAX(i.applydt) FROM " & T_LAB001 & " i" & _
                 "                    WHERE i.testcd = a.testcd and i.applydt<=a.orddt)" & _
                 " AND   x.testcd = a.testcd" & _
                 " AND   x.spccd = a.spccd" & _
                 " AND   x.applydt = (SELECT MAX(applydt) FROM " & T_LAB004 & _
                 "                    WHERE  testcd = x.testcd" & _
                 "                    AND    spccd  = x.spccd and applydt<=a.orddt)" & _
                 " AND " & DBJ("b.cdindex=*'" & LC2_ItemResult & "'") & _
                 " AND " & DBJ("a.testcd*=b.cdval1") & _
                 " AND " & DBJ("a.lastrst*=b.cdval2") & _
                 " ORDER BY h.rptseq"
    Else
        strWSql2 = strSQL & strWSql & _
                 " AND   h.testcd = a.testcd" & _
                 " AND   h.applydt = (SELECT MAX(i.applydt) FROM " & T_LAB001 & " i" & _
                 "                    WHERE i.testcd = a.testcd and i.applydt<=a.orddt)" & _
                 " AND   x.testcd = a.testcd" & _
                 " AND   x.spccd = a.spccd" & _
                 " AND   x.applydt = (SELECT MAX(applydt) FROM " & T_LAB004 & _
                 "                    WHERE  testcd = x.testcd" & _
                 "                    AND    spccd  = x.spccd and applydt<=a.orddt)" & _
                 " AND " & DBJ("b.cdindex=*'" & LC2_ItemResult & "'") & _
                 " AND " & DBJ("a.testcd*=b.cdval1") & _
                 " AND " & DBJ("a.lastrst*=b.cdval2") & _
                 " UNION "
                 
        strSQL = strWSql2 & strSQL & strWSql1 & _
                 " AND   h.testcd = a.testcd" & _
                 " AND   h.applydt = (SELECT MAX(i.applydt) FROM " & T_LAB001 & " i" & _
                 "                    WHERE i.testcd = a.testcd and i.applydt<=a.orddt)" & _
                 " AND   x.testcd = a.testcd" & _
                 " AND   x.spccd = a.spccd" & _
                 " AND   x.applydt = (SELECT MAX(applydt) FROM " & T_LAB004 & _
                 "                    WHERE  testcd = x.testcd" & _
                 "                    AND    spccd  = x.spccd and applydt<=a.orddt)" & _
                 " AND " & DBJ("b.cdindex=*'" & LC2_ItemResult & "'") & _
                 " AND " & DBJ("a.testcd*=b.cdval1") & _
                 " AND " & DBJ("a.lastrst*=b.cdval2") & _
                 " ORDER BY rptseq"
    End If
    
    DBOpenDyna strSQL
    Set pOraDynaset = m_RecordSet
End Sub

Public Sub GetItemResult(ByVal QueryType As String, _
                         ByRef pOraDynaset As Recordset, ByVal TestCd As String, _
                         ByVal SetCd As String, Optional BuildCd As String = "")
    Dim AccNo As String
    Dim FrNo As String
    Dim ToNo As String
    Dim WorkCd As String
    Dim WorkDt As String
    Dim strSQL As String
    Dim strWSql As String
    '
    If QueryType = "AccNo" Then
        AccNo = medGetP(SetCd, 2, vbTab)
        ToNo = medGetP(AccNo, 3, "-")
        AccNo = medGetP(SetCd, 1, vbTab)
        FrNo = medGetP(AccNo, 3, "-")
        WorkCd = "": WorkDt = ""
    Else
        AccNo = ""
        WorkCd = medGetP(SetCd, 1, vbTab)
        WorkDt = medGetP(SetCd, 2, vbTab)
        FrNo = medGetP(SetCd, 3, vbTab)
        ToNo = medGetP(SetCd, 4, vbTab)
    End If
    
    '## 5.0.16: 이상대(2004-12-31)
    '   - 최근결과값이 검사항목별 결과코드에 등록된 코드이면 결과명을 조회하도록 쿼리수정
    strSQL = " SELECT DISTINCT g.spcyy,g.spcno,a.workarea,a.accdt,a.accseq,a.testcd, a.rstval,a.rstcd," & _
             " a.rstunit,a.hldiv,a.dpdiv,a.spccd,a.statfg,a.lastrst,c.field1 as lastrstval, a.lastvfydt,a.lastvfytm," & _
             " a.lastvfyid,a.vfydt,a.vfytm,a.vfyid,a.mfyfg,a.grpfg,a.txtfg,a.rsttype,a.rstdiv," & _
             " a.ptid,a.orddt,a.ordno,a.ordseq,a.detailfg,a.autofg,a.eqpcd,g.coldt,g.rmkcd," & _
             " g.ageday,g.reqtotcnt,g.reqinputcnt,h.rptseq,h.txttype," & _
             " x.avalval,x.panicfg,x.panicfrval,x.panictoval,x.deltafg,x.deltaval,x.deltaval2," & _
             " i." & F_PTNM & " as PtNm," & F_DOB2("i") & " as DOB," & F_SEX2("i") & " as Sex"
    '
    '*/아이템별 결과등록 종류별 SQL조건문 기술 */
    ' 유의사항 : 1)상세항목은 제외 따라서 그룹 혹은 단일아이템만 대상이 되므로 Required결과만 처리.
    '                 2)상세항목이 제외됨에 따라 상세항목인 Required결과는 대상이 아님.
    Select Case QueryType
        Case "AccNo":
            strWSql = ", '' as workseq FROM " & T_HIS001 & " i, " & T_LAB031 & " c, " & T_LAB001 & " h," & T_LAB004 & " x, " & _
                                                T_LAB201 & " g," & T_LAB302 & " a, " & T_LAB102 & " b " & _
                      " WHERE  " & DBW("a.workarea", medGetP(AccNo, 1, "-"), 2) & _
                      " AND    " & DBW("a.accdt", medGetP(AccNo, 2, "-"), 2) & _
                      " AND    a.accseq BETWEEN " & DBV("accseq", FrNo) & " AND " & DBV("accseq", ToNo) & _
                      " AND    " & DBW("a.testcd", TestCd, 2) & _
                      " AND   (a.vfydt = ''  or  a.vfydt is null) " & _
                      " AND   (a.detailfg = ''  or  a.detailfg is null) " & _
                      " AND    g.workarea = a.workarea " & _
                      " AND    g.accdt    = a.accdt" & _
                      " AND    g.accseq   = a.accseq" & _
                      " AND    g.stscd IN (" & DBV("stscd", enStsCd.StsCd_LIS_Accession) & "," & _
                                               DBV("stscd", enStsCd.StsCd_LIS_InProcess) & ")" & _
                      " AND    i." & F_PTID & " = a.ptid"
               '" AND g.testdiv = '" & TST_RouTest & "'"
            If BuildCd <> "" Then strWSql = strWSql & " AND " & DBW("g.buildcd", BuildCd, 2)  '건물코드
        Case "WorkSheet":
            strWSql = ", k.workseq  FROM " & T_HIS001 & " i," & T_LAB031 & " c, " & T_LAB001 & " h," & T_LAB004 & " x, " & _
                                             T_LAB201 & " g," & T_LAB302 & " a," & T_LAB301 & " k, " & T_LAB102 & " b " & _
                      " WHERE  " & DBW("k.workdt", WorkDt, 2) & _
                      " AND    " & DBW("k.workcd", WorkCd, 2) & _
                      " AND    k.workseq BETWEEN " & DBV("workseq", FrNo) & " AND " & DBV("workseq", ToNo) & _
                      " AND    a.workarea = k.workarea" & _
                      " AND    a.accdt    = k.accdt" & _
                      " AND    a.accseq   = k.accseq" & _
                      " AND    " & DBW("a.testcd", TestCd, 2) & _
                      " AND   (a.vfydt = ''  or  a.vfydt is null ) " & _
                      " AND   (a.detailfg = ''  or  a.detailfg is null)" & _
                      " AND    g.workarea = a.workarea" & _
                      " AND    g.accdt    = a.accdt" & _
                      " AND    g.accseq   = a.accseq" & _
                      " AND    " & DBW("g.stscd", enStsCd.StsCd_LIS_InProcess, 2) & _
                      " AND    EXISTS(SELECT f.testcd FROM " & T_LAB008 & " f" & _
                                   " WHERE  " & DBW("f.workcd", WorkCd, 2) & _
                                   " AND    f.testcd = a.testcd " & _
                                   " AND    f.spccd = a.spccd" & _
                                   " AND    f.workarea = a.workarea)" & _
                                   " AND    i." & F_PTID & " = a.ptid"
    End Select
    '
    '*/ 화면 SORT
    strSQL = strSQL & strWSql & _
                    " AND   h.testcd  = a.testcd" & _
                    " AND   h.applydt = (SELECT MAX(j.applydt) FROM " & T_LAB001 & " j" & _
                                        " WHERE j.testcd  = a.testcd" & _
                                        " AND   j.applydt <=  g.coldt)" & _
                    " AND   x.testcd = a.testcd" & _
                    " AND   x.spccd  = a.spccd" & _
                    " AND   x.applydt = (SELECT MAX(applydt) FROM " & T_LAB004 & _
                                       " WHERE  testcd = x.testcd" & _
                                       " AND    spccd  = x.spccd)" & _
                    " AND   b.workarea = a.workarea AND b.accdt = a.accdt AND b.accseq = a.accseq " & _
                    " AND   " & DBJ(DBW("c.cdindex=*", LC2_ItemResult)) & _
                    " AND   " & DBJ("a.testcd*=c.cdval1") & _
                    " AND   " & DBJ("a.lastrst*=c.cdval2") & _
                    " ORDER BY workseq, a.workarea,a.accdt,a.accseq,h.rptseq"
                    
     Debug.Print strSQL
    DBOpenDyna strSQL
    Set pOraDynaset = m_RecordSet
    '
End Sub

Public Sub GetItemResult_Wm(ByVal QueryType As String, _
                         ByRef pOraDynaset As Recordset, ByVal TestCd As String, _
                         ByVal SetCd As String, Optional BuildCd As String = "")
    Dim AccNo As String
    Dim FrNo As String
    Dim ToNo As String
    Dim WorkCd As String
    Dim WorkDt As String
    Dim strSQL As String
    Dim strWSql As String
    Dim strDeptCd As String
    '
    If QueryType = "AccNo" Then
        AccNo = medGetP(SetCd, 2, vbTab)
        ToNo = medGetP(SetCd, 2, vbTab)
        AccNo = medGetP(SetCd, 1, vbTab)
        FrNo = medGetP(AccNo, 3, "-")
        WorkCd = "": WorkDt = ""
    Else
        AccNo = ""
        WorkCd = medGetP(SetCd, 1, vbTab)
        WorkDt = medGetP(SetCd, 2, vbTab)
        FrNo = medGetP(SetCd, 3, vbTab)
        ToNo = medGetP(SetCd, 4, vbTab)
    End If
    
    If TestCd = "ER" Then
        TestCd = "EM"
    End If
        
        
    Select Case TestCd
        Case "EM"
            strDeptCd = "j.deptcd"
        Case Else
            strDeptCd = "j.wardid"
    End Select
    
   
    '## 5.0.16: 이상대(2004-12-31)
    '   - 최근결과값이 검사항목별 결과코드에 등록된 코드이면 결과명을 조회하도록 쿼리수정
    strSQL = " SELECT DISTINCT g.spcyy,g.spcno,a.workarea,a.accdt,a.accseq,a.testcd, a.rstval,a.rstcd," & _
             " a.rstunit,a.hldiv,a.dpdiv,a.spccd,a.statfg,a.lastrst,c.field1 as lastrstval, a.lastvfydt,a.lastvfytm," & _
             " a.lastvfyid,a.vfydt,a.vfytm,a.vfyid,a.mfyfg,a.grpfg,a.txtfg,a.rsttype,a.rstdiv," & _
             " a.ptid,a.orddt,a.ordno,a.ordseq,a.detailfg,a.autofg,a.eqpcd,g.coldt,g.rmkcd," & _
             " g.ageday,g.reqtotcnt,g.reqinputcnt,h.rptseq,h.txttype," & _
             " x.avalval,x.panicfg,x.panicfrval,x.panictoval,x.deltafg,x.deltaval,x.deltaval2," & _
             " i." & F_PTNM & " as PtNm," & F_DOB2("i") & " as DOB," & F_SEX2("i") & " as Sex"
    '
    '*/아이템별 결과등록 종류별 SQL조건문 기술 */
    ' 유의사항 : 1)상세항목은 제외 따라서 그룹 혹은 단일아이템만 대상이 되므로 Required결과만 처리.
    '                 2)상세항목이 제외됨에 따라 상세항목인 Required결과는 대상이 아님.
    Select Case QueryType
        Case "AccNo":
            strWSql = ", '' as workseq FROM " & T_HIS001 & " i, " & T_LAB031 & " c, " & T_LAB001 & " h," & T_LAB004 & " x, " & _
                                                T_LAB201 & " g," & T_LAB302 & " a, " & T_LAB102 & " b, ORAS1.S2ORD101_V j " & _
                      " WHERE  a.workarea = '14' " & _
                      " AND   " & strDeptCd & " =  '" & TestCd & "' " & _
                      " AND    a.accdt = '" & AccNo & "' " & _
                      " AND    a.testcd = '" & ToNo & "' " & _
                      " AND   (a.vfydt = ''  or  a.vfydt is null) " & _
                      " AND   (a.detailfg = ''  or  a.detailfg is null) " & _
                      " AND    g.workarea = a.workarea " & _
                      " AND    g.accdt    = a.accdt" & _
                      " AND    g.accseq   = a.accseq AND a.PTID = j.PTID  AND a.ORDDT = j.ORDDT AND a.ORDNO = j.ORDNO " & _
                      " AND    g.stscd IN (" & DBV("stscd", enStsCd.StsCd_LIS_Accession) & "," & _
                                               DBV("stscd", enStsCd.StsCd_LIS_InProcess) & ")" & _
                      " AND    i." & F_PTID & " = a.ptid"
               '" AND g.testdiv = '" & TST_RouTest & "'"
            If BuildCd <> "" Then strWSql = strWSql & " AND " & DBW("g.buildcd", BuildCd, 2)  '건물코드
        Case "WorkSheet":
            strWSql = ", k.workseq  FROM " & T_HIS001 & " i," & T_LAB031 & " c, " & T_LAB001 & " h," & T_LAB004 & " x, " & _
                                             T_LAB201 & " g," & T_LAB302 & " a," & T_LAB301 & " k, " & T_LAB102 & " b " & _
                      " WHERE  " & DBW("k.workdt", WorkDt, 2) & _
                      " AND    " & DBW("k.workcd", WorkCd, 2) & _
                      " AND    k.workseq BETWEEN " & DBV("workseq", FrNo) & " AND " & DBV("workseq", ToNo) & _
                      " AND    a.workarea = k.workarea" & _
                      " AND    a.accdt    = k.accdt" & _
                      " AND    a.accseq   = k.accseq" & _
                      " AND    " & DBW("a.testcd", TestCd, 2) & _
                      " AND   (a.vfydt = ''  or  a.vfydt is null ) " & _
                      " AND   (a.detailfg = ''  or  a.detailfg is null)" & _
                      " AND    g.workarea = a.workarea" & _
                      " AND    g.accdt    = a.accdt" & _
                      " AND    g.accseq   = a.accseq" & _
                      " AND    " & DBW("g.stscd", enStsCd.StsCd_LIS_InProcess, 2) & _
                      " AND    EXISTS(SELECT f.testcd FROM " & T_LAB008 & " f" & _
                                   " WHERE  " & DBW("f.workcd", WorkCd, 2) & _
                                   " AND    f.testcd = a.testcd " & _
                                   " AND    f.spccd = a.spccd" & _
                                   " AND    f.workarea = a.workarea)" & _
                                   " AND    i." & F_PTID & " = a.ptid"
    End Select
    '
    '*/ 화면 SORT
'    strSQL = strSQL & strWSql & _
'                    " AND   h.testcd  = a.testcd" & _
'                    " AND   h.applydt = (SELECT MAX(j.applydt) FROM " & T_LAB001 & " j" & _
'                                        " WHERE j.testcd  = a.testcd" & _
'                                        " AND   j.applydt <=  g.coldt)" & _
'                    " AND   x.testcd = a.testcd" & _
'                    " AND   x.spccd  = a.spccd" & _
'                    " AND   x.applydt = (SELECT MAX(applydt) FROM " & T_LAB004 & _
'                                       " WHERE  testcd = x.testcd" & _
'                                       " AND    spccd  = x.spccd)" & _
'                    " AND   b.workarea = a.workarea AND b.accdt = a.accdt AND b.accseq = a.accseq " & _
'                    " AND   " & DBJ(DBW("c.cdindex=*", LC2_ItemResult)) & _
'                    " AND   " & DBJ("a.testcd*=c.cdval1") & _
'                    " AND   " & DBJ("a.lastrst*=c.cdval2") & _
'                    " ORDER BY workseq, a.workarea,a.accdt,a.accseq,h.rptseq"

    strSQL = strSQL & strWSql & _
                    " AND   h.testcd  = a.testcd" & _
                    " AND   x.testcd = a.testcd" & _
                    " AND   x.spccd  = a.spccd" & _
                    " AND   b.workarea = a.workarea AND b.accdt = a.accdt AND b.accseq = a.accseq " & _
                    " AND   " & DBJ(DBW("c.cdindex=*", LC2_ItemResult)) & _
                    " AND   " & DBJ("a.testcd*=c.cdval1") & _
                    " AND   " & DBJ("a.lastrst*=c.cdval2") & _
                    " ORDER BY workseq, a.workarea,a.accdt,a.accseq,h.rptseq"
                    
     Debug.Print strSQL
    DBOpenDyna strSQL
    Set pOraDynaset = m_RecordSet
    '
End Sub

Public Sub GetItemResult_Wm1(ByVal QueryType As String, _
                         ByRef pOraDynaset As Recordset, ByVal TestCd As String, _
                         ByVal SetCd As String, Optional BuildCd As String = "")
    Dim AccNo As String
    Dim FrNo As String
    Dim ToNo As String
    Dim WorkCd As String
    Dim WorkDt As String
    Dim strSQL As String
    Dim strWSql As String
    Dim strDeptCd As String
    
    If QueryType = "AccNo" Then
        AccNo = medGetP(SetCd, 2, vbTab)
        ToNo = medGetP(SetCd, 2, vbTab)
        AccNo = medGetP(SetCd, 1, vbTab)
        FrNo = medGetP(AccNo, 3, "-")
        WorkCd = "": WorkDt = ""
    Else
        AccNo = ""
        WorkCd = medGetP(SetCd, 1, vbTab)
        WorkDt = medGetP(SetCd, 2, vbTab)
        FrNo = medGetP(SetCd, 3, vbTab)
        ToNo = medGetP(SetCd, 4, vbTab)
    End If
    
    If TestCd = "ER" Then
        TestCd = "EM"
    End If
        
        
    Select Case TestCd
        Case "EM"
            strDeptCd = "j.deptcd"
        Case Else
            strDeptCd = "j.wardid"
    End Select
        
    
    '## 5.0.16: 이상대(2004-12-31)
    '   - 최근결과값이 검사항목별 결과코드에 등록된 코드이면 결과명을 조회하도록 쿼리수정
    strSQL = " SELECT DISTINCT g.spcyy,g.spcno,a.workarea,a.accdt,a.accseq,a.testcd, a.rstval,a.rstcd," & _
             " a.rstunit,a.hldiv,a.dpdiv,a.spccd,a.statfg,a.lastrst,c.field1 as lastrstval, a.lastvfydt,a.lastvfytm," & _
             " a.lastvfyid,a.vfydt,a.vfytm,a.vfyid,a.mfyfg,a.grpfg,a.txtfg,a.rsttype,a.rstdiv," & _
             " a.ptid,a.orddt,a.ordno,a.ordseq,a.detailfg,a.autofg,a.eqpcd,g.coldt,g.rmkcd," & _
             " g.ageday,g.reqtotcnt,g.reqinputcnt,h.rptseq,h.txttype," & _
             " x.avalval,x.panicfg,x.panicfrval,x.panictoval,x.deltafg,x.deltaval,x.deltaval2," & _
             " i." & F_PTNM & " as PtNm," & F_DOB2("i") & " as DOB," & F_SEX2("i") & " as Sex"
    '
    '*/아이템별 결과등록 종류별 SQL조건문 기술 */
    ' 유의사항 : 1)상세항목은 제외 따라서 그룹 혹은 단일아이템만 대상이 되므로 Required결과만 처리.
    '                 2)상세항목이 제외됨에 따라 상세항목인 Required결과는 대상이 아님.
    Select Case QueryType
        Case "AccNo":
            strWSql = ", '' as workseq FROM " & T_HIS001 & " i, " & T_LAB031 & " c, " & T_LAB001 & " h," & T_LAB004 & " x, " & _
                                                T_LAB201 & " g," & T_LAB302 & " a, " & T_LAB102 & " b, ORAS1.S2ORD101_V j " & _
                      " WHERE  a.workarea = '14' " & _
                      " AND   " & strDeptCd & " =  '" & TestCd & "' " & _
                      " AND    a.accdt = '" & AccNo & "' " & _
                      " AND    a.testcd = '" & ToNo & "' " & _
                      " AND   (a.detailfg = ''  or  a.detailfg is null) " & _
                      " AND    g.workarea = a.workarea " & _
                      " AND    g.accdt    = a.accdt" & _
                      " AND    g.accseq   = a.accseq AND a.PTID = j.PTID  AND a.ORDDT = j.ORDDT AND a.ORDNO = j.ORDNO " & _
                      " AND    g.stscd IN (" & DBV("stscd", enStsCd.StsCd_LIS_FinRst) & "," & _
                                               DBV("stscd", enStsCd.StsCd_LIS_FinRst) & ")" & _
                      " AND    i." & F_PTID & " = a.ptid"
               '" AND g.testdiv = '" & TST_RouTest & "'"
            If BuildCd <> "" Then strWSql = strWSql & " AND " & DBW("g.buildcd", BuildCd, 2)  '건물코드
        Case "WorkSheet":
            strWSql = ", k.workseq  FROM " & T_HIS001 & " i," & T_LAB031 & " c, " & T_LAB001 & " h," & T_LAB004 & " x, " & _
                                             T_LAB201 & " g," & T_LAB302 & " a," & T_LAB301 & " k, " & T_LAB102 & " b " & _
                      " WHERE  " & DBW("k.workdt", WorkDt, 2) & _
                      " AND    " & DBW("k.workcd", WorkCd, 2) & _
                      " AND    k.workseq BETWEEN " & DBV("workseq", FrNo) & " AND " & DBV("workseq", ToNo) & _
                      " AND    a.workarea = k.workarea" & _
                      " AND    a.accdt    = k.accdt" & _
                      " AND    a.accseq   = k.accseq" & _
                      " AND    " & DBW("a.testcd", TestCd, 2) & _
                      " AND   (a.vfydt = ''  or  a.vfydt is null ) " & _
                      " AND   (a.detailfg = ''  or  a.detailfg is null)" & _
                      " AND    g.workarea = a.workarea" & _
                      " AND    g.accdt    = a.accdt" & _
                      " AND    g.accseq   = a.accseq" & _
                      " AND    " & DBW("g.stscd", enStsCd.StsCd_LIS_InProcess, 2) & _
                      " AND    EXISTS(SELECT f.testcd FROM " & T_LAB008 & " f" & _
                                   " WHERE  " & DBW("f.workcd", WorkCd, 2) & _
                                   " AND    f.testcd = a.testcd " & _
                                   " AND    f.spccd = a.spccd" & _
                                   " AND    f.workarea = a.workarea)" & _
                                   " AND    i." & F_PTID & " = a.ptid"
    End Select
    '
    '*/ 화면 SORT
'    strSQL = strSQL & strWSql & _
'                    " AND   h.testcd  = a.testcd" & _
'                    " AND   h.applydt = (SELECT MAX(j.applydt) FROM " & T_LAB001 & " j" & _
'                                        " WHERE j.testcd  = a.testcd" & _
'                                        " AND   j.applydt <=  g.coldt)" & _
'                    " AND   x.testcd = a.testcd" & _
'                    " AND   x.spccd  = a.spccd" & _
'                    " AND   x.applydt = (SELECT MAX(applydt) FROM " & T_LAB004 & _
'                                       " WHERE  testcd = x.testcd" & _
'                                       " AND    spccd  = x.spccd)" & _
'                    " AND   b.workarea = a.workarea AND b.accdt = a.accdt AND b.accseq = a.accseq " & _
'                    " AND   " & DBJ(DBW("c.cdindex=*", LC2_ItemResult)) & _
'                    " AND   " & DBJ("a.testcd*=c.cdval1") & _
'                    " AND   " & DBJ("a.lastrst*=c.cdval2") & _
'                    " ORDER BY workseq, a.workarea,a.accdt,a.accseq,h.rptseq"
                    
    strSQL = strSQL & strWSql & _
                    " AND   h.testcd  = a.testcd" & _
                    " AND   x.testcd = a.testcd" & _
                    " AND   x.spccd  = a.spccd" & _
                    " AND   b.workarea = a.workarea AND b.accdt = a.accdt AND b.accseq = a.accseq " & _
                    " AND   " & DBJ(DBW("c.cdindex=*", LC2_ItemResult)) & _
                    " AND   " & DBJ("a.testcd*=c.cdval1") & _
                    " AND   " & DBJ("a.lastrst*=c.cdval2") & _
                    " ORDER BY workseq, a.workarea,a.accdt,a.accseq,h.rptseq"
                    
    Debug.Print strSQL
    DBOpenDyna strSQL
    Set pOraDynaset = m_RecordSet
    '
End Sub

Public Sub GetItemResult_2014(ByVal QueryType As String, _
                         ByRef pOraDynaset As Recordset, ByVal TestCd As String, _
                         ByVal SetCd As String, ByVal PoctCd As String, Optional BuildCd As String = "")
    Dim AccNo As String
    Dim FrNo As String
    Dim ToNo As String
    Dim WorkCd As String
    Dim WorkDt As String
    Dim strSQL As String
    Dim strWSql As String
    '
    If QueryType = "AccNo" Then
        AccNo = medGetP(SetCd, 2, vbTab)
        ToNo = medGetP(AccNo, 3, "-")
        AccNo = medGetP(SetCd, 1, vbTab)
        FrNo = medGetP(AccNo, 3, "-")
        WorkCd = "": WorkDt = ""
    Else
        AccNo = ""
        WorkCd = medGetP(SetCd, 1, vbTab)
        WorkDt = medGetP(SetCd, 2, vbTab)
        FrNo = medGetP(SetCd, 3, vbTab)
        ToNo = medGetP(SetCd, 4, vbTab)
    End If
    
    '## 5.0.16: 이상대(2004-12-31)
    '   - 최근결과값이 검사항목별 결과코드에 등록된 코드이면 결과명을 조회하도록 쿼리수정
    strSQL = " SELECT DISTINCT g.spcyy,g.spcno,a.workarea,a.accdt,a.accseq,a.testcd, a.rstval,a.rstcd," & _
             " a.rstunit,a.hldiv,a.dpdiv,a.spccd,a.statfg,a.lastrst,c.field1 as lastrstval, a.lastvfydt,a.lastvfytm," & _
             " a.lastvfyid,a.vfydt,a.vfytm,a.vfyid,a.mfyfg,a.grpfg,a.txtfg,a.rsttype,a.rstdiv," & _
             " a.ptid,a.orddt,a.ordno,a.ordseq,a.detailfg,a.autofg,a.eqpcd,g.coldt,g.rmkcd," & _
             " g.ageday,g.reqtotcnt,g.reqinputcnt,h.rptseq,h.txttype," & _
             " x.avalval,x.panicfg,x.panicfrval,x.panictoval,x.deltafg,x.deltaval,x.deltaval2," & _
             " i." & F_PTNM & " as PtNm," & F_DOB2("i") & " as DOB," & F_SEX2("i") & " as Sex"
    '       " i." & F_PTNM & " as PtNm," & F_DOB2 & " as DOB," & F_SEX2 & " as Sex"
    
    '
    '*/아이템별 결과등록 종류별 SQL조건문 기술 */
    ' 유의사항 : 1)상세항목은 제외 따라서 그룹 혹은 단일아이템만 대상이 되므로 Required결과만 처리.
    '                 2)상세항목이 제외됨에 따라 상세항목인 Required결과는 대상이 아님.
    Select Case QueryType
        Case "AccNo":
            strWSql = ", '' as workseq FROM " & T_HIS001 & " i, " & T_LAB031 & " c, " & T_LAB001 & " h," & T_LAB004 & " x, " & _
                                                T_LAB201 & " g," & T_LAB302 & " a, " & T_LAB102 & " b " & _
                      " WHERE  " & DBW("a.workarea", medGetP(AccNo, 1, "-"), 2) & _
                      " AND    " & DBW("a.accdt", medGetP(AccNo, 2, "-"), 2) & _
                      " AND    a.accseq BETWEEN " & DBV("accseq", FrNo) & " AND " & DBV("accseq", ToNo) & _
                      " AND    " & DBW("a.testcd", TestCd, 2) & _
                      " AND   (a.vfydt = ''  or  a.vfydt is null) " & _
                      " AND   (a.detailfg = ''  or  a.detailfg is null) " & _
                      " AND    g.workarea = a.workarea " & _
                      " AND    g.accdt    = a.accdt" & _
                      " AND    g.accseq   = a.accseq" & _
                      " AND    g.stscd IN (" & DBV("stscd", enStsCd.StsCd_LIS_Accession) & "," & _
                                               DBV("stscd", enStsCd.StsCd_LIS_InProcess) & ")" & _
                      " AND    i." & F_PTID & " = a.ptid"
               '" AND g.testdiv = '" & TST_RouTest & "'"
            If BuildCd <> "" Then strWSql = strWSql & " AND " & DBW("g.buildcd", BuildCd, 2)  '건물코드
        Case "WorkSheet":
            strWSql = ", k.workseq  FROM " & T_HIS001 & " i," & T_LAB031 & " c, " & T_LAB001 & " h," & T_LAB004 & " x, " & _
                                             T_LAB201 & " g," & T_LAB302 & " a," & T_LAB301 & " k, " & T_LAB102 & " b " & _
                      " WHERE  " & DBW("k.workdt", WorkDt, 2) & _
                      " AND    " & DBW("k.workcd", WorkCd, 2) & _
                      " AND    k.workseq BETWEEN " & DBV("workseq", FrNo) & " AND " & DBV("workseq", ToNo) & _
                      " AND    a.workarea = k.workarea" & _
                      " AND    a.accdt    = k.accdt" & _
                      " AND    a.accseq   = k.accseq" & _
                      " AND    " & DBW("a.testcd", TestCd, 2) & _
                      " AND   (a.vfydt = ''  or  a.vfydt is null ) " & _
                      " AND   (a.detailfg = ''  or  a.detailfg is null)" & _
                      " AND    g.workarea = a.workarea" & _
                      " AND    g.accdt    = a.accdt" & _
                      " AND    g.accseq   = a.accseq "
            Select Case PoctCd
             Case "0" '전체
             Case "1" '제활센타
                  strWSql = strWSql & " AND    g.wardid In ('94','95','96') "
             Case "2" '본관
                  strWSql = strWSql & " AND    g.wardid Not In ('94','95','96') "
            End Select
                      
            strWSql = strWSql & " AND    " & DBW("g.stscd", enStsCd.StsCd_LIS_InProcess, 2) & _
                                " AND    EXISTS(SELECT f.testcd FROM " & T_LAB008 & " f" & _
                                " WHERE  " & DBW("f.workcd", WorkCd, 2) & _
                                " AND    f.testcd = a.testcd " & _
                                " AND    f.spccd = a.spccd" & _
                                " AND    f.workarea = a.workarea)" & _
                                " AND    i." & F_PTID & " = a.ptid"
    End Select
    '
    '*/ 화면 SORT
    strSQL = strSQL & strWSql & _
                    " AND   h.testcd  = a.testcd" & _
                    " AND   h.applydt = (SELECT MAX(j.applydt) FROM " & T_LAB001 & " j" & _
                                        " WHERE j.testcd  = a.testcd" & _
                                        " AND   j.applydt <=  g.coldt)" & _
                    " AND   x.testcd = a.testcd" & _
                    " AND   x.spccd  = a.spccd" & _
                    " AND   x.applydt = (SELECT MAX(applydt) FROM " & T_LAB004 & _
                                       " WHERE  testcd = x.testcd" & _
                                       " AND    spccd  = x.spccd)" & _
                    " AND   b.workarea = a.workarea AND b.accdt = a.accdt AND b.accseq = a.accseq " & _
                    " AND   " & DBJ(DBW("c.cdindex=*", LC2_ItemResult)) & _
                    " AND   " & DBJ("a.testcd*=c.cdval1") & _
                    " AND   " & DBJ("a.lastrst*=c.cdval2") & _
                    " ORDER BY workseq, a.workarea,a.accdt,a.accseq,h.rptseq"
                    
     Debug.Print strSQL
    DBOpenDyna strSQL
    Set pOraDynaset = m_RecordSet
    '
End Sub

Public Sub GetItemResult_2015(ByVal QueryType As String, _
                         ByRef pOraDynaset As Recordset, ByVal TestCd As String, _
                         ByVal SetCd As String, ByVal PoctCd As String, Optional BuildCd As String = "")
    Dim AccNo As String
    Dim FrNo As String
    Dim ToNo As String
    Dim WorkCd As String
    Dim WorkDt As String
    Dim strSQL As String
    Dim strWSql As String
    '
    If QueryType = "AccNo" Then
        AccNo = medGetP(SetCd, 2, vbTab)
        ToNo = medGetP(AccNo, 3, "-")
        AccNo = medGetP(SetCd, 1, vbTab)
        FrNo = medGetP(AccNo, 3, "-")
        WorkCd = "": WorkDt = ""
    Else
        AccNo = ""
        WorkCd = medGetP(SetCd, 1, vbTab)
        WorkDt = medGetP(SetCd, 2, vbTab)
        FrNo = medGetP(SetCd, 3, vbTab)
        ToNo = medGetP(SetCd, 4, vbTab)
    End If
    
    '## 5.0.16: 이상대(2004-12-31)
    '   - 최근결과값이 검사항목별 결과코드에 등록된 코드이면 결과명을 조회하도록 쿼리수정
    strSQL = " SELECT DISTINCT g.spcyy,g.spcno,a.workarea,a.accdt,a.accseq,a.testcd, a.rstval,a.rstcd," & _
             " a.rstunit,a.hldiv,a.dpdiv,a.spccd,a.statfg,a.lastrst,c.field1 as lastrstval, a.lastvfydt,a.lastvfytm," & _
             " a.lastvfyid,a.vfydt,a.vfytm,a.vfyid,a.mfyfg,a.grpfg,a.txtfg,a.rsttype,a.rstdiv," & _
             " a.ptid,a.orddt,a.ordno,a.ordseq,a.detailfg,a.autofg,a.eqpcd,g.coldt,g.rmkcd," & _
             " g.ageday,g.reqtotcnt,g.reqinputcnt,h.rptseq,h.txttype," & _
             " x.avalval,x.panicfg,x.panicfrval,x.panictoval,x.deltafg,x.deltaval,x.deltaval2," & _
             " i." & F_PTNM & " as PtNm," & F_DOB2("i") & " as DOB," & F_SEX2("i") & " as Sex"
    '       " i." & F_PTNM & " as PtNm," & F_DOB2 & " as DOB," & F_SEX2 & " as Sex"
    
    '
    '*/아이템별 결과등록 종류별 SQL조건문 기술 */
    ' 유의사항 : 1)상세항목은 제외 따라서 그룹 혹은 단일아이템만 대상이 되므로 Required결과만 처리.
    '                 2)상세항목이 제외됨에 따라 상세항목인 Required결과는 대상이 아님.
    Select Case QueryType
        Case "AccNo":
            strWSql = ", '' as workseq FROM " & T_HIS001 & " i, " & T_LAB031 & " c, " & T_LAB001 & " h," & T_LAB004 & " x, " & _
                                                T_LAB201 & " g," & T_LAB302 & " a, " & T_LAB102 & " b " & _
                      " WHERE  " & DBW("a.workarea", medGetP(AccNo, 1, "-"), 2) & _
                      " AND    " & DBW("a.accdt", medGetP(AccNo, 2, "-"), 2) & _
                      " AND    a.accseq BETWEEN " & DBV("accseq", FrNo) & " AND " & DBV("accseq", ToNo) & _
                      " AND    " & DBW("a.testcd", TestCd, 2) & _
                      " AND   (a.vfydt = ''  or  a.vfydt is null) " & _
                      " AND   (a.detailfg = ''  or  a.detailfg is null) " & _
                      " AND    g.workarea = a.workarea " & _
                      " AND    g.accdt    = a.accdt" & _
                      " AND    g.accseq   = a.accseq" & _
                      " AND    g.stscd IN (" & DBV("stscd", enStsCd.StsCd_LIS_Accession) & "," & _
                                               DBV("stscd", enStsCd.StsCd_LIS_InProcess) & ")" & _
                      " AND    i." & F_PTID & " = a.ptid"
               '" AND g.testdiv = '" & TST_RouTest & "'"
            If BuildCd <> "" Then strWSql = strWSql & " AND " & DBW("g.buildcd", BuildCd, 2)  '건물코드
        Case "WorkSheet":
            strWSql = ", k.workseq  FROM " & T_HIS001 & " i," & T_LAB031 & " c, " & T_LAB001 & " h," & T_LAB004 & " x, " & _
                                             T_LAB201 & " g," & T_LAB302 & " a," & T_LAB301 & " k, " & T_LAB102 & " b " & _
                      " WHERE  " & DBW("k.workdt", WorkDt, 2) & _
                      " AND    " & DBW("k.workcd", WorkCd, 2) & _
                      " AND    k.workseq BETWEEN " & DBV("workseq", FrNo) & " AND " & DBV("workseq", ToNo) & _
                      " AND    a.workarea = k.workarea" & _
                      " AND    a.accdt    = k.accdt" & _
                      " AND    a.accseq   = k.accseq" & _
                      " AND    " & DBW("a.testcd", TestCd, 2) & _
                      " AND   (a.vfydt = ''  or  a.vfydt is null ) " & _
                      " AND   (a.detailfg = ''  or  a.detailfg is null)" & _
                      " AND    g.workarea = a.workarea" & _
                      " AND    g.accdt    = a.accdt" & _
                      " AND    g.accseq   = a.accseq "
            Select Case PoctCd
             Case "0" '전체
             Case "1" '제활센타
                  strWSql = strWSql & " AND    g.wardid In ('94','95','96') "
             Case "2" '본관
                  strWSql = strWSql & " AND    g.wardid Not In ('94','95','96') "
            End Select
                      
            strWSql = strWSql & " AND    " & DBW("g.stscd", enStsCd.StsCd_LIS_InProcess, 2) & _
                                " AND    EXISTS(SELECT f.testcd FROM " & T_LAB008 & " f" & _
                                " WHERE  " & DBW("f.workcd", WorkCd, 2) & _
                                " AND    f.testcd = a.testcd " & _
                                " AND    f.spccd = a.spccd" & _
                                " AND    f.workarea = a.workarea)" & _
                                " AND    i." & F_PTID & " = a.ptid"
    End Select
    '
    '*/ 화면 SORT
    strSQL = strSQL & strWSql & _
                    " AND   h.testcd  = a.testcd" & _
                    " AND   h.applydt = (SELECT MAX(j.applydt) FROM " & T_LAB001 & " j" & _
                                        " WHERE j.testcd  = a.testcd" & _
                                        " AND   j.applydt <=  g.coldt)" & _
                    " AND   x.testcd = a.testcd" & _
                    " AND   x.spccd  = a.spccd" & _
                    " AND   x.applydt = (SELECT MAX(applydt) FROM " & T_LAB004 & _
                                       " WHERE  testcd = x.testcd" & _
                                       " AND    spccd  = x.spccd)" & _
                    " AND   b.workarea = a.workarea AND b.accdt = a.accdt AND b.accseq = a.accseq " & _
                    " AND   " & DBJ(DBW("c.cdindex=*", LC2_ItemResult)) & _
                    " AND   " & DBJ("a.testcd*=c.cdval1") & _
                    " AND   " & DBJ("a.lastrst*=c.cdval2") & _
                    " ORDER BY workseq, a.workarea,a.accdt,a.accseq,h.rptseq"
                    
     Debug.Print strSQL
    DBOpenDyna strSQL
    Set pOraDynaset = m_RecordSet
    '
End Sub

Public Function GetRstTxt(ByVal WorkArea As String, _
                          ByVal AccDt As String, ByVal accseq As String, _
                          Optional TestCd As String = "Default") As String
    Dim strSQL As String
    Dim ii As Integer
    '
    GetRstTxt = ""
    If TestCd = "Default" Then
        'FootNote내역
        strSQL = " SELECT rsttxt FROM " & T_LAB304 & _
                 " WHERE  " & DBW("workarea", WorkArea, 2) & _
                 " AND    " & DBW("accdt", AccDt, 2) & _
                 " AND    " & DBW("accseq", accseq, 2) & _
                 " ORDER BY seq"
    Else
        'Text결과
        strSQL = " SELECT rsttxt FROM " & T_LAB303 & _
                 " WHERE  " & DBW("workarea", WorkArea, 2) & _
                 " AND    " & DBW("accdt", AccDt, 2) & _
                 " AND    " & DBW("accseq", accseq, 2) & _
                 " AND    " & DBW("testcd", TestCd, 2)
    End If
    '
    DBOpenDyna strSQL
    With m_RecordSet
        If .RecordCount > 0 Then
            .MoveFirst
            ii = 1
            Do Until .EOF
                If .RecordCount = 1 Or ii = 1 Then
                    GetRstTxt = "" & m_RecordSet.Fields("RSTTXT").Value
                Else
                    GetRstTxt = GetRstTxt & vbNewLine & String(65, "-") & vbNewLine & _
                                "" & m_RecordSet.Fields("RSTTXT").Value
                End If
                .MoveNext
                ii = ii + 1
            Loop
        End If
    End With
    '
End Function

Public Function GetSuppRstTxt(ByVal WorkArea As String, _
                              ByVal AccDt As String, ByVal accseq As String, _
                              ByVal TestCd As String) As String
    Dim strSQL As String
    '
    GetSuppRstTxt = ""
    'Supplementaryt 결과
    strSQL = " SELECT rsttxt,mfydt,mfytm FROM " & T_LAB305 & _
             " WHERE  " & DBW("workarea", WorkArea, 2) & _
             " AND    " & DBW("accdt", AccDt, 2) & _
             " AND    " & DBW("accseq", accseq, 2) & _
             " AND    " & DBW("testcd", TestCd, 2) & _
             " ORDER BY seq desc"
    '
    DBOpenDyna strSQL
    With m_RecordSet
        If .RecordCount > 0 Then
'            .MoveFirst
'            GetSuppRstTxt = vbNewLine
'            Do Until .EOF
                'GetSuppRstTxt = GetSuppRstTxt & vbNewLine & String(67, "-") & _
                                vbNewLine & "" & m_RecordSet.Fields("RSTTXT").Value
                GetSuppRstTxt = GetSuppRstTxt & m_RecordSet.Fields("rsttxt").Value & vbCr
'                GetSuppRstTxt = GetSuppRstTxt & "♣ 수정일시 : " & Format(m_RecordSet.Fields("mfydt").Value & "", "####-##-##") & "  "
'                GetSuppRstTxt = GetSuppRstTxt & Format(m_RecordSet.Fields("mfytm").Value & "", "0#:##:##") & vbNewLine
'                .MoveNext
'            Loop
            GetSuppRstTxt = Mid(GetSuppRstTxt, 1, Len(GetSuppRstTxt) - 1)
        End If
    End With
    '
End Function
 
Public Function UpdatableLAB205(ByVal WorkArea As String, ByVal AccDt As String, _
                                ByVal accseq As String, ByVal TestCd As String) As Boolean
    Dim SqlStmt As String
    '
    SqlStmt = " SELECT * FROM " & T_LAB205 & _
              " WHERE  " & DBW("workarea", WorkArea, 2) & _
              " AND    " & DBW("accdt", AccDt, 2) & _
              " AND    " & DBW("accseq", accseq, 2) & _
              " AND    " & DBW("testcd", TestCd, 2)
    UpdatableLAB205 = DBExists(SqlStmt)
    '
End Function
 
Public Function UpdatableLAB303(ByVal WorkArea As String, ByVal AccDt As String, _
                                ByVal accseq As String, ByVal TestCd As String) As Boolean
    Dim SqlStmt As String
    '
    SqlStmt = " SELECT * FROM " & T_LAB303 & _
              " WHERE  " & DBW("workarea", WorkArea, 2) & _
              " AND    " & DBW("accdt", AccDt, 2) & _
              " AND    " & DBW("accseq", accseq, 2) & _
              " AND    " & DBW("testcd", TestCd, 2)
    UpdatableLAB303 = DBExists(SqlStmt)
    '
End Function
 
Public Function UpdatableLAB304(ByVal WorkArea As String, ByVal AccDt As String, _
                                ByVal accseq As String, ByVal Seq As String) As Boolean
    Dim SqlStmt As String
    '
    SqlStmt = " SELECT * FROM " & T_LAB304 & _
              " WHERE  " & DBW("workarea", WorkArea, 2) & _
              " AND    " & DBW("accdt", AccDt, 2) & _
              " AND    " & DBW("accseq", accseq, 2) & _
              " AND    " & DBW("seq", Seq, 2)
    UpdatableLAB304 = DBExists(SqlStmt)
    '
End Function
Public Function GetResultModifySeq(ByVal TableNm As String, ByVal WorkArea As String, _
                                   ByVal AccDt As String, ByVal accseq As String, ByVal sTestCd As String, _
                                   ByRef mVfyDt As String, ByRef mVfyTm As String, ByRef mVfyID As String, _
                                   ByRef mSeq As String) As String
    Dim strSQL As String
    Dim strMaxSeq As String
    
    GetResultModifySeq = "1"
    strSQL = " SELECT seq AS MAXSEQ,mfydt,mfytm,mfyid FROM " & TableNm & _
             " WHERE  " & DBW("workarea", WorkArea, 2) & _
             " AND    " & DBW("accdt", AccDt, 2) & _
             " AND    " & DBW("accseq", accseq, 2) & _
             " AND    " & DBW(" testcd", sTestCd, 2) & _
             " ORDER BY seq DESC"
             
    DBOpenDyna strSQL
    If m_RecordSet.RecordCount > 0 Then
        strMaxSeq = "" & m_RecordSet.Fields("MAXSEQ").Value
        If strMaxSeq <> "" Then
            GetResultModifySeq = CStr(Val(strMaxSeq) + 1)
        End If
        mVfyDt = m_RecordSet.Fields("mfydt").Value & ""
        mVfyTm = m_RecordSet.Fields("mfytm").Value & ""
        mVfyID = m_RecordSet.Fields("mfyid").Value & ""
    End If
    
    mSeq = GetResultModifySeq
End Function

Public Function GetResultSeq(ByVal TableNm As String, ByVal WorkArea As String, _
                             ByVal AccDt As String, ByVal accseq As String) As String
    Dim strSQL As String
    Dim strMaxSeq As String
    GetResultSeq = "1"
    strSQL = " SELECT MAX(seq) AS MAXSEQ FROM " & TableNm & _
             " WHERE  " & DBW("workarea", WorkArea, 2) & _
             " AND    " & DBW("accdt", AccDt, 2) & _
             " AND    " & DBW("accseq", accseq, 2)
    DBOpenDyna strSQL
    If m_RecordSet.RecordCount > 0 Then
        strMaxSeq = "" & m_RecordSet.Fields("MAXSEQ").Value
        If strMaxSeq <> "" Then
            GetResultSeq = CStr(Val(strMaxSeq) + 1)
        End If
    End If
End Function

Public Function IsOrderFollowUp(ByVal PtId As String, ByVal OrdDt As String, _
                                ByVal OrdNo As String, ByVal OrdSeq As String) As Boolean
    Dim SqlStmt As String
    Dim vfyTotCnt As Long
    Dim vfyInputCnt As Long
    IsOrderFollowUp = False
    vfyTotCnt = 0: vfyInputCnt = 0
    SqlStmt = " SELECT COUNT(a.testcd) AS VfyTotCnt FROM " & T_LAB302 & " a" & _
              " WHERE  " & DBW("a.ptid", PtId, 2) & _
              " AND    " & DBW("a.orddt", OrdDt, 2) & _
              " AND    " & DBW("a.ordno", OrdNo, 2) & _
              " AND    " & DBW("a.ordseq", OrdSeq, 2) & _
              " AND   (a.vfydt = '' or a.vfydt is null) "
       '" AND a.rstdiv <> '*'"
    DBOpenDyna SqlStmt
    If m_RecordSet.RecordCount > 0 Then
        vfyTotCnt = Val("" & m_RecordSet.Fields("VFYTOTCNT").Value)
        If vfyTotCnt = 0 Then
            IsOrderFollowUp = True
        End If
    End If
'    m_RecordSet.Close
    Set m_RecordSet = Nothing
    '
End Function
 
'Public Function ResultDataEntry(ByRef objRst As clsPatientInfo, ByVal VfyDt As String, _
                                ByVal VfyTm As String, ByVal VfyId As String, _
                                ByVal FootNoteId As String) As Boolean

Public Function ResultDataEntry(ByRef objRst As Object, ByVal VfyDt As String, _
                                ByVal VfyTm As String, ByVal VfyID As String, _
                                ByVal FootNoteId As String) As Boolean
    
    Dim SqlStmt() As String
    Dim aryLab102() As String
    Dim strtmp As String
    Dim strSQL As String
    Dim ii As Long
    Dim intSql As Integer
    Dim intOrdSql As Integer
    Dim VerifyFg As Boolean
    Dim tmpVfyDt As String
    Dim tmpVfyTm As String
    Dim tmpVfyId As String
    
    '-- 전주예수병원 추가 변수=============
    Dim SqlTransfer()   As String
    Dim intTempSql      As Integer
    Dim clsTransfer     As New clsLISTransfer
    Dim strVfyDt        As String
    Dim strRstVal       As String
    Dim strKey          As String
    Dim strWrkDiv       As String
    Dim strOCSNote      As String
    Dim strFromRef      As String
    Dim strToRef        As String
    '======================================
    
    '/*  --    결과등록    -- */
    ' LAB102 : StsCd : 전체 Verify시 수정됨. :UPDATE
    ' LAB201 : Status,ReqinputCnt,AltInputCnt,FootNoteFg,VfyDt,VfyTm,VfyId,
    '              *RmkCd : UPDATE
    ' LAB205 :
    ' LAB302 : *RstVal,*RstCd,*HLDiv,*DPDiv,VfyDt,VfyTm,VfyId,TxtFg,*RstType :UPDATE
    ' LAB303 : *RstTxt : INSERT
    ' LAB304 : *RstText : INSERT
    '
    '접수내역
    On Error GoTo DBExecError
        
    DBErrClear
    VerifyFg = False
    intSql = 0: ReDim SqlStmt(0)
    
    '-- 전주예수병원 추가 변수=============
    intTempSql = 0: ReDim SqlTransfer(0)
    '======================================
    
    '결과내역
    For ii = 1 To objRst.TestCount
        With objRst.Result.Item(ii)
            '결과내역 수정
            
            If .RstDiv <> "*" Then
                intSql = intSql + 1: ReDim Preserve SqlStmt(intSql)
                SqlStmt(intSql) = DBW("rstval", .RstVal, 3) & _
                                  DBW("rstcd", .RstCd, 3) & _
                                  DBW("hldiv", .HLDiv, 3) & _
                                  DBW("dpdiv", .DPDiv, 3) & _
                                  DBW("eqpcd", .EqpCd, 3) & _
                                  DBW("txtfg", .TxtFg, 3)
                If .ExcFg = "0" Then
                    SqlStmt(intSql) = SqlStmt(intSql) & DBW("rsttype", .RstType, 3) & _
                                                        DBW("vfydt", .VfyDt, 3) & _
                                                        DBW("vfytm", .VfyTm, 3) & _
                                                        DBW("vfyid", .VfyID, 2)
                Else
                    SqlStmt(intSql) = SqlStmt(intSql) & DBW("rsttype", .RstType, 2)
                End If
                SqlStmt(intSql) = " UPDATE " & T_LAB302 & " SET " & SqlStmt(intSql) & _
                                  " WHERE  " & DBW("workarea", .WorkArea, 2) & _
                                  " AND    " & DBW("accdt", .AccDt, 2) & _
                                  " AND    " & DBW("accseq", .accseq, 2) & _
                                  " AND    " & DBW("testcd", .TestCd, 2) & _
                                  " AND   (vfydt = ''  or  vfydt is null) "  '저장시점에 결과확인 된 아이템은 제외한다(시점차로 인한 문제가 많음)... 99.12.10 by KMK
            Else
                If .VfyDt <> "" Then
                    intSql = intSql + 1: ReDim Preserve SqlStmt(intSql)
                    If .TxtFg <> "1" Then
                        SqlStmt(intSql) = DBW("vfydt", .VfyDt, 3) & _
                                          DBW("vfytm", .VfyTm, 3) & _
                                          DBW("vfyid", .VfyID, 2)
                    Else
                        SqlStmt(intSql) = DBW("vfydt", .VfyDt, 3) & _
                                          DBW("vfytm", .VfyTm, 3) & _
                                          DBW("vfyid", .VfyID, 3) & _
                                          DBW("txtfg", .TxtFg, 2)
                    End If
                    SqlStmt(intSql) = " UPDATE " & T_LAB302 & " SET " & SqlStmt(intSql) & _
                                      " WHERE  " & DBW("workarea", .WorkArea, 2) & _
                                      " AND    " & DBW("accdt", .AccDt, 2) & _
                                      " AND    " & DBW("accseq", .accseq, 2) & _
                                      " AND    " & DBW("testcd", .TestCd, 2) & _
                                      " AND    (vfydt = ''  or vfydt is null) "   '저장시점에 결과확인 된 아이템은 제외한다(시점차로 인한 문제가 많음)... 99.12.10 by KMK
                End If
            End If
            '
            
           '20030730: kjg
           '대표코드에도 텍스트결과를 입력할수 있게 한다.
'            If .RstDiv <> "*" AND .TxtFg = "1" Then
            If .TxtFg = "1" Then
                '텍스트결과 등록
                intSql = intSql + 1: ReDim Preserve SqlStmt(intSql)
                If UpdatableLAB303(.WorkArea, .AccDt, .accseq, .TestCd) = False Then
                    SqlStmt(intSql) = " INSERT INTO " & T_LAB303 & " (workarea,accdt,accseq,testcd, rsttxt)" & _
                                      " VALUES(" & _
                                      DBV("workarea", .WorkArea, 1) & _
                                      DBV("accdt", .AccDt, 1) & _
                                      DBV("accseq", .accseq, 1) & _
                                      DBV("testcd", .TestCd, 1) & _
                                      DBV("rsttxt", .textrst) & ")"
                Else
                    SqlStmt(intSql) = " UPDATE " & T_LAB303 & _
                                      " SET    " & _
                                                   DBW("rsttxt", .textrst, 2) & _
                                      " WHERE  " & DBW("workarea", .WorkArea, 2) & _
                                      " AND    " & DBW("accdt", .AccDt, 2) & _
                                      " AND    " & DBW("accseq", .accseq, 2) & _
                                      " AND    " & DBW("testcd", .TestCd, 2)
                End If
            End If
           '
            If .VfyDt <> "" Then
                '외부의뢰내역 Status 반영
                If UpdatableLAB205(.WorkArea, .AccDt, .accseq, .TestCd) Then
                    intSql = intSql + 1: ReDim Preserve SqlStmt(intSql)
                    SqlStmt(intSql) = " UPDATE " & T_LAB205 & _
                                      " SET    " & _
                                                   DBW("stscd", enStsCd.StsCd_LIS_MidRst, 2) & _
                                      " WHERE  " & DBW("workarea", .WorkArea, 2) & _
                                      " AND    " & DBW("accdt", .AccDt, 2) & _
                                      " AND    " & DBW("accseq", .accseq, 2) & _
                                      " AND    " & DBW("testcd", .TestCd, 2)
                End If
                If Not VerifyFg Then
                    VerifyFg = True
                    tmpVfyDt = .VfyDt: tmpVfyTm = .VfyTm: tmpVfyId = .VfyID
                End If
            End If
            
            
            '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.06
            If .VfyDt <> "" Then
                
                If .RstDiv = "*" Then
                    GoTo OCSSKIP
                End If
                
                If objRst.FootNote <> "" Then
                    strOCSNote = "[FootNote]" & Chr(13) & objRst.FootNote & Chr(13)
                Else
                    strOCSNote = ""
                End If
                
                If objRst.rmknm <> "" Then
                    strOCSNote = strOCSNote & "◈ Remark" & Chr(13) & objRst.rmknm
                End If
                
                strVfyDt = .VfyDt & .VfyTm
                
                '-- 검사항목별 결과코드(값) 존재 여부
                Call clsTransfer.ResultValue_Find(.TestCd, .RstCd)
                
                If strKey <> objRst.PtId & .OrdDt & .OrdNo Then
                    '-- 결과전송 시 필요정보
                    Call clsTransfer.SqlOCS_Info_New(.WorkArea, .AccDt, .accseq)
                
                    intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
                    
                    '-- 1. '3' 종합건진, '4' 일반검사, 이외검사로 구분을 한다.
                    strWrkDiv = clsTransfer.SqlOCBussdiv_Flag(objRst.PtId, .OrdDt, .OrdNo)
                    If Mid(objRst.PtId, 1, 1) <> "L" Then
                        If strWrkDiv = "3" Then
                            If clsTransfer.ItemVal <> "" Then
                                If clsTransfer.SqlOCSSU2EXAMT_Find(objRst.PtId, .OrdDt, .OrdNo) = True Then
                                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSSU2EXAMT_UPDATE_New2(objRst.PtId, .OrdDt, .OrdNo, _
                                                            "N", clsTransfer.ItemVal, strOCSNote, clsTransfer.RcvDt & clsTransfer.RcvTm, _
                                                            strVfyDt, .VfyID, strVfyDt, .VfyID, "", "", "", .textrst)
                                Else
                                    GoTo DBExecError
                                End If
                            Else
                                If clsTransfer.SqlOCSSU2EXAMT_Find(objRst.PtId, .OrdDt, .OrdNo) = True Then
                                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSSU2EXAMT_UPDATE_New2(objRst.PtId, .OrdDt, .OrdNo, _
                                                            "N", .RstCd, strOCSNote, clsTransfer.RcvDt & clsTransfer.RcvTm, _
                                                            strVfyDt, .VfyID, strVfyDt, .VfyID, "", "", "", .textrst)
                                Else
                                    GoTo DBExecError
                                End If
                            End If
                        ElseIf strWrkDiv = "4" Then
                                If clsTransfer.ItemVal <> "" Then
                                    If clsTransfer.SqlOCSSG2EXAMT_Find(objRst.PtId, .OrdDt, .OrdNo) = True Then
                                        SqlTransfer(intTempSql) = clsTransfer.SqlOCSSG2EXAMT_UPDATE_New2(objRst.PtId, .OrdDt, .OrdNo, _
                                                                "N", clsTransfer.ItemVal, strOCSNote, clsTransfer.RcvDt & clsTransfer.RcvTm, _
                                                                strVfyDt, .VfyID, strVfyDt, .VfyID, "", "", "", .textrst)
                                    Else
                                        GoTo DBExecError
                                    End If
                                Else
                                    If clsTransfer.SqlOCSSG2EXAMT_Find(objRst.PtId, .OrdDt, .OrdNo) = True Then
                                        SqlTransfer(intTempSql) = clsTransfer.SqlOCSSG2EXAMT_UPDATE_New2(objRst.PtId, .OrdDt, .OrdNo, _
                                                                "N", .RstCd, strOCSNote, clsTransfer.RcvDt & clsTransfer.RcvTm, _
                                                                strVfyDt, .VfyID, strVfyDt, .VfyID, "", "", "", .textrst)
                                    Else
                                        GoTo DBExecError
                                    End If
                                End If
                        Else
                            If clsTransfer.SqlOCSMDEXMORT_Find(objRst.PtId, .OrdDt, .OrdNo) = True Then
                                SqlTransfer(intTempSql) = clsTransfer.SqlOCSMDEXMORT_UPDATE(objRst.PtId, .OrdDt, .OrdNo, _
                                                        "N", strVfyDt, .VfyID, "", "")
                            Else
                                GoTo DBExecError
                            End If
                        End If
                    End If
                    strKey = objRst.PtId & .OrdDt & .OrdNo
                    
                End If
                
                
                '** 종건/일건 이외의 검사만 등록.
                If strWrkDiv <> "" Then
                    GoTo OCSSKIP
                End If
                
                '-- 참고치값 Check
                If .RefValFrom = "0" And .RefValTo = "0" Then
                    strFromRef = ""
                    strToRef = ""
                    If .RefCd <> "" Then
                        strFromRef = .RefCd
                    End If
                Else
                    strFromRef = .RefValFrom
                    strToRef = .RefValTo
                End If
                
'''                '-- 2. 건진외결과 INSERT
'''                '** 검사항목별 결과값 체크
'''                intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
'''                If Mid(objRst.PtId, 1, 1) <> "L" Then
'''                    If clsTransfer.ItemVal <> "" Then
'''                        If clsTransfer.SqlOCSRESULT_FLAG(objRst.PtId, .OrdDt, .OrdNo, .TestCd) = False Then
'''                            SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_INSERT(objRst.PtId, .OrdDt, .OrdNo, _
'''                                                    .TestCd, clsTransfer.SlipCd, clsTransfer.ItemVal, .textrst & vbNewLine & .SuppText, strFromRef, strToRef, .RstUnit, _
'''                                                    strOCSNote, clsTransfer.RcvDt & clsTransfer.RcvTm, strVfyDt, .VfyID, _
'''                                                    strVfyDt, .VfyID, "", clsTransfer.RsltType, "", "", "")
'''                        Else
'''                            SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_SUB_UPDATE(objRst.PtId, .OrdDt, .OrdNo, _
'''                                                    .TestCd, clsTransfer.ItemVal, .textrst & vbNewLine & .SuppText, strOCSNote, strVfyDt, .VfyID)
'''                        End If
'''                    Else
'''                        If clsTransfer.SqlOCSRESULT_FLAG(objRst.PtId, .OrdDt, .OrdNo, .TestCd) = False Then
'''                            SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_INSERT(objRst.PtId, .OrdDt, .OrdNo, _
'''                                                    .TestCd, clsTransfer.SlipCd, .RstCd, .textrst & vbNewLine & .SuppText, strFromRef, strToRef, .RstUnit, _
'''                                                    strOCSNote, clsTransfer.RcvDt & clsTransfer.RcvTm, strVfyDt, .VfyID, _
'''                                                    strVfyDt, .VfyID, "", clsTransfer.RsltType, "", "", "")
'''                        Else
'''                            SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_SUB_UPDATE(objRst.PtId, .OrdDt, .OrdNo, _
'''                                                    .TestCd, .RstCd, .textrst & vbNewLine & .SuppText, strOCSNote, strVfyDt, .VfyID)
'''                        End If
'''                    End If
'''                End If

                '-- 2. 건진외결과 INSERT
                '** 검사항목별 결과값 체크
                Select Case .TestCd
                    Case "ABOC22B", "ABOC22C", "ABOB22B", "ABOB22C", "RHD22B", "ABSC22B", "ABSC22C", _
                         "LABOC22B", "LABOC22C", "LABOB22B", "LABOB22C", "LRHD22B", "LABSC22B", "LABSC22C"
                    Case Else
                        intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
                        If Mid(objRst.PtId, 1, 1) <> "L" Then
                            If clsTransfer.ItemVal <> "" Then
                                If clsTransfer.SqlOCSRESULT_FLAG(objRst.PtId, .OrdDt, .OrdNo, .TestCd) = False Then
                                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_INSERT(objRst.PtId, .OrdDt, .OrdNo, _
                                                            .TestCd, clsTransfer.SlipCd, clsTransfer.ItemVal, .textrst & vbNewLine & .SuppText, strFromRef, strToRef, .RstUnit, _
                                                            strOCSNote, clsTransfer.RcvDt & clsTransfer.RcvTm, strVfyDt, .VfyID, _
                                                            strVfyDt, .VfyID, "", clsTransfer.RsltType, "", "", "")
                                Else
                                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_SUB_UPDATE(objRst.PtId, .OrdDt, .OrdNo, _
                                                            .TestCd, clsTransfer.ItemVal, .textrst & vbNewLine & .SuppText, strOCSNote, strVfyDt, .VfyID)
                                End If
                            Else
                                If clsTransfer.SqlOCSRESULT_FLAG(objRst.PtId, .OrdDt, .OrdNo, .TestCd) = False Then
                                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_INSERT(objRst.PtId, .OrdDt, .OrdNo, _
                                                            .TestCd, clsTransfer.SlipCd, .RstCd, .textrst & vbNewLine & .SuppText, strFromRef, strToRef, .RstUnit, _
                                                            strOCSNote, clsTransfer.RcvDt & clsTransfer.RcvTm, strVfyDt, .VfyID, _
                                                            strVfyDt, .VfyID, "", clsTransfer.RsltType, "", "", "")
                                Else
                                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_SUB_UPDATE(objRst.PtId, .OrdDt, .OrdNo, _
                                                            .TestCd, .RstCd, .textrst & vbNewLine & .SuppText, strOCSNote, strVfyDt, .VfyID)
                                End If
                            End If
                        End If
                End Select
            End If
            '/////////////////////////////////////////////////////////////////////////////////////////
            
OCSSKIP:

        End With
    Next ii
    '
    '결과보고 대기내역(LAB202) Insert
    If VerifyFg Then
        Dim tmpSql As String
        Dim tmpRs As Recordset
        Dim tmpDept As String
        Dim tmpBussDiv As String
        
        Dim KK As Integer
        
        For KK = 1 To objRst.TestCount
            With objRst.Result.Item(KK)
                tmpSql = " SELECT bussdiv FROM " & T_LAB101 & _
                       " WHERE " & DBW("ptid=", objRst.PtId) & _
                       " AND " & DBW("orddt=", .OrdDt) & _
                       " AND " & DBW("ordno=", .OrdNo)
                Set tmpRs = Nothing
                Set tmpRs = New Recordset
                tmpRs.Open tmpSql, DBConn
                If Not tmpRs.EOF Then
                    tmpBussDiv = Trim(tmpRs.Fields("bussdiv").Value & "")
                    If objRst.WardId = "" Then
                        tmpDept = objRst.DeptCd
                    Else
                        tmpDept = objRst.WardId
                    End If
'                    tmpRs.RsClose
                    Set tmpRs = Nothing
                    Exit For
                End If
            End With
        Next
        
        tmpSql = " SELECT * FROM " & T_LAB202 & _
                 " WHERE  " & DBW("deptcd", tmpDept, 2) & _
                 " AND    " & DBW("vfydt", tmpVfyDt, 2) & _
                 " AND    " & DBW("ptid", objRst.PtId, 2) & _
                 " AND    " & DBW("mfyfg", "0", 2)
        Set tmpRs = Nothing
        Set tmpRs = New Recordset
        tmpRs.Open tmpSql, DBConn
        
        If tmpRs.EOF Then
            intSql = intSql + 1: ReDim Preserve SqlStmt(intSql)
            SqlStmt(intSql) = "INSERT INTO " & T_LAB202 & " (deptcd, vfydt, ptid, mfyfg, vfytm, vfyid, donefg, doneid, majdoct, bussdiv) " & _
                              " VALUES ( " & _
                              DBV("deptcd", tmpDept, 1) & _
                              DBV("vfydt", tmpVfyDt, 1) & _
                              DBV("ptid", objRst.PtId, 1) & _
                              DBV("mfyfg", "0", 1) & _
                              DBV("vfytm", tmpVfyTm, 1) & _
                              DBV("vfyid", tmpVfyId, 1) & _
                              DBV("donefg", "", 1) & _
                              DBV("doneid", "0", 1) & _
                              DBV("majdoct", objRst.OrdDoct, 1) & _
                              DBV("bussdiv", tmpBussDiv) & " ) "
        Else
            If "" & tmpRs.Fields("DoneFg").Value = "1" Then
                intSql = intSql + 1: ReDim Preserve SqlStmt(intSql)
                SqlStmt(intSql) = " UPDATE " & T_LAB202 & _
                                  " set    " & _
                                               DBW("vfytm", tmpVfyTm, 3) & _
                                               DBW("vfyid", tmpVfyId, 3) & _
                                               DBW("majdoct", objRst.OrdDoct, 3) & _
                                               DBW("donefg", "", 3) & _
                                               DBW("doneid", "0", 2) & _
                                  " WHERE  " & DBW("deptcd", tmpDept, 2) & _
                                  " AND    " & DBW("vfydt", tmpVfyDt, 2) & _
                                  " AND    " & DBW("ptid", objRst.PtId, 2) & _
                                  " AND    " & DBW("mfyfg", "0", 2)
            End If
        End If

'        tmpRs.RsClose
        Set tmpRs = Nothing
    End If
   
    DBConn.BeginTrans
    For intSql = LBound(SqlStmt) To UBound(SqlStmt)
        intSql = intSql
        
        If Trim(SqlStmt(intSql)) <> "" Then DBConn.Execute (SqlStmt(intSql))
    Next intSql
    
    
    '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.06
'    온승호
    For intTempSql = LBound(SqlTransfer) To UBound(SqlTransfer)
        intTempSql = intTempSql

        Debug.Print SqlTransfer(intTempSql)
        If Trim(SqlTransfer(intTempSql)) <> "" Then DBConn.Execute (SqlTransfer(intTempSql))
    Next intTempSql
    '============================================================================================
    
    
    '
    '접수내역(LAB201) Update
    '** 결과내역을 띄워놓은 상태에서 다른 PC에서 입력될수 있으므로 최종 Input Item Count를 다시 읽어온다.
    '** 99.12.10 by KMK
    With objRst
      
        Dim tmpInputCnt As Integer
        
        tmpInputCnt = GetInputCnt(.WorkArea, .AccDt, .accseq)
        
        '접수내역 수정
        'SqlStmt(intSql) = "reqinputcnt" & DBNumConv(.ReqInputCnt, 3) &
        strSQL = DBW("reqinputcnt", tmpInputCnt, 3) & _
                 DBW("rmkcd", .RmkCd, 3)
        'If .StsCd = enStsCd.StsCd_LIS_FinRst Then
        If .ReqTotCnt = tmpInputCnt Then
            '전체 Verify시 Update
            strSQL = strSQL & DBW("stscd", .StsCd, 3) & _
                              DBW("vfydt", VfyDt, 3) & _
                              DBW("vfytm", VfyTm, 3) & _
                              DBW("vfyid", VfyID, 3)
        End If
        strSQL = strSQL & DBW("footnotefg", .FootNoteFg, 2)
        strSQL = " UPDATE " & T_LAB201 & _
                 " SET    " & strSQL & _
                 " WHERE  " & DBW("workarea", .WorkArea, 2) & _
                 " AND    " & DBW("accdt", .AccDt, 2) & _
                 " AND    " & DBW("accseq", .accseq, 2)
        DBConn.Execute (strSQL)
           
        If .FootNoteFg <> "" Then
            'FOOTNOTE 등록
            strSQL = ""
            If UpdatableLAB304(.WorkArea, .AccDt, .accseq, "1") = False Then
                strSQL = " INSERT INTO " & T_LAB304 & " (workarea,accdt,accseq,seq,vfyid,rsttxt) " & _
                         " VALUES(" & _
                         DBV("workarea", .WorkArea, 1) & _
                         DBV("accdt", .AccDt, 1) & _
                         DBV("accseq", .accseq, 1) & _
                         DBV("seq", "1", 1) & _
                         DBV("vfyid", FootNoteId, 1) & _
                         DBV("rsttxt", .FootNote) & ")"
            Else
                strSQL = " UPDATE " & T_LAB304 & _
                         " SET    " & _
                                    DBW("rsttxt", .FootNote, 3) & _
                                    DBW("vfyid", FootNoteId, 2) & _
                         " WHERE  " & DBW("workarea", .WorkArea, 2) & _
                         " AND    " & DBW("accdt", .AccDt, 2) & _
                         " AND    " & DBW("accseq", .accseq, 2) & _
                         " AND    " & DBW("seq", "1", 2)
            End If
            DBConn.Execute (strSQL)
        End If
    End With
   
    intOrdSql = intSql
    Dim tmpStsCd As String
    Dim tmpPtKey As String
    For ii = 1 To objRst.TestCount
        With objRst.Result.Item(ii)
            '처방BODY 수정
            
            If tmpPtKey = objRst.PtId & .OrdDt & .OrdNo & .OrdSeq & .TestCd Then GoTo Skip
            If (.ExcFg = "0") And (.VfyDt <> "") Then
                If .DetailFg <> "" Then
                    '상세항목인경우 RstDiv가 "*"인 행 기준으로 처방Body Update
                    If .RstDiv = "*" Then
                        If IsOrderFollowUp(objRst.PtId, .OrdDt, .OrdNo, .OrdSeq) = True Then
                            tmpStsCd = enStsCd.StsCd_LIS_FinRst   '전체 Verify
                        Else
                            tmpStsCd = enStsCd.StsCd_LIS_MidRst  '부분 Verify
                        End If
                        intSql = UBound(SqlStmt) + 1
                        ReDim Preserve SqlStmt(intSql)
                        
                        '** 원본 ==========================================================================
'                        SqlStmt(intSql) = " UPDATE " & T_LAB102 & _
'                                          " SET    " & _
'                                                       DBW("stscd", tmpStsCd, 3) & _
'                                                       DBW("examdt", .VfyDt, 3) & _
'                                                       DBW("examtm", .VfyTm, 3) & _
'                                                       DBW("examdoct", .VfyID, 2) & _
'                                          " WHERE  " & DBW("ptid", objRst.PtId, 2) & _
'                                          " AND    " & DBW("orddt", .OrdDt, 2) & _
'                                          " AND    " & DBW("ordno", .OrdNo, 2) & _
'                                          " AND    " & DBW("ordseq", .OrdSeq, 2)
                        '===================================================================================
                        
                        '** 전주 예수 병원 변경 루틴========================================================
                        '-- 1. '3' 종합건진, '4' 일반검사, 이외검사로 구분을 한다.
                        strWrkDiv = clsTransfer.SqlOCBussdiv_Flag(objRst.PtId, .OrdDt, .OrdNo)
                        
                        If strWrkDiv = "3" Then         '종건
                            SqlStmt(intSql) = " UPDATE su2examt " & _
                                              " SET    " & _
                                                           DBW("stscd", tmpStsCd, 2) & _
                                              " WHERE  patno = " & DBS(objRst.PtId) & _
                                              " AND    orddate = TO_DATE(" & DBS(.OrdDt) & ", 'yyyymmdd')" & _
                                              " AND    ordseqno = " & DBN(.OrdNo)
                        ElseIf strWrkDiv = "4" Then     '일건
                                SqlStmt(intSql) = " UPDATE sg2examt " & _
                                                  " SET    " & _
                                                               DBW("stscd", tmpStsCd, 2) & _
                                                  " WHERE  patno = " & DBS(objRst.PtId) & _
                                                  " AND    orddate = TO_DATE(" & DBS(.OrdDt) & ", 'yyyymmdd')" & _
                                                  " AND    ordseqno = " & DBN(.OrdNo)
                        Else
                            SqlStmt(intSql) = " UPDATE mdexmort " & _
                                              " SET    " & _
                                                           DBW("stscd", tmpStsCd, 2) & _
                                              " WHERE  patno = " & DBS(objRst.PtId) & _
                                              " AND    orddate = TO_DATE(" & DBS(.OrdDt) & ", 'yyyymmdd')" & _
                                              " AND    ordseqno = " & DBN(.OrdNo)
                        End If
                        '===================================================================================
                        ' 외부정도관리 STSCD 업데이트
                        If Mid(objRst.PtId, 1, 1) = "L" Then
                            SqlStmt(intSql) = " UPDATE s2ord999 " & _
                                              " SET    " & _
                                                           DBW("stscd", tmpStsCd, 2) & _
                                              " WHERE  patno = " & DBS(objRst.PtId) & _
                                              " AND    orddate = TO_DATE(" & DBS(.OrdDt) & ", 'yyyymmdd')" & _
                                              " AND    ordseqno = " & DBN(.OrdNo)
                        
                        End If
                    
                        DBConn.Execute (SqlStmt(intSql))
                    End If
                Else
                    '그룹코드 혹은 개별 ITEM인경우
                    If IsOrderFollowUp(objRst.PtId, .OrdDt, .OrdNo, .OrdSeq) = True Then
                        tmpStsCd = enStsCd.StsCd_LIS_FinRst   '전체 Verify
                    Else
                        tmpStsCd = enStsCd.StsCd_LIS_MidRst  '부분 Verify
                    End If
                    intSql = UBound(SqlStmt) + 1
                    ReDim Preserve SqlStmt(intSql)
                    
                    '** 원본 ==========================================================================
'                    SqlStmt(intSql) = " UPDATE " & T_LAB102 & _
'                                      " SET    " & _
'                                                   DBW("stscd", tmpStsCd, 3) & _
'                                                   DBW("examdt", .VfyDt, 3) & _
'                                                   DBW("examtm", .VfyTm, 3) & _
'                                                   DBW("examdoct", .VfyID, 2) & _
'                                      " WHERE  " & DBW("ptid", objRst.PtId, 2) & _
'                                      " AND    " & DBW("orddt", .OrdDt, 2) & _
'                                      " AND    " & DBW("ordno", .OrdNo, 2) & _
'                                      " AND    " & DBW("ordseq", .OrdSeq, 2)
                    '===================================================================================
                    
                    '** 전주 예수 병원 변경 루틴========================================================
                    '-- 1. '3' 종합건진, '4' 일반검사, 이외검사로 구분을 한다.
                    strWrkDiv = clsTransfer.SqlOCBussdiv_Flag(objRst.PtId, .OrdDt, .OrdNo)
                    
                    If strWrkDiv = "3" Then         '종건
                        SqlStmt(intSql) = " UPDATE su2examt " & _
                                          " SET    " & _
                                                       DBW("stscd", tmpStsCd, 2) & _
                                          " WHERE  patno = " & DBS(objRst.PtId) & _
                                          " AND    orddate = TO_DATE(" & DBS(.OrdDt) & ", 'yyyymmdd')" & _
                                          " AND    ordseqno = " & DBN(.OrdNo)
                    ElseIf strWrkDiv = "4" Then     '일건
                            SqlStmt(intSql) = " UPDATE sg2examt " & _
                                              " SET    " & _
                                                           DBW("stscd", tmpStsCd, 2) & _
                                              " WHERE  patno = " & DBS(objRst.PtId) & _
                                              " AND    orddate = TO_DATE(" & DBS(.OrdDt) & ", 'yyyymmdd')" & _
                                              " AND    ordseqno = " & DBN(.OrdNo)
                    Else
                        SqlStmt(intSql) = " UPDATE mdexmort " & _
                                          " SET    " & _
                                                       DBW("stscd", tmpStsCd, 2) & _
                                          " WHERE  patno = " & DBS(objRst.PtId) & _
                                          " AND    orddate = TO_DATE(" & DBS(.OrdDt) & ", 'yyyymmdd')" & _
                                          " AND    ordseqno = " & DBN(.OrdNo)
                    End If
                    '===================================================================================
                    ' 외부정도관리 STSCD 업데이트
                    If Mid(objRst.PtId, 1, 1) = "L" Then
                        SqlStmt(intSql) = " UPDATE s2ord999 " & _
                                          " SET    " & _
                                                       DBW("stscd", tmpStsCd, 2) & _
                                          " WHERE  patno = " & DBS(objRst.PtId) & _
                                          " AND    orddate = TO_DATE(" & DBS(.OrdDt) & ", 'yyyymmdd')" & _
                                          " AND    ordseqno = " & DBN(.OrdNo)
                    
                    End If
                    
                    DBConn.Execute (SqlStmt(intSql))
                End If
            End If
            tmpPtKey = objRst.PtId & .OrdDt & .OrdNo & .OrdSeq & .TestCd
            
Skip:
        End With
    Next ii
    
    DBConn.CommitTrans
    
    '감염관리결과등록
    Call ICSGeneralResultCheck(objRst)
    
    ResultDataEntry = True
    Exit Function

DBExecError:
'    Call Error_Routine
    MsgBox "처리 도중 오류가 발생하였습니다.", vbExclamation
    ResultDataEntry = False
    'Resume Next

End Function
 
'Public Function ItemDataEntry(ByRef objRst As clsPatientInfo, ByVal VfyDt As String, _
                              ByVal VfyTm As String, ByVal VfyId As String, _
                              ByVal FootNoteId As String) As Boolean

Public Function ItemDataEntry(ByRef objRst As Object, ByVal VfyDt As String, _
                              ByVal VfyTm As String, ByVal VfyID As String, _
                              ByVal FootNoteId As String) As Boolean
    
    Dim SqlStmt() As String
    Dim aryLab102() As String
    Dim strtmp As String
    Dim strSQL As String
    Dim ii As Long
    Dim intSql As Integer
    Dim intOrdSql As Integer
    Dim strFootNoteFg As String
    Dim blnFirst As Boolean
    Dim tmpInputCnt As Integer
    Dim tmpStsCd As String
    Dim sWA As String, sAD As String, sAS As String
    
    '-- 전주예수병원 추가 변수=============
    Dim SqlTransfer()   As String
    Dim intTempSql      As Integer
    Dim clsTransfer     As New clsLISTransfer
    Dim strVfyDt        As String
    Dim strRstVal       As String
    Dim strKey          As String
    Dim strWrkDiv       As String
    Dim strOCSNote      As String
    Dim strFromRef      As String
    Dim strToRef        As String
    '======================================
    
    '/*  --    결과등록    -- */
    ' LAB102 : StsCd : 전체 Verify시 수정됨. :UPDATE
    ' LAB201 : Status,ReqinputCnt,AltInputCnt,FootNoteFg,VfyDt,VfyTm,VfyId,
    '              *RmkCd : UPDATE
    ' LAB302 : *RstVal,*RstCd,*HLDiv,*DPDiv,VfyDt,VfyTm,VfyId,TxtFg,*RstType :UPDATE
    ' LAB303 : *RstTxt : INSERT
    ' LAB304 : *RstText : INSERT
    '
    '접수내역
    On Error GoTo DBExecError
    DBErrClear
    blnFirst = False
    intSql = 0: ReDim SqlStmt(0)
    
    '-- 전주예수병원 추가 변수=============
    intTempSql = 0: ReDim SqlTransfer(0)
    '======================================
    
    For ii = 1 To objRst.TestCount
       
        With objRst.Result.Item(ii)
              
            '결과내역 수정
            
            'If (.ExcFg = "1") AND (.VfyDt = "") Then GoTo Skip
            
            If (.ExcFg = "1") And (.RstCd = "") Then GoTo Skip
            
            sWA = .WorkArea
            sAD = .AccDt
            sAS = .accseq
            
            If .RstDiv <> "*" Then
                intSql = intSql + 1: ReDim Preserve SqlStmt(intSql)
                SqlStmt(intSql) = DBW("rstval", .RstVal, 3) & _
                                  DBW("rstcd", .RstCd, 3) & _
                                  DBW("hldiv", .HLDiv, 3) & _
                                  DBW("dpdiv", .DPDiv, 3) & _
                                  DBW("eqpcd", .EqpCd, 3) & _
                                  DBW("txtfg", .TxtFg, 3)
                If .ExcFg = "0" Then
                    SqlStmt(intSql) = SqlStmt(intSql) & DBW("rsttype", .RstType, 3) & _
                                                       DBW("vfydt", .VfyDt, 3) & _
                                                       DBW("vfytm", .VfyTm, 3) & _
                                                       DBW("vfyid", .VfyID, 2)
                Else
                    SqlStmt(intSql) = SqlStmt(intSql) & DBW("rsttype", .RstType, 2)
                End If
                SqlStmt(intSql) = " UPDATE " & T_LAB302 & _
                                  " SET    " & SqlStmt(intSql) & _
                                  " WHERE  " & DBW("workarea", .WorkArea, 2) & _
                                  " AND    " & DBW("accdt", .AccDt, 2) & _
                                  " AND    " & DBW("accseq", .accseq, 2) & _
                                  " AND    " & DBW("testcd", .TestCd, 2)
            End If
            '
            '텍스트결과 등록
            If .RstDiv <> "*" And .TxtFg = "1" Then
                intSql = intSql + 1: ReDim Preserve SqlStmt(intSql)
                If UpdatableLAB303(.WorkArea, .AccDt, .accseq, .TestCd) = False Then
                    SqlStmt(intSql) = "INSERT INTO " & T_LAB303 & " (workarea,accdt,accseq,testcd, rsttxt)" & _
                                      " VALUES(" & _
                                      DBV("workarea", .WorkArea, 1) & _
                                      DBV("accdt", .AccDt, 1) & _
                                      DBV("accseq", .accseq, 1) & _
                                      DBV("testcd", .TestCd, 1) & _
                                      DBV("rsttxt", .textrst) & ")"
                Else
                    SqlStmt(intSql) = " UPDATE " & T_LAB303 & _
                                      " SET    " & _
                                                   DBW("rsttxt", .textrst, 2) & _
                                      " WHERE  " & DBW("workarea", .WorkArea, 2) & _
                                      " AND    " & DBW("accdt", .AccDt, 2) & _
                                      " AND    " & DBW("accseq", .accseq, 2) & _
                                      " AND    " & DBW("testcd", .TestCd, 2)
                End If
            End If
            '
        
            '외부의뢰내역 Status 반영
            If .VfyDt <> "" Then
                If UpdatableLAB205(.WorkArea, .AccDt, .accseq, .TestCd) Then
                    intSql = intSql + 1: ReDim Preserve SqlStmt(intSql)
                        SqlStmt(intSql) = " UPDATE " & T_LAB205 & _
                                          " SET    " & _
                                                       DBW("stscd", enStsCd.StsCd_LIS_MidRst, 2) & _
                                          " WHERE  " & DBW("workarea", .WorkArea, 2) & _
                                          " AND    " & DBW("accdt", .AccDt, 2) & _
                                          " AND    " & DBW("accseq", .accseq, 2) & _
                                          " AND    " & DBW("testcd", .TestCd, 2) '& "'"
                End If
            End If
            
            '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.06
            If .VfyDt <> "" Then
                
                If .RstDiv = "*" Then
                    GoTo OCSSKIP
                End If
                
                If .FootNote <> "" Then
                    strOCSNote = "[FootNote]" & Chr(13) & .FootNote & Chr(13)
                Else
                    strOCSNote = ""
                End If
                
                If objRst.rmknm <> "" Then
                    strOCSNote = strOCSNote & "◈ Remark" & Chr(13) & objRst.rmknm
                End If
                
                strVfyDt = .VfyDt & .VfyTm
                
                '-- 검사항목별 결과코드(값) 존재 여부
                Call clsTransfer.ResultValue_Find(.TestCd, .RstCd)
'온승호
                If strKey <> medGetP(.PtInfo, 2, vbTab) & .OrdDt & .OrdNo Then
                    '-- 결과전송 시 필요정보
                    Call clsTransfer.SqlOCS_Info_New(.WorkArea, .AccDt, .accseq)

                    intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)

                    '-- 1. '3' 종합건진, '4' 일반검사, 이외검사로 구분을 한다.
                    strWrkDiv = clsTransfer.SqlOCBussdiv_Flag(medGetP(.PtInfo, 2, vbTab), .OrdDt, .OrdNo)
                    If strWrkDiv = "3" Then
                        If clsTransfer.ItemVal <> "" Then
                            If clsTransfer.SqlOCSSU2EXAMT_Find(medGetP(.PtInfo, 2, vbTab), .OrdDt, .OrdNo) = True Then
                                SqlTransfer(intTempSql) = clsTransfer.SqlOCSSU2EXAMT_UPDATE_New2(medGetP(.PtInfo, 2, vbTab), .OrdDt, .OrdNo, _
                                                        "N", clsTransfer.ItemVal, strOCSNote, clsTransfer.RcvDt & clsTransfer.RcvTm, _
                                                        strVfyDt, .VfyID, strVfyDt, .VfyID, "", "", "", .textrst)
                            Else
                                GoTo DBExecError
                            End If
                        Else
                            If clsTransfer.SqlOCSSU2EXAMT_Find(medGetP(.PtInfo, 2, vbTab), .OrdDt, .OrdNo) = True Then
                                SqlTransfer(intTempSql) = clsTransfer.SqlOCSSU2EXAMT_UPDATE_New2(medGetP(.PtInfo, 2, vbTab), .OrdDt, .OrdNo, _
                                                        "N", .RstCd, strOCSNote, clsTransfer.RcvDt & clsTransfer.RcvTm, _
                                                        strVfyDt, .VfyID, strVfyDt, .VfyID, "", "", "", .textrst)
                            Else
                                GoTo DBExecError
                            End If
                        End If
                    ElseIf strWrkDiv = "4" Then
                            If clsTransfer.ItemVal <> "" Then
                                If clsTransfer.SqlOCSSG2EXAMT_Find(medGetP(.PtInfo, 2, vbTab), .OrdDt, .OrdNo) = True Then
                                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSSG2EXAMT_UPDATE_New2(medGetP(.PtInfo, 2, vbTab), .OrdDt, .OrdNo, _
                                                            "N", clsTransfer.ItemVal, strOCSNote, clsTransfer.RcvDt & clsTransfer.RcvTm, _
                                                            strVfyDt, .VfyID, strVfyDt, .VfyID, "", "", "", .textrst)
                                Else
                                    GoTo DBExecError
                                End If
                            Else
                                If clsTransfer.SqlOCSSG2EXAMT_Find(medGetP(.PtInfo, 2, vbTab), .OrdDt, .OrdNo) = True Then
                                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSSG2EXAMT_UPDATE_New2(medGetP(.PtInfo, 2, vbTab), .OrdDt, .OrdNo, _
                                                            "N", .RstCd, strOCSNote, clsTransfer.RcvDt & clsTransfer.RcvTm, _
                                                            strVfyDt, .VfyID, strVfyDt, .VfyID, "", "", "", .textrst)
                                Else
                                    GoTo DBExecError
                                End If
                            End If
                    Else
                        If clsTransfer.SqlOCSMDEXMORT_Find(medGetP(.PtInfo, 2, vbTab), .OrdDt, .OrdNo) = True Then
                            SqlTransfer(intTempSql) = clsTransfer.SqlOCSMDEXMORT_UPDATE(medGetP(.PtInfo, 2, vbTab), .OrdDt, .OrdNo, _
                                                    "N", strVfyDt, .VfyID, "", "")
                        Else
                            GoTo DBExecError
                        End If
                    End If

                    strKey = medGetP(.PtInfo, 2, vbTab) & .OrdDt & .OrdNo

                End If
'온승호
                '** 종건/일건 이외의 검사만 등록.
                If strWrkDiv <> "" Then
                    GoTo OCSSKIP
                End If
                
                '-- 참고치값 Check
                If .RefValFrom = "0" And .RefValTo = "0" Then
                    strFromRef = ""
                    strToRef = ""
                    If .RefCd <> "" Then
                        strFromRef = .RefCd
                    End If
                Else
                    strFromRef = .RefValFrom
                    strToRef = .RefValTo
                End If
                
                '-- 2. 건진외결과 INSERT
                '** 검사항목별 결과값 체크
                intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
'온승호
                If clsTransfer.ItemVal <> "" Then
                    If clsTransfer.SqlOCSRESULT_FLAG(medGetP(.PtInfo, 2, vbTab), .OrdDt, .OrdNo, .TestCd) = False Then
                        SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_INSERT(medGetP(.PtInfo, 2, vbTab), .OrdDt, .OrdNo, _
                                                .TestCd, clsTransfer.SlipCd, clsTransfer.ItemVal, .textrst & vbNewLine & .SuppText, strFromRef, strToRef, .RstUnit, _
                                                strOCSNote, clsTransfer.RcvDt & clsTransfer.RcvTm, strVfyDt, .VfyID, _
                                                strVfyDt, .VfyID, "", clsTransfer.RsltType, "", "", "")
                    Else
                        GoTo DBExecError
                    End If
                Else
                    If clsTransfer.SqlOCSRESULT_FLAG(medGetP(.PtInfo, 2, vbTab), .OrdDt, .OrdNo, .TestCd) = False Then
                        SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_INSERT(medGetP(.PtInfo, 2, vbTab), .OrdDt, .OrdNo, _
                                                .TestCd, clsTransfer.SlipCd, .RstCd, .textrst & vbNewLine & .SuppText, strFromRef, strToRef, .RstUnit, _
                                                strOCSNote, clsTransfer.RcvDt & clsTransfer.RcvTm, strVfyDt, .VfyID, _
                                                strVfyDt, .VfyID, "", clsTransfer.RsltType, "", "", "")
                    Else
                        GoTo DBExecError
                    End If
                End If
'온승호
            End If
            '========================================================================================
            
OCSSKIP:

            DBConn.BeginTrans
            
            For intSql = 1 To UBound(SqlStmt)
                intSql = intSql
                DBConn.Execute (SqlStmt(intSql))
            Next intSql
   '
            
            '** 전주예수병원 추가 루틴 =============================================================
            For intTempSql = 1 To UBound(SqlTransfer)
                intTempSql = intTempSql
                'Debug.Print SqlTransfer(intTempSql)
                If Trim(SqlTransfer(intTempSql)) <> "" Then DBConn.Execute (SqlTransfer(intTempSql))
                
            Next intTempSql
            '=======================================================================================
            
            intOrdSql = intSql
         
            '처방BODY 수정
            '그룹코드 혹은 개별 ITEM인경우
            If .RstDiv = "R" Then
                If IsOrderFollowUp(medGetP(.PtInfo, 2, vbTab), .OrdDt, .OrdNo, .OrdSeq) = True Then
                    tmpStsCd = enStsCd.StsCd_LIS_FinRst   '전체 Verify
                Else
                    tmpStsCd = enStsCd.StsCd_LIS_MidRst  '부분 Verify
                End If
                intSql = UBound(SqlStmt) + 1
                ReDim Preserve SqlStmt(intSql)
                
                '** 원본 ==============================================================================
'                SqlStmt(intSql) = " UPDATE " & T_LAB102 & _
'                                  " SET    " & _
'                                               DBW("stscd", tmpStsCd, 3) & _
'                                               DBW("examdt", .VfyDt, 3) & _
'                                               DBW("examtm", .VfyTm, 3) & _
'                                               DBW("examdoct", .VfyID, 2) & _
'                                  " WHERE  " & DBW("ptid", medGetP(.PtInfo, 2, vbTab), 2) & _
'                                  " AND    " & DBW("orddt", .OrdDt, 2) & _
'                                  " AND    " & DBW("ordno", .OrdNo, 2) & _
'                                  " AND    " & DBW("ordseq", .OrdSeq, 2)
                '=======================================================================================
                
                '** 전주예수병원 추가 루틴 =============================================================
                '-- 결과전송 시 필요정보
                'strWrkDiv = clsTransfer.SqlOCBussdiv_Flag(medGetP(.PtInfo, 2, vbTab), .OrdDt, .OrdNo)
                
                If strWrkDiv = "3" Then         '종건
                    SqlStmt(intSql) = " UPDATE su2examt " & _
                                      " SET    " & _
                                                   DBW("stscd", tmpStsCd, 2) & _
                                      " WHERE  patno = " & DBS(medGetP(.PtInfo, 2, vbTab)) & _
                                      " AND    orddate = TO_DATE(" & DBS(.OrdDt) & ", 'yyyymmdd')" & _
                                      " AND    ordseqno = " & DBN(.OrdNo)
                ElseIf strWrkDiv = "4" Then     '일건
                        SqlStmt(intSql) = " UPDATE sg2examt " & _
                                          " SET    " & _
                                                       DBW("stscd", tmpStsCd, 2) & _
                                          " WHERE  patno = " & DBS(medGetP(.PtInfo, 2, vbTab)) & _
                                          " AND    orddate = TO_DATE(" & DBS(.OrdDt) & ", 'yyyymmdd')" & _
                                          " AND    ordseqno = " & DBN(.OrdNo)
                Else
                    SqlStmt(intSql) = " UPDATE mdexmort " & _
                                      " SET    " & _
                                                   DBW("stscd", tmpStsCd, 2) & _
                                      " WHERE  patno = " & DBS(medGetP(.PtInfo, 2, vbTab)) & _
                                      " AND    orddate = TO_DATE(" & DBS(.OrdDt) & ", 'yyyymmdd')" & _
                                      " AND    ordseqno = " & DBN(.OrdNo)
                End If
                '=======================================================================================
                
                DBConn.Execute (SqlStmt(intSql))
            End If
   
   
            '접수내역 수정
        
            intSql = UBound(SqlStmt) + 1
            ReDim Preserve SqlStmt(intSql)
         
            SqlStmt(intSql) = DBW("rmkcd", .OTmpCd, 3)
            If .VfyDt <> "" Then
                '결과가 입력되고 확인되면 ReqInputCnt Update
            
                tmpInputCnt = GetInputCnt(.WorkArea, .AccDt, .accseq)
            
                SqlStmt(intSql) = SqlStmt(intSql) & DBW("reqinputcnt", tmpInputCnt, 3)
                'SqlStmt(intSql) = SqlStmt(intSql) & "reqinputcnt" & DBNumConv(.ReqAltCnt, 3)
                'If .DataFg = "1" Then
                If .ReqTotCnt = tmpInputCnt Then
                '전체 Verify시 Update
                    SqlStmt(intSql) = SqlStmt(intSql) & DBW("stscd", enStsCd.StsCd_LIS_FinRst, 3) & _
                                                        DBW("vfydt", VfyDt, 3) & _
                                                        DBW("vfytm", VfyTm, 3) & _
                                                        DBW("vfyid", VfyID, 3)
               End If
            End If
            '
            strFootNoteFg = "0"
            strtmp = GetRstTxt(.WorkArea, .AccDt, .accseq)
            If strtmp <> "" Then
                strFootNoteFg = "1"
            Else
                strFootNoteFg = "0"
                If .FootNote <> "" Then
                    strFootNoteFg = "1"
                End If
            End If
            SqlStmt(intSql) = SqlStmt(intSql) & DBW("footnotefg", strFootNoteFg, 2)
            SqlStmt(intSql) = " UPDATE " & T_LAB201 & _
                              " SET    " & SqlStmt(intSql) & _
                              " WHERE  " & DBW("workarea", .WorkArea, 2) & _
                              " AND    " & DBW("accdt", .AccDt, 2) & _
                              " AND    " & DBW("accseq", .accseq, 2)
            blnFirst = True
            '
            DBConn.Execute (SqlStmt(intSql))
        
            'FOOTNOTE 등록
            If strFootNoteFg = "1" Then
                intSql = intSql + 1: ReDim Preserve SqlStmt(intSql)
                If UpdatableLAB304(.WorkArea, .AccDt, .accseq, "1") = False Then
                    SqlStmt(intSql) = "INSERT INTO " & T_LAB304 & " (workarea,accdt,accseq,seq,vfyid,rsttxt) " & _
                                      " VALUES(" & _
                                      DBV("workarea", .WorkArea, 1) & _
                                      DBV("accdt", .AccDt, 1) & _
                                      DBV("accseq", .accseq, 1) & _
                                      DBV("seq", "1", 1) & _
                                      DBV("vfyid", FootNoteId, 1) & _
                                      DBV("rsttxt", .FootNote) & ")"
                Else
                    SqlStmt(intSql) = " UPDATE " & T_LAB304 & _
                                      " SET    " & _
                                                   DBW("rsttxt", .FootNote, 3) & _
                                                   DBW("vfyid", FootNoteId, 2) & _
                                      " WHERE  " & DBW("workarea", .WorkArea, 2) & _
                                      " AND    " & DBW("accdt", .AccDt, 2) & _
                                      " AND    " & DBW("accseq", .accseq, 2) & _
                                      " AND    " & DBW("seq", "1", 2)
                End If
            End If
        
            DBConn.Execute (SqlStmt(intSql))
        
            DBConn.CommitTrans
        
            Erase SqlStmt
            intSql = 0: ReDim SqlStmt(0)
            
            '** 전주예수병원 추가 루틴 =============================================================
            Erase SqlTransfer
            intTempSql = 0: ReDim SqlTransfer(0)
            '=======================================================================================
            
            blnFirst = False
   '
Skip:
        End With

    Next ii
    
    '감염관리결과등록
    Call ICSGeneralResultCheck(objRst)
    
    Exit Function

DBExecError:
'    Call Error_Routine
    MsgBox "처리 도중 오류가 발생하였습니다.", vbExclamation
    m_LastServerErrText = "접수번호 " & sWA & "-" & sAD & "-" & sAS & " 등록시 오류가 발생했습니다." & vbCrLf & _
                          "접수번호를 메모하신 후 전산실 혹은 임상병리과로 연락바랍니다. (3577)"

End Function

'Public Function ResultModifyEntry(ByRef objRst As clsPatientInfo, ByVal MfyDt As String, _
                                  ByVal MfyTm As String, ByVal MfyId As String) As Boolean

Public Function ResultModifyEntry(ByRef objRst As Object, ByVal MfyDt As String, _
                                  ByVal MfyTm As String, ByVal MfyID As String) As Boolean
    
    Dim SqlStmt()       As String
    Dim aryLab102()     As String
    Dim strtmp          As String
    Dim strSQL          As String
    
    Dim strValues       As String
    Dim strWhere        As String
    Dim strLab302Where  As String
    Dim strSEQ          As String
    Dim strLastVfyDt    As String
    Dim strLastVfyTm    As String
    Dim strLastVfyID    As String
    Dim strMSeq         As String
    
    
    Dim blnTxtFgChange  As Boolean

    Dim intSql          As Integer
    Dim intOrdSql       As Integer
    Dim ii              As Long
    
    '-- 전주예수병원 추가 변수=============
    Dim SqlTransfer()   As String
    Dim intTempSql      As Integer
    Dim clsTransfer     As New clsLISTransfer
    Dim strVfyDt        As String
    Dim strRstVal       As String
    Dim strIPAdr        As String
    Dim strMfyDt        As String
    Dim strOCSSEQ       As String
    Dim strMKey         As String
    Dim strSKey         As String
    Dim strWrkDiv       As String
    Dim strOCSNote      As String
    Dim strFromRef      As String
    Dim strToRef        As String
    Dim bINUPFlag       As Boolean
    '======================================
    
   '/* 결과수정 */
   ' LAB102 : StsCd : 처방BODY별로 Stauts가 '4'이면 '5'로 Follow-Up
   ' LAB302 : *RstVal,*RstCd,*HLDiv,*DPDiv,TxtFg,*RstType :UPDATE
   ' LAB304 : *RstText : INSERT
   ' LAB305 : *RstText : INSERT
   ' LAB308 :      *      : INSERT
   ' *표시된 필드는 스프레드화면 동작시 이미 셋팅되어 있는 필드.
   '
    On Error GoTo DBExecError
    
    '-- 전주예수병원IP Address =============
    strIPAdr = clsTransfer.LocalIP_Address
    '=======================================
    
    DBErrClear
    intSql = 0: ReDim SqlStmt(0)
    
    '-- 전주예수병원 추가 변수=============
    intTempSql = 0: ReDim SqlTransfer(0)
    '======================================
    
    With objRst
        If .FootNoteFg = "0" And Trim(.mFootNote) <> "" Then
            '결과수정 이전에 FootNote내역이 없으면 접수내역에 FootNoteFg를 '1'로 셋팅.
            .FootNoteFg = "1"
        End If
        SqlStmt(intSql) = " UPDATE " & T_LAB201 & _
                          " SET    " & _
                                       DBW("footnotefg", .FootNoteFg, 3) & _
                                       DBW("rmkcd", .RmkCd, 2) & _
                          " WHERE  " & DBW("workarea", .WorkArea, 2) & _
                          " AND    " & DBW("accdt", .AccDt, 2) & _
                          " AND    " & DBW("accseq", .accseq, 2)
        If Trim(.mFootNote) <> "" Then
            'FOOTNOTE 수정내역
            intSql = intSql + 1: ReDim Preserve SqlStmt(intSql)
            strSEQ = GetResultSeq(T_LAB304, .WorkArea, .AccDt, .accseq)
            If strSEQ = "1" Or .StsCd = "5" Then
                '기존에 FOOTNOTE없거나 STATUS가 결과상태면 : INSERT
                SqlStmt(intSql) = " INSERT INTO " & T_LAB304 & " (workarea,accdt,accseq,seq,vfyid,rsttxt) " & _
                                  " VALUES(" & _
                                  DBV("workarea", .WorkArea, 1) & _
                                  DBV("accdt", .AccDt, 1) & _
                                  DBV("accseq", .accseq, 1) & _
                                  DBV("seq", strSEQ, 1) & _
                                  DBV("vfyid", MfyID, 1) & _
                                  DBV("rsttxt", .mFootNote) & ")"
            Else
                '기존에 FOOTNOTE가 있으나 STATUS가 결과이전이면 : UPDATE
                SqlStmt(intSql) = " UPDATE " & T_LAB304 & _
                                  " SET    " & _
                                               DBW("rsttxt", .mFootNote, 3) & _
                                               DBW("vfyid", MfyID, 2) & _
                                  " WHERE  " & DBW("workarea", .WorkArea, 2) & _
                                  " AND    " & DBW("accdt", .AccDt, 2) & _
                                  " AND    " & DBW("accseq", .accseq, 2) & _
                                  " AND    " & DBW("seq", "1", 2)
            End If
        End If
    End With
   '
    For ii = 1 To objRst.TestCount
        With objRst.Result.Item(ii)
            '결과내역 수정
            If .RstDiv <> "*" Then
                '결과내역 수정의 Where절
                strWhere = " WHERE  " & DBW("workarea", .WorkArea, 2) & _
                           " AND    " & DBW("accdt", .AccDt, 2) & _
                           " AND    " & DBW("accseq", .accseq, 2) & _
                           " AND    " & DBW("testcd", .TestCd, 2)
                If .DataFg <> "0" And .MfyFg = "0" Then
                    .MfyFg = "1"
                End If
                If .TxtFg <> "2" And Trim(.SuppText) <> "" Then
                    blnTxtFgChange = False
                    '텍스트결과가 수정되어 Supplementary Report가 추가되었으면 LAB305에 INSERT
                    .TxtFg = "2"
                    blnTxtFgChange = True
                End If
                '결과내역 : DataFg - '0':변화없음,'1':결과변화,,'2':결과,장비변화,'3':장비변화
                '              : 결과가 수정되면 LAB308(수정HISTORY내역) INSERT
                '                & LAB302(기존결과내역) UPDATE
                
                
                Select Case Val(.DataFg)
                    Case 0
                        strSQL = " UPDATE " & T_LAB303 & _
                                          " SET    " & _
                                                       DBW("rsttxt", .textrst, 2) & _
                                          " WHERE  " & DBW("workarea", .WorkArea, 2) & _
                                          " AND    " & DBW("accdt", .AccDt, 2) & _
                                          " AND    " & DBW("accseq", .accseq, 2) & _
                                          " AND    " & DBW("testcd", .TestCd, 2)
                        DBConn.Execute strSQL
                    Case 1 To 2:
                        '결과수정 or 결과,장비수정
                        
                        strValues = DBW("rstval", .mRstval, 3) & _
                                    DBW("rstcd", .mRstCd, 3) & _
                                    DBW("hldiv", .MHLDiv, 3) & _
                                    DBW("dpdiv", .MDPDiv, 3) & _
                                    DBW("txtfg", .TxtFg, 3) & _
                                    DBW("mfyfg", .MfyFg, 3) & _
                                    DBW("rptfg", "", 3) & _
                                    DBW("rsttype", .MRstType, 3)   '2000.12.29 김미경 막음 : 최근결과 가져올때 영향을 미치기 때문& _
                                    dbw("vfydt",MfyDt, 3) & _
                                    dbw("vfytm",MfyTm, 3)

                        If .DataFg = "1" Then
                            '결과만 수정
                            strValues = strValues & DBW("vfyid", MfyID, 2)
                        Else
                            '결과와 장비 모두 수정
                            strValues = strValues & DBW("vfyid", MfyID, 3) & _
                                                    DBW("eqpcd", .EqpCd, 2)
                        End If
                        intSql = intSql + 1: ReDim Preserve SqlStmt(intSql)
                        SqlStmt(intSql) = "UPDATE " & T_LAB302 & " SET " & strValues & strWhere
                        '결과HISTORY내역(LAB308) 등록
                        intSql = intSql + 1: ReDim Preserve SqlStmt(intSql)
                        
                        strLastVfyDt = .VfyDt
                        strLastVfyTm = .VfyTm
                        strLastVfyID = .VfyID
                        Call GetResultModifySeq(T_LAB308, .WorkArea, .AccDt, .accseq, .TestCd, strLastVfyDt, strLastVfyTm, strLastVfyID, strMSeq)
                        

                        SqlStmt(intSql) = "INSERT INTO " & T_LAB308 & _
                                          "(workarea,accdt,accseq,testcd,seq,rstval,rstcd,rstunit," & _
                                          "hldiv,dpdiv,mfydt,mfytm,mfyid,ptid,vfydt,vfytm,vfyid,rsttype)  VALUES(" & _
                                          DBV("workarea", .WorkArea, 1) & _
                                          DBV("accdt", .AccDt, 1) & _
                                          DBV("accseq", .accseq, 1) & _
                                          DBV("testcd", .TestCd, 1) & _
                                          DBV("seq", strMSeq, 1) & _
                                          DBV("rstval", .RstVal, 1) & _
                                          DBV("rstcd", .RstCd, 1) & _
                                          DBV("rstunit", .RstUnit, 1) & _
                                          DBV("hldiv", .HLDiv, 1) & _
                                          DBV("dpdiv", .DPDiv, 1) & _
                                          DBV("mfydt", MfyDt, 1) & _
                                          DBV("mfytm", MfyTm, 1) & _
                                          DBV("mfyid", MfyID, 1) & _
                                          DBV("ptid", objRst.PtId, 1) & _
                                          DBV("vfydt", strLastVfyDt, 1) & _
                                          DBV("vfytm", strLastVfyTm, 1) & _
                                          DBV("vfyid", strLastVfyID, 1) & _
                                          DBV("rsttype", .RstType) & ")"
                        
                        '처방BODY내역 STATUS UPDATE
                        '2000.12.11 김미경 수정 : 수정일, 수정시간, 수정자 반영
                        intSql = intSql + 1: ReDim Preserve SqlStmt(intSql)
                        
                        '** 원본 ==============================================================================
'                        SqlStmt(intSql) = " UPDATE " & T_LAB102 & _
'                                          " SET    " & _
'                                                       DBW("stscd", enStsCd.StsCd_LIS_Modify, 3) & _
'                                                       DBW("examdt", MfyDt, 3) & _
'                                                       DBW("examtm", MfyTm, 3) & _
'                                                       DBW("examdoct", MfyID, 2) & _
'                                          " WHERE  " & DBW("ptid", objRst.PtId, 2) & _
'                                          " AND    " & DBW("orddt", .OrdDt, 2) & _
'                                          " AND    " & DBW("ordno", .OrdNo, 2) & _
'                                          " AND    " & DBW("ordseq", .OrdSeq, 2)
                        '=======================================================================================
                        
                        '** 전주예수병원 추가 루틴 =============================================================
                        SqlStmt(intSql) = " UPDATE mdexmort " & _
                                          " SET    " & _
                                                       DBW("stscd", enStsCd.StsCd_LIS_Modify, 2) & _
                                          " WHERE  patno = " & DBS(objRst.PtId) & _
                                          " AND    orddate = TO_DATE(" & DBS(.OrdDt) & ", 'yyyymmdd') " & _
                                          " AND    ordseqno = " & DBN(.OrdNo)
                        '=======================================================================================
                        '
                        strSQL = " UPDATE " & T_LAB303 & _
                                          " SET    " & _
                                                       DBW("rsttxt", .textrst, 2) & _
                                          " WHERE  " & DBW("workarea", .WorkArea, 2) & _
                                          " AND    " & DBW("accdt", .AccDt, 2) & _
                                          " AND    " & DBW("accseq", .accseq, 2) & _
                                          " AND    " & DBW("testcd", .TestCd, 2)
                        DBConn.Execute strSQL
                    Case 3:
                        '장비수정
                        strValues = DBW("eqpcd", .EqpCd, 2)
                        intSql = intSql + 1: ReDim Preserve SqlStmt(intSql)
                        SqlStmt(intSql) = "UPDATE " & T_LAB302 & " SET " & strValues & strWhere
                End Select
                '
                If Trim(.SuppText) <> "" Then
'                    If .txtfg = "1" Then
'                        .SuppText = .SuppText & vbNewLine & "♣ 수정사유 : " & Format(MfyDt, "####-##-##") & "  " & Format(MfyTm, "0#:##:##")
'                    End If
                    'Supplementary결과 등록
                    intSql = intSql + 1: ReDim Preserve SqlStmt(intSql)
                    SqlStmt(intSql) = "INSERT INTO " & T_LAB305 & " (workarea,accdt,accseq,testcd, seq," & _
                                      "mfydt,mfytm,mfyid,mfyrsn,rsttxt) VALUES(" & _
                                      DBV("workarea", .WorkArea, 1) & _
                                      DBV("accdt", .AccDt, 1) & _
                                      DBV("accseq", .accseq, 1) & _
                                      DBV("testcd", .TestCd, 1) & _
                                      DBV("seq", GetResultSeq(T_LAB305, .WorkArea, .AccDt, .accseq), 1) & _
                                      DBV("mfydt", MfyDt, 1) & _
                                      DBV("mfytm", MfyTm, 1) & _
                                      DBV("mfyid", MfyID, 1) & _
                                      DBV("mfyrsn", .MfyRsn, 1) & _
                                      DBV("rsttxt", .SuppText) & ")"
                End If
                If .DataFg = "0" Or .DataFg = "3" Then
                    If blnTxtFgChange = True Then
                        'Supplementary Report만 추가되고 결과수정이 않되면 결과내역에 TxtFg만 UPDATE
                        intSql = intSql + 1: ReDim Preserve SqlStmt(intSql)
                        SqlStmt(intSql) = "UPDATE " & T_LAB302 & " SET  " & DBW("txtfg", "2", 2) & strWhere
                    End If
                End If
                
                '** 전주예수병원 추가 루틴 =============================================================
                '** 전 결과 가져오기(OCS측 결과 History 전결과 등록에 필요)
                'bINUPFlag = clsTransfer.SqlOCSRESULT_FLAG(objRst.PtId, .OrdDt, .OrdNo, .TestCd)
                
                If .RstDiv = "*" Then
                    GoTo OCSSKIP
                End If
                
                '-- Comments
                If objRst.FootNote <> "" Then
                    strOCSNote = "[FootNote]" & Chr(13) & objRst.FootNote & Chr(13)
                Else
                    strOCSNote = ""
                End If
                
                '-- 수정사유
                If objRst.mFootNote <> "" Then
                    strOCSNote = strOCSNote & "<<수정사유>>" & Chr(13) & objRst.mFootNote & Chr(13)
                End If
                
                '-- Remark
                If objRst.rmknm <> "" Then
                    strOCSNote = strOCSNote & "◈ Remark" & Chr(13) & objRst.rmknm
                End If
                
                strVfyDt = .VfyDt & .VfyTm
                strMfyDt = MfyDt & MfyTm
                
                Call clsTransfer.ResultValue_Find(.TestCd, .mRstCd)
                
                If strSKey <> objRst.PtId & .OrdDt & .OrdNo Then
                    '-- 결과전송 시 필요정보
                    Call clsTransfer.SqlOCS_Info_New(.WorkArea, .AccDt, .accseq)
                
                    intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
                    
                    '-- 1. '3' 종합건진, '4' 일반검사, 이외검사로 구분을 한다.
                    strWrkDiv = clsTransfer.SqlOCBussdiv_Flag(objRst.PtId, .OrdDt, .OrdNo)
                    If strWrkDiv = "3" Then
                        If clsTransfer.ItemVal <> "" Then
                            If clsTransfer.SqlOCSSU2EXAMT_Find(objRst.PtId, .OrdDt, .OrdNo) = True Then
                                SqlTransfer(intTempSql) = clsTransfer.SqlOCSSU2EXAMT_UPDATE_New3(objRst.PtId, .OrdDt, .OrdNo, _
                                                        "N", clsTransfer.ItemVal, strOCSNote, clsTransfer.RcvDt & clsTransfer.RcvTm, _
                                                        strVfyDt, .VfyID, strVfyDt, .VfyID, MfyID, strIPAdr, strMfyDt, enStsCd.StsCd_LIS_Modify)
                            Else
                                GoTo DBExecError
                            End If
                        Else
                            If clsTransfer.SqlOCSSU2EXAMT_Find(objRst.PtId, .OrdDt, .OrdNo) = True Then
                                SqlTransfer(intTempSql) = clsTransfer.SqlOCSSU2EXAMT_UPDATE_New3(objRst.PtId, .OrdDt, .OrdNo, _
                                                        "N", .mRstCd, strOCSNote, clsTransfer.RcvDt & clsTransfer.RcvTm, _
                                                        strVfyDt, .VfyID, strVfyDt, .VfyID, MfyID, strIPAdr, strMfyDt, enStsCd.StsCd_LIS_Modify)
                            Else
                                GoTo DBExecError
                            End If

                        End If
                    ElseIf strWrkDiv = "4" Then
                            If clsTransfer.ItemVal <> "" Then
                                If clsTransfer.SqlOCSSG2EXAMT_Find(objRst.PtId, .OrdDt, .OrdNo) = True Then
                                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSSG2EXAMT_UPDATE_NEW(objRst.PtId, .OrdDt, .OrdNo, _
                                                            "N", clsTransfer.ItemVal, strOCSNote, clsTransfer.RcvDt & clsTransfer.RcvTm, _
                                                            strVfyDt, .VfyID, strVfyDt, .VfyID, MfyID, strIPAdr, strMfyDt, enStsCd.StsCd_LIS_Modify)
                                Else
                                    GoTo DBExecError
                                End If
                            Else
                                If clsTransfer.SqlOCSSG2EXAMT_Find(objRst.PtId, .OrdDt, .OrdNo) = True Then
                                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSSG2EXAMT_UPDATE_NEW(objRst.PtId, .OrdDt, .OrdNo, _
                                                            "N", .mRstCd, strOCSNote, clsTransfer.RcvDt & clsTransfer.RcvTm, _
                                                            strVfyDt, .VfyID, strVfyDt, .VfyID, MfyID, strIPAdr, strMfyDt, enStsCd.StsCd_LIS_Modify)
                                Else
                                    GoTo DBExecError
                                End If
                            End If
                    Else
                        If clsTransfer.SqlOCSMDEXMORT_Find(objRst.PtId, .OrdDt, .OrdNo) = True Then
                            SqlTransfer(intTempSql) = clsTransfer.SqlOCSMDEXMORT_UPDATE_NEW(objRst.PtId, .OrdDt, .OrdNo, _
                                                    "N", strVfyDt, .VfyID, MfyID, strIPAdr, strMfyDt)
                        Else
                            GoTo DBExecError
                        End If
                    End If
                    
                    strSKey = objRst.PtId & .OrdDt & .OrdNo
                    
                End If
                
                '** 종건/일건 이외의 검사만 등록.
                If strWrkDiv <> "" Then
                    GoTo OCSSKIP
                End If
                
                '-- 참고치값 Check
                If .RefValFrom = "0" And .RefValTo = "0" Then
                    strFromRef = ""
                    strToRef = ""
                    If .RefCd <> "" Then
                        strFromRef = .RefCd
                    End If
                Else
                    strFromRef = .RefValFrom
                    strToRef = .RefValTo
                End If
                
                '-- 결과History내역 순번 가져오기
                strOCSSEQ = clsTransfer.SqlOCSRESULT_HISTORY(objRst.PtId, .OrdDt, .OrdNo, .TestCd)
                
                If strOCSSEQ = "" Then
                    strOCSSEQ = "1"
                Else
                    strOCSSEQ = CInt(strOCSSEQ) + 1
                End If
                
                '** 검사항목별 결과값 체크
                intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
                
                If clsTransfer.ItemVal <> "" Then
                    Select Case .TestCd
                        Case "ABOC22B", "ABOC22C", "ABOB22B", "ABOB22C", "RHD22B", "ABSC22B", "ABSC22C", _
                             "LABOC22B", "LABOC22C", "LABOB22B", "LABOB22C", "LRHD22B", "LABSC22B", "LABSC22C"
                        Case Else
                            If clsTransfer.SqlOCSRESULT_FLAG(objRst.PtId, .OrdDt, .OrdNo, .TestCd) = True Then
                                SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_UPDATE(objRst.PtId, .OrdDt, .OrdNo, _
                                                        .TestCd, clsTransfer.SlipCd, clsTransfer.ItemVal, .textrst & vbNewLine & .SuppText, strFromRef, strToRef, .RstUnit, _
                                                        strOCSNote, clsTransfer.RcvDt & clsTransfer.RcvTm, strVfyDt, .VfyID, "", _
                                                        strVfyDt, clsTransfer.RsltType, MfyID, strIPAdr, strMfyDt)
                            Else
                                GoTo DBExecError
                            End If
                    End Select
                
'''                    If clsTransfer.SqlOCSRESULT_FLAG(objRst.PtId, .OrdDt, .OrdNo, .TestCd) = True Then
'''                        SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_UPDATE(objRst.PtId, .OrdDt, .OrdNo, _
'''                                                .TestCd, clsTransfer.SlipCd, clsTransfer.ItemVal, .textrst & vbNewLine & .SuppText, strFromRef, strToRef, .RstUnit, _
'''                                                strOCSNote, clsTransfer.RcvDt & clsTransfer.RcvTm, strVfyDt, .VfyID, "", _
'''                                                strVfyDt, clsTransfer.RsltType, MfyID, strIPAdr, strMfyDt)
'''                    Else
'''                        GoTo DBExecError
'''                    End If
                    
                Else
                    If clsTransfer.SqlOCSRESULT_FLAG(objRst.PtId, .OrdDt, .OrdNo, .TestCd) = True Then
                        SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_UPDATE(objRst.PtId, .OrdDt, .OrdNo, _
                                                .TestCd, clsTransfer.SlipCd, .mRstCd, .textrst & vbNewLine & .SuppText, strFromRef, strToRef, .RstUnit, _
                                                strOCSNote, clsTransfer.RcvDt & clsTransfer.RcvTm, strVfyDt, .VfyID, "", _
                                                strVfyDt, clsTransfer.RsltType, MfyID, strIPAdr, strMfyDt)
                    Else
                        GoTo DBExecError
                    End If

                End If
                
                '-- 전결과 Set
                Call clsTransfer.SqlOCSRESULT_OLD(objRst.PtId, .OrdDt, .OrdNo, .TestCd)
                
                intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
                
                '-- OCS 결과History테이블 INSERT
                SqlTransfer(intTempSql) = clsTransfer.SqlOCSRSLTHT_INSERT(objRst.PtId, .OrdDt, .OrdNo, _
                                        .TestCd, strOCSSEQ, clsTransfer.SlipCd, clsTransfer.RstVal, clsTransfer.RstText, _
                                        clsTransfer.RefValTo, clsTransfer.RefValFrom, clsTransfer.RstUnit, _
                                        clsTransfer.FootNote, clsTransfer.RcvDt & clsTransfer.RcvTm, clsTransfer.VfyDt & clsTransfer.VfyTm, _
                                        clsTransfer.VfyID, clsTransfer.VfyDt & clsTransfer.VfyTm, clsTransfer.VfyID, _
                                        clsTransfer.MfyDt & clsTransfer.MfyTm, clsTransfer.RsltType, _
                                        clsTransfer.MfyID, clsTransfer.IPAdr, clsTransfer.MfyDt & clsTransfer.MfyTm)
                            
                '=======================================================================================
                
            Else
                '** 전주예수병원 추가 루틴 =============================================================
                '-- 대표항목 처리루틴
                
                If .RstDiv = "*" Then
                    GoTo OCSSKIP
                End If
                
                If objRst.FootNote <> "" Then
                    strOCSNote = "[FootNote]" & Chr(13) & objRst.FootNote & Chr(13)
                Else
                    strOCSNote = ""
                End If
                
                If objRst.rmknm <> "" Then
                    strOCSNote = strOCSNote & "◈ Remark" & Chr(13) & objRst.rmknm
                End If
                
                strVfyDt = .VfyDt & .VfyTm
                strMfyDt = MfyDt & MfyTm
                
                '** 검사항목별 결과값 체크
                Call clsTransfer.ResultValue_Find(.TestCd, .mRstCd)
                
                If strMKey <> objRst.PtId & .OrdDt & .OrdNo Then
                    '-- 결과전송 시 필요정보
                    Call clsTransfer.SqlOCS_Info_New(.WorkArea, .AccDt, .accseq)
                
                    intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
                    
                    '-- 1. '3' 종합건진과 '4' 일반검사 구분을 한다.
                    strWrkDiv = clsTransfer.SqlOCBussdiv_Flag(objRst.PtId, .OrdDt, .OrdNo)
                    If strWrkDiv = "3" Then
                        If clsTransfer.ItemVal <> "" Then
                            SqlTransfer(intTempSql) = clsTransfer.SqlOCSSU2EXAMT_UPDATE_New(objRst.PtId, .OrdDt, .OrdNo, _
                                                    "N", clsTransfer.ItemVal, strOCSNote, clsTransfer.RcvDt & clsTransfer.RcvTm, _
                                                    strVfyDt, .VfyID, strVfyDt, .VfyID, MfyID, strIPAdr, strMfyDt, enStsCd.StsCd_LIS_Modify)
                        Else
                            SqlTransfer(intTempSql) = clsTransfer.SqlOCSSU2EXAMT_UPDATE_New(objRst.PtId, .OrdDt, .OrdNo, _
                                                    "N", .mRstCd, strOCSNote, clsTransfer.RcvDt & clsTransfer.RcvTm, _
                                                    strVfyDt, .VfyID, strVfyDt, .VfyID, MfyID, strIPAdr, strMfyDt, enStsCd.StsCd_LIS_Modify)
                        End If
                    ElseIf strWrkDiv = "4" Then
                            If clsTransfer.ItemVal <> "" Then
                                SqlTransfer(intTempSql) = clsTransfer.SqlOCSSG2EXAMT_UPDATE_NEW(objRst.PtId, .OrdDt, .OrdNo, _
                                                        "N", clsTransfer.ItemVal, strOCSNote, clsTransfer.RcvDt & clsTransfer.RcvTm, _
                                                        strVfyDt, .VfyID, strVfyDt, .VfyID, MfyID, strIPAdr, strMfyDt, enStsCd.StsCd_LIS_Modify)
                            Else
                                SqlTransfer(intTempSql) = clsTransfer.SqlOCSSG2EXAMT_UPDATE_NEW(objRst.PtId, .OrdDt, .OrdNo, _
                                                        "N", .mRstCd, strOCSNote, clsTransfer.RcvDt & clsTransfer.RcvTm, _
                                                        strVfyDt, .VfyID, strVfyDt, .VfyID, MfyID, strIPAdr, strMfyDt, enStsCd.StsCd_LIS_Modify)
                            End If
                    Else
                        SqlTransfer(intTempSql) = clsTransfer.SqlOCSMDEXMORT_UPDATE_NEW(objRst.PtId, .OrdDt, .OrdNo, _
                                                    "N", strVfyDt, .VfyID, MfyID, strIPAdr, strMfyDt)
                    End If
                    
                    strMKey = objRst.PtId & .OrdDt & .OrdNo
                    
                End If
                
                '** 종건/일건 이외의 검사만 등록.
                If strWrkDiv <> "" Then
                    GoTo OCSSKIP
                End If
                
                '-- 참고치값 Check
                If .RefValFrom = "0" And .RefValTo = "0" Then
                    strFromRef = ""
                    strToRef = ""
                    If .RefCd <> "" Then
                        strFromRef = .RefCd
                    End If
                Else
                    strFromRef = .RefValFrom
                    strToRef = .RefValTo
                End If
                
                '-- 2. 건진외결과 INSERT
                strOCSSEQ = clsTransfer.SqlOCSRESULT_HISTORY(objRst.PtId, .OrdDt, .OrdNo, .TestCd)
                
                If strOCSSEQ = "" Then
                    strOCSSEQ = "1"
                Else
                    strOCSSEQ = CInt(strOCSSEQ) + 1
                End If
                
                intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)

                If clsTransfer.ItemVal <> "" Then
                    If clsTransfer.SqlOCSRESULT_FLAG(objRst.PtId, .OrdDt, .OrdNo, .TestCd) = True Then
                        SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_UPDATE(objRst.PtId, .OrdDt, .OrdNo, _
                                                .TestCd, clsTransfer.SlipCd, clsTransfer.ItemVal, .textrst & vbNewLine & .SuppText, .RefValTo, .RefValFrom, .RstUnit, _
                                                strOCSNote, clsTransfer.RcvDt & clsTransfer.RcvTm, strVfyDt, .VfyID, _
                                                "", strMfyDt, clsTransfer.RsltType, MfyID, strIPAdr, strMfyDt)
                    Else
                        GoTo DBExecError
                    End If

                Else
                    If clsTransfer.SqlOCSRESULT_FLAG(objRst.PtId, .OrdDt, .OrdNo, .TestCd) = True Then
                        SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_UPDATE(objRst.PtId, .OrdDt, .OrdNo, _
                                                .TestCd, clsTransfer.SlipCd, .mRstCd, .textrst & vbNewLine & .SuppText, .RefValTo, .RefValFrom, .RstUnit, _
                                                strOCSNote, clsTransfer.RcvDt & clsTransfer.RcvTm, strVfyDt, .VfyID, _
                                                "", strMfyDt, clsTransfer.RsltType, MfyID, strIPAdr, strMfyDt)
                    Else
                        GoTo DBExecError
                    End If

                End If

                intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
                
                '-- OCS 결과History테이블 INSERT
                SqlTransfer(intTempSql) = clsTransfer.SqlOCSRSLTHT_INSERT(objRst.PtId, .OrdDt, .OrdNo, _
                                        .TestCd, strOCSSEQ, .SlipCd, clsTransfer.RstVal, clsTransfer.RstText, _
                                        clsTransfer.RefValTo, clsTransfer.RefValFrom, clsTransfer.RstUnit, _
                                        clsTransfer.FootNote, clsTransfer.RcvDt & clsTransfer.RcvTm, clsTransfer.VfyDt & clsTransfer.VfyTm, _
                                        clsTransfer.VfyID, clsTransfer.VfyDt & clsTransfer.VfyTm, clsTransfer.VfyID, _
                                        clsTransfer.VfyDt & clsTransfer.VfyTm, clsTransfer.RsltType, _
                                        clsTransfer.MfyID, clsTransfer.IPAdr, clsTransfer.MfyDt & clsTransfer.MfyTm)
                
                '=======================================================================================
            End If
            
OCSSKIP:

        End With
    Next ii
    '
    '결과보고 대기내역(LAB202) Insert
    Dim tmpSql As String
    Dim tmpRs As Recordset
    Dim tmpDept As String
    Dim tmpBussDiv As String
    
    Dim KK As Integer
    
    For KK = 1 To objRst.TestCount
       With objRst.Result.Item(KK)
           tmpSql = " SELECT bussdiv FROM " & T_LAB101 & _
                  " WHERE " & DBW("ptid=", objRst.PtId) & _
                  " AND " & DBW("orddt=", .OrdDt) & _
                  " AND " & DBW("ordno=", .OrdNo)
           Set tmpRs = Nothing
           Set tmpRs = New Recordset
           tmpRs.Open tmpSql, DBConn
           
           If Not tmpRs.EOF Then
               tmpBussDiv = Trim(tmpRs.Fields("bussdiv").Value & "")
               If objRst.WardId = "" Then
                   tmpDept = objRst.DeptCd
               Else
                   tmpDept = objRst.WardId
               End If
'               tmpRs.RsClose
               Set tmpRs = Nothing
               Exit For
           End If
       End With
    Next

    tmpSql = " SELECT * FROM " & T_LAB202 & _
             " WHERE  " & DBW("deptcd", tmpDept, 2) & _
             " AND    " & DBW("vfydt", MfyDt, 2) & _
             " AND    " & DBW("ptid", objRst.PtId, 2) & _
             " AND    " & DBW("mfyfg", "1", 2)
'    Set tmpRs = OpenRecordSet(tmpSql)
    Set tmpRs = Nothing
    Set tmpRs = New Recordset
    tmpRs.Open tmpSql, DBConn
    
    If tmpRs.EOF Then
        intSql = intSql + 1: ReDim Preserve SqlStmt(intSql)
        
        SqlStmt(intSql) = " INSERT INTO " & T_LAB202 & " (deptcd, vfydt, ptid, mfyfg, vfytm, vfyid, donefg, doneid, majdoct, bussdiv) " & _
                          " VALUES ( " & _
                          DBV("deptcd", tmpDept, 1) & _
                          DBV("vfydt", MfyDt, 1) & _
                          DBV("ptid", objRst.PtId, 1) & _
                          DBV("mfyfg", "1", 1) & _
                          DBV("vfytm", MfyTm, 1) & _
                          DBV("vfyid", MfyID, 1) & _
                          DBV("donefg", "", 1) & _
                          DBV("doneid", "", 1) & _
                          DBV("majdoct", objRst.OrdDoct, 1) & _
                          DBV("bussdiv", tmpBussDiv) & ") "
    Else
        If "" & tmpRs.Fields("DoneFg").Value = "1" Then
            intSql = intSql + 1: ReDim Preserve SqlStmt(intSql)
            SqlStmt(intSql) = " UPDATE " & T_LAB202 & _
                              " set    " & _
                                           DBW("vfytm", MfyTm, 3) & _
                                           DBW("vfyid", MfyID, 3) & _
                                           DBW("majdoct", objRst.OrdDoct, 3) & _
                                           DBW("donefg", "", 3) & _
                                           DBW("doneid", "0", 2) & _
                              " WHERE  " & DBW("deptcd", tmpDept, 2) & _
                              " AND    " & DBW("vfydt", MfyDt, 2) & _
                              " AND    " & DBW("ptid", objRst.PtId, 2) & _
                              " AND    " & DBW("mfyfg", "1", 2)
        End If
    End If
    
'    tmpRs.RsClose
    Set tmpRs = Nothing
   
    DBConn.BeginTrans
    
    For intSql = LBound(SqlStmt) To UBound(SqlStmt)
        DBConn.Execute (SqlStmt(intSql))
        Debug.Print SqlStmt(intSql)
    Next intSql
    
    '** 전주예수병원 추가 루틴 =============================================================
    For intTempSql = LBound(SqlTransfer) To UBound(SqlTransfer)
        intTempSql = intTempSql
        'Debug.Print SqlTransfer(intTempSql)
        If Trim(SqlTransfer(intTempSql)) <> "" Then DBConn.Execute (SqlTransfer(intTempSql))
        
    Next intTempSql
    '=======================================================================================
    '
   
    DBConn.CommitTrans

    '감염관리결과등록
    Call ICSGeneralResultCheck(objRst, True)
    
    Exit Function
DBExecError:
'    Call Error_Routine
    MsgBox "처리 도중 오류가 발생하였습니다.", vbExclamation
   'Resume Next
End Function

Public Function GetLab201CountInfo(ByVal WorkArea As String, ByVal AccDt As String, _
                                   ByVal accseq As String) As String
    Dim strSQL As String
   '
    strSQL = " SELECT reqtotcnt,reqinpucnt FROM " & T_LAB201 & _
             " WHERE  " & DBW("workarea", WorkArea, 2) & _
             " AND    " & DBW("accdt", AccDt, 2) & _
             " AND    " & DBW("accseq", accseq, 2)
    DBExec strSQL
    '
    If m_RecordSet.RecordCount > 1 Then
        GetLab201CountInfo = "" & m_RecordSet.Fields("REQTOTCNT").Value & vbTab & _
                             "" & m_RecordSet.Fields("REQINPUTCNT").Value & vbTab & _
                             "" & m_RecordSet.Fields("ALTTOTCNT").Value & vbTab & _
                             "" & m_RecordSet.Fields("ALTINPUTCNT").Value
    Else
        GetLab201CountInfo = ""
    End If
   '
End Function

Public Function SqlStmtInstrument() As String
    SqlStmtInstrument = " SELECT a.eqpcd,a.eqpnm,a.remark" & _
                        " FROM   " & T_LAB006 & " a " & _
                        " WHERE  " & DBW("a.eqpdiv", "E", 2) & _
                        " ORDER BY eqpnm"
End Function

Public Function SqlStmtRstCode(ByVal Item As String) As String
    SqlStmtRstCode = " SELECT a.cdval2,a.field1" & _
                     " FROM   " & T_LAB031 & " a " & _
                     " WHERE  " & DBW("a.cdindex", LC2_ItemResult, 2) & _
                     " AND    " & DBW("a.cdval1", Item, 2) & _
                     " ORDER BY a.cdval1"
End Function

Public Function SqlStmtTransfer(ByVal EqpCd As String) As String
    
    Dim strSQL As String
    Dim sqlTransDtTm As String
    Dim sqlSpcNo As String
    'sqlTransDtTm = "" & FUNC_SUBSTR & "(a.transdt,3,8)+' '+" & FUNC_SUBSTR & "(a.transtm,1,2)+':'+" & FUNC_SUBSTR & "(a.transtm,3,2)+':'+" & FUNC_SUBSTR & "(a.transtm,5,2)"
    
    sqlSpcNo = "a.spcyy" & FUNC_CONCAT & "'-'" & FUNC_CONCAT & "" & FUNC_CONVERT("char", "a.spcno")
    SqlStmtTransfer = "SELECT /*+ RULE */ a.transno," & _
                              sqlSpcNo & " AS SpcNo, a.transdt, a.transtm FROM " & T_LAB306 & " a" & _
                      " WHERE " & DBW("a.transdt >=", Format(Date - 7, "yyyymmdd")) & _
                      " AND   exists(SELECT c.accseq FROM " & T_LAB201 & " c, " & T_LAB302 & " d " & _
                                    " WHERE c.spcyy = a.spcyy AND c.spcno = a.spcno " & _
                                    " AND   d.workarea = c.workarea AND d.accdt = c.accdt AND d.accseq = c.accseq " & _
                                    " AND  (d.vfydt = ''  or  d.vfydt is null) AND d.eqpcd=a.eqpcd)" & _
                      " AND   " & DBW("a.eqpcd", EqpCd, 2)
        
'--2000.12.5 김미경 수정 : QC 전송내역도 조회하기 위하여 UNION 걸어줌...
    'SqlStmtTransfer = SqlStmtTransfer & " UNION SELECT a.transno," & _
                      sqlSpcNo & " AS SpcNo, a.transdt, a.transtm FROM " & T_LAB306 & " a" & _
                      " WHERE  " & DBW("a.transdt >=", Format(Date - 1, "yyyymmdd")) & _
                      " AND    exists(SELECT c.accseq FROM " & T_LAB201 & " c, " & T_LAB026 & " d " & _
                                    " WHERE  c.spcyy = a.spcyy AND c.spcno = a.spcno " & _
                                    " AND    d.workarea = c.workarea AND d.accdt = c.accdt AND d.accseq = c.accseq " & _
                                    " AND   (d.vfydt = ''  or  d.vfydt is null) AND d.eqpcd=a.eqpcd)" & _
                      " AND  " & DBW("a.eqpcd", Eqpcd, 2)
    SqlStmtTransfer = SqlStmtTransfer & " ORDER BY transdt, transtm, transno "
    
    'SqlStmtTransfer = "SELECT " & sqlTransDtTm & " AS TransDtTm," & _
       sqlSpcNo & " AS SpcNo FROM " & T_LAB306 & " a" & _
       " WHERE a.transdt >='" & Format(Date - 1, "yyyymmdd") & "'" & _
       " AND exists(SELECT c.accseq FROM " & T_LAB201 & " c, " & T_LAB302 & " d WHERE c.spcyy = a.spcyy AND c.spcno = a.spcno " & _
       " AND d.workarea = c.workarea AND d.accdt = c.accdt AND d.accseq = c.accseq AND d.vfydt = '' AND d.eqpcd=a.eqpcd)" & _
       " AND a.eqpcd = '" & EqpCd & "'"
End Function
      
Public Function IsRstCode(ByVal Item As String) As Boolean
    Dim strSQL As String
    IsRstCode = False
     strSQL = " SELECT COUNT(a.cdval2) AS ITEMCNT" & _
              " FROM   " & T_LAB031 & " a " & _
              " WHERE  " & DBW("a.cdindex", LC2_ItemResult, 2) & _
              " AND    " & DBW("a.cdval1", Item, 2)
    DBOpenDyna strSQL
    If m_RecordSet.RecordCount > 0 Then
       If Val("" & m_RecordSet.Fields("ITEMCNT").Value) > 0 Then
          IsRstCode = True
       End If
    End If

End Function


Public Function GetInputCnt(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Long)

    Dim tmpRs As Recordset
    
    GetInputCnt = 0
    
    Set tmpRs = Nothing
    Set tmpRs = New Recordset
    tmpRs.Open " SELECT count(*) as InputCnt FROM " & T_LAB302 & _
                              " WHERE  " & DBW("workarea", pWorkArea, 2) & _
                              " AND    " & DBW("accdt", pAccDt, 2) & _
                              " AND    " & DBW("accseq", pAccSeq, 2) & _
                              " AND    vfydt <> ' '  AND  vfydt is not null " & _
                              " AND   (detailfg = ''  or  detailfg is null  or rstdiv = '*')", DBConn
    GetInputCnt = GetInputCnt + Val("" & tmpRs.Fields("InputCnt").Value)
    
    Set tmpRs = Nothing
    Set tmpRs = New Recordset
    tmpRs.Open " SELECT count(*) as InputCnt FROM " & T_LAB351 & _
                              " WHERE  " & DBW("workarea", pWorkArea, 2) & _
                              " AND    " & DBW("accdt", pAccDt, 2) & _
                              " AND    " & DBW("accseq", pAccSeq, 2) & _
                              " AND    vfydt <> ' '  AND  vfydt is not null", DBConn
    GetInputCnt = GetInputCnt + Val("" & tmpRs.Fields("InputCnt").Value)
    
    Set tmpRs = Nothing
    Set tmpRs = New Recordset
    tmpRs.Open " SELECT count(*) as InputCnt FROM " & T_LAB404 & _
                              " WHERE  " & DBW("workarea", pWorkArea, 2) & _
                              " AND    " & DBW("accdt", pAccDt, 2) & _
                              " AND    " & DBW("accseq", pAccSeq, 2) & _
                              " AND    vfydt <> ' '  AND  vfydt is not null" & _
                              " AND   (detailfg = ''  or  detailfg is null  or  rstdiv = '*')", DBConn
    GetInputCnt = GetInputCnt + Val("" & tmpRs.Fields("InputCnt").Value)
    Set tmpRs = Nothing
                                         
End Function

Public Function GetUnverifiedList(ByVal pWorkArea As String, ByVal pRcvDt As String, _
                                  ByVal pBuildingCd As String) As String
    Dim pRcvTm  As String
    Dim pFRcvDt As String
    
    'to
    pRcvTm = Mid(pRcvDt, 9, 6)
    'FROM
    pFRcvDt = Mid(pRcvDt, 15)
    
    pRcvDt = Mid(pRcvDt, 1, 8)
    
    '** 원본 -------------------------------------------------------------------------------------
'    GetUnverifiedList = " SELECT a.ptid,a.spcyy,a.spcno,e." & F_PTNM & " as ptnm,a.workarea, a.accdt, a.accseq,a.wardid,a.deptcd,a.hosilid," & _
'                        " a.statfg, a.reqtotcnt, a.reqinputcnt, '0' as TestDiv, c.abbrnm5, b.rstcd," & _
'                        " a.coldt, a.coltm, a.colid, a.rcvdt, a.rcvtm, a.rcvid" & _
'                        " FROM " & T_HIS001 & " e," & T_LAB201 & " a, " & T_LAB302 & " b, " & T_LAB001 & " c, " & T_LAB101 & " d " & _
'                        " WHERE  " & DBW("a.workarea = ", pWorkArea) & _
'                        " AND    " & DBW("a.rcvdt    >= ", pFRcvDt) & _
'                        " AND    a.rcvdt " & FUNC_CONCAT & " a.rcvtm <= " & DBS(pRcvDt & pRcvTm) & _
'                        " AND    " & DBW("a.stscd   >= ", enStsCd.StsCd_LIS_Accession) & _
'                        " AND    " & DBW("a.stscd   <= ", enStsCd.StsCd_LIS_MidRst) & _
'                        " AND    " & DBW("a.buildcd  = ", pBuildingCd) & _
'                        " AND   b.workarea=a.workarea AND b.accdt=a.accdt AND b.accseq=a.accseq " & _
'                        " AND  (b.vfydt='' or b.vfydt is null) " & _
'                        " AND   c.testcd = b.testcd  AND  (c.detailfg = ''  or  c.detailfg is null) " & _
'                        " AND   c.applydt = ( SELECT max(applydt) FROM " & T_LAB001 & _
'                                            " WHERE testcd = b.testcd )" & _
'                        " AND   d.ptid  = b.ptid " & _
'                        " AND   d.orddt = b.orddt " & _
'                        " AND   d.ordno = b.ordno " & _
'                        " AND   " & DBW("d.bussdiv <> ", enBussDiv.BussDiv_ICU) & _
'                        " AND   a.ptid=e." & F_PTID
'    If ObjSysInfo.projectid <> "BBS" Then
'
'        GetUnverifiedList = GetUnverifiedList & " UNION ALL " & _
'                            " SELECT a.ptid,a.spcyy,a.spcno,e." & F_PTNM & " as ptnm, a.workarea, a.accdt, a.accseq," & _
'                            " a.wardid,a.deptcd,a.hosilid,a.statfg, a.reqtotcnt, a.reqinputcnt, '1' as TestDiv," & _
'                            " c.abbrnm5, '' as rstcd, a.coldt, a.coltm, a.colid, a.rcvdt, a.rcvtm, a.rcvid" & _
'                            " FROM " & T_HIS001 & " e," & T_LAB201 & " a, " & T_LAB351 & " b, " & T_LAB001 & " c " & _
'                            " WHERE  " & DBW("a.workarea = ", pWorkArea) & _
'                            " AND    " & DBW("a.rcvdt    >= ", pFRcvDt) & _
'                            " AND    a.rcvdt " & FUNC_CONCAT & " a.rcvtm <= " & DBS(pRcvDt & pRcvTm) & _
'                            " AND    " & DBW("a.stscd   >= ", enStsCd.StsCd_LIS_Accession) & _
'                            " AND    " & DBW("a.stscd   <= ", enStsCd.StsCd_LIS_MidRst) & _
'                            " AND    " & DBW("a.buildcd  = ", pBuildingCd) & _
'                            " AND   b.workarea=a.workarea AND b.accdt=a.accdt AND b.accseq=a.accseq " & _
'                            " AND  (b.vfydt='' or b.vfydt is null) " & _
'                            " AND   c.testcd = b.testcd  AND  (c.detailfg = ''  or  c.detailfg is null) " & _
'                            " AND   c.applydt = ( SELECT max(applydt) FROM " & T_LAB001 & _
'                                                " WHERE testcd = b.testcd )" & _
'                            " AND   a.ptid=e." & F_PTID
'
'        GetUnverifiedList = GetUnverifiedList & " UNION ALL " & _
'                            " SELECT a.ptid,a.spcyy,a.spcno,e." & F_PTNM & " as ptnm, a.workarea, a.accdt, a.accseq," & _
'                            " a.wardid,a.deptcd,a.hosilid,a.statfg, a.reqtotcnt, a.reqinputcnt, '2' as TestDiv, c.abbrnm5, b.rstcd," & _
'                            " a.coldt, a.coltm, a.colid, a.rcvdt, a.rcvtm, a.rcvid" & _
'                            " FROM " & T_HIS001 & " e," & T_LAB201 & " a, " & T_LAB404 & " b, " & T_LAB001 & " c " & _
'                            " WHERE  " & DBW("a.workarea = ", pWorkArea) & _
'                            " AND    " & DBW("a.rcvdt    >= ", pFRcvDt) & _
'                            " AND    a.rcvdt " & FUNC_CONCAT & " a.rcvtm <= " & DBS(pRcvDt & pRcvTm) & _
'                            " AND    " & DBW("a.stscd   >= ", enStsCd.StsCd_LIS_Accession) & _
'                            " AND    " & DBW("a.stscd   <= ", enStsCd.StsCd_LIS_MidRst) & _
'                            " AND    " & DBW("a.buildcd  = ", pBuildingCd) & _
'                            " AND   b.workarea=a.workarea AND b.accdt=a.accdt AND b.accseq=a.accseq " & _
'                            " AND  (b.vfydt='' or b.vfydt is null) " & _
'                            " AND   c.testcd = b.testcd  AND  (c.detailfg = ''  or  c.detailfg is null) " & _
'                            " AND   c.applydt = ( SELECT max(applydt) FROM " & T_LAB001 & _
'                                                " WHERE testcd = b.testcd )" & _
'                            " AND   a.ptid=e." & F_PTID & _
'                            " order by accdt, accseq"
'    End If
    '---------------------------------------------------------------------------------------------
    
    '** 변경 By M.G.Choi
    ' 08.10.22 - 24 양성현 장기간 검색시 검색 시간 단축을 위해 인데스를 변경 추가함.
    ' 기 s2lab201에 걸린 index 2는 검색 시간이 너무 오래걸림. 그 원인은 검색기간을 먼저 설정한 다음 WA를 선택 하는데 있음.
    ' 향후 index를 수정 해야 할 필요가 있음.
'                         " AND    a.rcvdt " & FUNC_CONCAT & " a.rcvtm <= " & DBS(pRcvDt & pRcvTm) & _

       
       
   ' " AND    " & DBW("a.rcvdt <= ", pRcvDt) & " AND " & DBW("a.rcvtm <= ", pRcvTm)
   
   ' 2009.07.17 양성현 검체정보 추가, Order 순서 변경
   
       GetUnverifiedList = " SELECT/*+ INDEX(a S2LAB201_IDX2) */ " & _
                        " a.ptid,a.spcyy,a.spcno,e." & F_PTNM & " as ptnm,a.workarea, a.accdt, a.accseq,a.wardid,a.deptcd,a.hosilid," & _
                        " a.statfg, a.reqtotcnt, a.reqinputcnt, '0' as TestDiv, c.abbrnm5, b.spccd, g.field3, b.rstcd," & _
                        " a.coldt, a.coltm, a.colid, a.rcvdt, a.rcvtm, a.rcvid, i.wardid as wardid1" & _
                        " FROM " & T_LAB201 & " a, " & T_HIS001 & " e," & T_LAB302 & " b, " & T_LAB001 & " c, " & T_LAB032 & " g, ORAS1.S2ORD101_V I  " & _
                        " WHERE  " & DBW("a.workarea = ", pWorkArea) & _
                        " AND   b.workarea=a.workarea AND b.accdt=a.accdt AND b.accseq=a.accseq " & _
                        " AND    a.ptid=e." & F_PTID & _
                        " AND    b.PTID = i.PTID  AND b.orddt = i.ORDDT  AND b.ORDNO = i.ORDNO " & _
                        " AND    " & DBW("a.rcvdt >= ", pFRcvDt) & _
                        " AND    a.rcvdt <= '" & pRcvDt & "' " & _
                        " AND    " & DBW("a.stscd   >= ", enStsCd.StsCd_LIS_Accession) & _
                        " AND    " & DBW("a.stscd   < ", enStsCd.StsCd_LIS_MidRst) & _
                        " AND    " & DBW("a.buildcd  = ", pBuildingCd) & _
                        " AND  (b.vfydt='' or b.vfydt is null) and g.cdindex (+)= '" & LC3_Specimen & "' and g.cdval1 (+)= b.spccd " & _
                        " AND   c.testcd = b.testcd  AND  (c.detailfg = ''  or  c.detailfg is null) " & _
                        " AND   c.applydt = ( SELECT max(applydt) FROM " & T_LAB001 & _
                                            " WHERE testcd = b.testcd )"
    
   ' 2011.07.07 온승호 미생물/기타 조회 시 검체명 삭제
   ' 테스트 중
    
'    If ObjSysInfo.projectid <> "BBS" Then
'
'        GetUnverifiedList = GetUnverifiedList & " UNION ALL " & _
'                            " SELECT/*+ INDEX(a S2LAB201_IDX2) */ " & _
'                            " a.ptid,a.spcyy,a.spcno,e." & F_PTNM & " as ptnm, a.workarea, a.accdt, a.accseq," & _
'                            " a.wardid,a.deptcd,a.hosilid,a.statfg, a.reqtotcnt, a.reqinputcnt, '1' as TestDiv," & _
'                            " c.abbrnm5, f.spccd, g.field3, '' as rstcd, a.coldt, a.coltm, a.colid, a.rcvdt, a.rcvtm, a.rcvid" & _
'                            " FROM " & T_LAB201 & " a, " & T_HIS001 & " e," & T_LAB351 & " b, " & T_LAB001 & " c " & " , " & T_LAB102 & " f, " & T_LAB032 & " g " & _
'                            " WHERE  " & DBW("a.workarea = ", pWorkArea) & _
'                            " AND   b.workarea=a.workarea AND b.accdt=a.accdt AND b.accseq=a.accseq " & _
'                            " AND   a.ptid=e." & F_PTID & _
'                            " AND    " & DBW("a.rcvdt    >= ", pFRcvDt) & _
'                            " AND    a.rcvdt <= '" & pRcvDt & "' " & _
'                            " AND    " & DBW("a.stscd   >= ", enStsCd.StsCd_LIS_Accession) & _
'                            " AND    " & DBW("a.stscd   < ", enStsCd.StsCd_LIS_MidRst) & _
'                            " AND    " & DBW("a.buildcd  = ", pBuildingCd) & _
'                            " AND  (b.vfydt='' or b.vfydt is null) and g.cdindex (+)= '" & LC3_Specimen & "' and g.cdval1 (+)= f.spccd " & _
'                            " AND   c.testcd = b.testcd  AND  (c.detailfg = ''  or  c.detailfg is null) " & _
'                            " AND   c.applydt = ( SELECT max(applydt) FROM " & T_LAB001 & _
'                                                " WHERE testcd = b.testcd )" & _
'                            " and f.ptid = a.ptid and f.workarea = a.workarea and f.accdt = a.accdt and f.accseq = a.accseq "
'        GetUnverifiedList = GetUnverifiedList & " UNION ALL " & _
'                            " SELECT/*+ INDEX(a S2LAB201_IDX2) */ " & _
'                            " a.ptid,a.spcyy,a.spcno,e." & F_PTNM & " as ptnm, a.workarea, a.accdt, a.accseq," & _
'                            " a.wardid,a.deptcd,a.hosilid,a.statfg, a.reqtotcnt, a.reqinputcnt, '2' as TestDiv, c.abbrnm5, f.spccd, g.field3, b.rstcd," & _
'                            " a.coldt, a.coltm, a.colid, a.rcvdt, a.rcvtm, a.rcvid" & _
'                            " FROM " & T_LAB201 & " a, " & T_HIS001 & " e," & T_LAB404 & " b, " & T_LAB001 & " c " & " , " & T_LAB102 & " f, " & T_LAB032 & " g " & _
'                            " WHERE  " & DBW("a.workarea = ", pWorkArea) & _
'                            " AND   b.workarea=a.workarea AND b.accdt=a.accdt AND b.accseq=a.accseq " & _
'                            " AND   a.ptid=e." & F_PTID & _
'                            " AND    " & DBW("a.rcvdt    >= ", pFRcvDt) & _
'                            " AND    a.rcvdt <= '" & pRcvDt & "' " & _
'                            " AND    " & DBW("a.stscd   >= ", enStsCd.StsCd_LIS_Accession) & _
'                            " AND    " & DBW("a.stscd   < ", enStsCd.StsCd_LIS_MidRst) & _
'                            " AND    " & DBW("a.buildcd  = ", pBuildingCd) & _
'                            " AND  (b.vfydt='' or b.vfydt is null) and g.cdindex (+)= '" & LC3_Specimen & "' and g.cdval1 (+)= f.spccd " & _
'                            " AND   c.testcd = b.testcd  AND  (c.detailfg = ''  or  c.detailfg is null) " & _
'                            " AND   c.applydt = ( SELECT max(applydt) FROM " & T_LAB001 & _
'                                                " WHERE testcd = b.testcd )" & _
'                            " and f.ptid = a.ptid and f.workarea = a.workarea and f.accdt = a.accdt and f.accseq = a.accseq " & _
'                            " order by accdt, accseq "
                            
    If ObjSysInfo.projectid <> "BBS" Then
                                                
        GetUnverifiedList = GetUnverifiedList & " UNION ALL " & _
                            " SELECT/*+ INDEX(a S2LAB201_IDX2) */ " & _
                            " a.ptid,a.spcyy,a.spcno,e." & F_PTNM & " as ptnm, a.workarea, a.accdt, a.accseq," & _
                            " a.wardid,a.deptcd,a.hosilid,a.statfg, a.reqtotcnt, a.reqinputcnt, '1' as TestDiv," & _
                            " c.abbrnm5, '' as spccd, g.field3, '' as rstcd, a.coldt, a.coltm, a.colid, a.rcvdt, a.rcvtm, a.rcvid, '' wardid1" & _
                            " FROM " & T_LAB201 & " a, " & T_HIS001 & " e," & T_LAB351 & " b, " & T_LAB001 & " c " & " , " & T_LAB032 & " g " & _
                            " WHERE  " & DBW("a.workarea = ", pWorkArea) & _
                            " AND   b.workarea=a.workarea AND b.accdt=a.accdt AND b.accseq=a.accseq " & _
                            " AND   a.ptid=e." & F_PTID & _
                            " AND    " & DBW("a.rcvdt    >= ", pFRcvDt) & _
                            " AND    a.rcvdt <= '" & pRcvDt & "' " & _
                            " AND    " & DBW("a.stscd   >= ", enStsCd.StsCd_LIS_Accession) & _
                            " AND    " & DBW("a.stscd   < ", enStsCd.StsCd_LIS_MidRst) & _
                            " AND    " & DBW("a.buildcd  = ", pBuildingCd) & _
                            " AND  (b.vfydt='' or b.vfydt is null) and g.cdindex (+)= '" & LC3_Specimen & "' and g.cdval1 (+)= a.workarea " & _
                            " AND   c.testcd = b.testcd  AND  (c.detailfg = ''  or  c.detailfg is null) " & _
                            " AND   c.applydt = ( SELECT max(applydt) FROM " & T_LAB001 & _
                                                " WHERE testcd = b.testcd )"

        GetUnverifiedList = GetUnverifiedList & " UNION ALL " & _
                            " SELECT/*+ INDEX(a S2LAB201_IDX2) */ " & _
                            " a.ptid,a.spcyy,a.spcno,e." & F_PTNM & " as ptnm, a.workarea, a.accdt, a.accseq," & _
                            " a.wardid,a.deptcd,a.hosilid,a.statfg, a.reqtotcnt, a.reqinputcnt, '2' as TestDiv, c.abbrnm5, '' as spccd, g.field3, b.rstcd," & _
                            " a.coldt, a.coltm, a.colid, a.rcvdt, a.rcvtm, a.rcvid, '' wardid1" & _
                            " FROM " & T_LAB201 & " a, " & T_HIS001 & " e," & T_LAB404 & " b, " & T_LAB001 & " c " & " , " & T_LAB032 & " g " & _
                            " WHERE  " & DBW("a.workarea = ", pWorkArea) & _
                            " AND   b.workarea=a.workarea AND b.accdt=a.accdt AND b.accseq=a.accseq " & _
                            " AND   a.ptid=e." & F_PTID & _
                            " AND    " & DBW("a.rcvdt    >= ", pFRcvDt) & _
                            " AND    a.rcvdt <= '" & pRcvDt & "' " & _
                            " AND    " & DBW("a.stscd   >= ", enStsCd.StsCd_LIS_Accession) & _
                            " AND    " & DBW("a.stscd   < ", enStsCd.StsCd_LIS_MidRst) & _
                            " AND    " & DBW("a.buildcd  = ", pBuildingCd) & _
                            " AND  (b.vfydt='' or b.vfydt is null) and g.cdindex (+)= '" & LC3_Specimen & "' and g.cdval1 (+)= a.workarea " & _
                            " AND   c.testcd = b.testcd  AND  (c.detailfg = ''  or  c.detailfg is null) " & _
                            " AND   c.applydt = ( SELECT max(applydt) FROM " & T_LAB001 & _
                                                " WHERE testcd = b.testcd )" & _
                            " order by accdt, accseq "
                            
    End If
    Debug.Print GetUnverifiedList
End Function

Public Function GetUnverifiedList_WM(ByVal pWorkArea As String, ByVal pRcvDt As String, _
                                  ByVal pBuildingCd As String) As String
    Dim pRcvTm  As String
    Dim pFRcvDt As String
    
    'to
    pRcvTm = Mid(pRcvDt, 9, 6)
    'FROM
    pFRcvDt = Mid(pRcvDt, 15)
    
    pRcvDt = Mid(pRcvDt, 1, 8)
    
       GetUnverifiedList_WM = " SELECT/*+ INDEX(a S2LAB201_IDX2) */ " & _
                        " a.ptid,a.spcyy,a.spcno,e." & F_PTNM & " as ptnm,a.workarea, a.accdt, a.accseq,a.wardid,a.deptcd,a.hosilid," & _
                        " a.statfg, a.reqtotcnt, a.reqinputcnt, '0' as TestDiv, c.abbrnm5, b.spccd, g.field3, b.rstcd," & _
                        " a.coldt, a.coltm, a.colid, a.rcvdt, a.rcvtm, a.rcvid" & _
                        " FROM " & T_LAB201 & " a, " & T_HIS001 & " e," & T_LAB302 & " b, " & T_LAB001 & " c, " & T_LAB032 & " g " & _
                        " WHERE  " & DBW("a.workarea = ", pWorkArea) & _
                        " AND   b.workarea=a.workarea AND b.accdt=a.accdt AND b.accseq=a.accseq " & _
                        " AND    a.ptid=e." & F_PTID & _
                        " AND    " & DBW("a.rcvdt >= ", pFRcvDt) & _
                        " AND    a.rcvdt <= '" & pRcvDt & "' " & _
                        " AND    " & DBW("a.stscd   >= ", enStsCd.StsCd_LIS_Accession) & _
                        " AND    " & DBW("a.stscd   < ", enStsCd.StsCd_LIS_MidRst) & _
                        " AND    " & DBW("a.buildcd  = ", pBuildingCd) & _
                        " AND  (b.vfydt='' or b.vfydt is null) and g.cdindex (+)= '" & LC3_Specimen & "' and g.cdval1 (+)= b.spccd " & _
                        " AND   c.testcd = b.testcd  AND  (c.detailfg = ''  or  c.detailfg is null) " & _
                        " AND   c.applydt = ( SELECT max(applydt) FROM " & T_LAB001 & _
                                            " WHERE testcd = b.testcd )" & _
                        " AND  b.testcd in ('C3710-1','C3710','C37101','C37102','C3710B','C3710C','C3710D','C3710E')                                             "
                                
    If ObjSysInfo.projectid <> "BBS" Then
                                                
        GetUnverifiedList_WM = GetUnverifiedList_WM & " UNION ALL " & _
                            " SELECT/*+ INDEX(a S2LAB201_IDX2) */ " & _
                            " a.ptid,a.spcyy,a.spcno,e." & F_PTNM & " as ptnm, a.workarea, a.accdt, a.accseq," & _
                            " a.wardid,a.deptcd,a.hosilid,a.statfg, a.reqtotcnt, a.reqinputcnt, '1' as TestDiv," & _
                            " c.abbrnm5, '' as spccd, g.field3, '' as rstcd, a.coldt, a.coltm, a.colid, a.rcvdt, a.rcvtm, a.rcvid" & _
                            " FROM " & T_LAB201 & " a, " & T_HIS001 & " e," & T_LAB351 & " b, " & T_LAB001 & " c " & " , " & T_LAB032 & " g " & _
                            " WHERE  " & DBW("a.workarea = ", pWorkArea) & _
                            " AND   b.workarea=a.workarea AND b.accdt=a.accdt AND b.accseq=a.accseq " & _
                            " AND   a.ptid=e." & F_PTID & _
                            " AND    " & DBW("a.rcvdt    >= ", pFRcvDt) & _
                            " AND    a.rcvdt <= '" & pRcvDt & "' " & _
                            " AND    " & DBW("a.stscd   >= ", enStsCd.StsCd_LIS_Accession) & _
                            " AND    " & DBW("a.stscd   < ", enStsCd.StsCd_LIS_MidRst) & _
                            " AND    " & DBW("a.buildcd  = ", pBuildingCd) & _
                            " AND  (b.vfydt='' or b.vfydt is null) and g.cdindex (+)= '" & LC3_Specimen & "' and g.cdval1 (+)= a.workarea " & _
                            " AND   c.testcd = b.testcd  AND  (c.detailfg = ''  or  c.detailfg is null) " & _
                            " AND   c.applydt = ( SELECT max(applydt) FROM " & T_LAB001 & _
                                                " WHERE testcd = b.testcd )"

        GetUnverifiedList_WM = GetUnverifiedList_WM & " UNION ALL " & _
                            " SELECT/*+ INDEX(a S2LAB201_IDX2) */ " & _
                            " a.ptid,a.spcyy,a.spcno,e." & F_PTNM & " as ptnm, a.workarea, a.accdt, a.accseq," & _
                            " a.wardid,a.deptcd,a.hosilid,a.statfg, a.reqtotcnt, a.reqinputcnt, '2' as TestDiv, c.abbrnm5, '' as spccd, g.field3, b.rstcd," & _
                            " a.coldt, a.coltm, a.colid, a.rcvdt, a.rcvtm, a.rcvid" & _
                            " FROM " & T_LAB201 & " a, " & T_HIS001 & " e," & T_LAB404 & " b, " & T_LAB001 & " c " & " , " & T_LAB032 & " g " & _
                            " WHERE  " & DBW("a.workarea = ", pWorkArea) & _
                            " AND   b.workarea=a.workarea AND b.accdt=a.accdt AND b.accseq=a.accseq " & _
                            " AND   a.ptid=e." & F_PTID & _
                            " AND    " & DBW("a.rcvdt    >= ", pFRcvDt) & _
                            " AND    a.rcvdt <= '" & pRcvDt & "' " & _
                            " AND    " & DBW("a.stscd   >= ", enStsCd.StsCd_LIS_Accession) & _
                            " AND    " & DBW("a.stscd   < ", enStsCd.StsCd_LIS_MidRst) & _
                            " AND    " & DBW("a.buildcd  = ", pBuildingCd) & _
                            " AND  (b.vfydt='' or b.vfydt is null) and g.cdindex (+)= '" & LC3_Specimen & "' and g.cdval1 (+)= a.workarea " & _
                            " AND   c.testcd = b.testcd  AND  (c.detailfg = ''  or  c.detailfg is null) " & _
                            " AND   c.applydt = ( SELECT max(applydt) FROM " & T_LAB001 & _
                                                " WHERE testcd = b.testcd )" & _
                            " order by accdt, accseq "
                            
    End If
    Debug.Print GetUnverifiedList_WM
End Function

Public Function GetUnverifiedList_WM1(ByVal pWorkArea As String, ByVal pRcvDt As String, _
                                  ByVal pBuildingCd As String, ByVal pWardId As String) As String
    Dim pRcvTm  As String
    Dim pFRcvDt As String
    Dim strDeptCd As String
    
    'to
    pRcvTm = Mid(pRcvDt, 9, 6)
    'FROM
    pFRcvDt = Mid(pRcvDt, 15)
    
    pRcvDt = Mid(pRcvDt, 1, 8)
    
    If pWardId = "ER" Then
        strDeptCd = "i.deptcd"
        pWardId = "EM"
    Else
        strDeptCd = "i.wardid"
    End If
    
       GetUnverifiedList_WM1 = " SELECT/*+ INDEX(a S2LAB201_IDX2) */ " & _
                        " a.ptid,a.spcyy,a.spcno,e." & F_PTNM & " as ptnm,a.workarea, a.accdt, a.accseq,a.wardid,a.deptcd,a.hosilid," & _
                        " a.statfg, a.reqtotcnt, a.reqinputcnt, '0' as TestDiv, c.abbrnm5, b.spccd, g.field3, b.rstcd," & _
                        " a.coldt, a.coltm, a.colid, a.rcvdt, a.rcvtm, a.rcvid, i.wardid" & _
                        " FROM " & T_LAB201 & " a, " & T_HIS001 & " e," & T_LAB302 & " b, " & T_LAB001 & " c, " & T_LAB032 & " g, ORAS1.S2ORD101_V I " & _
                        " WHERE  " & DBW("a.workarea = ", pWorkArea) & _
                        " AND   b.workarea=a.workarea AND b.accdt=a.accdt AND b.accseq=a.accseq " & _
                        " AND    a.ptid=e." & F_PTID & _
                        " AND   b.PTID = i.PTID  AND b.orddt = i.ORDDT  AND b.ORDNO = i.ORDNO " & _
                        " AND    " & DBW("a.rcvdt >= ", pFRcvDt) & _
                        " AND    a.rcvdt <= '" & pRcvDt & "' " & _
                        " AND    " & DBW("a.stscd   >= ", enStsCd.StsCd_LIS_Accession) & _
                        " AND    " & DBW("a.stscd   < ", enStsCd.StsCd_LIS_MidRst) & _
                        " AND    " & DBW("a.buildcd  = ", pBuildingCd) & _
                        " AND  (b.vfydt='' or b.vfydt is null) and g.cdindex (+)= '" & LC3_Specimen & "' and g.cdval1 (+)= b.spccd " & _
                        " AND   c.testcd = b.testcd  AND  (c.detailfg = ''  or  c.detailfg is null) " & _
                        " AND   c.applydt = ( SELECT max(applydt) FROM " & T_LAB001 & _
                                            " WHERE testcd = b.testcd )" & _
                        " AND  b.testcd in ('C3710-1','C3710','C37101','C37102','C3710B','C3710C','C3710D','C3710E','C3710A') "
                        
                        If pWardId <> "전 체" Then
'                            GetUnverifiedList_WM1 = GetUnverifiedList_WM1 & " AND  i.wardid = '" & pWardId & "'         "
                            GetUnverifiedList_WM1 = GetUnverifiedList_WM1 & " AND  " & strDeptCd & " = '" & pWardId & "'         "

                        End If
                        
    If ObjSysInfo.projectid <> "BBS" Then
                                                
        GetUnverifiedList_WM1 = GetUnverifiedList_WM1 & " UNION ALL " & _
                            " SELECT/*+ INDEX(a S2LAB201_IDX2) */ " & _
                            " a.ptid,a.spcyy,a.spcno,e." & F_PTNM & " as ptnm, a.workarea, a.accdt, a.accseq," & _
                            " a.wardid,a.deptcd,a.hosilid,a.statfg, a.reqtotcnt, a.reqinputcnt, '1' as TestDiv," & _
                            " c.abbrnm5, '' as spccd, g.field3, '' as rstcd, a.coldt, a.coltm, a.colid, a.rcvdt, a.rcvtm, a.rcvid,'' wardid" & _
                            " FROM " & T_LAB201 & " a, " & T_HIS001 & " e," & T_LAB351 & " b, " & T_LAB001 & " c " & " , " & T_LAB032 & " g " & _
                            " WHERE  " & DBW("a.workarea = ", pWorkArea) & _
                            " AND   b.workarea=a.workarea AND b.accdt=a.accdt AND b.accseq=a.accseq " & _
                            " AND   a.ptid=e." & F_PTID & _
                            " AND    " & DBW("a.rcvdt    >= ", pFRcvDt) & _
                            " AND    a.rcvdt <= '" & pRcvDt & "' " & _
                            " AND    " & DBW("a.stscd   >= ", enStsCd.StsCd_LIS_Accession) & _
                            " AND    " & DBW("a.stscd   < ", enStsCd.StsCd_LIS_MidRst) & _
                            " AND    " & DBW("a.buildcd  = ", pBuildingCd) & _
                            " AND  (b.vfydt='' or b.vfydt is null) and g.cdindex (+)= '" & LC3_Specimen & "' and g.cdval1 (+)= a.workarea " & _
                            " AND   c.testcd = b.testcd  AND  (c.detailfg = ''  or  c.detailfg is null) " & _
                            " AND   c.applydt = ( SELECT max(applydt) FROM " & T_LAB001 & _
                                                " WHERE testcd = b.testcd )"

        GetUnverifiedList_WM1 = GetUnverifiedList_WM1 & " UNION ALL " & _
                            " SELECT/*+ INDEX(a S2LAB201_IDX2) */ " & _
                            " a.ptid,a.spcyy,a.spcno,e." & F_PTNM & " as ptnm, a.workarea, a.accdt, a.accseq," & _
                            " a.wardid,a.deptcd,a.hosilid,a.statfg, a.reqtotcnt, a.reqinputcnt, '2' as TestDiv, c.abbrnm5, '' as spccd, g.field3, b.rstcd," & _
                            " a.coldt, a.coltm, a.colid, a.rcvdt, a.rcvtm, a.rcvid,'' wardid" & _
                            " FROM " & T_LAB201 & " a, " & T_HIS001 & " e," & T_LAB404 & " b, " & T_LAB001 & " c " & " , " & T_LAB032 & " g " & _
                            " WHERE  " & DBW("a.workarea = ", pWorkArea) & _
                            " AND   b.workarea=a.workarea AND b.accdt=a.accdt AND b.accseq=a.accseq " & _
                            " AND   a.ptid=e." & F_PTID & _
                            " AND    " & DBW("a.rcvdt    >= ", pFRcvDt) & _
                            " AND    a.rcvdt <= '" & pRcvDt & "' " & _
                            " AND    " & DBW("a.stscd   >= ", enStsCd.StsCd_LIS_Accession) & _
                            " AND    " & DBW("a.stscd   < ", enStsCd.StsCd_LIS_MidRst) & _
                            " AND    " & DBW("a.buildcd  = ", pBuildingCd) & _
                            " AND  (b.vfydt='' or b.vfydt is null) and g.cdindex (+)= '" & LC3_Specimen & "' and g.cdval1 (+)= a.workarea " & _
                            " AND   c.testcd = b.testcd  AND  (c.detailfg = ''  or  c.detailfg is null) " & _
                            " AND   c.applydt = ( SELECT max(applydt) FROM " & T_LAB001 & _
                                                " WHERE testcd = b.testcd )" & _
                            " order by accdt, accseq "
                            
    End If
    Debug.Print GetUnverifiedList_WM1
End Function

Public Function GetWorkArea() As String

    GetWorkArea = " SELECT cdval1, field1 " & _
                  " FROM   " & T_LAB032 & _
                  " WHERE  " & DBW("cdindex =", lc3_workarea)
                  
    GetWorkArea = GetWorkArea & " ORDER BY TEXT1 " '2014-09-29 PSK 정렬처리 미사용필드 정렬처리...
    Debug.Print GetWorkArea
End Function

Public Function GetWorksheetCode(ByVal pBuildingCd As String) As String

    GetWorksheetCode = " SELECT a.cdval1 as WorkCd, a.field1 as WorkNm, b.workarea, count(b.testcd) as TestCnt " & _
                       " FROM   " & T_LAB032 & " a, " & T_LAB008 & " b " & _
                       " WHERE  " & DBW("a.cdindex = ", LC3_WorkSheetName) & _
                       " AND    b.workcd = a.cdval1 " & _
                       " AND    " & DBW("a.field2  = ", pBuildingCd) & _
                       " group by a.cdval1, a.field1, b.workarea "

End Function
    

Public Function GetLabNo(ByVal pSpcYY As String, ByVal pSpcNo As String) As String
   
    GetLabNo = " SELECT workarea, " & FUNC_SUBSTR & "(accdt,3,6) as accdt, accseq " & _
               " FROM " & T_LAB201 & _
               " WHERE  " & DBW("spcyy =", pSpcYY) & _
               " AND    " & DBW("spcno =", pSpcNo)

End Function

Public Function GetEqpList(ByVal pBuildingCd As String) As String

    GetEqpList = " SELECT eqpcd, eqpnm FROM   " & T_LAB006 & _
                 " WHERE  " & DBW("location = ", pBuildingCd) & _
                 " order by eqpcd"

End Function


Public Function GetMaxSeq(ByVal pWorkCd As String, ByRef pWorkDt As String) As String
    
    GetMaxSeq = " SELECT max(workseq) as MaxSeq FROM " & T_LAB301 & _
                " WHERE  " & DBW("workcd = ", pWorkCd) & _
                " AND    " & DBW("workdt = ", pWorkDt)
             
End Function


Public Function DeleteWorkSheet(ByVal WorkDt As String, _
                                ByVal WorkCd As String, _
                                ByVal DeleteList As String) As Boolean

    Dim SqlStmt() As String
    Dim aryTmp() As String
    Dim ii As Long
    Dim intSqlCnt As Integer
   '
    'DBOpenDyna strSql
    aryTmp = Split(DeleteList, vbTab)
    ReDim SqlStmt(UBound(aryTmp) - 1)
    '/*워크쉬트내역 생성
    On Error GoTo DBExecError
   '
    For ii = 0 To UBound(aryTmp) - 1
        SqlStmt(ii) = " DELETE " & T_LAB301 & _
                      " WHERE  " & DBW("workdt = ", WorkDt) & _
                      " AND    " & DBW("workcd = ", WorkCd) & _
                      " AND    " & DBW("workseq = ", aryTmp(ii))
    Next ii

    DBErrClear
    '
    'DBMultiExec SqlStmt
    DBConn.BeginTrans
    For intSqlCnt = LBound(SqlStmt) To UBound(SqlStmt)
        DBConn.Execute (SqlStmt(intSqlCnt))
    Next intSqlCnt
    DBConn.CommitTrans
    '
    DeleteWorkSheet = True
    Exit Function
DBExecError:
    DeleteWorkSheet = False

End Function


Public Function GetWorksheetTestCd(ByVal pWorkCd As String, ByVal pCase As String) As String

    If pCase = "1" Then 'Worksheet
        GetWorksheetTestCd = " SELECT distinct a.testcd, b.testnm " & _
                             " FROM " & T_LAB008 & " a, " & T_LAB001 & " b " & _
                             " WHERE  " & DBW("a.workcd = ", pWorkCd) & _
                             " AND    b.testcd = a.testcd " & _
                             " AND    b.applydt = ( SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = b.testcd ) " & _
                             " order  by a.testcd"
    Else    '접수번호
        GetWorksheetTestCd = " SELECT distinct a.testcd, a.testnm " & _
                             " FROM " & T_LAB001 & " a " & _
                             " WHERE  " & DBW("a.workarea = ", pWorkCd) & _
                             " AND    ( a.panelfg='' or a.panelfg is null ) " & _
                             " AND    ( a.detailfg='' or a.detailfg is null ) " & _
                             " AND    a.applydt = ( SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = a.testcd ) " & _
                             " order  by a.testcd"
    End If
    
End Function

Public Function WarnMsgChk(ByVal pTestCd As String, ByVal pRstCd As String) As String
    Dim strSQL As String
    
    strSQL = " SELECT text1 FROM " & T_LAB031 & _
             " WHERE  cdinex = " & DBS(LC2_Warning) & _
             "   AND  cdval1 = " & DBS(pTestCd) & _
             "   AND  cdval2 = " & DBS(pRstCd)
    
End Function

