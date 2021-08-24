Attribute VB_Name = "Module4"
Option Explicit
  
    Global RstTmp   As String
    Global PreRst   As String   '이전 결과값
    
    Global ChkInsert    As Integer  'SE-9000



'
'   기존의 결과가 존재할 때 새로운 결과로 수정할 것인지 확인
'   (TestNm : 수정을 원하는 검사항목명)
'
Public Function Chk_Update(TestNm As String, GData() As String)
'
'    Dim Msg As String
'    Dim Style   As String
'    Dim Response    As Integer
'
'    Chk_Update = True
'    Chk_Exist = False
'
'    With Insert_Server(iResCnt)
'        '기존 결과가 존재하는지 체크
'        If Trim$(.Result) = "" Or Trim$(.LstRes) = "" Then
'            Exit Function
'        Else
'            Chk_Exist = True        '기등록된 결과
'        End If
'        '기존 결과와 새로운 결과가 동일한지 검사
'        If Trim$(.Result) = Trim$(.LstRes) Then
'            Chk_Update = False
'            Exit Function
'        End If
'
'        Msg = "Lab ID : " & Right$(GData(1), 6) & "-" & GData(2) & "-" & GData(3) & _
'              " 의 검사항목 " & "[" & Trim$(TestNm) & "]" & _
'              " 의 결과값이 이미존재합니다." & Chr(13) & Chr(13) & Chr(13) & Chr(10) & _
'              "이전 결과값     " & Trim$(.LstRes) & " 을" & _
'              Chr(13) & Chr(13) & Chr(10) & "새로운 결과값  " & Trim$(.Result) & _
'              " 로 수정하시겠습니까?"
'        Style = 4 + 48
'
'        Response = MsgBox(Msg, Style, "결과수정확인")
'        If Response = vbNo Then
'            Chk_Update = False
'        End If
'    End With
    
End Function


'
'   Temp DB 에서 RstDate조회
'   (para1 : LabDate,  para2 : SlipCd,  para3 : LabSqNo,  para4 : OrdCd + SubSqNo)
'
Public Function Get_RstDate(para1 As String, para2 As String, para3 As String, para4 As String) As String
    
    Dim sStr    As String
    Dim tStr    As String
    Dim rData() As String
    
    Get_RstDate = ""
    PreRst = ""
    
    tStr = "      Select RSTDATE, SYSTIME, RSTVAL1 "
    tStr = tStr & " from #TEMPRESULT "
    tStr = tStr & "where LABDATE = '" & para1 & "'"
    tStr = tStr & "  and SLIPCD  = '" & para2 & "'"
    tStr = tStr & "  and LABSQNO = '" & para3 & "'"
    tStr = tStr & "  and ORDCD = '" & Left$(para4, 8) & "'"
    tStr = tStr & "  and SUBSQNO = '" & Trim$(Mid(para4, 9, 2)) & "'"

    ret = QSqlDBExec(tStr, QsqlCon1%)
    If ret = QSQL_SUCCESS Then
        ret = QSqlGetRow(sStr, QsqlCon1%)
        If ret = QSQL_SUCCESS Then
            QSqlGetField 3, sStr, rData()
            
            Get_RstDate = rData(1) & rData(2)
            PreRst = rData(3)   '이전 결과값
        End If
    End If
    
    ret = QSqlSelectFree(QsqlCon1%)
    
End Function


Public Function DROP_TEMP_TABLE() As Integer

    DROP_TEMP_TABLE = Qsqlclose(QsqlCon1%, ONECLOSE)

End Function

Public Function Get_RstDate_Batch(para1 As String, para2 As String, para3 As String, para4 As String) As String
    
    Dim sStr    As String
    Dim tStr    As String
    Dim rData() As String
    
    Get_RstDate_Batch = ""
    PreRst = ""
    
    tStr = "      Select RSTDATE, SYSTIME "
    tStr = tStr & " from #TEMPRESULT "
    tStr = tStr & "where LABDATE = '" & para1 & "'"
    tStr = tStr & "  and SLIPCD  = '" & para2 & "'"
    tStr = tStr & "  and LABSQNO = '" & para3 & "'"
    tStr = tStr & "  and ORDCD = '" & Left$(para4, 8) & "'"
    tStr = tStr & "  and SUBSQNO = '" & Trim$(Mid(para4, 9, 2)) & "'"

    ret = QSqlDBExec(tStr, QsqlCon1%)
    If ret = QSQL_SUCCESS Then
        ret = QSqlGetRow(sStr, QsqlCon1%)
        If ret = QSQL_SUCCESS Then
            QSqlGetField 2, sStr, rData()
            
            Get_RstDate_Batch = rData(1) & rData(2)
        End If
    End If
    
    ret = QSqlSelectFree(QsqlCon1%)
    
End Function

'-----------------------------
'   Local DB의 결과값을 Server에 등록
'   para      0 : CstIdNo
'               1 : LabDate
'               2 : SlipCd
'               3 : LabSqNo
'               4 : OrdCd (10자리 : OrdCd(8) + SubSqNo(2))  yk
'               5 : RstVal1
'               6 : RtnCd
'               7 : RecLabNo
'               8 : Age
'               9 : LabTime
'              10 : OrdId
'              11 : DeltaChk
'              12 : RecChk      접수구분
'              13 : PanicChk
'              14 : OrdStat     응급구분
'              15 : PanjChk     참고치 체크(Blast,LUC 등)
'   NextPara
'               0 : DeltaChk
'               1 : LabDate
'               2 : SlipCd
'               3 : LabSqNo
'               4 : OrdCd
'----------------------
Public Function Insert_Result(sw As Integer, para() As String, NextPara() As String, RstDate As String) As Integer
   
    '----- Sub 항목 추가
    Dim SubSqNo     As String       'Sub 항목순번
    Dim SubSqNo2    As String       '      "     (NextPara())
    SubSqNo = Trim$(Mid(para(4), 9, 2))
    SubSqNo2 = Trim$(Mid(NextPara(4), 9, 2))
    '-------------------
    
    Insert_Result = True
    
    RstDate = Format$(Now, "YYYYMMDDHHMMDD")
    If sw = True Then
            
        SqlStr = "UPDATE LAB01_DB..SLC010M SET " _
                & "RSTVAL1  = '" & para(5) & "'," _
                & "DELTACHK = '" & para(11) & "'," _
                & "PANICCHK = '" & para(13) & "'," _
                & "PANJCHK = '" & para(15) & "'," _
                & "ORDID = '" & para(10) & "'" _
                & " WHERE LABDATE = '" & para(1) & "'" _
                & " AND SLIPCD = '" & para(2) & "'" _
                & " AND LABSQNO = '" & para(3) & "'" _
                & " AND ORDCD = '" & Trim$(Left(para(4), 8)) & "'" _
                & " AND SUBSQNO = '" & SubSqNo & "'"    '8/4추가 yk
    
    Else
    
        SqlStr = "INSERT INTO LAB01_DB..SLC010M ( " _
                & "LABDATE, SLIPCD,  LABSQNO, ORDCD,  SUBSQNO, " _
                & "RSTDATE, RSTVAL1, RSTVAL2, RSTETC, DELTACHK,PANICCHK, " _
                & "PANJCHK, ORDSTAT, " _
                & "RTNCD,   RECLABNO,AGE,ORDID, CFMID, " _
                & "LABTIME, CSTIDNO, SYSDATE, SYSTIME) VALUES ( " _
                & "'" & para(1) & "', " _
                & "'" & para(2) & "', " _
                & "'" & para(3) & "', " _
                & "'" & Trim$(Left(para(4), 8)) & "', " _
                & "'" & SubSqNo & "', " _
                & "'" & Left(RstDate, 8) & "', " _
                & "'" & para(5) & "', " _
                & " 0, " _
                & "'', " _
                & "'" & para(11) & "', " _
                & "'" & para(13) & "', "
        SqlStr = SqlStr _
                & "'" & para(15) & "', " _
                & "'" & para(14) & "', " _
                & "'" & para(6) & "', " _
                & "'" & para(7) & "', " _
                & "'" & para(8) & "', " _
                & "'" & para(10) & "', " _
                & "'', " _
                & "'" & para(9) & "', " _
                & "'" & para(0) & "', " _
                & "'" & Left(RstDate, 8) & "', " _
                & "'" & Right(RstDate, 6) & "') " _

    End If

    If QSqlDBExec(SqlStr, QsqlConn%) <> QSQL_SUCCESS Then
        Insert_Result = False
        GoTo Insert_Result_End
    End If
    
        
    If sw = True Then
    
        '-----------
        '   다음결과 DeltaCheck
        '----------
        If NextPara(1) <> "" Then
            SqlStr = "UPDATE LAB01_DB..SLC010M SET " _
                    & " DELTACHK = '" & NextPara(0) & Chr$(39) _
                    & " WHERE LABDATE  = '" & NextPara(1) & Chr$(39) _
                    & " AND SLIPCD = '" & NextPara(2) & Chr$(39) _
                    & " AND LABSQNO = '" & NextPara(3) & Chr$(39) _
                    & " AND ORDCD = '" & Left$(NextPara(4), 8) & Chr$(39) _
                    & " AND SUBSQNO = '" & SubSqNo2 & Chr$(39) _
                    
            If QSqlDBExec(SqlStr, QsqlConn%) <> QSQL_SUCCESS Then
                Insert_Result = False
                GoTo Insert_Result_End
            End If
        End If
    Else
    
        '----------------------
        '   결과구분 Check
        '   table : LAB01_DB..SLB020M
        '   key   : LABDATE, SLIPCD, LABSQNO, ORDCD, (SUBSQNO)
        '   결과등록하면 RSTCHK  field에 'Y' setting
        '-----------------------
                                
        SqlStr = "UPDATE LAB01_DB..SLB020M SET " _
                    & " RSTCHK = 'Y' " _
                    & " WHERE LABDATE = '" & para(1) & Chr$(39) _
                    & " AND SLIPCD = '" & para(2) & Chr$(39) _
                    & " AND LABSQNO = '" & para(3) & Chr$(39) _
                    & " AND ORDCD = '" & Left$(para(4), 8) & Chr$(39)
                                        
        If QSqlDBExec(SqlStr, QsqlConn%) <> QSQL_SUCCESS Then
           ' ** Comment 97.08.19 ** Insert_Result = False
            GoTo Insert_Result_End
        End If
    
        '-----------------
        '   환자별 임상병리검사 내역에 결과 '03'
        '   table : WD01A_DB..WD1A050M_TBL
        '   key : RcptYmd (접수일자), RcptNo(RECLABNO)
        '-----------------
        
        If para(12) = "21" Then
            SqlStr = "UPDATE WD01A_DB..WD1A050M_TBL SET" _
                        & " OrdComm = '03'" _
                        & " WHERE RcptYmd = '" & Left(para(7), 8) & Chr$(39) _
                        & " AND RcptNo = '" & Right(para(7), 5) & Chr$(39)

            If QSqlDBExec(SqlStr, QsqlConn%) <> QSQL_SUCCESS Then Insert_Result = False
        End If
    End If

Insert_Result_End:
    'YK (8/18)
    'If Insert_Result = False Then MsgBox "접수번호 " & Right(para(1), 6) & "-" & para(2) & "-" & para(3) & "의 결과입력에 오류가 있습니다.", MB_ICONEXCLAMATION

End Function

Public Function Get_Pre_Result(CstIDNo As String) As Integer

'    Dim SData() As String
'    Dim sStr    As String
'
'    With Pre_Res
'        .CstIDNo = ""
'        .LabDate = ""
'        .LabSqNo = ""
'        .OrdCd = ""
'        .RstDate = ""
'        .RstVal = ""
'        .SlipCd = ""
'        .SysTime = ""
'        .SubSqNo = ""
'
'        Get_Pre_Result = False
'
'        SqlStr = "    Select LABDATE, SLIPCD,  LABSQNO, ORDCD, SUBSQNO " _
'                    & "      RSTDATE, RSTVAL1, SYSTIME, CSTIDNO " _
'                    & " from LAB01_DB..SLC010M " _
'                    & "where CSTIDNO = '" & CstIDNo & "'"
'
'        If QSqlDBExec(SqlStr, QsqlConn) = QSQL_SUCCESS Then
'            If QSqlGetRow(sStr, QsqlConn) = QSQL_SUCCESS Then
'                QSqlGetField 9, sStr, SData()
'
'                .LabDate = SData(1)
'                .SlipCd = SData(2)
'                .LabSqNo = SData(3)
'                .OrdCd = SData(4)
'                .SubSqNo = SData(5)
'                .RstDate = SData(6)
'                .RstVal = SData(7)
'                .SysTime = SData(8)
'                .CstIDNo = SData(9)
'
'                Get_Pre_Result = True
'            End If
'        End If
'        QSqlSelectFree (QsqlConn)
'    End With
    
End Function
'
'   해당 레코드의 존재 여부 검사
'
Function RecordExist(Tb As Recordset, IndexName As String, para As String) As Integer

    Dim CurrRecord As Variant
    
    If Tb.RecordCount < 1 Or Tb.BOF Or Tb.EOF Then
        RecordExist = False
        Exit Function
    End If
    
    CurrRecord = Tb.Bookmark
    Tb.MoveFirst
    Tb.Index = IndexName
    Tb.Seek "=", para
    
    If Tb.NoMatch Then
        Tb.Bookmark = CurrRecord
        RecordExist = False
    Else
        RecordExist = True
    End If

End Function
'
'   해당 레코드의 존재여부 검사(조건 추가)
'
Function RecordExists(Tb As Recordset, IndexName As String, slip As String, tcode As String) As Integer
    
    Dim CurrRecord As Variant
    
    If Tb.RecordCount < 1 Or Tb.BOF Or Tb.EOF Then
        RecordExists = False
        Exit Function
    End If
    
    CurrRecord = Tb.Bookmark
    Tb.MoveFirst
    Tb.Index = IndexName
    Tb.Seek "=", slip, tcode
    
    If Tb.NoMatch Then
        Tb.Bookmark = CurrRecord
        RecordExists = False
    Else
        RecordExists = True
    End If

End Function

'
'   Delta/Panic Check하여 구조체에 저장 ==> Stored Procedure 호출
'   (GetData(1), GetData(2), GetData(3), ItemCd, Result)
'
Public Sub SetServer_DeltaChk(sDate As String, sSlip As String, sSqNo As String, sCstIDNo As String, sOrdCd As String, sResult As String, sCurDate As String, Optional sPanj As String)

'    Dim SqlDoc  As String
'    Dim SubSqNo As String
'    Dim iRes    As Integer
'    Dim return_cd   As Integer
'    Dim sStr    As String
'    Dim tData() As String
'
'    '--- 검사항목/SUB 검사항목 코드 편집
'    If Trim(Mid(sOrdCd, 9, 2)) <> "" Then
'        SubSqNo = Trim(Mid(sOrdCd, 9, 2))
'    Else
'        SubSqNo = ""
'    End If
'    sOrdCd = Left(sOrdCd, 8)
'
'    '--- Delta Check
'    SqlDoc = " Select Order_Register( '" _
'            & sDate & "','" _
'            & sSlip & "','" _
'            & sSqNo & "','" _
'            & sCstIDNo & "','" _
'            & sOrdCd & "','" _
'            & ssubsqno & "','" _
'            & sResult & "','" _
'            & sCurDate & "') from LAB01_DB "
'
'    If QSqlDBExec(QsqlConn) <> QSQL_SUCCESS Then
'        return_cd = QSqlSelectFree(QsqlConn)
'        Exit Sub
'    End If
'
'    If QSqlGetRow(sStr, QsqlConn) = QSQL_SUCCESS Then
'        QSqlGetField 3, sStr, tData()
'
'        If tData(3) = "Y" Then
'            '--- Delta/panic/이전 결과값 저장
'            With Insert_Server(iResCnt)
'                .DeltaChk = tData(1)
'                .LstRes = tData(2)
'                .OrdCd = sOrdCd
'                .SubNo = SubSqNo
'                .Result = sResult
'                .PanicChk = Panic_Check(sResult, sOrdCd)    'Panic Check
'                .PanjChk = sPanj                ''B','W'값 저장(혈액학)
'            End With
'            iResCnt = iResCnt + 1
'        End If
'    End If
'
'    return_cd = QSqlSelectFree(QsqlConn)
        
End Sub

