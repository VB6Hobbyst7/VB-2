Attribute VB_Name = "DB_Query"
Option Explicit



'결과 저장시 G/M/B 코드 비교
Public gState_G     As String       '//// 그룹코드
Public gState_M     As String       '//// 멀티코드
Public gState_B     As String       '//// 배터리코드


Function Insert_Data_QC(ByVal argSpcRow As Integer) As Integer
    Dim iRow            As Integer
    Dim i               As Integer
    Dim j               As Integer
    Dim k               As Integer
    Dim lsID            As String
    Dim lsSpecNo        As String
    Dim lsPid           As String
    Dim sResult         As String
    Dim sCnt            As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim ExamCnt         As String
    Dim ExamCode_Spec   As String
    Dim lsQC_Date       As String
    
    Dim insCnt          As String
    Dim strQCcode As String
    Dim varQcCode As Variant
        
    With frmInterface
        Insert_Data_QC = -1
        ExamCode_Spec = ""
        lsID = ""
        lsID = Trim(GetText(.vasID, argSpcRow, colBarcode))
        lsSpecNo = Trim(GetText(.vasID, argSpcRow, colSpecNo))
        lsPid = Trim(GetText(.vasID, argSpcRow, colPID))
        
        lsQC_Date = Format(GetDateFull, "yyyymmdd")

        'Local에서 환자별로 결과값 가져오기
        ClearSpread .vasTemp

        SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag, RESDATE, EXAMDATE" & vbCrLf & _
              " From pat_res " & vbCrLf & _
              " Where equipno = '" & gEquip & "' " & vbCrLf & _
              " And examdate = '" & Format(CDate(.dtpToday.Value), "yyyymmdd") & "'  " & vbCrLf & _
              " And barcode = '" & Trim(GetText(.vasID, argSpcRow, colBarcode)) & "' " & vbCrLf & _
              " And diskno = '" & Trim(GetText(.vasID, argSpcRow, colRack)) & "' " & vbCrLf & _
              " And posno = '" & Trim(GetText(.vasID, argSpcRow, colPos)) & "' " & vbCrLf & _
              " And sendflag < '2' "
        
        res = db_select_Vas(gLocal, SQL, .vasTemp)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
        
        For i = 1 To frmInterface.vasTemp.DataRowCnt
            If ExamCode_Spec <> "" Then
                ExamCode_Spec = ExamCode_Spec & ",'" & Trim(GetText(frmInterface.vasTemp, i, 2)) & "'"
            Else
                ExamCode_Spec = "'" & Trim(GetText(frmInterface.vasTemp, i, 2)) & "'"
            End If
        Next i
        
        If ExamCode_Spec = "" Then: ExamCode_Spec = "''"
        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1

    
        sCnt = ""
        sResult1 = ""
        sResult2 = ""
        
        SQL = "SELECT EXMN_CD FROM SPSLMQCTM "
        If lsID = "9904800111" Then
            SQL = SQL & vbCrLf & " WHERE eqpm_cd = '048'"   'ABL800 BASIC QC
        Else
            SQL = SQL & vbCrLf & " WHERE eqpm_cd = '036'"   'ABL800 BASIC QC
        End If
        res = db_select_Row(gServer, SQL)
        
        strQCcode = ""
        
        For k = 0 To UBound(gReadBuf)
            If gReadBuf(k) <> "" Then
                strQCcode = strQCcode & Trim(gReadBuf(k)) & ","
            Else
                Exit For
            End If
        Next
        
        varQcCode = Split(strQCcode, ",")

        cn_Ser.BeginTrans
        '서버로 결과값 저장하기
        For iRow = 1 To .vasTemp.DataRowCnt
            sCnt = ""
            
            sResult1 = Trim(GetText(.vasTemp, iRow, 4))
            sResult2 = Trim(GetText(.vasTemp, iRow, 3))
            
            If Mid(sResult1, 1, 3) = "-99" Then: sResult1 = " "
            
            If sResult1 <> "" Then
                SQL = "SELECT MAX(RSLT_SQNO) FROM SPSLHQRST "
                SQL = SQL & vbCrLf & "WHERE EQPM_CD = '" & Mid(lsID, 3, 3) & "' "
                SQL = SQL & vbCrLf & "  AND SBSN_CD = '" & Mid(lsID, 6, 3) & "' "
                SQL = SQL & vbCrLf & "  AND LVL_CD = '" & Mid(lsID, 9, 1) & "' "
                SQL = SQL & vbCrLf & "  AND EXMN_CD  = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "
                SQL = SQL & vbCrLf & "  AND EXMN_DY = '" & Trim(GetText(.vasTemp, iRow, 9)) & "' "
                SQL = SQL & vbCrLf & "  AND RSLT_VALU IS NULL "
                SQL = SQL & vbCrLf & "GROUP BY RSLT_SQNO "
                res = db_select_Col(gServer, SQL)
                sCnt = gReadBuf(0)
                
                If IsNumeric(sCnt) = True Then
                    SQL = "UPDATE SPSLHQRST "
                    SQL = SQL & vbCrLf & "  SET RSLT_VALU = '" & sResult1 & "', "                        '결과(장비결과)
                    SQL = SQL & vbCrLf & "      RSLT_DT = sysdate, "                                     '결과(수정결과)"
                    SQL = SQL & vbCrLf & "      RSLT_RPTR_ID = '" & gUserID & "', "                                                           'Delta 체크"
                    SQL = SQL & vbCrLf & "      AMEN_ID = '" & gUserID & "', "                                                           'Panic 체크"
                    SQL = SQL & vbCrLf & "      UPDT_DT = sysdate "                                     '결과입력자"
                    SQL = SQL & vbCrLf & "WHERE EQPM_CD = '" & Mid(lsID, 3, 3) & "' "
                    SQL = SQL & vbCrLf & "  AND SBSN_CD = '" & Mid(lsID, 6, 3) & "' "
                    SQL = SQL & vbCrLf & "  AND LVL_CD = '" & Mid(lsID, 9, 1) & "' "
                    SQL = SQL & vbCrLf & "  AND EXMN_CD  = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "
                    SQL = SQL & vbCrLf & "  AND EXMN_DY = '" & Trim(GetText(.vasTemp, iRow, 9)) & "' "
                    SQL = SQL & vbCrLf & "  AND RSLT_SQNO = '" & sCnt & "' "
                    SQL = SQL & vbCrLf & "  AND RSLT_VALU IS NULL "
                    res = SendQuery(gServer, SQL)
                    If res < 0 Then
                        SaveQuery SQL
                       ' db_RollBack gServer
                       cn_Ser.RollbackTrans
                        Exit Function
                    End If
                
                Else
                    If insCnt = "" Then
                        SQL = "SELECT MAX(RSLT_SQNO) FROM SPSLHQRST "
                        SQL = SQL & vbCrLf & "WHERE EQPM_CD = '" & Mid(lsID, 3, 3) & "' "
                        SQL = SQL & vbCrLf & "  AND SBSN_CD = '" & Mid(lsID, 6, 3) & "' "
                        SQL = SQL & vbCrLf & "  AND LVL_CD = '" & Mid(lsID, 9, 1) & "' "
                        'SQL = SQL & vbCrLf & "  AND EXMN_CD  = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "
                        SQL = SQL & vbCrLf & "  AND EXMN_DY = '" & Trim(GetText(.vasTemp, iRow, 9)) & "' "
                        res = db_select_Col(gServer, SQL)
                    
                        If gReadBuf(0) = "" Then
                            insCnt = "1"
                        Else
                            insCnt = CLng(gReadBuf(0)) + 1
                        End If
                    End If
                    
                    For k = 0 To UBound(varQcCode) - 1
                        If Trim(GetText(.vasTemp, iRow, 2)) <> "" And Trim(GetText(.vasTemp, iRow, 2)) = varQcCode(k) Then
                            SQL = ""
                            SQL = SQL & vbCrLf & "INSERT INTO SPSLHQRST(EXMN_DY   ,EQPM_CD ,SBSN_CD ,LVL_CD  "
                            SQL = SQL & vbCrLf & "                     ,RSLT_SQNO ,EXMN_CD ,RSLT_DT ,RSLT_RPTR_ID "
                            SQL = SQL & vbCrLf & "                     ,RSLT_VALU ,SPCM_NO ,DEL_YN "
                            SQL = SQL & vbCrLf & "                     ,REGI_ID   ,RGST_DT ,AMEN_ID ,UPDT_DT) "
                            SQL = SQL & vbCrLf & "               VALUES('" & Trim(GetText(.vasTemp, iRow, 9)) & "', '" & Mid(lsID, 3, 3) & "', '" & Mid(lsID, 6, 3) & "', '" & Mid(lsID, 9, 1) & "', "
                            SQL = SQL & vbCrLf & "                      " & insCnt & ", '" & Trim(GetText(.vasTemp, iRow, 2)) & "', sysdate, '" & gUserID & "', "
                            'SQL = SQL & vbCrLf & "                      " & sCnt & ", '" & Trim(GetText(.vasTemp, iRow, 2)) & "', sysdate, '" & gUserID & "', "
                            SQL = SQL & vbCrLf & "                      '" & sResult1 & "', '" & lsID & "', 'N', "
                            SQL = SQL & vbCrLf & "                      '" & gUserID & "', sysdate, '" & gUserID & "', sysdate ) "
                            res = SendQuery(gServer, SQL)
                            If res = -1 Then
                                SaveQuery SQL
                                cn_Ser.RollbackTrans
                                Exit Function
                            End If
                            Exit For
                        End If
                    Next
                        
                End If
                
            End If
            
        Next iRow
        
        cn_Ser.CommitTrans
        Insert_Data_QC = 1
    End With
    
End Function

'-- 해당 환자 검사의 H/L, Delta, Panic 판정하기
Function GetDecision(ByVal argSpcRow As Integer, ByVal strBarno As String, ByVal strExamCode As String, ByVal strResult As String) As String
    Dim rs_Delta        As ADODB.Recordset
    Dim rs_DPRef        As ADODB.Recordset
    Dim strBefoRslt     As String
    Dim strDestRslt     As String
    Dim strHLVal        As String
    Dim strDelta        As String
    Dim strPanic        As String
    Dim strSex          As String
    Dim strHVal         As String
    Dim strLVal         As String
                
             
    '-- 환자의 성별
    strSex = Trim(GetText(frmInterface.vasID, argSpcRow, colSex))
    
    '##### 바인딩 수정 - 11 ##############################################
''    '-- 해당 환자의 참고치,델타,패닉 찾아오기
''    SQL = "SELECT MALE_HIGH,MALE_LOW,FEML_HIGH,FEML_LOW,DELT_DVSN,DELT_HIGH,DELT_LOW,DELT_DD,PANC_DVSN,PANC_HIGH,PANC_LOW                 "
''    SQL = SQL & vbCrLf & " FROM SPSLMFBIF                                                                                                                      "
''    SQL = SQL & vbCrLf & " WHERE USE_STR_DY <= SYSDATE                                                                                                         "
''    SQL = SQL & vbCrLf & "   AND USE_END_DY >= SYSDATE                                                                                                         "
''    SQL = SQL & vbCrLf & "   and EXMN_CD = '" & Trim(strExamCode) & "' "
''    Set rs_DPRef = cn_Ser.Execute(SQL)
    SQL = "SELECT MALE_HIGH,MALE_LOW,FEML_HIGH,FEML_LOW,DELT_DVSN,DELT_HIGH,DELT_LOW,DELT_DD,PANC_DVSN,PANC_HIGH,PANC_LOW "
    SQL = SQL & vbCrLf & " FROM SPSLMFBIF               "
    SQL = SQL & vbCrLf & " WHERE USE_STR_DY <= ?  "
    SQL = SQL & vbCrLf & "   AND USE_END_DY >= ?  "
    SQL = SQL & vbCrLf & "   and EXMN_CD = ? "
    
    Set AdoCmd_ORACLE = New ADODB.Command
    Set AdoRs_ORACLE = New ADODB.Recordset
    Set AdoCmd_ORACLE.ActiveConnection = cn_Ser 'ADOConnection
    
    AdoCmd_ORACLE.CommandType = adCmdText
    AdoCmd_ORACLE.CommandText = SQL
    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("USE_STR_DY", adDBDate, , , gsDBDateTime)
    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("USE_END_DY", adDBDate, , , gsDBDateTime)
    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 12, Trim(strExamCode))
    
    Set AdoRs_ORACLE = New ADODB.Recordset
    AdoRs_ORACLE.Open AdoCmd_ORACLE, , adOpenStatic, adLockBatchOptimistic
    
    If AdoRs_ORACLE.BOF = False Then
        '레코드셋 전달
        Set rs_DPRef = AdoRs_ORACLE
        Do Until rs_DPRef.EOF
            '-- 성별로 판정결과 비교
            '-- 결과값이 수치일 경우에만 비교한다.
            If IsNumeric(strResult) Then
                strHLVal = ""
                If strSex = "M" Then
                    If IsNumeric(rs_DPRef.Fields("MALE_HIGH")) Then
                        If CDbl(strResult) > CDbl(rs_DPRef.Fields("MALE_HIGH")) Then
                            strHLVal = "H"
                        Else
                            strHLVal = " "
                        End If
                    Else
                        strHLVal = ""
                    End If
                    
                    If IsNumeric(rs_DPRef.Fields("MALE_LOW")) Then
                        If Trim(strHLVal) = "" Then
                            If CDbl(strResult) < CDbl(rs_DPRef.Fields("MALE_LOW")) Then
                                strHLVal = "L"
                            Else
                                strHLVal = " "
                            End If
                        End If
                    Else
                        strHLVal = ""
                    End If
                
                Else
                    If IsNumeric(rs_DPRef.Fields("FEML_HIGH")) Then
                        If CDbl(strResult) > CDbl(rs_DPRef.Fields("FEML_HIGH")) Then
                            strHLVal = "H"
                        Else
                            strHLVal = " "
                        End If
                    Else
                        strHLVal = ""
                    End If
                    If IsNumeric(rs_DPRef.Fields("FEML_LOW")) Then
                        If Trim(strHLVal) = "" Then
                            If (CDbl(strResult) < CDbl(rs_DPRef.Fields("FEML_LOW"))) Then
                                strHLVal = "L"
                            Else
                                strHLVal = " "
                            End If
                        End If
                    Else
                        strHLVal = ""
                    End If
                End If
            Else
                strHLVal = ""
            End If
            
            '-- Panic 구분
            '-- 결과값이 수치일 경우에만 비교한다.
            If IsNumeric(strResult) Then
                strPanic = ""
                Select Case Trim(rs_DPRef.Fields("PANC_DVSN"))
                    Case 0:     '0 사용안함
                            strPanic = ""
                    Case 1:     '1 상한만
                            If IsNumeric(rs_DPRef.Fields("PANC_HIGH")) Then
                                If CDbl(strResult) > rs_DPRef.Fields("PANC_HIGH") Then
                                    strPanic = "P"
                                Else
                                    strPanic = " "
                                End If
                            Else
                                strPanic = ""
                            End If
                    Case 2:     '2 하한만
                            If IsNumeric(rs_DPRef.Fields("PANC_LOW")) Then
                                If CDbl(strResult) < rs_DPRef.Fields("PANC_LOW") Then
                                    strPanic = "P"
                                Else
                                    strPanic = " "
                                End If
                            Else
                                strPanic = ""
                            End If
                    Case 3:     '3 모두 사용
                            If IsNumeric(rs_DPRef.Fields("PANC_LOW")) And IsNumeric(rs_DPRef.Fields("PANC_HIGH")) Then
                                If (CDbl(strResult) < rs_DPRef.Fields("PANC_LOW") Or _
                                    CDbl(strResult) > rs_DPRef.Fields("PANC_HIGH")) Then
                                    strPanic = "P"
                                Else
                                    strPanic = " "
                                End If
                            Else
                                strPanic = ""
                            End If
                    Case Else:
                            strPanic = ""
                End Select
            Else
                strPanic = ""
            End If
            
    
        
            '** 이전결과 조회 시작
            '-- 델타값을 계산하기 위한 이전결과 조회 (한달이내 결과값중 최근값만 조회한다.)
            SQL = ""
            SQL = SQL & vbCrLf & "SELECT B.SPCM_NO           BEFO_BCNO                                                               "
            SQL = SQL & vbCrLf & "     , B.EXMN_CD           BEFO_EXMN_CD                                                            "
            SQL = SQL & vbCrLf & "     , B.REAL_RSLT         BEFO_REAL_RSLT                                                          "
            SQL = SQL & vbCrLf & "     , B.VIEW_RSLT         BEFO_VIEW_RSLT                                                          "
            SQL = SQL & vbCrLf & "     , B.LAST_RPTG_DT     BEFO_FINL_DT                                                             "
            SQL = SQL & vbCrLf & "     , (SYSDATE - B.LAST_RPTG_DT)  DELTA_TERM_DT                                                   "  '오늘부터의 이전결과 기간을 구한다.
            SQL = SQL & vbCrLf & "     , B.PID               PID                                                                     "
            SQL = SQL & vbCrLf & "  FROM (SELECT MAX(B.LAST_RPTG_DT) LAST_RPTG_DT                                                    "
            SQL = SQL & vbCrLf & "             , B.EXMN_CD                                                                           "
            SQL = SQL & vbCrLf & "             , B.PID                                                                               "
            SQL = SQL & vbCrLf & "          FROM SPSLHRRST A, SPSLHRRST B                                                            "
            SQL = SQL & vbCrLf & "         WHERE A.SPCM_NO   <> B.SPCM_NO                                                            "
            SQL = SQL & vbCrLf & "           AND A.PID        = B.PID                                                                "
            SQL = SQL & vbCrLf & "           AND A.EXMN_CD    = B.EXMN_CD                                                            "
            SQL = SQL & vbCrLf & "           AND A.RCPN_DT   >= B.RCPN_DT                                                            "
            SQL = SQL & vbCrLf & "           AND B.LAST_RPTG_DT IS NOT NULL                                                          "
            'SQL = SQL & vbCrLf & "           AND A.RSLT_STAT < '3'                                                                   "
            SQL = SQL & vbCrLf & "           AND A.SPCM_NO = FN_LABCVTBCNO('" & strBarno & "')                                       "
            SQL = SQL & vbCrLf & "         GROUP BY B.PID, B.EXMN_CD ) A, SPSLHRRST B                                                "
            SQL = SQL & vbCrLf & " WHERE A.PID = B.PID                                                                               "
            SQL = SQL & vbCrLf & "   AND A.LAST_RPTG_DT = B.LAST_RPTG_DT                                                             "
            SQL = SQL & vbCrLf & "   AND A.EXMN_CD = B.EXMN_CD                                                                       "
            SQL = SQL & vbCrLf & "   AND A.EXMN_CD = '" & Trim(strExamCode) & "' "         '검사코드"
            SQL = SQL & vbCrLf & "   AND B.LAST_RPTG_DT BETWEEN (SYSDATE-30) AND SYSDATE                "           '-- 30일 이내
            Set rs_Delta = cn_Ser.Execute(SQL)
            Do Until rs_Delta.EOF
                strBefoRslt = rs_Delta.Fields("BEFO_REAL_RSLT")             '이전결과
                strDestRslt = Trim(strResult)  '현재결과
                If IsNumeric(strBefoRslt) = False Then '///////////////////// 이전결과가 문자가 섞였을때
                    Do
                        If Trim(strBefoRslt) = "" Then Exit Do
                        strBefoRslt = Mid(strBefoRslt, 2)
                        If IsNumeric(Mid(strBefoRslt, 1, 1)) = True Then
                            If InStr(1, strBefoRslt, ")") > 0 Then: strBefoRslt = Mid(strBefoRslt, 1, InStr(1, strBefoRslt, ")") - 1)
                            Exit Do
                        End If
                    Loop
                End If
                
                '-- Delta 구분  (아래 로직이 맞는지 검증 필요함...必)
                '-- 결과값이 수치일 경우에만 비교한다.
                If IsNumeric(strDestRslt) And IsNumeric(strBefoRslt) = True Then
                    strDelta = ""
                    Select Case Trim(rs_DPRef.Fields("DELT_DVSN"))
                        Case 0:     '0 사용안함
                                strDelta = ""
                        Case 1:     '1 변화차 = 현재결과 - 이전결과
                                strDelta = ""
                                strDelta = CDbl(strDestRslt) - CDbl(strBefoRslt)                    '변화차
                        Case 2:     '2 변화비율 = 변화차 / 이전결과 * 100
                                strDelta = ""
                                strDelta = CDbl(strDestRslt) - CDbl(strBefoRslt)                    '변화차
                                strDelta = (CDbl(strDelta) / CDbl(strBefoRslt)) * 100               '변화비율
                        Case 3:     '3 기간당 변화비율 = 변화비율 / 기간
                                strDelta = ""
                                strDelta = CDbl(strDestRslt) - CDbl(strBefoRslt)                    '변화차
                                strDelta = (CDbl(strDelta) / CDbl(strBefoRslt)) * 100               '변화비율
                                strDelta = strDelta / CDbl(rs_Delta.Fields("DELTA_TERM_DT"))        '기간당 변화비율
                        Case 4:     '4 기간당 변화차 = 변화차 / 기간
                                strDelta = ""
                                strDelta = CDbl(strDestRslt) - CDbl(strBefoRslt)                    '변화차
                                strDelta = CDbl(strDelta) / CDbl(rs_Delta.Fields("DELTA_TERM_DT"))  '기간당 변화차
                        Case 5:     '5 절대변화비율 = 변화차 / 이전결과
                                strDelta = ""
                                strDelta = CDbl(strDestRslt) - CDbl(strBefoRslt)                    '변화차
                                strDelta = CDbl(strDelta) / CDbl(strBefoRslt)                       '절대변화비율
                        Case Else:
                                strDelta = ""
                    End Select
                Else
                    strDelta = ""
                End If
                '-- Delta 판정
                If IsNumeric(rs_DPRef.Fields("DELT_HIGH")) And IsNumeric(rs_DPRef.Fields("DELT_LOW")) Then
                    If (CDbl(strDestRslt) > rs_DPRef.Fields("DELT_HIGH") Or CDbl(strDestRslt) < rs_DPRef.Fields("DELT_LOW")) Then
                        strDelta = "D"
                    Else
                        strDelta = " "
                    End If
                Else
                    strPanic = ""
                End If
    
                rs_Delta.MoveNext
            Loop
            
            rs_DPRef.MoveNext
        Loop
    End If
    
    GetDecision = strHLVal & "|" & strDelta & "|" & strPanic
    
    Set rs_DPRef = Nothing
    Set AdoCmd_ORACLE = Nothing
    Set AdoRs_ORACLE = Nothing
    
    '##### 바인딩 수정 - 11 ##############################################
        
    GetDecision = strHLVal & "|" & strDelta & "|" & strPanic
    
End Function



Function RsltState_Check(asSpecNo As String, asExamCode As String) As String '/// 결과 형태 : (그룹코드/멀티코드) : 상태가 중간보고 이하일때
    Dim PRSC_CD_G       As String
    Dim EXMN_CD         As String
    Dim PRSC_CD_M       As String
    Dim PRSC_CD_B       As String
    
    Dim strExam         As Variant
    Dim varExam         As Variant
    Dim i               As Integer
    Dim strWhereSQL     As String
    
    RsltState_Check = ""
    PRSC_CD_G = " "
    PRSC_CD_M = " "
    PRSC_CD_B = " "
    
    '##### 바인딩 수정 - 37 ##############################################
''    SQL = ""
''    SQL = SQL & vbCrLf & "SELECT DISTINCT "
''    SQL = SQL & vbCrLf & "       R1.PRSC_CD "
''    SQL = SQL & vbCrLf & "      ,NVL(R1.RSLT_STAT, '0') RSLT_FLAG "
''    SQL = SQL & vbCrLf & "  FROM SPSLHRRST R1 "
''    SQL = SQL & vbCrLf & "      ,SPSLMFBIF F1 "
''    SQL = SQL & vbCrLf & " WHERE R1.EXMN_CD = F1.EXMN_CD "
''    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT >= F1.USE_STR_DY "
''    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT <  F1.USE_END_DY "
''    SQL = SQL & vbCrLf & "   AND R1.SPCM_NO  = '" & asSpecNo & "' "
''    SQL = SQL & vbCrLf & "   AND R1.EXMN_CD = '" & asExamCode & "' "
''    SQL = SQL & vbCrLf & "   AND R1.PRSC_CD LIKE ('%G%') "
''    SQL = SQL & vbCrLf & " GROUP BY R1.PRSC_CD, R1.RSLT_STAT "
''    res = db_select_Col(gServer, SQL)
''    If gReadBuf(0) <> "" Then: PRSC_CD_G = gReadBuf(0)
''    gReadBuf(0) = ""

    SQL = ""
    SQL = SQL & vbCrLf & "SELECT DISTINCT R1.PRSC_CD, NVL(R1.RSLT_STAT, '0') RSLT_FLAG "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST R1, SPSLMFBIF F1 "
    SQL = SQL & vbCrLf & " WHERE R1.EXMN_CD = F1.EXMN_CD "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT >= F1.USE_STR_DY "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT <  F1.USE_END_DY "
    SQL = SQL & vbCrLf & "   AND R1.SPCM_NO  = ? "
    SQL = SQL & vbCrLf & "   AND R1.EXMN_CD = ? "
    SQL = SQL & vbCrLf & "   AND R1.PRSC_CD LIKE (?) "
    SQL = SQL & vbCrLf & " GROUP BY R1.PRSC_CD, R1.RSLT_STAT "
    
    Set AdoCmd_ORACLE = New ADODB.Command
    Set AdoRs_ORACLE = New ADODB.Recordset
    Set AdoCmd_ORACLE.ActiveConnection = cn_Ser 'ADOConnection
    
    AdoCmd_ORACLE.CommandType = adCmdText
    AdoCmd_ORACLE.CommandText = SQL
    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, asSpecNo)
    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 12, asExamCode)
    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("PRSC_CD", adVarChar, , 12, "%G%")   '처방코드
    
    Set AdoRs_ORACLE = New ADODB.Recordset
    AdoRs_ORACLE.Open AdoCmd_ORACLE, , adOpenStatic, adLockBatchOptimistic
    
    If AdoRs_ORACLE.BOF = False Then
        PRSC_CD_G = AdoRs_ORACLE.Fields(0) & ""
    End If
    Set AdoCmd_ORACLE = Nothing
    Set AdoRs_ORACLE = Nothing
    '##### 바인딩 수정 - 37 ##############################################
    
    '##### 바인딩 수정 - 28 ##############################################
''    SQL = ""
''    SQL = SQL & vbCrLf & "SELECT DISTINCT "
''    SQL = SQL & vbCrLf & "      ,R1.EXMN_CD "
''    SQL = SQL & vbCrLf & "      ,NVL(R1.RSLT_STAT, '0') RSLT_FLAG "
''    SQL = SQL & vbCrLf & "  FROM SPSLHRRST R1 "
''    SQL = SQL & vbCrLf & "      ,SPSLMFBIF F1 "
''    SQL = SQL & vbCrLf & " WHERE R1.EXMN_CD = F1.EXMN_CD "
''    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT >= F1.USE_STR_DY "
''    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT <  F1.USE_END_DY "
''    SQL = SQL & vbCrLf & "   AND R1.SPCM_NO  = '" & asSpecNo & "' "
''    SQL = SQL & vbCrLf & "   AND R1.EXMN_CD IN (" & gAllExam & ") "
''    SQL = SQL & vbCrLf & "   AND F1.CD_DVSN IN ('M') "
''    SQL = SQL & vbCrLf & " GROUP BY R1.EXMN_CD, R1.RSLT_STAT "
''    res = db_select_Col(gServer, SQL)
''    If gReadBuf(0) <> "" Then: PRSC_CD_M = gReadBuf(0)
''    gReadBuf(0) = ""
    strExam = Replace(gAllExam, "'", "")
    varExam = Split(strExam, ",")
    strWhereSQL = ""
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT DISTINCT "
    SQL = SQL & vbCrLf & "       R1.EXMN_CD "
    SQL = SQL & vbCrLf & "      ,NVL(R1.RSLT_STAT, '0') RSLT_FLAG "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST R1 "
    SQL = SQL & vbCrLf & "      ,SPSLMFBIF F1 "
    SQL = SQL & vbCrLf & " WHERE R1.EXMN_CD = F1.EXMN_CD "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT >= F1.USE_STR_DY "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT <  F1.USE_END_DY "
    SQL = SQL & vbCrLf & "   AND R1.SPCM_NO  = ? "
    SQL = SQL & vbCrLf & "   AND F1.CD_DVSN IN (?) "
    For i = 0 To UBound(varExam)
        strWhereSQL = strWhereSQL & "R1.EXMN_CD = ? OR "
    Next
    If strWhereSQL <> "" Then
        SQL = SQL & vbCrLf & "AND (" & Mid(strWhereSQL, 1, Len(strWhereSQL) - 3) & ")"
    End If
    
    SQL = SQL & vbCrLf & " GROUP BY R1.EXMN_CD, R1.RSLT_STAT "
    
    Set AdoCmd_ORACLE = New ADODB.Command
    Set AdoRs_ORACLE = New ADODB.Recordset
    Set AdoCmd_ORACLE.ActiveConnection = cn_Ser 'ADOConnection
    
    AdoCmd_ORACLE.CommandType = adCmdText
    AdoCmd_ORACLE.CommandText = SQL
    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, asSpecNo)
    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("CD_DVSN", adVarChar, , 5, "M")
    For i = 0 To UBound(varExam)
        AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 12, varExam(i))
    Next
    Set AdoRs_ORACLE = New ADODB.Recordset
    AdoRs_ORACLE.Open AdoCmd_ORACLE, , adOpenStatic, adLockBatchOptimistic
    
    If AdoRs_ORACLE.BOF = False Then
        PRSC_CD_M = AdoRs_ORACLE.Fields(0) & ""
    End If
    Set AdoCmd_ORACLE = Nothing
    Set AdoRs_ORACLE = Nothing
    '##### 바인딩 수정 - 28 ##############################################

    '##### 바인딩 수정 - 18 ##############################################
''    SQL = ""
''    SQL = SQL & vbCrLf & "SELECT DISTINCT "
''    SQL = SQL & vbCrLf & "      ,R1.EXMN_CD "
''    SQL = SQL & vbCrLf & "      ,NVL(R1.RSLT_STAT, '0') RSLT_FLAG "
''    SQL = SQL & vbCrLf & "  FROM SPSLHRRST R1 "
''    SQL = SQL & vbCrLf & "      ,SPSLMFBIF F1 "
''    SQL = SQL & vbCrLf & " WHERE R1.EXMN_CD = F1.EXMN_CD "
''    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT >= F1.USE_STR_DY "
''    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT <  F1.USE_END_DY "
''    SQL = SQL & vbCrLf & "   AND R1.SPCM_NO  = '" & asSpecNo & "' "
''    SQL = SQL & vbCrLf & "   AND R1.EXMN_CD IN (" & gAllExam & ") "
''    SQL = SQL & vbCrLf & "   AND F1.CD_DVSN IN ('B') "
''    SQL = SQL & vbCrLf & " GROUP BY R1.EXMN_CD, R1.RSLT_STAT "
''    res = db_select_Col(gServer, SQL)
''    If gReadBuf(0) <> "" Then: PRSC_CD_B = gReadBuf(0)
''    gReadBuf(0) = ""

    Erase varExam
    strExam = Replace(gAllExam, "'", "")
    varExam = Split(strExam, ",")
    strWhereSQL = ""
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT DISTINCT R1.EXMN_CD,NVL(R1.RSLT_STAT, '0') RSLT_FLAG "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST R1, SPSLMFBIF F1 "
    SQL = SQL & vbCrLf & " WHERE R1.EXMN_CD = F1.EXMN_CD "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT >= F1.USE_STR_DY "
    SQL = SQL & vbCrLf & "   AND R1.RCPN_DT <  F1.USE_END_DY "
    SQL = SQL & vbCrLf & "   AND R1.SPCM_NO = ? "
    SQL = SQL & vbCrLf & "   AND F1.CD_DVSN IN (?) "
    For i = 0 To UBound(varExam)
        strWhereSQL = strWhereSQL & "R1.EXMN_CD = ? OR "
    Next
    If strWhereSQL <> "" Then
        SQL = SQL & vbCrLf & "AND (" & Mid(strWhereSQL, 1, Len(strWhereSQL) - 3) & ")"
    End If
    SQL = SQL & vbCrLf & " GROUP BY R1.EXMN_CD, R1.RSLT_STAT "
    
    Set AdoCmd_ORACLE = New ADODB.Command
    Set AdoRs_ORACLE = New ADODB.Recordset
    Set AdoCmd_ORACLE.ActiveConnection = cn_Ser 'ADOConnection
    
    AdoCmd_ORACLE.CommandType = adCmdText
    AdoCmd_ORACLE.CommandText = SQL
    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, asSpecNo)
    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("CD_DVSN", adVarChar, , 5, "B")
    For i = 0 To UBound(varExam)
        AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 12, varExam(i))
    Next
    
    Set AdoRs_ORACLE = New ADODB.Recordset
    AdoRs_ORACLE.Open AdoCmd_ORACLE, , adOpenStatic, adLockBatchOptimistic
    
    If AdoRs_ORACLE.BOF = False Then
        PRSC_CD_B = AdoRs_ORACLE.Fields(0) & ""
    End If
    Set AdoCmd_ORACLE = Nothing
    Set AdoRs_ORACLE = Nothing
    '##### 바인딩 수정 - 18 ##############################################
        
    RsltState_Check = PRSC_CD_G & "/" & PRSC_CD_M & "/" & PRSC_CD_B
    
End Function


Function Make_Remark_all(asExamCode As String, asSex As String, asResult As String)
'///////////// 코멘트 생성 (검체전체)
    Dim i As Integer
    
    Dim Comment_Gubun As String
    Dim Comment_MFGubun As String
    Dim Comment_Code As String      '///////// 판별이후
    Dim Comment_CodeH As String
    Dim Comment_CodeL As String

    Dim Comment_RefMH As String
    Dim Comment_RefML As String
    Dim Comment_RefFH As String
    Dim Comment_RefFL As String

    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT cmtdest, cmtflag, CMTCODE, cmtcodeSub, cmhigh, cmlow, cFhigh, cFlow "
    SQL = SQL & vbCrLf & "  FROM EQUIPEXAM "
    SQL = SQL & vbCrLf & " WHERE EXAMCODE IN (" & asExamCode & ") "
    SQL = SQL & vbCrLf & "   AND CMTDEST = '1' "
    
    res = db_select_Col(gLocal, SQL)
    
    If res < 1 Then: Exit Function
    
    Comment_Gubun = gReadBuf(0)
    Comment_MFGubun = gReadBuf(1)
    Comment_CodeH = gReadBuf(2)
    Comment_CodeL = gReadBuf(3)
    Comment_RefMH = gReadBuf(4)
    Comment_RefML = gReadBuf(5)
    Comment_RefFH = gReadBuf(6)
    Comment_RefFL = gReadBuf(7)

    gReadBuf(0) = ""
    gReadBuf(1) = ""
    gReadBuf(2) = ""
    gReadBuf(3) = ""
    gReadBuf(4) = ""
    gReadBuf(5) = ""
    gReadBuf(6) = ""
    gReadBuf(7) = ""
        
        
    '///// 0:공통, 1:남/여, 2:사용안함
    If Comment_MFGubun = "0" Then
        
        If asResult >= Comment_RefMH Then
            Comment_Code = Comment_CodeH
        ElseIf asResult <= Comment_RefML Then
            Comment_Code = Comment_CodeL
        End If
        
        SQL = ""
        SQL = SQL & vbCrLf & "SELECT CNTS "
        SQL = SQL & vbCrLf & "  FROM SPSLMFRMK "
        SQL = SQL & vbCrLf & " WHERE OPNN_CD = '" & Comment_Code & "' "
        SQL = SQL & vbCrLf & ""
        res = db_select_Col(gServer, SQL)
        
        
        
        
        If InStr(1, gComment_All, gReadBuf(0)) = 0 Then
            If gComment_All = "" Then
                gComment_All = gReadBuf(0)
            Else
                gComment_All = gComment_All & chrCR & gReadBuf(0)
            End If
        End If
    ElseIf Comment_MFGubun = "1" Then
        
        If asSex = "M" Then
            If asResult >= Comment_RefMH Then
                Comment_Code = Comment_CodeH
            ElseIf asResult <= Comment_RefML Then
                Comment_Code = Comment_CodeL
            End If
        ElseIf asSex = "F" Then
            If asResult >= Comment_RefFH Then
                Comment_Code = Comment_CodeH
            ElseIf asResult <= Comment_RefFL Then
                Comment_Code = Comment_CodeL
            End If
        End If

        SQL = ""
        SQL = SQL & vbCrLf & "SELECT CNTS "
        SQL = SQL & vbCrLf & "  FROM SPSLMFRMK "
        SQL = SQL & vbCrLf & " WHERE OPNN_CD = '" & Comment_Code & "' "
        SQL = SQL & vbCrLf & ""
        res = db_select_Col(gServer, SQL)
        
        If InStr(1, gComment_All, gReadBuf(0)) = 0 Then
            If gComment_All = "" Then
                gComment_All = gReadBuf(0)
            Else
                gComment_All = gComment_All & chrCR & gReadBuf(0)
            End If
        End If
        
    ElseIf Comment_MFGubun = "2" Then
        
        SQL = ""
        SQL = SQL & vbCrLf & "SELECT CNTS "
        SQL = SQL & vbCrLf & "  FROM SPSLMFRMK "
        SQL = SQL & vbCrLf & " WHERE OPNN_CD = '" & Comment_CodeH & "' "
        SQL = SQL & vbCrLf & ""
        res = db_select_Col(gServer, SQL)
        
        If InStr(1, gComment_All, gReadBuf(0)) = 0 Then
            If gComment_All = "" Then
                gComment_All = gReadBuf(0)
            Else
                gComment_All = gComment_All & chrCR & gReadBuf(0)
            End If
        End If
        
    End If

    
End Function
'''
'''Function Insert_Data(ByVal argSpcRow As Integer) As Integer
'''    Dim iRow            As Integer
'''    Dim i               As Integer
'''    Dim j               As Integer
'''    Dim lsID            As String
'''    Dim lsSpecNo        As String
'''    Dim lsPid           As String
'''    Dim sResult         As String
'''    Dim sCnt            As String
'''    Dim sResult1        As String
'''    Dim sResult2        As String
'''    Dim ExamCnt         As String
'''    Dim ExamCode_Spec   As String
'''    Dim ExamCode_Remark     As String
'''
'''    Dim State_GM    As String       '//// 그룹/멀티 코드
'''    Dim State_cnt   As Integer      '//// 그룹/멀티 코드 쪽 변수
'''    Dim State_G     As String       '//// 그룹코드
'''    Dim State_M     As String       '//// 멀티코드
'''    Dim State_B     As String       '//// 배터리코드
'''
'''    Dim Send_State      As String
'''    Dim SQL_LOCAL As String
'''
'''    Dim strWhereSQL     As String
'''    Dim sqlRet          As Integer
'''
'''    With frmInterface
'''        gComment_All = ""
'''        Insert_Data = -1
'''        ExamCode_Spec = ""
'''        ExamCode_Remark = ""
'''
'''        State_GM = ""
'''        State_cnt = 0
'''        State_G = ""
'''        State_M = ""
'''        lsID = ""
'''        lsID = Trim(GetText(.vasID, argSpcRow, colBarcode))
'''        lsSpecNo = Trim(GetText(.vasID, argSpcRow, colSpecNo))
'''        lsPid = Trim(GetText(.vasID, argSpcRow, colPID))
'''
'''        'Local에서 환자별로 결과값 가져오기
'''        ClearSpread .vasTemp
'''        ClearSpread .vasTemp1
'''        SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag, PSEX " & vbCrLf & _
'''              " From pat_res " & vbCrLf & _
'''              " Where equipno = '" & gEquip & "' " & vbCrLf & _
'''              " And examdate = '" & Format(CDate(.dtpToday.Value), "yyyymmdd") & "'  " & vbCrLf & _
'''              " And barcode = '" & Trim(GetText(.vasID, argSpcRow, colBarcode)) & "' " & vbCrLf & _
'''              " And diskno = '" & Trim(GetText(.vasID, argSpcRow, colRack)) & "' " & vbCrLf & _
'''              " And posno = '" & Trim(GetText(.vasID, argSpcRow, colPos)) & "' "
'''
'''        Save_Raw_Data "[로컬조회]  " & SQL
'''
'''        res = db_select_Vas(gLocal, SQL, .vasTemp)
'''
'''        If res = -1 Then
'''            SaveQuery SQL
'''            Exit Function
'''        End If
'''
'''        For i = 1 To frmInterface.vasTemp.DataRowCnt    '/// 실제 검사한 검사코드들
'''            If ExamCode_Spec <> "" Then
'''                ExamCode_Spec = ExamCode_Spec & ",'" & Trim(GetText(frmInterface.vasTemp, i, 2)) & "'"
'''            Else
'''                ExamCode_Spec = "'" & Trim(GetText(frmInterface.vasTemp, i, 2)) & "'"
'''            End If
'''        Next i
'''
'''        If ExamCode_Spec = "" Then: ExamCode_Spec = "''"
'''        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1
'''
'''        sCnt = ""
'''        sResult1 = ""
'''        sResult2 = ""
'''
'''
'''
'''        '/-------------------------------리마크 처리 때문에 인터페이스에 저장된 코드로 검체를 조회해서 리마크 표시해줄것을 찾음(필요한장비만 열기)
'''        SQL = "SELECT EXMN_CD "
'''        SQL = SQL & vbCrLf & "FROM SPSLHRRST "
'''        SQL = SQL & vbCrLf & "WHERE EXMN_CD IN (" & gAllExam & ")"
'''        SQL = SQL & vbCrLf & "  AND SPCM_NO = '" & lsSpecNo & "' "
'''        res = db_select_Vas(gServer, SQL, .vasTemp1)
'''
'''
'''        For i = 1 To frmInterface.vasTemp1.DataRowCnt    '/// 실제 검사한 검사코드들
'''            If ExamCode_Remark <> "" Then
'''                ExamCode_Remark = ExamCode_Remark & ",'" & Trim(GetText(frmInterface.vasTemp1, i, 1)) & "'"
'''            Else
'''                ExamCode_Remark = "'" & Trim(GetText(frmInterface.vasTemp, i, 1)) & "'"
'''            End If
'''        Next i
'''
'''        If ExamCode_Remark = "" Then ExamCode_Remark = "''"
'''
'''        For i = 1 To frmInterface.vasTemp.DataRowCnt
'''            Call Make_Remark_all(ExamCode_Remark, Trim(GetText(frmInterface.vasTemp, i, 8)), Trim(GetText(frmInterface.vasTemp, i, 8)))
'''        Next i
'''        '/--------------------------------------------------------------------------------------------------------------
'''
'''        cn_Ser.BeginTrans
'''        '서버로 결과값 저장하기
'''        For iRow = 1 To .vasTemp.DataRowCnt
'''
'''            sResult1 = Trim(GetText(.vasTemp, iRow, 4))
'''            sResult2 = Trim(GetText(.vasTemp, iRow, 3))
'''
'''            If sResult1 <> "" And Mid(sResult1, 1, 3) <> "-99" Then
'''                gComment_Code = ""
'''
'''
'''                SQL = "SELECT RSLT_NO FROM SPSLHRRST "
'''                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
'''                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                      '검사코드"
'''                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
'''                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
'''
'''                Save_Raw_Data "[SQL]  " & SQL
'''
'''                res = db_select_Col(gServer, SQL)
'''
'''                If gReadBuf(0) = "" Then: gReadBuf(0) = "0"
'''
'''                sCnt = CLng(gReadBuf(0)) + 1
'''
'''                '/----------------------------- 자동리마크 처리 (필요한장비만 열기)
''''                Call Make_Remark_all(ExamCode_Remark, Trim(GetText(frmInterface.vasTemp, i, 8)), Trim(GetText(frmInterface.vasTemp, i, 4)))
'''                '/-----------------------------
'''
'''                '-- 결과값이 숫자값일 경우만 델타/패닉 판정을 한다.
'''                sResult = Trim(GetText(frmInterface.vasTemp, iRow, 3))
'''                If IsNumeric(sResult) Then
'''                    Dim strDecision     As Variant
'''                    Dim strBarcode      As String
'''
'''                    strBarcode = Trim(GetText(frmInterface.vasID, argSpcRow, colBarcode))
'''                    'strDecision = GetDecision(argSpcRow, strBarcode, iRow)
'''                    strDecision = GetDecision(argSpcRow, strBarcode, Trim(GetText(frmInterface.vasTemp, iRow, 2)), sResult)
'''                    strDecision = Split(strDecision, "|")
'''                Else
'''                    strDecision = "||"
'''                    strDecision = Split(strDecision, "|")
'''                End If
'''
'''
'''                SQL = "UPDATE SPSLHRRST "   '-- 결과테이블
'''                SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & Trim(GetText(frmInterface.vasTemp, iRow, 4)) & "', "      '결과(장비결과)
'''                SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & Trim(GetText(frmInterface.vasTemp, iRow, 3)) & "', "      '결과(수정결과)"
'''                SQL = SQL & vbCrLf & "       DTRM_DVSN = '" & strDecision(0) & "', "                                    'H/L 체크"
'''                SQL = SQL & vbCrLf & "       DLTA_YN = '" & strDecision(1) & "', "                                      'Delta 체크"
'''                SQL = SQL & vbCrLf & "       PANC_YN = '" & strDecision(2) & "', "                                      'Panic 체크"
'''                SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
'''                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "                                      '결과수정자
'''                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
'''                SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '결과번호 (결과 넣을시에 증가)
'''
'''
'''                Send_State = "1" '/  <---------- 혈액학장비가 아니라서 상태가 1로만 들어감
'''
'''                '/----------------------------- 결과 상태 넣기
'''                If Send_State = "1" Then
'''
'''                    SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gUserID & "', "                                 '결과입력자"
'''                    SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'''                    SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "
'''                ElseIf Send_State = "2" Then
'''
'''                    SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gUserID & "', "                                 '결과입력자"
'''                    SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'''                    SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gUserID & "', "                                 '중간보고자"
'''                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
'''                    SQL = SQL & vbCrLf & "       RSLT_STAT = '2' "
'''                ElseIf Send_State = "3" Then
'''
'''                    SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gUserID & "', "                                 '결과입력자"
'''                    SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'''                    SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gUserID & "', "                                     '중간보고자"
'''                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
'''                    SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gUserID & "', "                                 '최종보고자"
'''                    SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
'''                    SQL = SQL & vbCrLf & "       RSLT_STAT = '3' "
'''                End If
'''
'''                '/----------------------------- 결과 상태 넣기
'''
'''                '/----------------------------- 자동리마크 처리 (필요한장비만 열기)
'''                If gComment_All <> "" Or gComment_Code <> "" Then
'''                    SQL = SQL & vbCrLf & "       ,EXMN_PER_OPNN = '" & gComment_All & chrCR & gComment_Code & "' "
'''                End If
'''                '/-----------------------------
'''                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
'''                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '검사코드"
'''                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
'''                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
'''
'''                Save_Raw_Data "[SQL]  " & SQL
'''
'''                res = SendQuery(gServer, SQL)
'''                If res < 0 Then
'''                    SaveQuery SQL
'''                   ' db_RollBack gServer
'''                   cn_Ser.RollbackTrans
'''                    Exit Function
'''                End If
'''
'''                State_GM = RsltState_Check(lsSpecNo, Trim(GetText(.vasTemp, iRow, 2)))
'''
'''                State_cnt = InStr(1, State_GM, "/")
'''                State_G = Mid(State_GM, 1, State_cnt - 1)
'''                State_GM = Mid(State_GM, State_cnt + 1)
'''                State_cnt = InStr(1, State_GM, "/")
'''                State_M = Mid(State_GM, 1, State_cnt - 1)
'''                State_B = Mid(State_GM, State_cnt + 1)
'''
'''
'''                '/------------------------------------ 결과테이블 그룹코드 상태 업데이트
'''                If Trim(State_G) <> "" Then
'''                    SQL = "UPDATE SPSLHRRST "
'''
'''                        '/////////  D/P/H 가 없을때 : 검사결과를 최종보고로 넣는다
'''                        If Send_State = "1" Then
'''
'''                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gUserID & "', "                                 '결과입력자"
'''                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'''                            SQL = SQL & vbCrLf & "       RSLT_STAT = '1', "
'''                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
'''                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "                                      '결과수정자
'''                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
'''                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
'''                        ElseIf Send_State = "2" Then
'''
'''                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gUserID & "', "                                 '결과입력자"
'''                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gUserID & "', "                                 '중간보고자"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
'''                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gUserID & "', "                                 '최종보고자"
'''                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
'''                            SQL = SQL & vbCrLf & "       RSLT_STAT = '2', "
'''                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
'''                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "                                      '결과수정자
'''                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
'''                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
'''                        ElseIf Send_State = "3" Then
'''
'''                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gUserID & "', "                                 '결과입력자"
'''                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                     '중간보고자"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
'''                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gUserID & "', "                                 '최종보고자"
'''                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
'''                            SQL = SQL & vbCrLf & "       RSLT_STAT = '3', "
'''                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
'''                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "                                      '결과수정자
'''                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
'''                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
'''                        End If
'''                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
'''                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(State_G) & "' "                                        '검사코드"
'''                    SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
'''                    SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
'''
'''                    Save_Raw_Data "[SQL]  " & SQL
'''
'''                    res = SendQuery(gServer, SQL)
'''                    If res = -1 Then
'''                        SaveQuery SQL
'''                        cn_Ser.RollbackTrans
'''                        Exit Function
'''                    End If
'''                End If
'''                '/------------------------------------
'''
'''                '/------------------------------------ 결과테이블 멀티코드 상태 업데이트
'''                If Trim(State_M) <> "" Then
'''                    SQL = "UPDATE SPSLHRRST "
'''
'''                        '/////////  D/P/H 가 없을때 : 검사결과를 최종보고로 넣는다
'''                        If Send_State = "1" Then
'''
'''                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gUserID & "', "                                 '결과입력자"
'''                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'''                            SQL = SQL & vbCrLf & "       RSLT_STAT = '1', "
'''                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
'''                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "                                      '결과수정자
'''                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
'''                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
'''                        ElseIf Send_State = "2" Then
'''
'''                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gUserID & "', "                                 '결과입력자"
'''                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gUserID & "', "                                 '중간보고자"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
'''                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gUserID & "', "                                 '최종보고자"
'''                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
'''                            SQL = SQL & vbCrLf & "       RSLT_STAT = '2', "
'''                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
'''                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "                                      '결과수정자
'''                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
'''                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
'''                        ElseIf Send_State = "3" Then
'''
'''                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gUserID & "', "                                 '결과입력자"
'''                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                     '중간보고자"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
'''                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gUserID & "', "                                 '최종보고자"
'''                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
'''                            SQL = SQL & vbCrLf & "       RSLT_STAT = '3', "
'''                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
'''                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "                                      '결과수정자
'''                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
'''                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
'''                        End If
'''                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
'''                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(State_M) & "' "                                        '검사코드"
'''                    SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
'''                    SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
'''
'''                    Save_Raw_Data "[SQL]  " & SQL
'''
'''                    res = SendQuery(gServer, SQL)
'''                    If res = -1 Then
'''                        SaveQuery SQL
'''                        cn_Ser.RollbackTrans
'''                        Exit Function
'''                    End If
'''                End If
'''            '/------------------------------------
'''
'''            '/------------------------------------ 결과테이블 그룹코드 상태 업데이트
'''                If Trim(State_B) <> "" Then
'''                    SQL = "UPDATE SPSLHRRST "
'''
'''                        '/////////  D/P/H 가 없을때 : 검사결과를 최종보고로 넣는다
'''                        If Send_State = "1" Then
'''
'''                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gUserID & "', "                                 '결과입력자"
'''                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'''                            SQL = SQL & vbCrLf & "       RSLT_STAT = '1', "
'''                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
'''                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "                                      '결과수정자
'''                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
'''                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
'''                        ElseIf Send_State = "2" Then
'''
'''                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gUserID & "', "                                 '결과입력자"
'''                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gUserID & "', "                                 '중간보고자"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
'''                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gUserID & "', "                                 '최종보고자"
'''                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
'''                            SQL = SQL & vbCrLf & "       RSLT_STAT = '2', "
'''                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
'''                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "                                      '결과수정자
'''                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
'''                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
'''                        ElseIf Send_State = "3" Then
'''
'''                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = '" & gUserID & "', "                                 '결과입력자"
'''                            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & gEquipCode & "', "                                     '중간보고자"
'''                            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
'''                            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & gUserID & "', "                                 '최종보고자"
'''                            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
'''                            SQL = SQL & vbCrLf & "       RSLT_STAT = '3', "
'''                            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
'''                            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "                                      '결과수정자
'''                            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
'''                            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "' "
'''                        End If
'''                    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
'''                    SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(State_B) & "' "                                        '검사코드"
'''                    SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
'''                    SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
'''
'''                    Save_Raw_Data "[SQL]  " & SQL
'''
'''                    res = SendQuery(gServer, SQL)
'''                    If res = -1 Then
'''                        SaveQuery SQL
'''                        cn_Ser.RollbackTrans
'''                        Exit Function
'''                    End If
'''                End If
'''            '/------------------------------------
'''
'''            '/------------------------------------ 접수테이블 STATE 업데이트
'''
'''                '##### 바인딩 수정 - 7 ##############################################
'''''                '////////// 접수 테이블
'''''                SQL = "UPDATE SPSLMJBDI "
'''''                If Send_State = "1" Then
'''''                    SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
'''''                    SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "
'''''                    SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
'''''                ElseIf Send_State = "2" Then
'''''                    SQL = SQL & vbCrLf & "   SET RSLT_STAT = '2', "
'''''                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
'''''                    SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "
'''''                    SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
'''''                ElseIf Send_State = "3" Then
'''''                    SQL = SQL & vbCrLf & "   SET RSLT_STAT = '3', "
'''''                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
'''''                    SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
'''''                    SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "
'''''                    SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
'''''                End If
'''''                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
'''''                SQL = SQL & vbCrLf & "   AND EXMN_CD IN ('" & Trim(State_G) & "','" & Trim(State_M) & "','" & Trim(State_B) & "','" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "') "                    '검사코드"
'''''                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
'''''                SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
'''''                res = SendQuery(gServer, SQL)
'''
'''                If Send_State = "1" Then
'''                    SQL = "UPDATE SPSLMJBDI SET RSLT_STAT = ?, AMEN_ID = ?, UPDT_DT = ? " & vbCrLf
'''                ElseIf Send_State = "2" Then
'''                    SQL = "UPDATE SPSLMJBDI SET RSLT_STAT = ?, AMEN_ID = ?, UPDT_DT = ?, MDDL_RPTG_DT = ? " & vbCrLf
'''                ElseIf Send_State = "3" Then
'''                    SQL = "UPDATE SPSLMJBDI SET RSLT_STAT = ?, AMEN_ID = ?, UPDT_DT = ?, MDDL_RPTG_DT = ?, LAST_RPTG_DT = ? " & vbCrLf
'''                End If
'''                SQL = SQL & " WHERE SPCM_NO = ? " & vbCrLf
'''                SQL = SQL & "   AND RSLT_STAT <> ? " & vbCrLf
'''                SQL = SQL & "   AND SPCM_STAT = ? " & vbCrLf
'''                strWhereSQL = ""
'''                If State_G <> "" Then
'''                    strWhereSQL = strWhereSQL & "   AND (EXMN_CD = ? "
'''                    If State_M <> "" Then
'''                        strWhereSQL = strWhereSQL & "   OR EXMN_CD = ? )"
'''                    Else
'''                        strWhereSQL = strWhereSQL & ")"
'''                    End If
'''                Else
'''                    If State_M <> "" Then
'''                        strWhereSQL = strWhereSQL & "   AND EXMN_CD = ? "
'''                    End If
'''                End If
'''                SQL = SQL & strWhereSQL
'''
'''                Set AdoCmd_ORACLE = New ADODB.Command
'''                Set AdoCmd_ORACLE.ActiveConnection = cn_Ser 'ADOConnection
'''
'''                AdoCmd_ORACLE.CommandType = adCmdText
'''                AdoCmd_ORACLE.CommandText = SQL
'''
'''                '-- 시스템 날짜 가져오기 함수 : gsDBDateTime
'''
'''                If Send_State = "1" Then
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "1")
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, gUserID & "")
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(lsSpecNo))
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_STAT", adVarChar, , 5, "2")
'''                    If State_G <> "" Then
'''                        AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 12, Trim(State_G))
'''                    End If
'''                    If State_M <> "" Then
'''                        AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 12, Trim(State_M))
'''                    End If
'''                ElseIf Send_State = "2" Then
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 15, "2")
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, gUserID & "")
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTG_DT", adDBDate, , 5, gsDBDateTime)
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(lsSpecNo))
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_STAT", adVarChar, , 5, "2")
'''                    If State_G <> "" Then
'''                        AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 5, Trim(State_G))
'''                    End If
'''                    If State_M <> "" Then
'''                        AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 5, Trim(State_M))
'''                    End If
'''                ElseIf Send_State = "3" Then
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, gUserID & "")
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTG_DT", adDBDate, , , gsDBDateTime)
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("LAST_RPTG_DT", adDBDate, , , gsDBDateTime)
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(lsSpecNo))
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
'''                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_STAT", adVarChar, , 5, "2")
'''                    If State_G <> "" Then
'''                        AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 5, Trim(State_G))
'''                    End If
'''                    If State_M <> "" Then
'''                        AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 5, Trim(State_M))
'''                    End If
'''                End If
'''
'''                AdoCmd_ORACLE.Execute sqlRet
'''                Set AdoCmd_ORACLE = Nothing
'''                '##### 바인딩 수정 - 7 ##############################################
'''
'''                If sqlRet < 0 Then
'''                    SaveQuery SQL
'''                    cn_Ser.RollbackTrans
'''                    Exit Function
'''                End If
'''
'''            '/------------------------------------
'''            End If
'''        Next iRow
'''
'''        '/------------------------------------ 처방테이블 STATE 업데이트
'''        '///////// 처방테이블
'''        '##### 바인딩 수정 - 19 ##############################################
'''''        SQL = "UPDATE SPSLMJBBI "
'''''        If Send_State = "1" Then
'''''            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
'''''            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "
'''''            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
'''''        ElseIf Send_State = "2" Then
'''''            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '2', "
'''''            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "
'''''            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
'''''        ElseIf Send_State = "3" Then
'''''            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '3', "
'''''            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gUserID & "', "
'''''            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
'''''        End If
'''''        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
'''''        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
'''''        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
'''''        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
'''''        res = SendQuery(gServer, SQL)
'''
'''        SQL = "UPDATE SPSLMJBBI "
'''        SQL = SQL & vbCrLf & "   SET RSLT_STAT = ?, AMEN_ID = ?, UPDT_DT = ? "
'''        SQL = SQL & vbCrLf & " WHERE SPCM_NO = ? "
'''        SQL = SQL & vbCrLf & "   AND PID = ? "
'''        SQL = SQL & vbCrLf & "   AND RSLT_STAT <> ? "
'''        SQL = SQL & vbCrLf & "   AND SPCM_STAT = ? "
'''
'''        Set AdoCmd_ORACLE = New ADODB.Command
'''        Set AdoCmd_ORACLE.ActiveConnection = cn_Ser 'ADOConnection
'''
'''        AdoCmd_ORACLE.CommandType = adCmdText
'''        AdoCmd_ORACLE.CommandText = SQL
'''
'''        If Send_State = "1" Then
'''            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "1")
'''        ElseIf Send_State = "2" Then
'''            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "2")
'''        ElseIf Send_State = "2" Then
'''            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
'''        End If
'''        AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, gUserID & "")
'''        AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
'''        AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(lsSpecNo))
'''        AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("PID", adVarChar, , 8, Trim(lsPid))
'''        AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
'''        AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_STAT", adVarChar, , 5, "2")
'''
'''        AdoCmd_ORACLE.Execute sqlRet
'''        Set AdoCmd_ORACLE = Nothing
'''        '##### 바인딩 수정 - 19 ##############################################
'''
'''        If sqlRet < 0 Then
'''            SaveQuery SQL
'''            cn_Ser.RollbackTrans
'''            Exit Function
'''        End If
'''        '/------------------------------------
'''        'db_Commit gServer
'''        cn_Ser.CommitTrans
'''        Insert_Data = 1
'''    End With
'''
'''End Function


'//////////////결과 저장 바꿈 (2011.10.11) - 효준
Function Insert_Data(ByVal argSpcRow As Integer) As Integer
    Dim iRow            As Integer
    Dim i               As Integer
    Dim j               As Integer
    Dim lsID            As String
    Dim lsSpecNo        As String
    Dim lsPid           As String
    Dim sResult         As String
    Dim sCnt            As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim ExamCnt         As String
    Dim ExamCode_Spec   As String
    Dim ExamCode_Remark As String
    
    Dim State_GM        As String       '//// 그룹/멀티 코드
    Dim State_cnt       As Integer      '//// 그룹/멀티 코드 쪽 변수
    Dim State_G         As String       '//// 그룹코드
    Dim State_M         As String       '//// 멀티코드
    Dim State_B         As String       '//// 배터리코드
    
    Dim Send_State      As String
    Dim SQL_LOCAL       As String
    Dim CODE_L8_cnt     As String
    
    Dim sqlRet          As Integer
    

    With frmInterface
        gComment_All = ""
        Insert_Data = -1
        ExamCode_Spec = ""
        ExamCode_Remark = ""
        CODE_L8_cnt = ""
        State_GM = ""
        State_cnt = 0
        State_G = ""
        State_M = ""
        gState_G = ""
        gState_M = ""
        gState_B = ""
        lsID = ""
        lsID = Trim(GetText(.vasID, argSpcRow, colBarcode))
        lsSpecNo = Trim(GetText(.vasID, argSpcRow, colSpecNo))
        lsPid = Trim(GetText(.vasID, argSpcRow, colPID))

        'Local에서 환자별로 결과값 가져오기
        ClearSpread .vasTemp

        SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag, PSEX " & vbCrLf & _
              " From pat_res " & vbCrLf & _
              " Where equipno = '" & gEquip & "' " & vbCrLf & _
              " And examdate = '" & Format(CDate(.dtpToday.Value), "yyyymmdd") & "'  " & vbCrLf & _
              " And barcode = '" & Trim(GetText(.vasID, argSpcRow, colBarcode)) & "' " & vbCrLf & _
              " And diskno = '" & Trim(GetText(.vasID, argSpcRow, colRack)) & "' " & vbCrLf & _
              " And posno = '" & Trim(GetText(.vasID, argSpcRow, colPos)) & "' "
        res = db_select_Vas(gLocal, SQL, .vasTemp)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
        
        For i = 1 To frmInterface.vasTemp.DataRowCnt    '/// 실제 검사한 검사코드들
            If ExamCode_Spec <> "" Then
                ExamCode_Spec = ExamCode_Spec & ",'" & Trim(GetText(frmInterface.vasTemp, i, 2)) & "'"
            Else
                ExamCode_Spec = "'" & Trim(GetText(frmInterface.vasTemp, i, 2)) & "'"
            End If
        Next i
        
        If ExamCode_Spec = "" Then: ExamCode_Spec = "''"
        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1

        sCnt = ""
        sResult1 = ""
        sResult2 = ""
        
        
        SQL = "SELECT COUNT(A.EXMN_CD), B.BLPS_ID FROM SPSLHRRST A, SPSLMJBDI B "
        SQL = SQL & vbCrLf & " WHERE A.SPCM_NO = B.SPCM_NO "                                             '검체번호"
        SQL = SQL & vbCrLf & "   AND A.SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
        SQL = SQL & vbCrLf & "   AND A.EXMN_CD IN (" & gAllExam & ") "                                              '검사코드"
        'SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
        SQL = SQL & vbCrLf & " GROUP BY  B.BLPS_ID "
        res = db_select_Col(gServer, SQL)
        
        If Val(gReadBuf(0)) = "0" Then Exit Function
        .lblUser.Caption = gReadBuf(1)

        
        
        cn_Ser.BeginTrans
        '서버로 결과값 저장하기
        For iRow = 1 To .vasTemp.DataRowCnt

            sResult1 = Trim(GetText(.vasTemp, iRow, 4))
            sResult2 = Trim(GetText(.vasTemp, iRow, 3))
            
            If sResult1 <> "" And Mid(sResult1, 1, 3) <> "-99" And Trim(GetText(.vasTemp, iRow, 2)) <> "" And lsSpecNo <> "" Then
                gComment_Code = ""
            
            
                SQL = "SELECT RSLT_NO FROM SPSLHRRST "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
                res = db_select_Col(gServer, SQL)
                 
                If gReadBuf(0) = "" Then: gReadBuf(0) = "0"
                
                sCnt = CLng(gReadBuf(0)) + 1

'                               SQL = "UPDATE SPSLHRRST "
'                SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & sResult1 & "', "                                          '결과(장비결과)
'                SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & sResult2 & "', "                                          '결과(수정결과)"
'                SQL = SQL & vbCrLf & "       DTRM_DVSN = '" & Trim(GetText(.vasTemp, iRow, 5)) & "', "                  'HL 체크"
'                SQL = SQL & vbCrLf & "       PANC_YN = '" & Trim(GetText(.vasTemp, iRow, 6)) & "', "                    'Delta 체크"
'                SQL = SQL & vbCrLf & "       DLTA_YN = '" & Trim(GetText(.vasTemp, iRow, 7)) & "', "                    'Panic 체크"
'                SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
'                SQL = SQL & vbCrLf & "       AMEN_ID = '" & .lblUser.Caption & "', "                                      '결과수정자
'                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
'                SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '결과번호 (결과 넣을시에 증가)
'
'
'                Send_State = "3" '/  <---------- 혈액학장비가 아니라서 상태가 3로만 들어감
'
'                '/----------------------------- 결과 상태 넣기
'                If Send_State = "1" Then
'
'                    SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & .lblUser.Caption & "', "                                 '결과입력자"
'                    SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'                    SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "
'                ElseIf Send_State = "2" Then
'
'                    SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & .lblUser.Caption & "', "                                 '결과입력자"
'                    SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'                    SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & .lblUser.Caption & "', "                                  '중간보고자"
'                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
'                    SQL = SQL & vbCrLf & "       RSLT_STAT = '2' "
'                ElseIf Send_State = "3" Then
'
'                    SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & .lblUser.Caption & "', "                                 '결과입력자"
'                    SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'                    SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = '" & .lblUser.Caption & "', "                                 '중간보고자"
'                    SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
'                    SQL = SQL & vbCrLf & "       LAST_RPTR_ID = '" & .lblUser.Caption & "', "                                 '최종보고자"
'                    SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
'                    SQL = SQL & vbCrLf & "       RSLT_STAT = '3' "
'                End If
'
'                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
'                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "                     '검사코드"
'                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
'                SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
'                res = SendQuery(gServer, SQL)
'                If res < 0 Then
'                    SaveQuery SQL
'                   ' db_RollBack gServer
'                   cn_Ser.RollbackTrans
'                    Exit Function
'                End If
                

                Send_State = "3"
                
                SQL = "UPDATE SPSLHRRST "
                SQL = SQL & vbCrLf & "   SET REAL_RSLT = ?, "                                          '결과(장비결과)
                SQL = SQL & vbCrLf & "       VIEW_RSLT = ?, "                                          '결과(수정결과)"
                SQL = SQL & vbCrLf & "       DTRM_DVSN = ?, "                  'HL 체크"
                SQL = SQL & vbCrLf & "       PANC_YN = ?, "                    'Delta 체크"
                SQL = SQL & vbCrLf & "       DLTA_YN = ?, "                    'Panic 체크"
                SQL = SQL & vbCrLf & "       EXMN_EQPM = ?, "                                        '장비코드
                SQL = SQL & vbCrLf & "       AMEN_ID = ?, "                                      '결과수정자
                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                SQL = SQL & vbCrLf & "       RSLT_NO = ?, "                                                '결과번호 (결과 넣을시에 증가)
                SQL = SQL & vbCrLf & "       RSLT_INPS_ID = ?, "                                 '결과입력자"
                SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = ?, "                                     '중간보고자"
                SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
                SQL = SQL & vbCrLf & "       LAST_RPTR_ID = ?, "                                 '최종보고자"
                SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
                SQL = SQL & vbCrLf & "       RSLT_STAT = ? "
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = ? "                                             '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = ? "                     '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = ? "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT < ? "                                                          '결과상태"
                
                Set AdoCmd_ORACLE = New ADODB.Command
                Set AdoCmd_ORACLE.ActiveConnection = cn_Ser 'ADOConnection
                AdoCmd_ORACLE.CommandType = adCmdText
                AdoCmd_ORACLE.CommandText = SQL
                
                
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("REAL_RSLT", adVarChar, , 20, sResult1)
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("VIEW_RSLT", adVarChar, , 20, sResult2)
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("DTRM_DVSN", adVarChar, , 20, Trim(GetText(.vasTemp, iRow, 5)))
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("PANC_YN", adVarChar, , 20, Trim(GetText(.vasTemp, iRow, 6)))
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("DLTA_YN", adVarChar, , 20, Trim(GetText(.vasTemp, iRow, 7)))
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_EQPM", adVarChar, , 20, gEquipCode)
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_NO", adVarChar, , 20, sCnt)
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPS_ID", adVarChar, , 20, .lblUser.Caption)
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTR_ID", adVarChar, , 20, .lblUser.Caption)
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("LAST_RPTR_ID", adVarChar, , 20, .lblUser.Caption)
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 20, Send_State)
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 20, lsSpecNo)
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 20, Trim(GetText(.vasTemp, iRow, 2)))
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("PID", adVarChar, , 20, lsPid)
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 20, "2")
                
                AdoCmd_ORACLE.Execute sqlRet
                Set AdoCmd_ORACLE = Nothing
                
                State_GM = RsltState_Check(lsSpecNo, Trim(GetText(.vasTemp, iRow, 2)))
                
                State_cnt = InStr(1, State_GM, "/")
                State_G = Mid(State_GM, 1, State_cnt - 1)
                State_GM = Mid(State_GM, State_cnt + 1)
                State_cnt = InStr(1, State_GM, "/")
                State_M = Mid(State_GM, 1, State_cnt - 1)
                State_B = Mid(State_GM, State_cnt + 1)
                
                '/코드 비교시 이용
                If InStr(1, gState_G, State_G) = 0 Then
                    gState_G = gState_G & "," & State_G
                End If
                If InStr(1, gState_M, State_M) = 0 Then
                    gState_M = gState_M & "," & State_M
                End If
                If InStr(1, gState_B, State_B) = 0 Then
                    gState_B = gState_B & "," & State_B
                End If
                
                If Mid(gState_G, 1, 1) = "," Then gState_G = Mid(gState_G, 2)
                
                If Mid(gState_M, 1, 1) = "," Then gState_M = Mid(gState_M, 2)
                
                If Mid(gState_B, 1, 1) = "," Then gState_B = Mid(gState_B, 2)
                
                
                '/------------------------------------ 결과테이블 그룹코드 상태 업데이트
                If Trim(State_G) <> "" Then
                    SQL = "UPDATE SPSLHRRST "
                    If Send_State = "1" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = ? "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "      ,RSLT_INPT_DT = ? "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "      ,RSLT_STAT = ? "
                            SQL = SQL & vbCrLf & "      ,EXMN_EQPM = ? "                                        '장비코드
                            SQL = SQL & vbCrLf & "      ,AMEN_ID = ? "                                      '결과수정자
                            SQL = SQL & vbCrLf & "      ,UPDT_DT = ? "                                                       '결과수정일시
                            
                            SQL = SQL & vbCrLf & "      ,RSLT_NO = ? "
                        ElseIf Send_State = "2" Then
                            
                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = ? "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "      ,RSLT_INPT_DT = ? "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "      ,MDDL_RPTR_ID = ? "                                 '중간보고자"
                            SQL = SQL & vbCrLf & "      ,MDDL_RPTG_DT = ? "                                                  '중간보고일시"
                            SQL = SQL & vbCrLf & "      ,RSLT_STAT = ? "
                            SQL = SQL & vbCrLf & "      ,EXMN_EQPM = ? "                                        '장비코드
                            SQL = SQL & vbCrLf & "      ,AMEN_ID = ? "                                      '결과수정자
                            SQL = SQL & vbCrLf & "      ,UPDT_DT = ? "                                                       '결과수정일시
                            
                            SQL = SQL & vbCrLf & "      ,RSLT_NO = ? "
                        ElseIf Send_State = "3" Then
                            
                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = ? "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "      ,RSLT_INPT_DT = ? "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "      ,MDDL_RPTR_ID = ? "                                 '중간보고자"
                            SQL = SQL & vbCrLf & "      ,MDDL_RPTG_DT = ? "                                                  '중간보고일시"
                            SQL = SQL & vbCrLf & "      ,LAST_RPTR_ID = ? "                                 '최종보고자"
                            SQL = SQL & vbCrLf & "      ,LAST_RPTG_DT = ? "                                                  '최종보고일시"
                            SQL = SQL & vbCrLf & "      ,RSLT_STAT = ? "
                            SQL = SQL & vbCrLf & "      ,EXMN_EQPM = ? "                                        '장비코드
                            SQL = SQL & vbCrLf & "      ,AMEN_ID = ? "                                      '결과수정자
                            SQL = SQL & vbCrLf & "      ,UPDT_DT = ? "                                                       '결과수정일시
                            
                            SQL = SQL & vbCrLf & "      ,RSLT_NO = ? "
                        End If
                        
                        SQL = SQL & vbCrLf & " WHERE SPCM_NO = ? "                                             '검체번호"
                        SQL = SQL & vbCrLf & "   AND EXMN_CD = ? "                                        '검사코드"
                        SQL = SQL & vbCrLf & "   AND PID = ? "                                                    '환자번호"
                        SQL = SQL & vbCrLf & "   AND RSLT_STAT < ? "                                                          '결과상태"

                        Set AdoCmd_ORACLE = New ADODB.Command
                        Set AdoCmd_ORACLE.ActiveConnection = cn_Ser 'ADOConnection
                        
                        AdoCmd_ORACLE.CommandType = adCmdText
                        AdoCmd_ORACLE.CommandText = SQL
                        
                        '-- 시스템 날짜 가져오기 함수 : gsDBDateTime
                        
                        If Send_State = "1" Then
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPS_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPT_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "1")
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_EQPM", adVarChar, , 10, gEquipCode)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
                            
                        ElseIf Send_State = "2" Then
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPS_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPT_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTR_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTG_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "2")
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_EQPM", adVarChar, , 10, gEquipCode)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
                            
                        ElseIf Send_State = "3" Then

                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPS_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPT_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTR_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTG_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("LAST_RPTR_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("LAST_RPTG_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_EQPM", adVarChar, , 10, gEquipCode)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
                            
                        End If
                            
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_NO", adVarChar, , 1, sCnt)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(lsSpecNo))
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_G))
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("PID", adVarChar, , 15, lsPid)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "2")
                        
                        AdoCmd_ORACLE.Execute sqlRet
                        Set AdoCmd_ORACLE = Nothing
                End If
                '/------------------------------------
                
                '/------------------------------------ 결과테이블 멀티코드 상태 업데이트
                If Trim(State_M) <> "" Then
                    SQL = "UPDATE SPSLHRRST "
                    If Send_State = "1" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = ? "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "      ,RSLT_INPT_DT = ? "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "      ,RSLT_STAT = ? "
                            SQL = SQL & vbCrLf & "      ,EXMN_EQPM = ? "                                        '장비코드
                            SQL = SQL & vbCrLf & "      ,AMEN_ID = ? "                                      '결과수정자
                            SQL = SQL & vbCrLf & "      ,UPDT_DT = ? "                                                       '결과수정일시
                            
                            SQL = SQL & vbCrLf & "      ,RSLT_NO = ? "
                        ElseIf Send_State = "2" Then
                            
                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = ? "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "      ,RSLT_INPT_DT = ? "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "      ,MDDL_RPTR_ID = ? "                                 '중간보고자"
                            SQL = SQL & vbCrLf & "      ,MDDL_RPTG_DT = ? "                                                  '중간보고일시"
                            SQL = SQL & vbCrLf & "      ,RSLT_STAT = ? "
                            SQL = SQL & vbCrLf & "      ,EXMN_EQPM = ? "                                        '장비코드
                            SQL = SQL & vbCrLf & "      ,AMEN_ID = ? "                                      '결과수정자
                            SQL = SQL & vbCrLf & "      ,UPDT_DT = ? "                                                       '결과수정일시
                            
                            SQL = SQL & vbCrLf & "      ,RSLT_NO = ? "
                        ElseIf Send_State = "3" Then
                            
                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = ? "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "      ,RSLT_INPT_DT = ? "                                                  '결과입력일시"
                            SQL = SQL & vbCrLf & "      ,MDDL_RPTR_ID = ? "                                 '중간보고자"
                            SQL = SQL & vbCrLf & "      ,MDDL_RPTG_DT = ? "                                                  '중간보고일시"
                            SQL = SQL & vbCrLf & "      ,LAST_RPTR_ID = ? "                                 '최종보고자"
                            SQL = SQL & vbCrLf & "      ,LAST_RPTG_DT = ? "                                                  '최종보고일시"
                            SQL = SQL & vbCrLf & "      ,RSLT_STAT = ? "
                            SQL = SQL & vbCrLf & "      ,EXMN_EQPM = ? "                                        '장비코드
                            SQL = SQL & vbCrLf & "      ,AMEN_ID = ? "                                      '결과수정자
                            SQL = SQL & vbCrLf & "      ,UPDT_DT = ? "                                                       '결과수정일시
                            
                            SQL = SQL & vbCrLf & "      ,RSLT_NO = ? "
                        End If
                        
                        SQL = SQL & vbCrLf & " WHERE SPCM_NO = ? "                                             '검체번호"
                        SQL = SQL & vbCrLf & "   AND EXMN_CD = ? "                                        '검사코드"
                        SQL = SQL & vbCrLf & "   AND PID = ? "                                                    '환자번호"
                        SQL = SQL & vbCrLf & "   AND RSLT_STAT < ? "                                                          '결과상태"

                        Set AdoCmd_ORACLE = New ADODB.Command
                        Set AdoCmd_ORACLE.ActiveConnection = cn_Ser 'ADOConnection
                        
                        AdoCmd_ORACLE.CommandType = adCmdText
                        AdoCmd_ORACLE.CommandText = SQL
                        
                        '-- 시스템 날짜 가져오기 함수 : gsDBDateTime
                        
                        If Send_State = "1" Then
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPS_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPT_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "1")
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_EQPM", adVarChar, , 10, gEquipCode)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
                            
                        ElseIf Send_State = "2" Then
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPS_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPT_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTR_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTG_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "2")
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_EQPM", adVarChar, , 10, gEquipCode)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
                            
                        ElseIf Send_State = "3" Then

                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPS_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPT_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTR_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTG_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("LAST_RPTR_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("LAST_RPTG_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_EQPM", adVarChar, , 10, gEquipCode)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
                            
                        End If
                            
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_NO", adVarChar, , 1, sCnt)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(lsSpecNo))
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_M))
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("PID", adVarChar, , 15, lsPid)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "2")
                            
                        AdoCmd_ORACLE.Execute sqlRet
                        Set AdoCmd_ORACLE = Nothing
                End If
            '/------------------------------------
            
            '/------------------------------------ 결과테이블 그룹코드 상태 업데이트
                If Trim(State_B) <> "" Then
                    SQL = "UPDATE SPSLHRRST "
                    If Send_State = "1" Then

                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = ? "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "      ,RSLT_INPT_DT = ? "                                 '결과입력일시"
                            SQL = SQL & vbCrLf & "      ,RSLT_STAT = ? "
                            SQL = SQL & vbCrLf & "      ,EXMN_EQPM = ? "                                    '장비코드
                            SQL = SQL & vbCrLf & "      ,AMEN_ID = ? "                                      '결과수정자
                            SQL = SQL & vbCrLf & "      ,UPDT_DT = ? "                                      '결과수정일시
                            
                            SQL = SQL & vbCrLf & "      ,RSLT_NO = ? "
                        ElseIf Send_State = "2" Then
                            
                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = ? "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "      ,RSLT_INPT_DT = ? "                                 '결과입력일시"
                            SQL = SQL & vbCrLf & "      ,MDDL_RPTR_ID = ? "                                 '중간보고자"
                            SQL = SQL & vbCrLf & "      ,MDDL_RPTG_DT = ? "                                 '중간보고일시"
                            SQL = SQL & vbCrLf & "      ,RSLT_STAT = ? "
                            SQL = SQL & vbCrLf & "      ,EXMN_EQPM = ? "                                    '장비코드
                            SQL = SQL & vbCrLf & "      ,AMEN_ID = ? "                                      '결과수정자
                            SQL = SQL & vbCrLf & "      ,UPDT_DT = ? "                                      '결과수정일시
                            
                            SQL = SQL & vbCrLf & "      ,RSLT_NO = ? "
                        ElseIf Send_State = "3" Then
                            
                            SQL = SQL & vbCrLf & " SET   RSLT_INPS_ID = ? "                                 '결과입력자"
                            SQL = SQL & vbCrLf & "      ,RSLT_INPT_DT = ? "                                 '결과입력일시"
                            SQL = SQL & vbCrLf & "      ,MDDL_RPTR_ID = ? "                                 '중간보고자"
                            SQL = SQL & vbCrLf & "      ,MDDL_RPTG_DT = ? "                                 '중간보고일시"
                            SQL = SQL & vbCrLf & "      ,LAST_RPTR_ID = ? "                                 '최종보고자"
                            SQL = SQL & vbCrLf & "      ,LAST_RPTG_DT = ? "                                 '최종보고일시"
                            SQL = SQL & vbCrLf & "      ,RSLT_STAT = ? "
                            SQL = SQL & vbCrLf & "      ,EXMN_EQPM = ? "                                    '장비코드
                            SQL = SQL & vbCrLf & "      ,AMEN_ID = ? "                                      '결과수정자
                            SQL = SQL & vbCrLf & "      ,UPDT_DT = ? "                                      '결과수정일시
                            
                            SQL = SQL & vbCrLf & "      ,RSLT_NO = ? "
                        End If
                        
                        SQL = SQL & vbCrLf & " WHERE SPCM_NO = ? "                                          '검체번호"
                        SQL = SQL & vbCrLf & "   AND EXMN_CD = ? "                                          '검사코드"
                        SQL = SQL & vbCrLf & "   AND PID = ? "                                              '환자번호"
                        SQL = SQL & vbCrLf & "   AND RSLT_STAT < ? "                                        '결과상태"

                        Set AdoCmd_ORACLE = New ADODB.Command
                        Set AdoCmd_ORACLE.ActiveConnection = cn_Ser                                         'ADOConnection
                        
                        AdoCmd_ORACLE.CommandType = adCmdText
                        AdoCmd_ORACLE.CommandText = SQL
                        
                        '-- 시스템 날짜 가져오기 함수 : gsDBDateTime
                        
                        If Send_State = "1" Then
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPS_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPT_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "1")
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_EQPM", adVarChar, , 10, gEquipCode)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
                            
                        ElseIf Send_State = "2" Then
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPS_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPT_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTR_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTG_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "2")
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_EQPM", adVarChar, , 10, gEquipCode)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
                            
                        ElseIf Send_State = "3" Then

                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPS_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_INPT_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTR_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTG_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("LAST_RPTR_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("LAST_RPTG_DT", adDBDate, , , gsDBDateTime)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_EQPM", adVarChar, , 10, gEquipCode)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
                            
                        End If
                            
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_NO", adVarChar, , 1, sCnt)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(lsSpecNo))
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_B))
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("PID", adVarChar, , 15, lsPid)
                            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "2")
                        
                        
                        
                        AdoCmd_ORACLE.Execute sqlRet
                        Set AdoCmd_ORACLE = Nothing
                End If
            '/------------------------------------
            
            '/------------------------------------ 접수테이블 STATE 업데이트
                If Send_State = "" Then cn_Ser.RollbackTrans: Exit Function
                
                If Send_State = "1" Then
                    SQL = "UPDATE SPSLMJBDI SET RSLT_STAT = ?, AMEN_ID = ?, UPDT_DT = ? " & vbCrLf
                ElseIf Send_State = "2" Then
                    SQL = "UPDATE SPSLMJBDI SET RSLT_STAT = ?, AMEN_ID = ?, UPDT_DT = ?, MDDL_RPTG_DT = ? " & vbCrLf
                ElseIf Send_State = "3" Then
                    SQL = "UPDATE SPSLMJBDI SET RSLT_STAT = ?, AMEN_ID = ?, UPDT_DT = ?, MDDL_RPTG_DT = ?, LAST_RPTG_DT = ? " & vbCrLf
                End If
                SQL = SQL & " WHERE SPCM_NO = ? " & vbCrLf
                SQL = SQL & "   AND RSLT_STAT <> ? " & vbCrLf
                SQL = SQL & "   AND SPCM_STAT = ? " & vbCrLf
                SQL = SQL & "   AND EXMN_CD IN (?,?,?,?) "
                
                Set AdoCmd_ORACLE = New ADODB.Command
                Set AdoCmd_ORACLE.ActiveConnection = cn_Ser 'ADOConnection
                
                AdoCmd_ORACLE.CommandType = adCmdText
                AdoCmd_ORACLE.CommandText = SQL
                
                '-- 시스템 날짜 가져오기 함수 : gsDBDateTime
                
                If Send_State = "1" Then
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "1")
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(lsSpecNo))
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_STAT", adVarChar, , 5, "2")
                    
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_G))
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_M))
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_B))
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(GetText(.vasTemp, iRow, 2)))
                    
                ElseIf Send_State = "2" Then
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 15, "2")
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTG_DT", adDBDate, , 5, gsDBDateTime)
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(lsSpecNo))
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_STAT", adVarChar, , 5, "2")
                    
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_G))
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_M))
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_B))
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(GetText(.vasTemp, iRow, 2)))

                ElseIf Send_State = "3" Then
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTG_DT", adDBDate, , , gsDBDateTime)
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("LAST_RPTG_DT", adDBDate, , , gsDBDateTime)
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(lsSpecNo))
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_STAT", adVarChar, , 5, "2")
                    
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_G))
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_M))
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_B))
                    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(GetText(.vasTemp, iRow, 2)))
                End If
                
                AdoCmd_ORACLE.Execute sqlRet
                Set AdoCmd_ORACLE = Nothing
                '##### 바인딩 수정 - 7 ##############################################
     
                If sqlRet < 0 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
            
            '/------------------------------------
            End If
        Next iRow
        
        '/------------------------------------ 접수테이블 STATE 업데이트
         If Send_State = "" Then cn_Ser.RollbackTrans: Exit Function
                
         Dim MDMD_EXAMCODE   As String
         Dim MDMD_LIST
         Dim sParam_string   As String
         
         MDMD_EXAMCODE = ""
         MDMD_LIST = Split(gState_G, ",")
         For i = 0 To UBound(MDMD_LIST) - 1
             MDMD_EXAMCODE = MDMD_EXAMCODE & ",'" & MDMD_LIST(i) & "'"
         Next i
         MDMD_LIST = Split(gState_M, ",")
         For i = 0 To UBound(MDMD_LIST) - 1
             MDMD_EXAMCODE = MDMD_EXAMCODE & ",'" & MDMD_LIST(i) & "'"
         Next i
         MDMD_LIST = Split(gState_B, ",")
         For i = 0 To UBound(MDMD_LIST) - 1
             MDMD_EXAMCODE = MDMD_EXAMCODE & ",'" & MDMD_LIST(i) & "'"
         Next i
         
         MDMD_EXAMCODE = MDMD_EXAMCODE & "," & ExamCode_Spec
         MDMD_EXAMCODE = Mid(MDMD_EXAMCODE, 2)
         MDMD_LIST = Split(Replace(MDMD_EXAMCODE, "'", ""), ",")
         
         For i = 0 To UBound(MDMD_LIST)
             If sParam_string <> "" Then
                 sParam_string = sParam_string & ",?"
             Else
                 sParam_string = ",?"
             End If
         Next i
         
         sParam_string = Mid(sParam_string, 2)
         
        If Send_State = "1" Then
            SQL = "UPDATE SPSLMJBDI SET RSLT_STAT = ?, AMEN_ID = ?, UPDT_DT = ? " & vbCrLf
        ElseIf Send_State = "2" Then
            SQL = "UPDATE SPSLMJBDI SET RSLT_STAT = ?, AMEN_ID = ?, UPDT_DT = ?, MDDL_RPTG_DT = ? " & vbCrLf
        ElseIf Send_State = "3" Then
            SQL = "UPDATE SPSLMJBDI SET RSLT_STAT = ?, AMEN_ID = ?, UPDT_DT = ?, MDDL_RPTG_DT = ?, LAST_RPTG_DT = ? " & vbCrLf
        End If
        SQL = SQL & " WHERE SPCM_NO = ? " & vbCrLf
        SQL = SQL & "   AND RSLT_STAT <> ? " & vbCrLf
        SQL = SQL & "   AND SPCM_STAT = ? " & vbCrLf
        SQL = SQL & "   AND EXMN_CD IN (" & sParam_string & ")"
        
        Set AdoCmd_ORACLE = New ADODB.Command
        Set AdoCmd_ORACLE.ActiveConnection = cn_Ser 'ADOConnection
        
        AdoCmd_ORACLE.CommandType = adCmdText
        AdoCmd_ORACLE.CommandText = SQL
        
        '-- 시스템 날짜 가져오기 함수 : gsDBDateTime
        
        If Send_State = "1" Then
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "1")
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(lsSpecNo))
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_STAT", adVarChar, , 5, "2")
            
        ElseIf Send_State = "2" Then
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 15, "2")
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTG_DT", adDBDate, , 5, gsDBDateTime)
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(lsSpecNo))
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_STAT", adVarChar, , 5, "2")

        ElseIf Send_State = "3" Then
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("UPDT_DT", adDBDate, , , gsDBDateTime)
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("MDDL_RPTG_DT", adDBDate, , , gsDBDateTime)
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("LAST_RPTG_DT", adDBDate, , , gsDBDateTime)
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(lsSpecNo))
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_STAT", adVarChar, , 5, "2")
            
        End If
             
             For i = 0 To UBound(MDMD_LIST)
                 AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(MDMD_LIST(i)))
             Next i
             
'                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_G))
'                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_M))
'                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(State_B))
'                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("EXMN_CD", adVarChar, , 10, Trim(GetText(.vasTemp, iRow, 2)))
        
        AdoCmd_ORACLE.Execute sqlRet
        Set AdoCmd_ORACLE = Nothing
        '##### 바인딩 수정 - 7 ##############################################

        If sqlRet < 0 Then
            SaveQuery SQL
            cn_Ser.RollbackTrans
            Exit Function
        End If
        '/------------------------------------

''        '/------------------------------------ 처방테이블 STATE 업데이트
''
''        '///////// 처방테이블
''        SQL = "UPDATE SPSLMJBBI "
''        If Send_State = "1" Then
''            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1', "
''            SQL = SQL & vbCrLf & "       AMEN_ID = '" & .lblUser.Caption & "', "
''            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
''        ElseIf Send_State = "2" Then
''            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '2', "
''            SQL = SQL & vbCrLf & "       AMEN_ID = '" & .lblUser.Caption & "', "
''            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
''        ElseIf Send_State = "3" Then
''            SQL = SQL & vbCrLf & "   SET RSLT_STAT = '3', "
''            SQL = SQL & vbCrLf & "       AMEN_ID = '" & .lblUser.Caption & "', "
''            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
''        End If
''        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
''        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
''        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "
''        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2' "
''        res = SendQuery(gServer, SQL)
''
''        If res = -1 Then
''            SaveQuery SQL
''            cn_Ser.RollbackTrans
''            Exit Function
''
''        End If
''        '/------------------------------------

        '/------------------------------------ 처방테이블 STATE 업데이트
            SQL = "UPDATE SPSLMJBBI SET RSLT_STAT = ?, AMEN_ID = ?, UPDT_DT = SYSDATE " & vbCrLf
            SQL = SQL & " WHERE SPCM_NO = ? " & vbCrLf
            SQL = SQL & "   AND RSLT_STAT < ? " & vbCrLf
            SQL = SQL & "   AND SPCM_STAT = ? " & vbCrLf
            SQL = SQL & "   AND PID = ? "
            
            Set AdoCmd_ORACLE = New ADODB.Command
            Set AdoCmd_ORACLE.ActiveConnection = cn_Ser 'ADOConnection
            AdoCmd_ORACLE.CommandType = adCmdText
            AdoCmd_ORACLE.CommandText = SQL
            
            If Send_State = "1" Then
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "1")
            ElseIf Send_State = "2" Then
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "2")
            ElseIf Send_State = "3" Then
                AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
            End If
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("AMEN_ID", adVarChar, , 20, .lblUser.Caption)
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(lsSpecNo))
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_STAT", adVarChar, , 5, "2")
            AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("PID", adVarChar, , 20, lsPid)
            
            AdoCmd_ORACLE.Execute sqlRet
            Set AdoCmd_ORACLE = Nothing
        '/------------------------------------
        
'''''        '/------------------------------------ 처방테이블 업데이트(MDMDHTORD)
'''''        If Send_State = "3" Then
'''''
'''''                           SQL = "UPDATE MDMDHTORD "
'''''            SQL = SQL & vbCrLf & "   SET PRSC_STAT = '51'"      '/50 예비보고, 51 최종보고
'''''            SQL = SQL & vbCrLf & "     , RPTG_DT = SYSDATE"
'''''            SQL = SQL & vbCrLf & "     , AMEN_ID = '" & .lblUser.Caption & "'"
'''''            SQL = SQL & vbCrLf & " WHERE (PRSC_SQNO, PRSC_CD) "
'''''            SQL = SQL & vbCrLf & "       IN (SELECT PRSC_SQNO, EXMN_CD "
'''''            SQL = SQL & vbCrLf & "             FROM SPSLMJBDI "
'''''            SQL = SQL & vbCrLf & "            WHERE SPCM_NO = '" & lsSpecNo & "' "
'''''            SQL = SQL & vbCrLf & "              AND EXMN_CD IN (" & Trim(MDMD_EXAMCODE) & ") "                     '검사코드"
'''''            SQL = SQL & vbCrLf & "              AND SPCM_STAT = '2') "
'''''            SQL = SQL & vbCrLf & "AND DC_DVSN = 'O' "
'''''
'''''            Save_Raw_Data "[처방업데이트]" & SQL
'''''
'''''            res = SendQuery(gServer, SQL)
'''''            If res = -1 Then
'''''                SaveQuery SQL
'''''                cn_Ser.RollbackTrans
'''''                Exit Function
'''''            End If
'''''        End If
        
                  SQL = "UPDATE MDMDHTORD "
            SQL = SQL & vbCrLf & "   SET PRSC_STAT = '51'"      '/50 예비보고, 51 최종보고
            SQL = SQL & vbCrLf & "     , RPTG_DT = SYSDATE"
            SQL = SQL & vbCrLf & "     , AMEN_ID = '" & .lblUser.Caption & "'"
            SQL = SQL & vbCrLf & "     , updt_DT = SYSDATE"
            SQL = SQL & vbCrLf & " WHERE PRSC_SQNO = "
            SQL = SQL & "                   (SELECT PRSC_SQNO "
            SQL = SQL & vbCrLf & "             FROM SPSLMJBDI "
            SQL = SQL & vbCrLf & "            WHERE BRCD_LABL_NO = '" & lsID & "') "
'            SQL = SQL & vbCrLf & "            WHERE SPCM_NO = '" & lsSpecNo & "' "
'            SQL = SQL & vbCrLf & "              AND EXMN_CD IN (" & Trim(MDMD_EXAMCODE) & ") "                     '검사코드"
'            SQL = SQL & vbCrLf & "              AND SPCM_STAT = '2') "
            SQL = SQL & vbCrLf & "AND DC_DVSN = 'O' "
            
            Save_Raw_Data "[처방업데이트]" & SQL
            
            res = SendQuery(gServer, SQL)
            If res = -1 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
        
        '/------------------------------------ 처방테이블 업데이트
        
        
        'db_Commit gServer
        cn_Ser.CommitTrans
        Insert_Data = 1
    End With
End Function


Public Function gsDBDateTime() As Date
Dim sRs     As ADODB.Recordset
Dim strSQL  As String

    Set sRs = New ADODB.Recordset
    
    strSQL = "select sysdate from dual"
    sRs.Open strSQL, cn_Ser, adOpenStatic, adLockReadOnly
    
    If Not sRs.EOF Then
        gsDBDateTime = sRs("SYSDATE")
    Else
        gsDBDateTime = Now
    End If
    sRs.Close
    Set sRs = Nothing

End Function


Function Insert_Data_R(ByVal argSpcRow As Long, Optional asSend As Integer = 0) As Integer
'서버의 데이타 베이스에 저장
    Dim iRow            As Integer
    Dim i               As Integer
    Dim j               As Integer
    Dim lsID            As String
    Dim lsSpecNo        As String
    Dim lsPid           As String
    Dim sResult         As String
    Dim lsInsertTime    As String
    Dim sCnt            As String
    

    Insert_Data_R = -1
    
    lsID = ""
    lsID = Trim(GetText(frmInterface.vasRID, argSpcRow, colBarcode))
    lsSpecNo = Trim(GetText(frmInterface.vasRID, argSpcRow, colSpecNo))
    lsPid = Trim(GetText(frmInterface.vasRID, argSpcRow, colPID))
    lsInsertTime = Trim(Format(GetDateFull, "mm")) & "/" & Trim(Format(GetDateFull, "dd")) & "/" & Trim(Format(GetDateFull, "yyyy")) & " " & Trim(Format(GetDateFull, "hh:mm:ss"))
    'lsInsertTime = Trim(Format(GetDateFull, "yyyymmddhhmmss"))
    
    
    'Local에서 환자별로 결과값 가져오기
    ClearSpread frmInterface.vasTemp
    
    SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag " & vbCrLf & _
          " From pat_res " & vbCrLf & _
          " Where equipno = '" & gEquip & "' " & vbCrLf & _
          " And examdate = '" & Format(CDate(frmInterface.dtpToday.Value), "yyyymmdd") & "'  " & vbCrLf & _
          " And barcode = '" & Trim(GetText(frmInterface.vasRID, argSpcRow, colBarcode)) & "' " & vbCrLf & _
          " And diskno = '" & Trim(GetText(frmInterface.vasRID, argSpcRow, colRack)) & "' " & vbCrLf & _
          " And posno = '" & Trim(GetText(frmInterface.vasRID, argSpcRow, colPos)) & "' "
    res = db_select_Vas(gLocal, SQL, frmInterface.vasTemp)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    frmInterface.vasTemp.MaxRows = frmInterface.vasTemp.DataRowCnt + 1
    
    gHIVPosFlag = -1
    
    sCnt = ""
    'db_BeginTran gServer
    '서버로 결과값 저장하기
    For iRow = 1 To frmInterface.vasTemp.DataRowCnt
        sCnt = ""
        
        SQL = "SELECT RSLT_NO FROM SPSLHRRST "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
        SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "' "                      '검사코드"
        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
        SQL = SQL & vbCrLf & "   AND RSLT_STAT = '0' "                                                          '결과상태"
        res = db_select_Col(gServer, SQL)
        sCnt = CLng(gReadBuf(0)) + 1
        
        SQL = "UPDATE SPSLHRRST "
        SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & Trim(GetText(frmInterface.vasTemp, iRow, 4)) & "', "                   '결과(장비결과)
        SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & Trim(GetText(frmInterface.vasTemp, iRow, 3)) & "', "                   '결과(수정결과)"
        SQL = SQL & vbCrLf & "       DLTA_YN = 'N', "                                                           'Delta 체크"
        SQL = SQL & vbCrLf & "       PANC_YN = 'N', "                                                           'Panic 체크"
        SQL = SQL & vbCrLf & "       RSLT_INPS_ID = 'test', "                                                   '결과입력자"
        SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
        SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = 'test', "                                                   '중간보고자"
        SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
        SQL = SQL & vbCrLf & "       LAST_RPTR_ID = 'test', "                                                   '최종보고자"
        SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
        SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
        SQL = SQL & vbCrLf & "       AMEN_ID = 'test', "                                                        '결과수정자
        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
        SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '결과번호 (결과 넣을시에 증가)
        SQL = SQL & vbCrLf & "       RSLT_STAT = '3' "                                                          '결과상태"
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
        SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "' "                      '검사코드"
        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
        SQL = SQL & vbCrLf & "   AND RSLT_STAT = '0' "                                                          '결과상태"
        res = SendQuery(gServer, SQL)
        If res < 0 Then
            SaveQuery SQL
           ' db_RollBack gServer
            Exit Function
        End If
        
    Next iRow
    
    
    
    
    SQL = "SELECT EXMN_CD FROM SPSLHRRST "
    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
    SQL = SQL & vbCrLf & "   AND AND EXMN_CD NOT LIKE '%G%' "
    SQL = SQL & vbCrLf & "   AND RSLT_STAT > '0' "
    SQL = SQL & vbCrLf & "   AND VIEW_RSLT IS NOT NULL "
    res = db_select_Vas(gServer, SQL, frmInterface.vasTemp1)
    
    If res = 0 Then                                                                 '///// 결과테이블에 결과가 다 들어가 있는 경우 (그룹코드제외)
        SQL = "Update SPSLMJBBI"
        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '3'"
        SQL = SQL & vbCrLf & "       AMEN_ID = 'test', "
        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT = '0'"
        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
        res = SendQuery(gServer, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
    
        SQL = "Update SPSLMJBDI"
        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '3'"
        SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
        SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
        SQL = SQL & vbCrLf & "       AMEN_ID = 'test', "
        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT = '0'"
        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
        res = SendQuery(gServer, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
        
    ElseIf res = -1 Then                                                             '///// 쿼리 에러인경우
        SaveQuery SQL
        Exit Function
    End If
    
    SQL = ""

       
    'db_Commit gServer
    Insert_Data_R = 1

End Function

'Function Get_Sample_Info_SPCMNO(ByVal asRow As Long) As Integer
'
'    Dim sBarcode As String
'    Dim sSpecNo As String
'    Dim sTestCd As String
'
'    Get_Sample_Info_SPCMNO = -1
'    '환자정보 가져오기
'    sSpecNo = Trim(GetText(frmInterface.vasResult, asRow, colSpecNo))
'    sTestCd = Trim(GetText(frmInterface.vasResult, asRow, colTestCd))
'
'    If sSpecNo = "" Then
'        Exit Function
'    End If
'    '바코드번호로 검체번호 불러오기FN_LABCVTPRTBCNO(SPCM_NO) --> 바코드라벨번호 리턴
'
'    SQL = "SELECT FN_LABCVTPRTBCNO('" & Trim(sSpecNo) & "') FROM DUAL "
'    res = db_select_Col(gServer, SQL)
'    sBarcode = Trim(gReadBuf(0))
'
'    '환자번호, 환자이름, 주민번호, 성별, 나이
'    SQL = "SELECT PID, PT_NM, SEX, AGE "
'    SQL = SQL & vbCrLf & " FROM SPSLMJBBI "
'    SQL = SQL & vbCrLf & "WHERE SPCM_NO = '" & sSpecNo & "' "
'    SQL = SQL & vbCrLf & "  AND SPCM_STAT = '2' "
'    SQL = SQL & vbCrLf & "  AND RSLT_STAT < '2' "
'    res = db_select_Col(gServer, SQL)
'
'    '///////// gAllExam 자리에 검사 코드 넣어줌 세부코드 도 붙어 있는게 B312001 , 02, 03
'
'    If res = 1 Then
'        SetText frmInterface.vasResult, Trim(sSpecNo), asRow, colSpecNo     '2
'        SetText frmInterface.vasResult, Trim(sBarcode), asRow, colBarcode   '3
'        SetText frmInterface.vasResult, Trim(sTestCd), asRow, colTestCd    '4
'        SetText frmInterface.vasResult, Trim(gReadBuf(0)), asRow, colPID    '6
'        SetText frmInterface.vasResult, Trim(gReadBuf(1)), asRow, colPName  '7
'        SetText frmInterface.vasResult, Trim(gReadBuf(2)), asRow, colSex    '8
'        SetText frmInterface.vasResult, Trim(gReadBuf(3)), asRow, colAge    '9
'        Get_Sample_Info_SPCMNO = 1
'    Else
'        Get_Sample_Info_SPCMNO = -1
'    End If
'
'End Function


Function Get_Sample_Info_QC(ByVal asRow As Long) As Integer

    Dim sBarcode As String
    Dim sQCdate  As String
    
    Get_Sample_Info_QC = -1
    '환자정보 가져오기
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBarcode))   '샘플 바코드 번호
    If sBarcode = "" Or IsNumeric(sBarcode) = False Then
        Exit Function
    End If
    
    sQCdate = Trim(Format(GetDateFull, "yyyymmdd"))
    
    '환자번호, 환자이름, 주민번호, 성별, 나이
    SQL = "SELECT SBSN_NO, '정도관리', '', "
    SQL = SQL & vbCrLf & "                 (SELECT MAX(RSLT_SQNO) + 1 FROM SPSLHQRST "
    SQL = SQL & vbCrLf & "                   WHERE EQPM_CD = '" & Mid(sBarcode, 3, 3) & "' "
    SQL = SQL & vbCrLf & "                     AND SBSN_CD = '" & Mid(sBarcode, 6, 3) & "' "
    SQL = SQL & vbCrLf & "                     AND LVL_CD  = '" & Mid(sBarcode, 9, 1) & "' "
    SQL = SQL & vbCrLf & "                     AND EXMN_DY = '" & sQCdate & "' )"
    SQL = SQL & vbCrLf & " FROM SPSLMQMST "
    SQL = SQL & vbCrLf & "WHERE EQPM_CD = '" & Mid(sBarcode, 3, 3) & "' "
    SQL = SQL & vbCrLf & "  AND SBSN_CD = '" & Mid(sBarcode, 6, 3) & "' "
    SQL = SQL & vbCrLf & "  AND LVL_CD = '" & Mid(sBarcode, 9, 1) & "' "
    res = db_select_Col(gServer, SQL)
    
    '///////// gAllExam 자리에 검사 코드 넣어줌 세부코드 도 붙어 있는게 B312001 , 02, 03
    
    If res = 1 Then
        SetText frmInterface.vasID, Trim(sBarcode), asRow, colSpecNo
        SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colPID
        SetText frmInterface.vasID, Trim(gReadBuf(1)), asRow, colPName
        SetText frmInterface.vasID, Trim(gReadBuf(2)), asRow, colSex
        SetText frmInterface.vasID, Trim(gReadBuf(3)), asRow, colAge
        
        SetText frmInterface.vasList, Trim(sBarcode), 1, colSpecNo
        SetText frmInterface.vasList, Trim(gReadBuf(0)), 1, colPID
        SetText frmInterface.vasList, Trim(gReadBuf(1)), 1, colPName
        SetText frmInterface.vasList, Trim(gReadBuf(2)), 1, colSex
        SetText frmInterface.vasList, Trim(gReadBuf(3)), 1, colAge
        
        Get_Sample_Info_QC = 1
    Else
    
        Get_Sample_Info_QC = -1
        Call SaveQuery(SQL)
    End If

End Function

Function Get_Sample_Info(ByVal asRow As Long) As Integer
    
    Dim sBarcode As String
    Dim sSpecNo As String
    Dim sRet    As String
    
On Error GoTo Err

    Get_Sample_Info = -1
    '환자정보 가져오기
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBarcode))   '샘플 바코드 번호
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    '##### 바인딩 수정 - 99 ##############################################
''    '바코드번호로 검체번호 불러오기
''    SQL = "SELECT FN_LABCVTBCNO('" & Trim(sBarcode) & "') FROM DUAL "
''    res = db_select_Col(gServer, SQL)
''    sSpecNo = Trim(gReadBuf(0))
    
    SQL = "SELECT FN_LABCVTBCNO(?) FROM DUAL "
    
    Save_Raw_Data "[Get_Sample_Info]" & SQL
    
    Set AdoCmd_ORACLE = New ADODB.Command
    Set AdoRs_ORACLE = New ADODB.Recordset
    Set AdoCmd_ORACLE.ActiveConnection = cn_Ser 'ADOConnection
    
    AdoCmd_ORACLE.CommandType = adCmdText
    AdoCmd_ORACLE.CommandText = SQL
    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("FN_LABCVTBCNO", adVarChar, , 10, Trim(sBarcode))
    Set AdoRs_ORACLE = New ADODB.Recordset
    AdoRs_ORACLE.Open AdoCmd_ORACLE, , adOpenStatic, adLockBatchOptimistic
    
    sSpecNo = AdoRs_ORACLE.Fields(0) & ""
    
    Set AdoCmd_ORACLE = Nothing
    Set AdoRs_ORACLE = Nothing
    

    '##### 바인딩 수정 - 99 ##############################################
    
    '##### 바인딩 수정 - 98 ##############################################
''    '환자번호, 환자이름, 주민번호, 성별, 나이
''    SQL = "SELECT PID, PT_NM, SEX, AGE "
''    SQL = SQL & vbCrLf & " FROM SPSLMJBBI "
''    SQL = SQL & vbCrLf & "WHERE SPCM_NO = '" & sSpecNo & "' "
''    SQL = SQL & vbCrLf & "  AND SPCM_STAT = '2' "
''    SQL = SQL & vbCrLf & "  AND RSLT_STAT <> '3' "
''    res = db_select_Col(gServer, SQL)
    
    '///////// gAllExam 자리에 검사 코드 넣어줌 세부코드 도 붙어 있는게 B312001 , 02, 03
    
''    If res = 1 Then
''        SetText frmInterface.vasID, Trim(sSpecNo), asRow, colSpecNo     '2
''        SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colPID    '6
''        SetText frmInterface.vasID, Trim(gReadBuf(1)), asRow, colPName  '7
''        SetText frmInterface.vasID, Trim(gReadBuf(2)), asRow, colSex    '8
''        SetText frmInterface.vasID, Trim(gReadBuf(3)), asRow, colAge    '9
''        Get_Sample_Info = 1
''    Else
''        Get_Sample_Info = -1
''    End If

    SQL = "SELECT PID, PT_NM, SEX, AGE "
    SQL = SQL & vbCrLf & " FROM SPSLMJBBI "
    SQL = SQL & vbCrLf & "WHERE SPCM_NO = ? "
    SQL = SQL & vbCrLf & "  AND SPCM_STAT = ? "
'    SQL = SQL & vbCrLf & "  AND RSLT_STAT <> ? "
    
    Save_Raw_Data "[Get_Sample_Info]" & SQL
    
    Set AdoCmd_ORACLE = New ADODB.Command
    Set AdoRs_ORACLE = New ADODB.Recordset
    Set AdoCmd_ORACLE.ActiveConnection = cn_Ser 'ADOConnection
    
    AdoCmd_ORACLE.CommandType = adCmdText
    AdoCmd_ORACLE.CommandText = SQL
    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_NO", adVarChar, , 15, Trim(sSpecNo))
    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("SPCM_STAT", adVarChar, , 5, "2")
'    AdoCmd_ORACLE.Parameters.Append AdoCmd_ORACLE.CreateParameter("RSLT_STAT", adVarChar, , 5, "3")
    Set AdoRs_ORACLE = New ADODB.Recordset
    AdoRs_ORACLE.Open AdoCmd_ORACLE, , adOpenStatic, adLockBatchOptimistic

    If AdoRs_ORACLE.BOF = False Then
        SetText frmInterface.vasID, Trim(sSpecNo), asRow, colSpecNo     '2
        SetText frmInterface.vasID, Trim(AdoRs_ORACLE.Fields(0) & ""), asRow, colPID    '6
        SetText frmInterface.vasID, Trim(AdoRs_ORACLE.Fields(1) & ""), asRow, colPName  '7
        SetText frmInterface.vasID, Trim(AdoRs_ORACLE.Fields(2) & ""), asRow, colSex    '8
        SetText frmInterface.vasID, Trim(AdoRs_ORACLE.Fields(3) & ""), asRow, colAge    '9
        Get_Sample_Info = 1
    Else
        '자동접수
        'wsdl url은 현재 운용중인 아래 주소로 사용하시면 됩니다. http://isis.nhimc:8800/service/PoctService?wsdl
        'PoctService 내의 서비스는 registSpcmRcpn(String sBcno, String sPoctDevModel ) 호출하시면 됩니다.
        'sBcno가 장비에서 넘어가는 검체번호(18로 시작하는 열자리숫자)이며, sPoctDevModel는 일반적으로 장비구분을 위한 장비명text 입니다.
        
'        sRet = Online_XML_Qry(sBarcode)
        
        
        Get_Sample_Info = -1
        
        
    End If

    Set AdoCmd_ORACLE = Nothing
    Set AdoRs_ORACLE = Nothing
    '##### 바인딩 수정 - 98 ##############################################

Exit Function

Err:
    Save_Raw_Data "[Get_Sample_Info]" & Err.Description

    Get_Sample_Info = -1
    
End Function



Public Function Online_XML_Qry(ByVal sBcno As String) As String
    Dim oSOAP   As MSSOAPLib30.SoapClient30
    Dim Send    As String
    Dim sParam  As String
    Dim txtSendXML  As String
    
    On Error GoTo ErrHandle
    
    Online_XML_Qry = ""
    
    Set oSOAP = New MSSOAPLib30.SoapClient30
    oSOAP.ClientProperty("ServerHTTPRequest") = True
    oSOAP.MSSoapInit gServerPath
            
    txtSendXML = "<?xml version='1.0' encoding='UTF-8'?>"
    txtSendXML = txtSendXML & vbCrLf & "<soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'>"
    txtSendXML = txtSendXML & vbCrLf & "<soapenv:Body>"
    txtSendXML = txtSendXML & vbCrLf & "<registSpcmRcpn xmlns='http://svc.poct.ws.nhimc/'>"
    txtSendXML = txtSendXML & vbCrLf & "<arg0 xmlns=''>" & Trim(sBcno) & "</arg0>"
    txtSendXML = txtSendXML & vbCrLf & "<arg1 xmlns=''>" & gEquip & "</arg1>"
    txtSendXML = txtSendXML & vbCrLf & "</registSpcmRcpn>"
    txtSendXML = txtSendXML & vbCrLf & "</soapenv:Body>"
    txtSendXML = txtSendXML & vbCrLf & "</soapenv:Envelope>" & vbCrLf


    Call Save_Raw_Data("[Send SOAP  => " & sBcno & " ]" & gEquip)
    
    Send = oSOAP.registSpcmRcpn(sBcno, gEquip)
    
    
    Call Save_Raw_Data("[Recv SOAP => " & sBcno & " ]" & Send)
    
    Online_XML_Qry = Send
    Set oSOAP = Nothing
    
    DoEvents
    
    Exit Function
    
ErrHandle:
    Online_XML_Qry = ""
    
    Save_Raw_Data "[Online_XML_Qry]" & Err.Description

    If oSOAP.FaultString <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf
    End If
    If Trim(Err.Description) <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[ERROR]" & Err.Description & vbCrLf
    End If


End Function

Function Get_Sample_InfoR(ByVal asRow As Long) As Integer
   Dim sBarcode As String
    Dim sSpecNo As String

    Get_Sample_InfoR = -1
    '환자정보 가져오기
    sBarcode = Trim(GetText(frmInterface.vasRID, asRow, colBarcode))   '샘플 바코드 번호
    If sBarcode = "" Then
        Exit Function
    End If
    '바코드번호로 검체번호 불러오기
    SQL = "SELECT FN_LABCVTBCNO(" & Trim(sBarcode) & ") FROM DUAL "
    res = db_select_Col(gServer, SQL)
    
    sSpecNo = Trim(gReadBuf(0))
    
    '환자번호, 환자이름, 주민번호, 성별, 나이
    SQL = "SELECT PID, PT_NM, SEX, AGE "
    SQL = SQL & vbCrLf & " FROM SPSLMJBBI "
    SQL = SQL & vbCrLf & "WHERE SPCM_NO = '" & sSpecNo & "' "
    SQL = SQL & vbCrLf & "  AND SPCM_STAT = '2' "
    SQL = SQL & vbCrLf & "  AND RSLT_STAT = '0' "
    res = db_select_Col(gServer, SQL)
    
    '///////// gAllExam 자리에 검사 코드 넣어줌 세부코드 도 붙어 있는게 B312001 , 02, 03
    
    If res = 1 Then
        SetText frmInterface.vasID, Trim(sSpecNo), asRow, colSpecNo
        SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colPID
        SetText frmInterface.vasID, Trim(gReadBuf(1)), asRow, colPName
        SetText frmInterface.vasID, Trim(gReadBuf(2)), asRow, colSex
        SetText frmInterface.vasID, Trim(gReadBuf(3)), asRow, colAge
        
        Get_Sample_InfoR = 1
    Else
    
        Get_Sample_InfoR = -1
    End If
End Function

Function EquipExamCode(argEquipCode As String, argPID As String, argSENO As String, argSEQN As String) As String
'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Dim i As Integer
Dim sExamCode As String

    EquipExamCode = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    ClearSpread frmInterface.vasTemp1
    sExamCode = ""
    
    SQL = " Select examcode From EquipExam " & vbCrLf & _
          " Where equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
          " And equipcode = '" & Trim(argEquipCode) & "' "
    res = db_select_Vas(gLocal, SQL, frmInterface.vasTemp1)
    
    If frmInterface.vasTemp1.DataRowCnt < 1 Then
        Exit Function
    End If
    
    For i = 1 To frmInterface.vasTemp1.DataRowCnt
        If sExamCode <> "" Then
            sExamCode = sExamCode & ",'" & Trim(GetText(frmInterface.vasTemp1, i, 1)) & "'"
        Else
            sExamCode = "'" & Trim(GetText(frmInterface.vasTemp1, i, 1)) & "'"
        End If
    Next i

    'SPSLHRRST
    SQL = " Select SUCD From LRESULT " & CR & _
          " Where PAID = '" & Trim(argPID) & "' " & vbCrLf & _
          "   and SENO = " & argSENO & vbCrLf & _
          "   and SEQN = " & argSEQN & vbCrLf & _
          "   and SUCD in ( " & sExamCode & ")  "
          
    res = db_select_Col(gServer, SQL)
  
    If gReadBuf(0) <> "" Then
        EquipExamCode = Trim(gReadBuf(0))
    End If
    
End Function


Function GetEquipExamCode_CA1500(argEquipCode As String, argPID As String) As String
'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Dim i As Integer
Dim sExamCode As String
Dim strExamCode As String
Dim strStatFg  As String
Dim sExamCd As String
Dim strItems As String
Dim strTemp As String
Dim strIntBase As String

    GetEquipExamCode_CA1500 = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
'    argPID = "1558200030"
    
    SQL = "SELECT FN_LABCVTBCNO('" & argPID & "') FROM DUAL"
    res = db_select_Col(gServer, SQL)
    GetEquipExamCode_CA1500 = ""
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            sExamCd = Trim(gReadBuf(i))
        Else
            Exit For
        End If
    Next
    
    SQL = " Select EXMN_CD From SPSLHRRST " & CR & _
          " Where SPCM_NO = '" & Trim(sExamCd) & "' " & vbCrLf & _
          "   and SUBSTR(exmn_cd,1,1) <> 'G'" & _
          "   and RSLT_NO IS NOT NULL"
          
    res = db_select_Row(gServer, SQL)
    strExamCode = ""
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            strExamCode = strExamCode & "'" & Trim(gReadBuf(i)) & "',"
        Else
            Exit For
        End If
    Next
    
    strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
    'EquipExamCode =
    
    ClearSpread frmInterface.vasTemp1
'    sExamCode = ""
    Erase gReadBuf
          SQL = "Select equipcode "
    SQL = SQL & "  From EquipExam "
    SQL = SQL & " Where equipno  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   and examcode in (" & Trim(strExamCode) & ")"
    SQL = SQL & " order by equipcode    "
    res = db_select_Row(gLocal, SQL)
    strExamCode = ""
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            strIntBase = Trim(gReadBuf(i))
            strIntBase = Mid(strIntBase, 1, Len(strIntBase) - 1) & "0" & Space$(6)
            If strIntBase <> strTemp Then
                strExamCode = strExamCode & strIntBase 'Mid(Trim(gReadBuf(i)), 1, Len(Trim(gReadBuf(i))) - 1) & "0" & Space$(6)
                strTemp = strIntBase
            End If

            'strExamCode = strExamCode & Mid(Trim(gReadBuf(i)), 1, Len(Trim(gReadBuf(i))) - 1) & "0" & Space$(6)
        Else
            Exit For
        End If
    Next
    
    '응급유무 (R:Routin, E:Stat)
    'strStatFg = IIf(pAccInfo.StatFg = "1", "E", "U")
    strStatFg = "U"
    
    
'    strExamCode = STX & "S2210101" & strStatFg & Space(6) & Space(4) & mOrder.RackNo & mOrder.TubePos & mOrder.BarNo & _
                "B" & Space(15) & strExamCode & ETX
    
    strExamCode = "" & "S2210101" & strStatFg & Space(6) & Space(4) & mResult.RackNo & mResult.TubePos & mResult.BarNo & _
                "B" & Space(15) & strExamCode & ""
    
    GetEquipExamCode_CA1500 = strExamCode
    
End Function

Function GetOrderExamCode(argEquipCode As String, argPID As String) As String
'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Dim i           As Integer
Dim sExamCode   As String
Dim strExamCode As String
Dim sExamCd     As String

    GetOrderExamCode = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    SQL = "SELECT FN_LABCVTBCNO('" & argPID & "') FROM DUAL"
    res = db_select_Col(gServer, SQL)
    GetOrderExamCode = ""
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            sExamCd = Trim(gReadBuf(i))
        Else
            Exit For
        End If
    Next
    
    '-- 검사코드 가져오기
    SQL = " Select EXMN_CD From SPSLHRRST " & CR & _
          " Where SPCM_NO = '" & Trim(sExamCd) & "' " & vbCrLf & _
          "   and RSLT_NO IS NOT NULL"
          
    res = db_select_Col(gServer, SQL)
    GetOrderExamCode = ""
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            GetOrderExamCode = GetOrderExamCode & "'" & Trim(gReadBuf(i)) & "',"
        Else
            Exit For
        End If
    Next
    
    If GetOrderExamCode <> "" Then
        GetOrderExamCode = Mid(GetOrderExamCode, 1, Len(GetOrderExamCode) - 1)
    End If
    
End Function

Function GetOrderExamCode_New(argEquipCode As String, argPID As String) As String
'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Dim i           As Integer
Dim sExamCode   As String
Dim strExamCode As String
Dim sExamCd     As String
Dim rs_svr As ADODB.Recordset

On Error GoTo Err

    GetOrderExamCode_New = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    SQL = "SELECT FN_LABCVTBCNO('" & argPID & "') FROM DUAL"
    Set rs_svr = cn_Ser.Execute(SQL)
    Do Until rs_svr.EOF
        sExamCd = Trim(rs_svr.Fields(0) & "")
        rs_svr.MoveNext
    Loop
    Set rs_svr = Nothing
    
    '-- 검사코드 가져오기
    SQL = " Select EXMN_CD From SPSLHRRST " & CR & _
          " Where SPCM_NO = '" & Trim(sExamCd) & "' " & vbCrLf & _
          "   and RSLT_NO IS NOT NULL"
    
    Set rs_svr = cn_Ser.Execute(SQL)
    Do Until rs_svr.EOF
        GetOrderExamCode_New = GetOrderExamCode_New & "'" & Trim(rs_svr.Fields(0)) & "',"
        rs_svr.MoveNext
    Loop
    
    If GetOrderExamCode_New <> "" Then
        GetOrderExamCode_New = Mid(GetOrderExamCode_New, 1, Len(GetOrderExamCode_New) - 1)
    End If
    
Exit Function

Err:
    GetOrderExamCode_New = ""

End Function

Function GetOrderExamCode_MIC(argEquipCode As String, argPID As String) As String
'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Dim i           As Integer
Dim sExamCode   As String
Dim strExamCode As String
Dim sExamCd     As String
Dim rs_svr As ADODB.Recordset

    GetOrderExamCode_MIC = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    sExamCd = argPID
    
    '-- 검사코드 가져오기
    SQL = " Select EXMN_CD From SPSLHRRST " & CR & _
          " Where SPCM_NO = '" & Trim(sExamCd) & "' " & vbCrLf & _
          "   and RSLT_NO IS NOT NULL"

    Set rs_svr = cn_Ser.Execute(SQL)
    Do Until rs_svr.EOF
        GetOrderExamCode_MIC = GetOrderExamCode_MIC & "'" & Trim(rs_svr.Fields(0)) & "',"
        rs_svr.MoveNext
    Loop

    If GetOrderExamCode_MIC <> "" Then
        GetOrderExamCode_MIC = Mid(GetOrderExamCode_MIC, 1, Len(GetOrderExamCode_MIC) - 1)
    End If

    '-- 임시 테스트용
'    GetOrderExamCode_MIC = "'L41000'"
    
End Function


Function GetEquipExamCode_E411(argEquipCode As String, argPID As String) As String
'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Dim i As Integer
Dim sExamCode As String
Dim strExamCode As String
Dim sSpecNo     As String

    GetEquipExamCode_E411 = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    '바코드번호로 검체번호 불러오기
    SQL = "SELECT FN_LABCVTBCNO('" & Trim(argPID) & "') FROM DUAL "
    res = db_select_Col(gServer, SQL)
    sSpecNo = Trim(gReadBuf(0))
    
    '-- 검사코드 가져오기
    SQL = " Select EXMN_CD From SPSLHRRST " & CR & _
          " Where SPCM_NO = '" & Trim(sSpecNo) & "' " & vbCrLf & _
          "   and RSLT_NO IS NOT NULL"
          
    res = db_select_Row(gServer, SQL)
    strExamCode = ""
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            strExamCode = strExamCode & "'" & Trim(gReadBuf(i)) & "',"
        Else
            Exit For
        End If
    Next
    
    If strExamCode = "" Then
'        MsgBox "미접수 환자"
        GetEquipExamCode_E411 = ""
        Exit Function
    End If
    strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
    'EquipExamCode =
    
    ClearSpread frmInterface.vasTemp1
'    sExamCode = ""
    
    '-- 가져온 검사코드의 채널 찾기
          SQL = "Select distinct equipcode "
    SQL = SQL & "  From EquipExam "
    SQL = SQL & " Where equipno  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   and examcode in (" & Trim(strExamCode) & ")"
    
    res = db_select_Row(gLocal, SQL)
    strExamCode = ""
    For i = 0 To UBound(gReadBuf)
    
        If gReadBuf(i) <> "" Then
            'gReadBuf(i) = Mid(gReadBuf(i), 1, Len(gReadBuf(i)) - 1)
            If Trim(gReadBuf(i)) <> "990" Then
                strExamCode = strExamCode & "\^^^" & Trim(gReadBuf(i))
            End If
        Else
            Exit For
        End If
    Next
    
    GetEquipExamCode_E411 = Mid(strExamCode, 2)
    
End Function

'Function GetEquipExamCode_CA1500(argEquipCode As String, argPID As String) As String
''검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
''한 장비 번호에 검사코드가 1개이상 존재
'Dim i As Integer
'Dim sExamCode As String
'Dim strExamCode As String
'Dim sSpecNo     As String
'
'    GetEquipExamCode_CA1500 = ""
'
'    If Trim(argEquipCode) = "" Then
'        Exit Function
'    End If
'
'    '바코드번호로 검체번호 불러오기
'    SQL = "SELECT FN_LABCVTBCNO('" & Trim(argPID) & "') FROM DUAL "
'    res = db_select_Col(gServer, SQL)
'    sSpecNo = Trim(gReadBuf(0))
'
'    '-- 검사코드 가져오기
'    SQL = " Select EXMN_CD From SPSLHRRST " & CR & _
'          " Where SPCM_NO = '" & Trim(sSpecNo) & "' " & vbCrLf & _
'          "   and RSLT_NO IS NOT NULL"
'
'    res = db_select_Row(gServer, SQL)
'    strExamCode = ""
'
'    For i = 0 To UBound(gReadBuf)
'        If gReadBuf(i) <> "" Then
'            strExamCode = strExamCode & "'" & Trim(gReadBuf(i)) & "',"
'        Else
'            Exit For
'        End If
'    Next
'
'    strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
'    'EquipExamCode =
'
'    ClearSpread frmInterface.vasTemp1
''    sExamCode = ""
'
'    '-- 가져온 검사코드의 채널 찾기
'          SQL = "Select distinct equipcode "
'    SQL = SQL & "  From EquipExam "
'    SQL = SQL & " Where equipno  = '" & Trim(gEquip) & "' "
'    SQL = SQL & "   and examcode in (" & Trim(strExamCode) & ")"
'
'    res = db_select_Row(gLocal, SQL)
'    strExamCode = ""
'    For i = 0 To UBound(gReadBuf)
'
'        If gReadBuf(i) <> "" Then
'            'gReadBuf(i) = Mid(gReadBuf(i), 1, Len(gReadBuf(i)) - 1)
'            strExamCode = strExamCode & "\^^^" & Trim(gReadBuf(i))
'        Else
'            Exit For
'        End If
'    Next
'
'    GetEquipExamCode_CA1500 = Mid(strExamCode, 2)
'
'End Function

Function GetEquipExamCode(argEquipCode As String, argPID As String) As String
'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Dim i As Integer
Dim sExamCode As String
Dim strExamCode As String

    GetEquipExamCode = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    SQL = " Select EXMN_CD From SPSLHRRST " & CR & _
          " Where SPCM_NO = '" & Trim(argPID) & "' " & vbCrLf & _
          "   and RSLT_NO IS NOT NULL"
          
    res = db_select_Col(gServer, SQL)
    strExamCode = ""
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            strExamCode = strExamCode & "'" & Trim(gReadBuf(i)) & "',"
        Else
            Exit For
        End If
    Next
    
    strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
    'EquipExamCode =
    
    ClearSpread frmInterface.vasTemp1
    sExamCode = ""
    
          SQL = "Select equipcode "
    SQL = SQL & "  From EquipExam "
    SQL = SQL & " Where equipno  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   and examcode in (" & Trim(argEquipCode) & ")"
    
    res = db_select_Col(gLocal, SQL)
    strExamCode = ""
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            strExamCode = strExamCode & Trim(gReadBuf(i)) & "0" & Space$(6)
        Else
            Exit For
        End If
    Next
    
    GetEquipExamCode = strExamCode
    
End Function


