Attribute VB_Name = "DB_Query"
Option Explicit

'-- 해당 환자 검사의 H/L, Delta, Panic 판정하기
Function GetDecision(ByVal argSpcRow As Integer, ByVal strBarNo As String, ByVal iRow As Integer) As String
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
    
    '-- 해당 환자의 참고치,델타,패닉 찾아오기
    SQL = "SELECT MALE_HIGH,MALE_LOW,FEML_HIGH,FEML_LOW,DELT_DVSN,DELT_HIGH,DELT_LOW,DELT_DD,PANC_DVSN,PANC_HIGH,PANC_LOW                 "
    SQL = SQL & vbCrLf & " FROM SPSLMFBIF                                                                                                                      "
    SQL = SQL & vbCrLf & " WHERE USE_STR_DY <= SYSDATE                                                                                                         "
    SQL = SQL & vbCrLf & "   AND USE_END_DY >= SYSDATE                                                                                                         "
    SQL = SQL & vbCrLf & "   and EXMN_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "'"
    Set rs_DPRef = cn_Ser.Execute(SQL)
    Do Until rs_DPRef.EOF
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
        SQL = SQL & vbCrLf & "           AND A.RSLT_STAT < '3'                                                                   "
        SQL = SQL & vbCrLf & "           AND A.SPCM_NO = FN_LABCVTBCNO('" & strBarNo & "')                                       "
        SQL = SQL & vbCrLf & "         GROUP BY B.PID, B.EXMN_CD ) A, SPSLHRRST B                                                "
        SQL = SQL & vbCrLf & " WHERE A.PID = B.PID                                                                               "
        SQL = SQL & vbCrLf & "   AND A.LAST_RPTG_DT = B.LAST_RPTG_DT                                                             "
        SQL = SQL & vbCrLf & "   AND A.EXMN_CD = B.EXMN_CD                                                                       "
        SQL = SQL & vbCrLf & "   AND A.EXMN_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "' "         '검사코드"
        SQL = SQL & vbCrLf & "   AND B.LAST_RPTG_DT BETWEEN (SYSDATE-30) AND SYSDATE                "           '-- 30일 이내
        Set rs_Delta = cn_Ser.Execute(SQL)
        Do Until rs_Delta.EOF
            strBefoRslt = rs_Delta.Fields("BEFO_VIEW_RSLT")             '이전결과
            strDestRslt = Trim(GetText(frmInterface.vasTemp, iRow, 3))  '현재결과
            
            '-- 성별로 판정결과 비교
            '-- 결과값이 수치일 경우에만 비교한다.
            If IsNumeric(strDestRslt) Then
                If strSex = "M" Then
                    If IsNumeric(rs_DPRef.Fields("MALE_HIGH")) Then
                        If CDbl(strDestRslt) > CDbl(rs_DPRef.Fields("MALE_HIGH")) Then
                            strHLVal = "H"
                        Else
                            strHLVal = ""
                        End If
                    Else
                        strHLVal = ""
                    End If
                    
                    If IsNumeric(rs_DPRef.Fields("MALE_LOW")) Then
                        If CDbl(strDestRslt) < CDbl(rs_DPRef.Fields("MALE_LOW")) Then
                            strHLVal = "L"
                        Else
                            strHLVal = ""
                        End If
                    Else
                        strHLVal = ""
                    End If
                
                Else
                    If IsNumeric(rs_DPRef.Fields("FEML_HIGH")) Then
                        If CDbl(strDestRslt) > CDbl(rs_DPRef.Fields("FEML_HIGH")) Then
                            strHLVal = "H"
                        Else
                            strHLVal = ""
                        End If
                    Else
                        strHLVal = ""
                    End If
                    If IsNumeric(rs_DPRef.Fields("FEML_LOW")) Then
                        If CDbl(strDestRslt) < CDbl(rs_DPRef.Fields("FEML_LOW")) Then
                            strHLVal = "L"
                        Else
                            strHLVal = ""
                        End If
                    Else
                        strHLVal = ""
                    End If
                End If
            Else
                strHLVal = ""
            End If
            
            '-- Delta 구분  (아래 로직이 맞는지 검증 필요함...必)
            '-- 결과값이 수치일 경우에만 비교한다.
            If IsNumeric(strDestRslt) Then
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
                            strDelta = strDelta / CInt(rs_Delta.Fields("DELTA_TERM_DT"))        '기간당 변화비율
                    Case 4:     '4 기간당 변화차 = 변화차 / 기간
                            strDelta = ""
                            strDelta = CDbl(strDestRslt) - CDbl(strBefoRslt)                    '변화차
                            strDelta = CDbl(strDelta) / CInt(rs_Delta.Fields("DELTA_TERM_DT"))  '기간당 변화차
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
            
            '-- Panic 구분
            '-- 결과값이 수치일 경우에만 비교한다.
            If IsNumeric(strDestRslt) Then
                Select Case Trim(rs_DPRef.Fields("PANC_DVSN"))
                    Case 0:     '0 사용안함
                            strPanic = ""
                    Case 1:     '1 상한만
                            If IsNumeric(rs_DPRef.Fields("PANC_HIGH")) Then
                                If CDbl(strDestRslt) > rs_DPRef.Fields("PANC_HIGH") Then
                                    strPanic = "P"
                                Else
                                    strPanic = ""
                                End If
                            Else
                                strPanic = ""
                            End If
                    Case 2:     '2 하한만
                            If IsNumeric(rs_DPRef.Fields("PANC_LOW")) Then
                                If CDbl(strDestRslt) < rs_DPRef.Fields("PANC_LOW") Then
                                    strPanic = "P"
                                Else
                                    strPanic = ""
                                End If
                            Else
                                strPanic = ""
                            End If
                    Case 3:     '3 모두 사용
                            If IsNumeric(rs_DPRef.Fields("PANC_LOW")) And IsNumeric(rs_DPRef.Fields("PANC_HIGH")) Then
                                If (CDbl(strDestRslt) < rs_DPRef.Fields("PANC_LOW") Or CDbl(strDestRslt) > rs_DPRef.Fields("PANC_HIGH")) Then
                                    strPanic = "P"
                                Else
                                    strPanic = ""
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
            rs_Delta.MoveNext
        Loop
        
        rs_DPRef.MoveNext
    Loop
    
    Set rs_DPRef = Nothing
        
    GetDecision = strHLVal & "|" & strDelta & "|" & strPanic


End Function

Function Insert_Data_Allergy(ByVal argSpcRow As Integer) As Integer
    Dim iRow            As Integer
    Dim i               As Integer
    Dim j               As Integer
    Dim lsID            As String
    Dim lsSpecNo        As String
    Dim lsPid           As String
    Dim sResult         As String
    Dim lsInsertTime    As String
    Dim sCnt            As String
        
On Error GoTo Err

    Insert_Data_Allergy = -1
    
    lsID = ""
    lsID = Trim(GetText(frmInterface.vasID, argSpcRow, colBarcode))
    lsSpecNo = Trim(GetText(frmInterface.vasID, argSpcRow, colSpecNo))
    lsPid = Trim(GetText(frmInterface.vasID, argSpcRow, colPID))
    lsInsertTime = Trim(Format(GetDateFull, "dd")) & "/" & Trim(Format(GetDateFull, "mm")) & "/" & Trim(Format(GetDateFull, "yyyy")) & " " & Trim(Format(GetDateFull, "hh:mm:ss"))
    
    If lsSpecNo = "" Then
        Exit Function
    End If
    
    'Local에서 환자별로 결과값 가져오기
    ClearSpread frmInterface.vasTemp
    
    SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag, ifgbn " & vbCrLf & _
          " From pat_res " & vbCrLf & _
          " Where equipno = '" & gEquip & "' " & vbCrLf & _
          " And examdate = '" & Format(CDate(frmInterface.dtpToday.Value), "yyyymmdd") & "'  " & vbCrLf & _
          " And barcode = '" & Trim(GetText(frmInterface.vasID, argSpcRow, colBarcode)) & "' " & vbCrLf & _
          " And diskno = '" & Trim(GetText(frmInterface.vasID, argSpcRow, colRack)) & "' " & vbCrLf & _
          " And posno = '" & Trim(GetText(frmInterface.vasID, argSpcRow, colPos)) & "' "
    res = db_select_Vas(gLocal, SQL, frmInterface.vasTemp)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    frmInterface.vasTemp.MaxRows = frmInterface.vasTemp.DataRowCnt + 1
    
    gHIVPosFlag = -1
    
    sCnt = ""
    
    cn_Ser.BeginTrans
    
    '서버로 결과값 저장하기
    For iRow = 1 To frmInterface.vasTemp.DataRowCnt
        sCnt = ""

        SQL = "SELECT SPCM_NO FROM SPSLHFOIN "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
        SQL = SQL & vbCrLf & "   AND RFVL_DVSN = '" & Trim(GetText(frmInterface.vasTemp, iRow, 8)) & "' "                                                     '결과상태"
        SQL = SQL & vbCrLf & "   AND ITEM_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "' "                      '검사코드"
        res = db_select_Col(gServer, SQL)
        If res > 0 Then
            SQL = "UPDATE SPSLHFOIN "   '-- 결과테이블
            SQL = SQL & vbCrLf & "   SET EXMN_RSLT01 = '" & Trim(GetText(frmInterface.vasTemp, iRow, 4)) & "', "                   '결과(장비결과)
            SQL = SQL & vbCrLf & "       EXMN_RSLT02 = '" & Trim(GetText(frmInterface.vasTemp, iRow, 5)) & "', "                   '결과(Class)"
            'SQL = SQL & vbCrLf & "       REGI_ID = '', "
            SQL = SQL & vbCrLf & "       RGST_DT = SysDate, "
            'SQL = SQL & vbCrLf & "       AMEN_ID = '', "
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
            SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
            SQL = SQL & vbCrLf & "   AND RFVL_DVSN = '" & Trim(GetText(frmInterface.vasTemp, iRow, 8)) & "' "
            SQL = SQL & vbCrLf & "   AND ITEM_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "' "
            res = SendQuery(gServer, SQL)
            
            If res < 0 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
        Else
            SQL = "INSERT INTO SPSLHFOIN (SPCM_NO, PID, RFVL_DVSN,ITEM_CD,EXMN_RSLT01,EXMN_RSLT02,REGI_ID,RGST_DT,AMEN_ID,UPDT_DT)"
            SQL = SQL & vbCrLf & " Values ( "
            SQL = SQL & vbCrLf & " '" & lsSpecNo & "', "
            SQL = SQL & vbCrLf & " '" & lsPid & "', "
            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 8)) & "', "
            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "', "
            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 3)) & "', "
            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 5)) & "', "
            SQL = SQL & vbCrLf & " '', "
            SQL = SQL & vbCrLf & " sysdate, "
            SQL = SQL & vbCrLf & " '', "
            SQL = SQL & vbCrLf & " sysdate) "
            res = SendQuery(gServer, SQL)
            
            If res < 0 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
        
        End If
    Next iRow
    
    SQL = "SELECT EXMN_CD FROM SPSLHRRST "  '-- 접수테이블
    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
    SQL = SQL & vbCrLf & "   AND EXMN_CD NOT LIKE '%G%' "
    SQL = SQL & vbCrLf & "   AND RSLT_STAT > '0' "
    SQL = SQL & vbCrLf & "   AND VIEW_RSLT IS NOT NULL "
    res = db_select_Vas(gServer, SQL, frmInterface.vasTemp1)
    
    If res = 0 Then                                                                 '///// 결과테이블에 결과가 다 들어가 있는 경우 (그룹코드제외)
        SQL = "Update SPSLMJBBI"    '-- 검체테이블
        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1',"
        SQL = SQL & vbCrLf & "       AMEN_ID = 'test', "
        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2'"
        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
        res = SendQuery(gServer, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            cn_Ser.RollbackTrans
            Exit Function
        End If
    
        SQL = "Update SPSLMJBDI"    '-- 처방테이블
        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1',"
'        SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
'        SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
        SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "', "
        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
'        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2'"
        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
        res = SendQuery(gServer, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            cn_Ser.RollbackTrans
            Exit Function
        End If
        
    ElseIf res = -1 Then                                                             '///// 쿼리 에러인경우
        SaveQuery SQL
        cn_Ser.RollbackTrans
        Exit Function
    
    Else
        SQL = "Update SPSLMJBBI"    '-- 검체테이블
        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1',"
        SQL = SQL & vbCrLf & "       AMEN_ID = 'test', "
        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2'"
        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
        res = SendQuery(gServer, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
    
        SQL = "Update SPSLMJBDI"    '-- 처방테이블
        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1',"
'        SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
'        SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
        SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "', "
        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
'        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2'"
        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
        res = SendQuery(gServer, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            cn_Ser.RollbackTrans
            Exit Function
        End If
    
    End If
    
    SQL = ""

    cn_Ser.CommitTrans
       
    Insert_Data_Allergy = 1
    
    Exit Function
    
Err:
    cn_Ser.RollbackTrans
    
End Function

Function Insert_Data(ByVal argSpcRow As Integer) As Integer
    Dim iRow            As Integer
    Dim i               As Integer
    Dim j               As Integer
    Dim lsID            As String
    Dim lsSpecNo        As String
    Dim lsPid           As String
    Dim sResult         As String
    Dim lsInsertTime    As String
    Dim sCnt            As String
        
On Error GoTo Err

    Insert_Data = -1
    
    lsID = ""
    lsID = Trim(GetText(frmInterface.vasID, argSpcRow, colBarcode))
    lsSpecNo = Trim(GetText(frmInterface.vasID, argSpcRow, colSpecNo))
    lsPid = Trim(GetText(frmInterface.vasID, argSpcRow, colPID))
    lsInsertTime = Trim(Format(GetDateFull, "dd")) & "/" & Trim(Format(GetDateFull, "mm")) & "/" & Trim(Format(GetDateFull, "yyyy")) & " " & Trim(Format(GetDateFull, "hh:mm:ss"))
    
    If lsSpecNo = "" Then
        Exit Function
    End If
    
    'Local에서 환자별로 결과값 가져오기
    ClearSpread frmInterface.vasTemp
    
    SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag " & vbCrLf & _
          " From pat_res " & vbCrLf & _
          " Where equipno = '" & gEquip & "' " & vbCrLf & _
          " And examdate = '" & Format(CDate(frmInterface.dtpToday.Value), "yyyymmdd") & "'  " & vbCrLf & _
          " And barcode = '" & Trim(GetText(frmInterface.vasID, argSpcRow, colBarcode)) & "' " & vbCrLf & _
          " And diskno = '" & Trim(GetText(frmInterface.vasID, argSpcRow, colRack)) & "' " & vbCrLf & _
          " And posno = '" & Trim(GetText(frmInterface.vasID, argSpcRow, colPos)) & "' "
    res = db_select_Vas(gLocal, SQL, frmInterface.vasTemp)
    
    If res = -1 Then
        SaveQuery SQL
        cn_Ser.RollbackTrans
        Exit Function
    End If
    
    frmInterface.vasTemp.MaxRows = frmInterface.vasTemp.DataRowCnt + 1
    
    gHIVPosFlag = -1
    
    sCnt = ""
    
    cn_Ser.BeginTrans
    
    '서버로 결과값 저장하기
    For iRow = 1 To frmInterface.vasTemp.DataRowCnt
        sCnt = ""
        
        SQL = "SELECT RSLT_NO FROM SPSLHRRST "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
        SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "' "                      '검사코드"
        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
        res = db_select_Col(gServer, SQL)
        If res > 0 Then
            sCnt = CLng(gReadBuf(0)) + 1
            
            '-- 결과값이 숫자값일 경우만 델타/패닉 판정을 한다.
            sResult = Trim(GetText(frmInterface.vasTemp, iRow, 3))
            If IsNumeric(sResult) Then
                Dim strDecision     As Variant
                Dim strBarcode      As String

                strBarcode = Trim(GetText(frmInterface.vasID, argSpcRow, colBarcode))
                strDecision = GetDecision(argSpcRow, strBarcode, iRow)
                strDecision = Split(strDecision, "|")
            Else
                strDecision = "||"
                strDecision = Split(strDecision, "|")
            End If
            
            SQL = "UPDATE SPSLHRRST "   '-- 결과테이블
            SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & Trim(GetText(frmInterface.vasTemp, iRow, 4)) & "', "                   '결과(장비결과)
            SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & Trim(GetText(frmInterface.vasTemp, iRow, 3)) & "', "                   '결과(수정결과)"
            SQL = SQL & vbCrLf & "       DTRM_DVSN = '" & strDecision(0) & "', "                                                         'H/L 체크"
            SQL = SQL & vbCrLf & "       DLTA_YN = '" & strDecision(1) & "', "                                                           'Delta 체크"
            SQL = SQL & vbCrLf & "       PANC_YN = '" & strDecision(2) & "', "                                                           'Panic 체크"
            SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "', "                                                   '결과입력자"
            SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'            SQL = SQL & vbCrLf & "       MDDL_RPTR_ID = 'test', "                                                   '중간보고자"
'            SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "                                                  '중간보고일시"
'            SQL = SQL & vbCrLf & "       LAST_RPTR_ID = 'test', "                                                   '최종보고자"
'            SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "                                                  '최종보고일시"
            SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
            SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "', "                                                        '결과수정자
            SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
            SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '결과번호 (결과 넣을시에 증가)
            SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "                                                          '결과상태"
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
            SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "' "         '검사코드"
            SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
            SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
            res = SendQuery(gServer, SQL)
            
            If res < 0 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
        End If
    Next iRow
    
    SQL = "SELECT EXMN_CD FROM SPSLHRRST "  '-- 접수테이블
    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
    SQL = SQL & vbCrLf & "   AND EXMN_CD NOT LIKE '%G%' "
    SQL = SQL & vbCrLf & "   AND RSLT_STAT > '0' "
    SQL = SQL & vbCrLf & "   AND VIEW_RSLT IS NOT NULL "
    res = db_select_Vas(gServer, SQL, frmInterface.vasTemp1)
    
    If res = 0 Then                                                                 '///// 결과테이블에 결과가 다 들어가 있는 경우 (그룹코드제외)
        SQL = "Update SPSLMJBBI"    '-- 검체테이블
        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1'"
        SQL = SQL & vbCrLf & "       AMEN_ID = 'test', "
        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2'"
        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
        res = SendQuery(gServer, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            cn_Ser.RollbackTrans
            Exit Function
        End If
    
        SQL = "Update SPSLMJBDI"    '-- 처방테이블
        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1',"
'        SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
'        SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
        SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "', "
        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
'        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2'"
        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
        res = SendQuery(gServer, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            cn_Ser.RollbackTrans
            Exit Function
        End If
        
    ElseIf res = -1 Then                                                             '///// 쿼리 에러인경우
        SaveQuery SQL
        cn_Ser.RollbackTrans
        Exit Function
    
    Else
        SQL = "Update SPSLMJBBI"    '-- 검체테이블
        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1',"
        SQL = SQL & vbCrLf & "       AMEN_ID = 'test', "
        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2'"
        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
        res = SendQuery(gServer, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            cn_Ser.RollbackTrans
            Exit Function
        End If
    
        SQL = "Update SPSLMJBDI"    '-- 처방테이블
        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1',"
'        SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
'        SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
        SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "', "
        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
'        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2'"
        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
        res = SendQuery(gServer, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            cn_Ser.RollbackTrans
            Exit Function
        End If
    
    End If
    
    SQL = ""

    cn_Ser.CommitTrans
       
    Insert_Data = 1
    
    Exit Function
    
Err:
    cn_Ser.RollbackTrans
    
End Function

'''Function Insert_Data_MIC(ByVal argSpcRow As Integer) As Integer
'''    Dim iRow            As Integer
'''    Dim i               As Integer
'''    Dim j               As Integer
'''    Dim lsID            As String
'''    Dim lsSpecNo        As String
'''    Dim lsPid           As String
'''    Dim sResult         As String
'''    Dim lsInsertTime    As String
'''    Dim sCnt            As String
'''
'''On Error GoTo Err
'''
'''    Insert_Data_MIC = -1
'''
'''    lsID = ""
'''    lsID = Trim(GetText(frmInterface.vasResult, argSpcRow, colBarcode))
'''    lsSpecNo = Trim(GetText(frmInterface.vasResult, argSpcRow, colSpecNo))
'''    lsPid = Trim(GetText(frmInterface.vasResult, argSpcRow, colPID))
'''    lsInsertTime = Trim(Format(GetDateFull, "dd")) & "/" & Trim(Format(GetDateFull, "mm")) & "/" & Trim(Format(GetDateFull, "yyyy")) & " " & Trim(Format(GetDateFull, "hh:mm:ss"))
'''
'''    If lsSpecNo = "" Then
'''        Exit Function
'''    End If
'''
'''    'Local에서 환자별로 결과값 가져오기
'''    ClearSpread frmInterface.vasTemp
'''
'''    SQL = " Select isocd, equipcode, examcode, result, antsize, EQUIPRESULT, refflag, panicflag, deltaflag,exmncd " & vbCrLf & _
'''          " From pat_res " & vbCrLf & _
'''          " Where equipno = '" & gEquip & "' " & vbCrLf & _
'''          " And examdate = '" & Format(CDate(frmInterface.dtpToday.Value), "yyyymmdd") & "'  " & vbCrLf & _
'''          " And barcode = '" & Trim(GetText(frmInterface.vasResult, argSpcRow, 3)) & "' " & vbCrLf & _
'''          " And examcode = '" & Trim(GetText(frmInterface.vasResult, argSpcRow, 4)) & "' " & vbCrLf & _
'''          " And receno = '" & Trim(GetText(frmInterface.vasResult, argSpcRow, 2)) & "' " & vbCrLf & _
'''          " And isocd = '" & Trim(GetText(frmInterface.vasResult, argSpcRow, 7)) & "' "
'''    res = db_select_Vas(gLocal, SQL, frmInterface.vasTemp)
'''
'''    If res = -1 Then
'''        SaveQuery SQL
'''        Exit Function
'''    End If
'''
'''    frmInterface.vasTemp.MaxRows = frmInterface.vasTemp.DataRowCnt + 1
'''
'''    gHIVPosFlag = -1
'''
'''    sCnt = ""
'''
'''    cn_Ser.BeginTrans
'''
'''    '서버로 결과값 저장하기
'''    For iRow = 1 To frmInterface.vasTemp.DataRowCnt
'''        sCnt = ""
'''
'''        If iRow = 1 Then
'''            '-- 미생물 세균결과
'''            SQL = "SELECT SPCM_NO FROM SPSLHMBAC "
'''            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "                                             '검체번호
'''            SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 3)) & "' "     '검사코드
'''            SQL = SQL & vbCrLf & "   AND BCTR_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 1)) & "' "     '세균코드
'''            SQL = SQL & vbCrLf & "   AND BCTR_SQNO = " & iRow                                                   '세균일련번호:번호-N5
'''            res = db_select_Col(gServer, SQL)
'''            If res > 0 Then
'''                '로직 변경 --> 삭제하고 인서트해야 할것 같음. 세균일련번호가 일정하지 않을것 같음. 덤프가 항상 같지 않을것임.
'''                '              미생물 담당자에게 업무 물어보고 정할것.
'''                '====================???????????????????????????????????????????????????????????
'''                SQL = "DELETE FROM SPSLHMBAC "
'''                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "                                             'spcm_no        검체번호
'''                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & lsPid & "' "                                            'exmn_cd        검사코드
'''                SQL = SQL & vbCrLf & "   AND BCTR_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 8)) & "' "     'bctr_cd        세균코드
'''                SQL = SQL & vbCrLf & "   AND BCTR_SQNO = '" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "' "   'bctr_sqno      세균일련번호:번호-N5
'''                '====================???????????????????????????????????????????????????????????
'''
'''                SQL = "UPDATE SPSLHMBAC SET "
'''                SQL = SQL & " SORT_SEQ = '" & iRow & "', "
'''                SQL = SQL & " SPCM_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 10)) & "', "
'''                SQL = SQL & " CLTR_VOL_CD = '', "
'''                SQL = SQL & " CLTR_PERD = '', "
'''                SQL = SQL & " PRE_RSLT_CD = '', "
'''                SQL = SQL & " MDDL_RPTR_ID = '" & gEquipCode & "', "
'''                SQL = SQL & " LAST_BCTR_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 1)) & "', "
'''                SQL = SQL & " MDDL_RPTG_DT = sysdate, "
'''                SQL = SQL & " LAST_RPTR_ID = '" & gEquipCode & "', "
'''                SQL = SQL & " LAST_RPTG_DT = sysdate, "
'''                SQL = SQL & " RSLT_STAT = '1', "
'''                SQL = SQL & " CMNT_DVSN = '', "
'''                SQL = SQL & " EQPM_CD = '" & gEquipCode & "', "
'''                SQL = SQL & " RMRK = '', "
'''                SQL = SQL & " REGI_ID = '" & gEquipCode & "', "
'''                SQL = SQL & " RGST_DT = sysdate, "
'''                SQL = SQL & " AMEN_ID = '" & gEquipCode & "', "
'''                SQL = SQL & " UPDT_DT = sysdate "
'''                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "                                             'spcm_no        검체번호
'''                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & lsPid & "' "                                            'exmn_cd        검사코드
'''                SQL = SQL & vbCrLf & "   AND BCTR_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 1)) & "' "     'bctr_cd        세균코드
'''                SQL = SQL & vbCrLf & "   AND BCTR_SQNO = '" & iRow & "' "                                           'bctr_sqno      세균일련번호:번호-N5
'''
'''                res = SendQuery(gServer, SQL)
'''
'''                If res < 0 Then
'''                    SaveQuery SQL
'''                    cn_Ser.RollbackTrans
'''                    Exit Function
'''                End If
'''            Else
'''                SQL = "INSERT INTO SPSLHMBAC (SPCM_NO,      EXMN_CD,        BCTR_CD,        BCTR_SQNO,      SORT_SEQ,"
'''                SQL = SQL & vbCrLf & "        SPCM_CD,      CLTR_VOL_CD,    CLTR_PERD,      PRE_RSLT_CD,    MDDL_RPTR_ID,"
'''                SQL = SQL & vbCrLf & "        LAST_BCTR_CD, MDDL_RPTG_DT,   LAST_RPTR_ID,   LAST_RPTG_DT,   RSLT_STAT, "
'''                SQL = SQL & vbCrLf & "        CMNT_DVSN,    EXMN_EQPM,      RMRK,           REGI_ID,        RGST_DT, "
'''                SQL = SQL & vbCrLf & "        AMEN_ID, UPDT_DT)"
'''                SQL = SQL & vbCrLf & " Values ( "
'''                SQL = SQL & vbCrLf & " '" & lsID & "', "                                            'spcm_no        검체번호
'''                SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 3)) & "', "    'exmn_cd        검사코드
'''                SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 1)) & "', "    'bctr_cd        세균코드
'''                SQL = SQL & vbCrLf & " '" & iRow & "', "                                            'bctr_sqno      세균일련번호:번호-N5
'''                SQL = SQL & vbCrLf & " '" & iRow & "', "                                            'sort_seq       정렬순서
'''                SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 10)) & "', "   'spcm_cd        검체코드
'''                SQL = SQL & vbCrLf & " '', "                                                        'CLTR_VOL_CD    배양량코드:구분코드
'''                SQL = SQL & vbCrLf & " '', "                                                        'CLTR_PERD      배양기간:내용-V200
'''                SQL = SQL & vbCrLf & " '', "                                                        'PRE_RSLT_CD    예비결과코드:구분코드
'''                SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'MDDL_RPTR_ID   중간보고자ID:직원번호
'''                SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 1)) & "', "    'LAST_BCTR_CD   최종세균코드
'''                SQL = SQL & vbCrLf & " sysdate, "                                                   'MDDL_RPTG_DT   중간보고일시:날짜-DT
'''                SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'LAST_RPTR_ID   최종보고자ID:직원번호
'''                SQL = SQL & vbCrLf & " sysdate, "                                                   'LAST_RPTG_DT   최종보고일시:날짜-DT
'''                SQL = SQL & vbCrLf & " '1', "                                                       'RSLT_STAT      결과상태:구분코드
'''                SQL = SQL & vbCrLf & " '', "                                                        'CMNT_DVSN      코멘트구분
'''                SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'EQPM_CD        장비코드:구분코드
'''                SQL = SQL & vbCrLf & " '', "                                                        'RMRK           비고
'''                SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'REGI_ID        등록자ID:직원번호
'''                SQL = SQL & vbCrLf & " sysdate, "                                                   'RGST_DT        등록일시:날짜-DT
'''                SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'AMEN_ID        수정자ID:직원번호
'''                SQL = SQL & vbCrLf & " sysdate) "                                                   'UPDT_DT        수정일시:날짜-DT
'''
'''                res = SendQuery(gServer, SQL)
'''
'''                If res < 0 Then
'''                    SaveQuery SQL
'''                    cn_Ser.RollbackTrans
'''                    Exit Function
'''                End If
'''
'''            End If
'''        End If
'''
'''        '-- 미생물 항생제결과
'''        SQL = "SELECT SPCM_NO FROM SPSLHMANT "
'''        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "                                             '검체번호"
'''        SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & lsPid & "' "                                                    '환자번호"
'''        SQL = SQL & vbCrLf & "   AND BCTR_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 8)) & "' "                                                     '결과상태"
'''        SQL = SQL & vbCrLf & "   AND BCTR_SQNO = '" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "' "                      '검사코드"
'''        SQL = SQL & vbCrLf & "   AND ANTB_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "' "                      '검사코드"
'''        res = db_select_Col(gServer, SQL)
'''        If res > 0 Then
'''            SQL = "UPDATE SPSLHMANT "
'''                SQL = SQL & " SPCM_CD = '', "
'''                SQL = SQL & " ANTB_RSLT = '', "
'''                SQL = SQL & " DTRM_RSLT = '', "
'''                SQL = SQL & " ANTB_EXMN_MTHD = '', "
'''                SQL = SQL & " RSLT_RPTR_ID = '', "
'''                SQL = SQL & " RSLT_RPTG_DT = '', "
'''                SQL = SQL & " MDDL_RPTG_ID = '', "
'''                SQL = SQL & " MDDL_RPTG_DT = '', "
'''                SQL = SQL & " LAST_RPTR_ID = '', "
'''                SQL = SQL & " LAST_RPTG_DT = '', "
'''                SQL = SQL & " RSLT_STAT = '', "
'''                SQL = SQL & " EXMN_EQPM = '', "
'''                SQL = SQL & " REGI_ID = '', "
'''                SQL = SQL & " RGST_DT = '', "
'''                SQL = SQL & " AMEN_ID = '', "
'''                SQL = SQL & " UPDT_DT = '' "
'''                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
'''                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & lsPid & "' "
'''                SQL = SQL & vbCrLf & "   AND BCTR_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 8)) & "' "
'''                SQL = SQL & vbCrLf & "   AND BCTR_SQNO = '" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "' "
'''                SQL = SQL & vbCrLf & "   AND ANTB_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "' "
'''            res = SendQuery(gServer, SQL)
'''
'''            If res < 0 Then
'''                SaveQuery SQL
'''                cn_Ser.RollbackTrans
'''                Exit Function
'''            End If
'''        Else
'''            SQL = "INSERT INTO SPSLHMANT (SPCM_NO,EXMN_CD,BCTR_CD,BCTR_SQNO,ANTB_CD,"
'''            SQL = SQL & vbCrLf & "SPCM_CD,ANTB_RSLT,DTRM_RSLT,ANTB_EXMN_MTHD,RSLT_RPTR_ID,RSLT_RPTG_DT,MDDL_RPTR_ID,MDDL_RPTG_DT,"
'''            SQL = SQL & vbCrLf & "LAST_RPTR_ID , LAST_RPTG_DT, RSLT_STAT, EXMN_EQPM, REGI_ID, RGST_DT, AMEN_ID, UPDT_DT)"
'''            SQL = SQL & vbCrLf & " Values ( "
'''            SQL = SQL & vbCrLf & " '" & lsSpecNo & "', "
'''            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 3)) & "', "
'''            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 1)) & "', "
'''            SQL = SQL & vbCrLf & " '" & iRow & "', "
'''            SQL = SQL & vbCrLf & " '" & iRow & "', "
'''            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 5)) & "', "    'spcm_cd
'''            SQL = SQL & vbCrLf & " '', "    'ANTB_RSLT
'''            SQL = SQL & vbCrLf & " '', "    'DTRM_RSLT
'''            SQL = SQL & vbCrLf & " '', "    'ANTB_EXMN_MTHD
'''            SQL = SQL & vbCrLf & " '', "    'RSLT_RPTR_ID
'''            SQL = SQL & vbCrLf & " '', "    'RSLT_RPTG_DT
'''            SQL = SQL & vbCrLf & " '', "    'MDDL_RPTR_ID
'''            SQL = SQL & vbCrLf & " '', "    'MDDL_RPTG_DT
'''            SQL = SQL & vbCrLf & " '', "    'LAST_RPTR_ID
'''            SQL = SQL & vbCrLf & " '', "    'LAST_RPTG_DT
'''            SQL = SQL & vbCrLf & " '', "    'RSLT_STAT
'''            SQL = SQL & vbCrLf & " '', "    'EQPM_CD
'''            SQL = SQL & vbCrLf & " '', "    'REGI_ID
'''            SQL = SQL & vbCrLf & " '', "    'RGST_DT
'''            SQL = SQL & vbCrLf & " '', "    'AMEN_ID
'''            SQL = SQL & vbCrLf & " sysdate) "   'UPDT_DT
'''            res = SendQuery(gServer, SQL)
'''
'''            If res < 0 Then
'''                SaveQuery SQL
'''                cn_Ser.RollbackTrans
'''                Exit Function
'''            End If
'''
'''        End If
'''
'''    Next iRow
'''
'''    SQL = "SELECT EXMN_CD FROM SPSLHRRST "  '-- 접수테이블
'''    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
'''    SQL = SQL & vbCrLf & "   AND EXMN_CD NOT LIKE '%G%' "
'''    SQL = SQL & vbCrLf & "   AND RSLT_STAT > '0' "
'''    SQL = SQL & vbCrLf & "   AND VIEW_RSLT IS NOT NULL "
'''    res = db_select_Vas(gServer, SQL, frmInterface.vasTemp1)
'''
'''    If res = 0 Then                                                                 '///// 결과테이블에 결과가 다 들어가 있는 경우 (그룹코드제외)
'''        SQL = "Update SPSLMJBBI"    '-- 검체테이블
'''        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1',"
'''        SQL = SQL & vbCrLf & "       AMEN_ID = 'test', "
'''        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
'''        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
'''        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
'''        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2'"
'''        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
'''        res = SendQuery(gServer, SQL)
'''
'''        If res = -1 Then
'''            SaveQuery SQL
'''            cn_Ser.RollbackTrans
'''            Exit Function
'''        End If
'''
'''        SQL = "Update SPSLMJBDI"    '-- 처방테이블
'''        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1',"
''''        SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
''''        SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
'''        SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "', "
'''        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
'''        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
''''        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
'''        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2'"
'''        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
'''        res = SendQuery(gServer, SQL)
'''
'''        If res = -1 Then
'''            SaveQuery SQL
'''            cn_Ser.RollbackTrans
'''            Exit Function
'''        End If
'''
'''    ElseIf res = -1 Then                                                             '///// 쿼리 에러인경우
'''        SaveQuery SQL
'''        cn_Ser.RollbackTrans
'''        Exit Function
'''
'''    Else
'''        SQL = "Update SPSLMJBBI"    '-- 검체테이블
'''        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1',"
'''        SQL = SQL & vbCrLf & "       AMEN_ID = 'test', "
'''        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
'''        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
'''        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
'''        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2'"
'''        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
'''        res = SendQuery(gServer, SQL)
'''
'''        If res = -1 Then
'''            SaveQuery SQL
'''            Exit Function
'''        End If
'''
'''        SQL = "Update SPSLMJBDI"    '-- 처방테이블
'''        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1',"
''''        SQL = SQL & vbCrLf & "       MDDL_RPTG_DT = SysDate, "
''''        SQL = SQL & vbCrLf & "       LAST_RPTG_DT = SysDate, "
'''        SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "', "
'''        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
'''        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsSpecNo & "' "
''''        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
'''        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2'"
'''        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
'''        res = SendQuery(gServer, SQL)
'''
'''        If res = -1 Then
'''            SaveQuery SQL
'''            cn_Ser.RollbackTrans
'''            Exit Function
'''        End If
'''
'''    End If
'''
'''    SQL = ""
'''
'''    cn_Ser.CommitTrans
'''
'''    Insert_Data_MIC = 1
'''
'''    Exit Function
'''
'''Err:
'''    cn_Ser.RollbackTrans
'''
'''
'''End Function


Function Insert_Data_QC(ByVal argSpcRow As Integer) As Integer
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
    Dim lsQC_Date       As String

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

        SQL = " Select equipcode, examcode, result, EQUIPRESULT, refflag, panicflag, deltaflag, RESDATE " & vbCrLf & _
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
        cn_Ser.BeginTrans
        '서버로 결과값 저장하기
        For iRow = 1 To .vasTemp.DataRowCnt
            sCnt = ""
            
            sResult1 = Trim(GetText(.vasTemp, iRow, 4))
            sResult2 = Trim(GetText(.vasTemp, iRow, 3))
            
            If Mid(sResult1, 1, 3) = "-99" Then: sResult1 = " "
            
            If sResult1 <> "" Then
                SQL = "SELECT RSLT_SQNO FROM SPSLHQRST "
                SQL = SQL & vbCrLf & "WHERE EQPM_CD = '" & Mid(lsID, 3, 3) & "' "
                SQL = SQL & vbCrLf & "  AND SBSN_CD = '" & Mid(lsID, 6, 3) & "' "
                SQL = SQL & vbCrLf & "  AND LVL_CD = '" & Mid(lsID, 9, 1) & "' "
                SQL = SQL & vbCrLf & "  AND EXMN_CD  = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "
                SQL = SQL & vbCrLf & "  AND EXMN_DY = '" & lsQC_Date & "' "
                SQL = SQL & vbCrLf & "  AND RSLT_VALU IS NULL "
                SQL = SQL & vbCrLf & "GROUP BY RSLT_SQNO "
                res = db_select_Col(gServer, SQL)
                sCnt = gReadBuf(0)
                
                If IsNumeric(sCnt) = True Then
                    SQL = "UPDATE SPSLHQRST "
                    SQL = SQL & vbCrLf & "  SET RSLT_VALU = '" & sResult1 & "', "                                       '결과(장비결과)
                    SQL = SQL & vbCrLf & "      RSLT_DT = TO_DATE('" & Trim(GetText(.vasTemp, iRow, 8)) & "'), "          '결과(수정결과)"
                    SQL = SQL & vbCrLf & "      RSLT_RPTR_ID = '" & gEquipCode & "_INF' "                               'Delta 체크"
                    SQL = SQL & vbCrLf & "      AMEN_ID = '" & gEquipCode & "_INF' "                                    'Panic 체크"
                    SQL = SQL & vbCrLf & "      UPDT_DT = sysdate, "                                                    '결과입력자"
                    SQL = SQL & vbCrLf & "WHERE EQPM_CD = '" & Mid(lsID, 3, 3) & "' "
                    SQL = SQL & vbCrLf & "  AND SBSN_CD = '" & Mid(lsID, 6, 3) & "' "
                    SQL = SQL & vbCrLf & "  AND LVL_CD = '" & Mid(lsID, 9, 1) & "' "
                    SQL = SQL & vbCrLf & "  AND EXMN_CD  = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "
                    SQL = SQL & vbCrLf & "  AND EXMN_DY = '" & lsQC_Date & "' "
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
                    SQL = "SELECT MAX(RSLT_SQNO) FROM SPSLHQRST "
                    SQL = SQL & vbCrLf & "WHERE EQPM_CD = '" & Mid(lsID, 3, 3) & "' "
                    SQL = SQL & vbCrLf & "  AND SBSN_CD = '" & Mid(lsID, 6, 3) & "' "
                    SQL = SQL & vbCrLf & "  AND LVL_CD = '" & Mid(lsID, 9, 1) & "' "
                    SQL = SQL & vbCrLf & "  AND EXMN_CD  = '" & Trim(GetText(.vasTemp, iRow, 2)) & "' "
                    SQL = SQL & vbCrLf & "  AND EXMN_DY = '" & lsQC_Date & "' "
                    'SQL = SQL & vbCrLf & "  AND RSLT_SQNO = '" & sCnt & "' "
                    'SQL = SQL & vbCrLf & "GROUP BY RSLT_SQNO "
                    res = db_select_Col(gServer, SQL)
                
                    If gReadBuf(0) = "" Then sCnt = "1"
                    sCnt = CLng(gReadBuf(0)) + 1
                    If Trim(GetText(.vasTemp, iRow, 2)) <> "" Then
                        SQL = ""
                        SQL = SQL & vbCrLf & "INSERT INTO SPSLHQRST(EXMN_DY   ,EQPM_CD ,SBSN_CD ,LVL_CD  "
                        SQL = SQL & vbCrLf & "                     ,RSLT_SQNO ,EXMN_CD ,RSLT_DT ,RSLT_RPTR_ID "
                        SQL = SQL & vbCrLf & "                     ,RSLT_VALU ,SPCM_NO ,DEL_YN "
                        SQL = SQL & vbCrLf & "                     ,REGI_ID   ,RGST_DT ,AMEN_ID ,UPDT_DT) "
                        SQL = SQL & vbCrLf & "               VALUES('" & lsQC_Date & "', '" & Mid(lsID, 3, 3) & "', '" & Mid(lsID, 6, 3) & "', '" & Mid(lsID, 9, 1) & "', "
                        SQL = SQL & vbCrLf & "                      " & sCnt & ", '" & Trim(GetText(.vasTemp, iRow, 2)) & "', TO_DATE('" & Trim(GetText(.vasTemp, iRow, 8)) & "'), '" & gEquipCode & "_INF', "
                        SQL = SQL & vbCrLf & "                      '" & sResult1 & "', '" & lsID & "', 'N', "
                        SQL = SQL & vbCrLf & "                      '" & gEquipCode & "_INF', sysdate, '" & gEquipCode & "_INF', sysdate ) "
                        res = SendQuery(gServer, SQL)
                        If res = -1 Then
                            SaveQuery SQL
                            cn_Ser.RollbackTrans
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next iRow
        
        cn_Ser.CommitTrans
        Insert_Data_QC = 1
    End With
End Function

'''Function Insert_Data_MIC(ByVal argSpcRow As Integer) As Integer
'''    Dim iRow            As Integer
'''    Dim i               As Integer
'''    Dim j               As Integer
'''    Dim lsID            As String
'''    Dim lsSpecNo        As String
'''    Dim lsPid           As String
'''    Dim sResult         As String
'''    Dim lsInsertTime    As String
'''    Dim sCnt            As String
'''    Dim rScnt           As Integer
'''
'''On Error GoTo Err
'''
'''    Insert_Data_MIC = -1
'''
'''    lsID = ""
'''    lsID = Trim(GetText(frmInterface.vasResult, argSpcRow, colBarcode))
'''    lsSpecNo = Trim(GetText(frmInterface.vasResult, argSpcRow, colSpecNo))
'''    lsPid = Trim(GetText(frmInterface.vasResult, argSpcRow, 5))
'''    lsInsertTime = Trim(Format(GetDateFull, "dd")) & "/" & Trim(Format(GetDateFull, "mm")) & "/" & Trim(Format(GetDateFull, "yyyy")) & " " & Trim(Format(GetDateFull, "hh:mm:ss"))
'''
'''    If lsSpecNo = "" Then
'''        Exit Function
'''    End If
'''
'''    'Local에서 환자별로 결과값 가져오기
'''    ClearSpread frmInterface.vasTemp
'''
''''    SQL = " Select isocd, equipcode, examcode, result, antsize, EQUIPRESULT, refflag, panicflag, deltaflag,exmncd " & vbCrLf
'''
'''    SQL = " Select isocd, equipcode, examcode, antrslt, antsize, EQUIPRESULT, refflag, panicflag, deltaflag,exmncd " & vbCrLf & _
'''          "   From pat_res " & vbCrLf & _
'''          " Where equipno = '" & gEquip & "' " & vbCrLf & _
'''          "   And examdate = '" & Format(CDate(frmInterface.dtpToday.Value), "yyyymmdd") & "'  " & vbCrLf & _
'''          "   And barcode = '" & Trim(GetText(frmInterface.vasResult, argSpcRow, 3)) & "' " & vbCrLf & _
'''          "   And examcode = '" & Trim(GetText(frmInterface.vasResult, argSpcRow, 4)) & "' " & vbCrLf & _
'''          "   And receno = '" & Trim(GetText(frmInterface.vasResult, argSpcRow, 2)) & "' " & vbCrLf & _
'''          "   And isocd = '" & Trim(GetText(frmInterface.vasResult, argSpcRow, 7)) & "' "
'''
'''    res = db_select_Vas(gLocal, SQL, frmInterface.vasTemp)
'''
'''    If res = -1 Then
'''        SaveQuery SQL
'''        Exit Function
'''    End If
'''
'''    frmInterface.vasTemp.MaxRows = frmInterface.vasTemp.DataRowCnt + 1
'''
'''    gHIVPosFlag = -1
'''
'''    sCnt = ""
'''
'''    cn_Ser.BeginTrans
'''
'''    '서버로 결과값 저장하기
'''    For iRow = 1 To frmInterface.vasTemp.DataRowCnt
'''        If iRow = 1 Then
'''            '-- 미생물 세균결과 조회
'''            SQL = "SELECT MAX(BCTR_SQNO) FROM SPSLHMBAC "
'''            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "                                                 'SPCM_NO    검체번호
'''            SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(frmInterface.vasTemp, iRow, 3)) & "' "         'EXMN_CD    검사코드
'''
'''            res = db_select_Col(gServer, SQL)
'''            If res > 0 And gReadBuf(0) <> "" And IsNumeric(gReadBuf(0)) Then
'''                sCnt = CLng(gReadBuf(0)) + 1
'''            Else
'''                sCnt = 1
'''            End If
'''
'''            '-- 미생물 세균결과 저장
'''            SQL = "INSERT INTO SPSLHMBAC (SPCM_NO,      EXMN_CD,        BCTR_CD,        BCTR_SQNO,      SORT_SEQ,"
'''            SQL = SQL & vbCrLf & "        SPCM_CD,      CLTR_VOL_CD,    CLTR_PERD,      PRE_RSLT_CD,    LAST_BCTR_CD,"
'''            SQL = SQL & vbCrLf & "        RSLT_RPTR_ID, RSLT_RPTG_DT,   MDDL_RPTR_ID,   MDDL_RPTG_DT,   LAST_RPTR_ID,"
'''            SQL = SQL & vbCrLf & "        LAST_RPTG_DT, RSLT_STAT,      CMNT_DVSN,      EXMN_EQPM,      RMRK,   REGI_ID,   RGST_DT,AMEN_ID,UPDT_DT) "
'''            SQL = SQL & vbCrLf & " Values ( "
'''            SQL = SQL & vbCrLf & " '" & lsID & "', "                                            'SPCM_NO        검체번호
'''            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 3)) & "', "    'EXMN_CD        검사코드
'''            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 1)) & "', "    'BCTR_CD        세균코드
'''            SQL = SQL & vbCrLf & " '" & sCnt & "', "                                            'BCTR_SQNO      세균일련번호:번호-N5
'''            SQL = SQL & vbCrLf & " '" & iRow & "', "                                            'SORT_SEQ       정렬순서
'''            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 10)) & "', "   'SPCM_CD        검체코드
'''            SQL = SQL & vbCrLf & " '', "                                                        'CLTR_VOL_CD    배양량코드:구분코드
'''            SQL = SQL & vbCrLf & " '', "                                                        'CLTR_PERD      배양기간:내용-V200
'''            SQL = SQL & vbCrLf & " '', "                                                        'PRE_RSLT_CD    예비결과코드:구분코드
'''            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 1)) & "', "    'LAST_BCTR_CD   최종세균코드
'''            SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'RSLT_RPTR_ID   결과보고자ID:직원번호
'''            SQL = SQL & vbCrLf & " sysdate, "                                                   'RSLT_RPTG_DT   결과보고일시:날짜-DT
'''            SQL = SQL & vbCrLf & " '', "                                                        'MDDL_RPTR_ID   중간보고자ID:직원번호
'''            SQL = SQL & vbCrLf & " '', "                                                        'MDDL_RPTG_DT   중간보고일시:날짜-DT
'''            SQL = SQL & vbCrLf & " '', "                                                        'LAST_RPTR_ID   최종보고자ID:직원번호
'''            SQL = SQL & vbCrLf & " '', "                                                        'LAST_RPTG_DT   최종보고일시:날짜-DT
'''            SQL = SQL & vbCrLf & " '1', "                                                       'RSLT_STAT      결과상태:구분코드 ==> 결과등록 : 1 [RSLT_RPTR_ID, RSLT_RPTG_DT 입력]    ?? 검사실 선생님만 보여야한다고 함.
'''                                                                                                '                                     예비보고 : 2 [RSLT_RPTR_ID, RSLT_RPTG_DT, MDDDL_RPTR_ID, MDDL_RPTG_DT 입력]
'''                                                                                                '                                     최종보고 : 3 [RSLT_RPTR_ID, RSLT_RPTG_DT, MDDDL_RPTR_ID, MDDL_RPTG_DT, LAST_RPTR_ID, LAST_RPTG_DT 입력]
'''            SQL = SQL & vbCrLf & " '', "                                                        'CMNT_DVSN      코멘트구분
'''            SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'EQPM_CD        장비코드:구분코드
'''            SQL = SQL & vbCrLf & " '', "                                                        'RMRK           비고
'''            SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'REGI_ID        등록자ID:직원번호
'''            SQL = SQL & vbCrLf & " sysdate, "                                                   'RGST_DT        등록일시:날짜-DT
'''            SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'AMEN_ID        결과수정자
'''            SQL = SQL & vbCrLf & " sysdate) "                                                   'UPDT_DT        결과수정일시
'''
'''            res = SendQuery(gServer, SQL)
'''
'''            If res < 0 Then
'''                SaveQuery SQL
'''                cn_Ser.RollbackTrans
'''                Exit Function
'''            End If
'''
'''            SQL = "SELECT RSLT_NO FROM SPSLHRRST "
'''            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "                                                 '검체번호"
'''            SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(frmInterface.vasResult, argSpcRow, 4)) & "' "  '검사코드"
'''            SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
'''            SQL = SQL & vbCrLf & "   AND RSLT_STAT <> '3' "                                                          '결과상태"
'''            res = db_select_Col(gServer, SQL)
'''            If res > 0 Then
''''                sCnt = CLng(gReadBuf(0)) + 1
'''
'''                '-- 결과값이 숫자값일 경우만 델타/패닉 판정을 한다.
'''                sResult = Trim(GetText(frmInterface.vasTemp, iRow, 3))
'''                If IsNumeric(sResult) Then
'''                    Dim strDecision     As Variant
'''                    Dim strBarcode      As String
'''
'''                    strBarcode = Trim(GetText(frmInterface.vasID, argSpcRow, colBarcode))
'''                    strDecision = GetDecision(argSpcRow, strBarcode, iRow)
'''                    strDecision = Split(strDecision, "|")
'''                Else
'''                    strDecision = "||"
'''                    strDecision = Split(strDecision, "|")
'''                End If
'''
'''                SQL = "UPDATE SPSLHRRST "   '-- 결과테이블
'''                SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & Trim(GetText(frmInterface.vasTemp, iRow, 1)) & "', "      '결과(장비결과)
'''                SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & Trim(GetText(frmInterface.vasTemp, iRow, 1)) & "', "      '결과(수정결과)"
'''                SQL = SQL & vbCrLf & "       DTRM_DVSN = '" & strDecision(0) & "', "                                    'H/L 체크"
'''                SQL = SQL & vbCrLf & "       DLTA_YN = '" & strDecision(1) & "', "                                      'Delta 체크"
'''                SQL = SQL & vbCrLf & "       PANC_YN = '" & strDecision(2) & "', "                                      'Panic 체크"
'''                SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "', "                                     '결과입력자"
'''                SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
'''                SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
'''                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "', "                                          '결과수정자
'''                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
'''                SQL = SQL & vbCrLf & "       RSLT_NO = '" & sCnt & "', "                                                '결과번호 (결과 넣을시에 증가)
'''                SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "                                                          '결과상태"
'''                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "                                                 '검체번호"
'''                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(frmInterface.vasResult, argSpcRow, 4)) & "' "  '검사코드"
'''                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
'''                SQL = SQL & vbCrLf & "   AND RSLT_STAT <> '3' "                                                          '결과상태"
'''                res = SendQuery(gServer, SQL)
'''
'''                If res < 0 Then
'''                    SaveQuery SQL
'''                    cn_Ser.RollbackTrans
'''                    Exit Function
'''                End If
'''
'''            End If
'''
'''        End If
'''
'''        '-- 미생물 항생제결과 저장
'''        If Trim(GetText(frmInterface.vasTemp, iRow, 2)) <> "" Then
'''            SQL = "INSERT INTO SPSLHMANT (SPCM_NO,      EXMN_CD,        BCTR_CD,        BCTR_SQNO,      ANTB_CD,"
'''            SQL = SQL & vbCrLf & "        SPCM_CD,      ANTB_RSLT,      DTRM_RSLT,      ANTB_EXMN_MTHD, RSLT_RPTR_ID,"
'''            SQL = SQL & vbCrLf & "        RSLT_RPTG_DT, MDDL_RPTR_ID,   MDDL_RPTG_DT,   LAST_RPTR_ID,   LAST_RPTG_DT,"
'''            SQL = SQL & vbCrLf & "        RSLT_STAT,    EXMN_EQPM,      REGI_ID,        RGST_DT,AMEN_ID,UPDT_DT)"
'''            SQL = SQL & vbCrLf & " Values ( "
'''            SQL = SQL & vbCrLf & " '" & lsID & "', "                                            'SPCM_NO        검체번호
'''            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 3)) & "', "    'EXMN_CD        검사코드
'''            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 1)) & "', "    'BCTR_CD        세균코드
'''            SQL = SQL & vbCrLf & " '" & sCnt & "', "                                            'BCTR_SQNO      세균일련번호:번호-N5
'''            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "', "    'ANTB_CD        항생제코드:구분코드
'''            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 10)) & "', "   'SPCM_CD        검체코드
'''            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 5)) & "', "    'ANTB_RSLT      항생제결과:분류-V50
'''            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 4)) & "', "    'DTRM_RSLT      판정결과
'''            SQL = SQL & vbCrLf & " 'M', "                                                       'ANTB_EXMN_MTHD 항생제검사방법:구분코드 ==> 검사방법 M : MICRO 법 의미
'''            SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'RSLT_RPTR_ID   결과보고자ID:직원번호
'''            SQL = SQL & vbCrLf & " sysdate, "                                                   'RSLT_RPTG_DT   결과보고일시:날짜-DT
'''            SQL = SQL & vbCrLf & " '', "                                                        'MDDL_RPTR_ID   중간보고자ID:직원번호
'''            SQL = SQL & vbCrLf & " '', "                                                        'MDDL_RPTG_DT   중간보고일시:날짜-DT
'''            SQL = SQL & vbCrLf & " '', "                                                        'LAST_RPTR_ID   최종보고자ID:직원번호
'''            SQL = SQL & vbCrLf & " '', "                                                        'LAST_RPTG_DT   최종보고일시:날짜-DT
'''            SQL = SQL & vbCrLf & " '1', "                                                       'RSLT_STAT      결과상태:구분코드 ==> 결과등록 : 1 [RSLT_RPTR_ID, RSLT_RPTG_DT 입력]    ?? 검사실 선생님만 보여야한다고 함.
'''                                                                                                '                                     예비보고 : 2 [RSLT_RPTR_ID, RSLT_RPTG_DT, MDDDL_RPTR_ID, MDDL_RPTG_DT 입력]
'''                                                                                                '                                     최종보고 : 3 [RSLT_RPTR_ID, RSLT_RPTG_DT, MDDDL_RPTR_ID, MDDL_RPTG_DT, LAST_RPTR_ID, LAST_RPTG_DT 입력]
'''            SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'EQPM_CD        장비코드:구분코드
'''            SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'REGI_ID        등록자ID:직원번호
'''            SQL = SQL & vbCrLf & " sysdate, "                                                   'RGST_DT        등록일시:날짜-DT
'''            SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'AMEN_ID        결과수정자
'''            SQL = SQL & vbCrLf & " sysdate) "                                                   'UPDT_DT        결과수정일시
'''            res = SendQuery(gServer, SQL)
'''
'''            If res < 0 Then
'''                SaveQuery SQL
'''                cn_Ser.RollbackTrans
'''                Exit Function
'''            End If
'''        End If
'''
'''    Next iRow
'''
'''
'''    SQL = "SELECT EXMN_CD FROM SPSLHRRST "  '-- 접수테이블
'''    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "
'''    SQL = SQL & vbCrLf & "   AND EXMN_CD NOT LIKE '%G%' "
'''    SQL = SQL & vbCrLf & "   AND RSLT_STAT > '0' "
'''    SQL = SQL & vbCrLf & "   AND VIEW_RSLT IS NOT NULL "
'''    res = db_select_Vas(gServer, SQL, frmInterface.vasTemp1)
'''
'''    If res = 0 Then                                                                 '///// 결과테이블에 결과가 다 들어가 있는 경우 (그룹코드제외)
'''        SQL = "Update SPSLMJBBI"    '-- 검체테이블
'''        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1',"
'''        SQL = SQL & vbCrLf & "       AMEN_ID = 'test', "
'''        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
'''        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "
'''        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
'''        SQL = SQL & vbCrLf & "   AND RSLT_STAT <> '3'"
'''        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
'''        res = SendQuery(gServer, SQL)
'''
'''        If res = -1 Then
'''            SaveQuery SQL
'''            cn_Ser.RollbackTrans
'''            Exit Function
'''        End If
'''
'''        SQL = "Update SPSLMJBDI"    '-- 처방테이블
'''        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1',"
'''        SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "', "
'''        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
'''        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "
'''        SQL = SQL & vbCrLf & "   AND RSLT_STAT <> '3'"
'''        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
'''        res = SendQuery(gServer, SQL)
'''
'''        If res = -1 Then
'''            SaveQuery SQL
'''            cn_Ser.RollbackTrans
'''            Exit Function
'''        End If
'''
'''    ElseIf res = -1 Then                                                             '///// 쿼리 에러인경우
'''        SaveQuery SQL
'''        cn_Ser.RollbackTrans
'''        Exit Function
'''
'''    Else
'''        SQL = "Update SPSLMJBBI"    '-- 검체테이블
'''        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1',"
'''        SQL = SQL & vbCrLf & "       AMEN_ID = 'test', "
'''        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
'''        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "
'''        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
'''        SQL = SQL & vbCrLf & "   AND RSLT_STAT <> '3'"
'''        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
'''        res = SendQuery(gServer, SQL)
'''
'''        If res = -1 Then
'''            SaveQuery SQL
'''            Exit Function
'''        End If
'''
'''        SQL = "Update SPSLMJBDI"    '-- 처방테이블
'''        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1',"
'''        SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "', "
'''        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
'''        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "
'''        SQL = SQL & vbCrLf & "   AND RSLT_STAT <> '3'"
'''        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
'''        res = SendQuery(gServer, SQL)
'''
'''        If res = -1 Then
'''            SaveQuery SQL
'''            cn_Ser.RollbackTrans
'''            Exit Function
'''        End If
'''
'''    End If
'''
'''    SQL = ""
'''
'''    cn_Ser.CommitTrans
'''
'''    Insert_Data_MIC = 1
'''
'''    Exit Function
'''
'''Err:
'''    cn_Ser.RollbackTrans
'''
'''
'''End Function

Function Insert_Data_MIC(ByVal argSpcRow As Integer) As Integer
    Dim iRow            As Integer
    Dim i               As Integer
    Dim j               As Integer
    Dim lsID            As String
    Dim lsSpecNo        As String
    Dim lsPid           As String
    Dim sResult         As String
    Dim lsInsertTime    As String
    Dim sCnt            As String
    Dim strCmntDvsn     As String
    Dim intCmnt         As Integer
    Dim rScnt           As Integer
    Dim strExmnCD       As String
    Dim strANAE         As String
    Dim sSortSeq        As String
    
On Error GoTo Err

    Insert_Data_MIC = -1
    
    lsID = ""
    lsID = Trim(GetText(frmInterface.vasResult, argSpcRow, colBarcode))
    lsSpecNo = Trim(GetText(frmInterface.vasResult, argSpcRow, colSpecNo))
    lsPid = Trim(GetText(frmInterface.vasResult, argSpcRow, 5))
    lsInsertTime = Trim(Format(GetDateFull, "dd")) & "/" & Trim(Format(GetDateFull, "mm")) & "/" & Trim(Format(GetDateFull, "yyyy")) & " " & Trim(Format(GetDateFull, "hh:mm:ss"))
    
    If lsSpecNo = "" Then
        Exit Function
    End If
    
    'Local에서 환자별로 결과값 가져오기
    ClearSpread frmInterface.vasTemp
    
    SQL = " Select isocd, equipcode, examcode, antrslt, antsize, EQUIPRESULT, refflag, panicflag, deltaflag,exmncd " & vbCrLf & _
          "   From pat_res " & vbCrLf & _
          " Where equipno = '" & gEquip & "' " & vbCrLf & _
          "   And examdate = '" & Format(CDate(frmInterface.dtpToday.Value), "yyyymmdd") & "'  " & vbCrLf & _
          "   And barcode = '" & Trim(GetText(frmInterface.vasResult, argSpcRow, 3)) & "' " & vbCrLf & _
          "   And examcode = '" & Trim(GetText(frmInterface.vasResult, argSpcRow, 4)) & "' " & vbCrLf & _
          "   And receno = '" & Trim(GetText(frmInterface.vasResult, argSpcRow, 2)) & "' " & vbCrLf & _
          "   And isocd = '" & Trim(GetText(frmInterface.vasResult, argSpcRow, 7)) & "' "
    res = db_select_Vas(gLocal, SQL, frmInterface.vasTemp)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    frmInterface.vasTemp.MaxRows = frmInterface.vasTemp.DataRowCnt + 1
    
    gHIVPosFlag = -1
    
    sCnt = ""
    
    cn_Ser.BeginTrans
    
    '서버로 결과값 저장하기
    For iRow = 1 To frmInterface.vasTemp.DataRowCnt
        If iRow = 1 Then
            If Len(Trim(GetText(frmInterface.vasTemp, iRow, 3))) = 8 Then
                If Mid(Trim(GetText(frmInterface.vasTemp, iRow, 3)), 7, 2) = "01" Then
                    strANAE = "AERO"    '호기성
                Else
                    strANAE = "ANAE"    '혐기성
                End If
            Else
                strANAE = ""
            End If
            
            strExmnCD = Mid(Trim(GetText(frmInterface.vasTemp, iRow, 3)), 1, 6)
            
            
            '-- 미생물 세균결과 조회
            SQL = "SELECT MAX(BCTR_SQNO) FROM SPSLHMBAC "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "                                                 'SPCM_NO    검체번호
            SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & strExmnCD & "' "         'EXMN_CD    검사코드
            If strANAE <> "" Then
                SQL = SQL & vbCrLf & "   AND CLTR_VOL_CD = '" & strANAE & "' "         '- 호기/혐기 구분
            End If
            
            res = db_select_Col(gServer, SQL)
            If res > 0 And gReadBuf(0) <> "" And IsNumeric(gReadBuf(0)) Then
                sCnt = CLng(gReadBuf(0)) + 1
            Else
                sCnt = 1
                If strANAE = "AERO" Then        '호기성
                    sCnt = "1001"
                ElseIf strANAE = "ANAE" Then    '혐기성
                    sCnt = "2001"
                End If
            End If
            
            '-- 미생물 세균결과 정렬번호 조회
            SQL = "SELECT MAX(SORT_SEQ) FROM SPSLHMBAC "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "                                                 'SPCM_NO    검체번호
            SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & strExmnCD & "' "         'EXMN_CD    검사코드
            res = db_select_Col(gServer, SQL)
            If res > 0 And gReadBuf(0) <> "" And IsNumeric(gReadBuf(0)) Then
                sSortSeq = CLng(gReadBuf(0)) + 1
            Else
                sSortSeq = 1
            End If
            
            '-- 미생물 세균결과 저장
            SQL = "INSERT INTO SPSLHMBAC (SPCM_NO,      EXMN_CD,        BCTR_CD,        BCTR_SQNO,      SORT_SEQ,"
            SQL = SQL & vbCrLf & "        SPCM_CD,      CLTR_VOL_CD,    CLTR_PERD,      PRE_RSLT_CD,    LAST_BCTR_CD,"
            SQL = SQL & vbCrLf & "        RSLT_RPTR_ID, RSLT_RPTG_DT,   MDDL_RPTR_ID,   MDDL_RPTG_DT,   LAST_RPTR_ID,"
            SQL = SQL & vbCrLf & "        LAST_RPTG_DT, RSLT_STAT,      CMNT_DVSN,      EXMN_EQPM,      RMRK,   REGI_ID,   RGST_DT,AMEN_ID,UPDT_DT) "
            SQL = SQL & vbCrLf & " Values ( "
            SQL = SQL & vbCrLf & " '" & lsID & "', "                                            'SPCM_NO        검체번호
            SQL = SQL & vbCrLf & " '" & strExmnCD & "', "                                       'EXMN_CD        검사코드
            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 1)) & "', "    'BCTR_CD        세균코드
            SQL = SQL & vbCrLf & " '" & sCnt & "', "                                            'BCTR_SQNO      세균일련번호:번호-N5
            SQL = SQL & vbCrLf & " '" & sSortSeq & "', "                                        'SORT_SEQ       정렬순서
            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 10)) & "', "   'SPCM_CD        검체코드
            SQL = SQL & vbCrLf & " '" & strANAE & "', "                                         'CLTR_VOL_CD    배양량코드:구분코드
            SQL = SQL & vbCrLf & " '', "                                                        'CLTR_PERD      배양기간:내용-V200
            SQL = SQL & vbCrLf & " '', "                                                        'PRE_RSLT_CD    예비결과코드:구분코드
            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 1)) & "', "    'LAST_BCTR_CD   최종세균코드
            SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'RSLT_RPTR_ID   결과보고자ID:직원번호
            SQL = SQL & vbCrLf & " sysdate, "                                                   'RSLT_RPTG_DT   결과보고일시:날짜-DT
            SQL = SQL & vbCrLf & " '', "                                                        'MDDL_RPTR_ID   중간보고자ID:직원번호
            SQL = SQL & vbCrLf & " '', "                                                        'MDDL_RPTG_DT   중간보고일시:날짜-DT
            SQL = SQL & vbCrLf & " '', "                                                        'LAST_RPTR_ID   최종보고자ID:직원번호
            SQL = SQL & vbCrLf & " '', "                                                        'LAST_RPTG_DT   최종보고일시:날짜-DT
            SQL = SQL & vbCrLf & " '1', "                                                       'RSLT_STAT      결과상태:구분코드 ==> 결과등록 : 1 [RSLT_RPTR_ID, RSLT_RPTG_DT 입력]    ?? 검사실 선생님만 보여야한다고 함.
                                                                                                '                                     예비보고 : 2 [RSLT_RPTR_ID, RSLT_RPTG_DT, MDDDL_RPTR_ID, MDDL_RPTG_DT 입력]
                                                                                                '                                     최종보고 : 3 [RSLT_RPTR_ID, RSLT_RPTG_DT, MDDDL_RPTR_ID, MDDL_RPTG_DT, LAST_RPTR_ID, LAST_RPTG_DT 입력]
            strCmntDvsn = ""
            For intCmnt = 1 To frmInterface.vasTemp.MaxRows
                strCmntDvsn = Trim(GetText(frmInterface.vasTemp, intCmnt, 10))
                If strCmntDvsn <> "" Then
                    Exit For
                End If
            Next
            SQL = SQL & vbCrLf & " '" & strCmntDvsn & "', "                                     'CMNT_DVSN      코멘트구분
            SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'EQPM_CD        장비코드:구분코드
            SQL = SQL & vbCrLf & " '', "                                                        'RMRK           비고
            SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'REGI_ID        등록자ID:직원번호
            SQL = SQL & vbCrLf & " sysdate, "                                                   'RGST_DT        등록일시:날짜-DT
            SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'AMEN_ID        결과수정자
            SQL = SQL & vbCrLf & " sysdate) "                                                   'UPDT_DT        결과수정일시
            
            res = SendQuery(gServer, SQL)
            
            If res < 0 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
        
            SQL = "SELECT RSLT_NO FROM SPSLHRRST "
            SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "                                                 '검체번호"
            SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & strExmnCD & "' "  '검사코드"
            SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
            SQL = SQL & vbCrLf & "   AND RSLT_STAT <> '3' "                                                          '결과상태"
            res = db_select_Col(gServer, SQL)
            If res > 0 Then
                rScnt = CLng(gReadBuf(0)) + 1
                
                '-- 결과값이 숫자값일 경우만 델타/패닉 판정을 한다.
                sResult = Trim(GetText(frmInterface.vasTemp, iRow, 3))
                If IsNumeric(sResult) Then
                    Dim strDecision     As Variant
                    Dim strBarcode      As String
        
                    strBarcode = Trim(GetText(frmInterface.vasID, argSpcRow, colBarcode))
                    strDecision = GetDecision(argSpcRow, strBarcode, iRow)
                    strDecision = Split(strDecision, "|")
                Else
                    strDecision = "||"
                    strDecision = Split(strDecision, "|")
                End If
                
                SQL = "UPDATE SPSLHRRST "   '-- 결과테이블
'                SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & Trim(GetText(frmInterface.vasTemp, iRow, 1)) & "', "      '결과(장비결과)
'                SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & Trim(GetText(frmInterface.vasTemp, iRow, 1)) & "', "      '결과(수정결과)"
                SQL = SQL & vbCrLf & "   SET REAL_RSLT = '" & Trim(GetText(frmInterface.vasRes, iRow, 9)) & "', "      '결과(장비결과)
                SQL = SQL & vbCrLf & "       VIEW_RSLT = '" & Trim(GetText(frmInterface.vasRes, iRow, 9)) & "', "      '결과(수정결과)"
                SQL = SQL & vbCrLf & "       DTRM_DVSN = '" & strDecision(0) & "', "                                    'H/L 체크"
                SQL = SQL & vbCrLf & "       DLTA_YN = '" & strDecision(1) & "', "                                      'Delta 체크"
                SQL = SQL & vbCrLf & "       PANC_YN = '" & strDecision(2) & "', "                                      'Panic 체크"
                SQL = SQL & vbCrLf & "       RSLT_INPS_ID = '" & gEquipCode & "', "                                     '결과입력자"
                SQL = SQL & vbCrLf & "       RSLT_INPT_DT = SysDate, "                                                  '결과입력일시"
                SQL = SQL & vbCrLf & "       EXMN_EQPM = '" & gEquipCode & "', "                                        '장비코드
                SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "', "                                          '결과수정자
                SQL = SQL & vbCrLf & "       UPDT_DT = SysDate, "                                                       '결과수정일시
                SQL = SQL & vbCrLf & "       RSLT_NO = '" & rScnt & "', "                                                '결과번호 (결과 넣을시에 증가)
                SQL = SQL & vbCrLf & "       RSLT_STAT = '1' "                                                          '결과상태"
                SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "                                                 '검체번호"
                SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & Trim(GetText(frmInterface.vasResult, argSpcRow, 4)) & "' "  '검사코드"
                'SQL = SQL & vbCrLf & "   AND EXMN_CD = '" & strExmnCD & "' "                                            '검사코드"
                SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "                                                    '환자번호"
                SQL = SQL & vbCrLf & "   AND RSLT_STAT <> '3' "                                                          '결과상태"
                res = SendQuery(gServer, SQL)
                
                If res < 0 Then
                    SaveQuery SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
            End If
        End If
        
        '-- 미생물 항생제결과 저장
        'Debug.Print Trim(GetText(frmInterface.vasTemp, iRow, 2))
        If Trim(GetText(frmInterface.vasTemp, iRow, 2)) <> "" And Trim(GetText(frmInterface.vasTemp, iRow, 2)) <> "ESBL" And _
            Len(Trim(GetText(frmInterface.vasTemp, iRow, 4))) <= 1 And Trim(GetText(frmInterface.vasTemp, iRow, 2)) <> "----" Then
            SQL = "INSERT INTO SPSLHMANT (SPCM_NO,      EXMN_CD,        BCTR_CD,        BCTR_SQNO,      ANTB_CD,"
            SQL = SQL & vbCrLf & "        SPCM_CD,      ANTB_RSLT,      DTRM_RSLT,      ANTB_EXMN_MTHD, RSLT_RPTR_ID,"
            SQL = SQL & vbCrLf & "        RSLT_RPTG_DT, MDDL_RPTR_ID,   MDDL_RPTG_DT,   LAST_RPTR_ID,   LAST_RPTG_DT,"
            SQL = SQL & vbCrLf & "        RSLT_STAT,    EXMN_EQPM,      REGI_ID,        RGST_DT,AMEN_ID,UPDT_DT)"
            SQL = SQL & vbCrLf & " Values ( "
            SQL = SQL & vbCrLf & " '" & lsID & "', "                                            'SPCM_NO        검체번호
            'SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 3)) & "', "    'EXMN_CD        검사코드
            SQL = SQL & vbCrLf & " '" & strExmnCD & "', "                                       'EXMN_CD        검사코드
            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 1)) & "', "    'BCTR_CD        세균코드
            SQL = SQL & vbCrLf & " '" & sCnt & "', "                                            'BCTR_SQNO      세균일련번호:번호-N5
            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 2)) & "', "    'ANTB_CD        항생제코드:구분코드
            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 10)) & "', "   'SPCM_CD        검체코드
            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 5)) & "', "    'ANTB_RSLT      항생제결과:분류-V50
            SQL = SQL & vbCrLf & " '" & Trim(GetText(frmInterface.vasTemp, iRow, 4)) & "', "    'DTRM_RSLT      판정결과
            SQL = SQL & vbCrLf & " 'M', "                                                       'ANTB_EXMN_MTHD 항생제검사방법:구분코드 ==> 검사방법 M : MICRO 법 의미
            SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'RSLT_RPTR_ID   결과보고자ID:직원번호
            SQL = SQL & vbCrLf & " sysdate, "                                                   'RSLT_RPTG_DT   결과보고일시:날짜-DT
            SQL = SQL & vbCrLf & " '', "                                                        'MDDL_RPTR_ID   중간보고자ID:직원번호
            SQL = SQL & vbCrLf & " '', "                                                        'MDDL_RPTG_DT   중간보고일시:날짜-DT
            SQL = SQL & vbCrLf & " '', "                                                        'LAST_RPTR_ID   최종보고자ID:직원번호
            SQL = SQL & vbCrLf & " '', "                                                        'LAST_RPTG_DT   최종보고일시:날짜-DT
            SQL = SQL & vbCrLf & " '1', "                                                       'RSLT_STAT      결과상태:구분코드 ==> 결과등록 : 1 [RSLT_RPTR_ID, RSLT_RPTG_DT 입력]    ?? 검사실 선생님만 보여야한다고 함.
                                                                                                '                                     예비보고 : 2 [RSLT_RPTR_ID, RSLT_RPTG_DT, MDDDL_RPTR_ID, MDDL_RPTG_DT 입력]
                                                                                                '                                     최종보고 : 3 [RSLT_RPTR_ID, RSLT_RPTG_DT, MDDDL_RPTR_ID, MDDL_RPTG_DT, LAST_RPTR_ID, LAST_RPTG_DT 입력]
            SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'EQPM_CD        장비코드:구분코드
            SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'REGI_ID        등록자ID:직원번호
            SQL = SQL & vbCrLf & " sysdate, "                                                   'RGST_DT        등록일시:날짜-DT
            SQL = SQL & vbCrLf & " '" & gEquipCode & "', "                                      'AMEN_ID        결과수정자
            SQL = SQL & vbCrLf & " sysdate) "                                                   'UPDT_DT        결과수정일시
            
            Save_Raw_Data "[SQL]" & SQL
            
            res = SendQuery(gServer, SQL)
            
            If res < 0 Then
                SaveQuery SQL
                cn_Ser.RollbackTrans
                Exit Function
            End If
        End If
    Next iRow
    
    SQL = "SELECT EXMN_CD FROM SPSLHRRST "  '-- 접수테이블
    SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "
    SQL = SQL & vbCrLf & "   AND EXMN_CD NOT LIKE '%G%' "
    SQL = SQL & vbCrLf & "   AND RSLT_STAT > '0' "
    SQL = SQL & vbCrLf & "   AND VIEW_RSLT IS NOT NULL "
    res = db_select_Vas(gServer, SQL, frmInterface.vasTemp1)
    
    If res = 0 Then                                                                 '///// 결과테이블에 결과가 다 들어가 있는 경우 (그룹코드제외)
        SQL = "Update SPSLMJBBI"    '-- 검체테이블
        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1',"
        SQL = SQL & vbCrLf & "       AMEN_ID = 'test', "
        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "
        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
'        SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2'"
        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
        res = SendQuery(gServer, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            cn_Ser.RollbackTrans
            Exit Function
        End If
    
        SQL = "Update SPSLMJBDI"    '-- 처방테이블
        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1',"
        SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "', "
        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT <> '3'"
        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
        res = SendQuery(gServer, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            cn_Ser.RollbackTrans
            Exit Function
        End If
        
    ElseIf res = -1 Then                                                             '///// 쿼리 에러인경우
        SaveQuery SQL
        cn_Ser.RollbackTrans
        Exit Function
    
    Else
        SQL = "Update SPSLMJBBI"    '-- 검체테이블
        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1',"
        SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "', "
        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "
        SQL = SQL & vbCrLf & "   AND PID = '" & lsPid & "' "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT <> '3'"
        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
        res = SendQuery(gServer, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
    
        SQL = "Update SPSLMJBDI"    '-- 처방테이블
        SQL = SQL & vbCrLf & "   SET RSLT_STAT = '1',"
        SQL = SQL & vbCrLf & "       AMEN_ID = '" & gEquipCode & "', "
        SQL = SQL & vbCrLf & "       UPDT_DT = SysDate "
        SQL = SQL & vbCrLf & " WHERE SPCM_NO = '" & lsID & "' "
        SQL = SQL & vbCrLf & "   AND RSLT_STAT <> '3'"
        SQL = SQL & vbCrLf & "   AND SPCM_STAT = '2'"
        res = SendQuery(gServer, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            cn_Ser.RollbackTrans
            Exit Function
        End If
    
    End If
    
    SQL = ""

    cn_Ser.CommitTrans
       
    Insert_Data_MIC = 1
    
    Exit Function
    
Err:
    cn_Ser.RollbackTrans
    
    
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

Function Get_Sample_Info_SPCMNO(ByVal asRow As Long) As Integer
    
    Dim sBarcode As String
    Dim sSpecNo As String
    Dim sTestCd As String
    
    Get_Sample_Info_SPCMNO = -1
    '환자정보 가져오기
    sSpecNo = Trim(GetText(frmInterface.vasResult, asRow, colSpecNo))
    sTestCd = Trim(GetText(frmInterface.vasResult, asRow, colTestCd))
    
    If sSpecNo = "" Then
        Exit Function
    End If
    '바코드번호로 검체번호 불러오기FN_LABCVTPRTBCNO(SPCM_NO) --> 바코드라벨번호 리턴

    SQL = "SELECT FN_LABCVTPRTBCNO('" & Trim(sSpecNo) & "') FROM DUAL "
    res = db_select_Col(gServer, SQL)
    sBarcode = Trim(gReadBuf(0))
    
    '환자번호, 환자이름, 주민번호, 성별, 나이
    SQL = "SELECT PID, PT_NM, SEX, AGE "
    SQL = SQL & vbCrLf & " FROM SPSLMJBBI "
    SQL = SQL & vbCrLf & "WHERE SPCM_NO = '" & sSpecNo & "' "
    SQL = SQL & vbCrLf & "  AND SPCM_STAT = '2' "
    SQL = SQL & vbCrLf & "  AND RSLT_STAT < '2' "
    res = db_select_Col(gServer, SQL)
    
    '///////// gAllExam 자리에 검사 코드 넣어줌 세부코드 도 붙어 있는게 B312001 , 02, 03
    
    If res = 1 Then
        SetText frmInterface.vasResult, Trim(sSpecNo), asRow, colSpecNo     '2
        SetText frmInterface.vasResult, Trim(sBarcode), asRow, colBarcode   '3
        SetText frmInterface.vasResult, Trim(sTestCd), asRow, colTestCd    '4
        SetText frmInterface.vasResult, Trim(gReadBuf(0)), asRow, colPID    '6
        SetText frmInterface.vasResult, Trim(gReadBuf(1)), asRow, colPName  '7
        SetText frmInterface.vasResult, Trim(gReadBuf(2)), asRow, colSex    '8
        SetText frmInterface.vasResult, Trim(gReadBuf(3)), asRow, colAge    '9
        Get_Sample_Info_SPCMNO = 1
    Else
        Get_Sample_Info_SPCMNO = -1
    End If

End Function

Function Get_Sample_Info(ByVal asRow As Long) As Integer
    
    Dim sBarcode As String
    Dim sSpecNo As String
    
    Get_Sample_Info = -1
    '환자정보 가져오기
    sBarcode = Trim(GetText(frmInterface.vasID, asRow, colBarcode))   '샘플 바코드 번호
    
    If sBarcode = "" Then
        Exit Function
    End If
    '바코드번호로 검체번호 불러오기
    SQL = "SELECT FN_LABCVTBCNO('" & Trim(sBarcode) & "') FROM DUAL "
    res = db_select_Col(gServer, SQL)
    sSpecNo = Trim(gReadBuf(0))
    
    '환자번호, 환자이름, 주민번호, 성별, 나이
    SQL = "SELECT PID, PT_NM, SEX, AGE "
    SQL = SQL & vbCrLf & " FROM SPSLMJBBI "
    SQL = SQL & vbCrLf & "WHERE SPCM_NO = '" & sSpecNo & "' "
    SQL = SQL & vbCrLf & "  AND SPCM_STAT = '2' "
    SQL = SQL & vbCrLf & "  AND RSLT_STAT < '2' "
    res = db_select_Col(gServer, SQL)
    
    '///////// gAllExam 자리에 검사 코드 넣어줌 세부코드 도 붙어 있는게 B312001 , 02, 03
    
    If res = 1 Then
        SetText frmInterface.vasID, Trim(sSpecNo), asRow, colSpecNo     '2
        SetText frmInterface.vasID, Trim(gReadBuf(0)), asRow, colPID    '6
        SetText frmInterface.vasID, Trim(gReadBuf(1)), asRow, colPName  '7
        SetText frmInterface.vasID, Trim(gReadBuf(2)), asRow, colSex    '8
        SetText frmInterface.vasID, Trim(gReadBuf(3)), asRow, colAge    '9
        Get_Sample_Info = 1
    Else
        Get_Sample_Info = -1
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
    'GetOrderExamCode_MIC = "'L41000'"
    
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


